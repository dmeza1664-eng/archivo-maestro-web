const express = require('express');
const { query, transaction } = require('../db');
const { requireAuth, requireRole, writeAudit } = require('../auth');
const {
  ensureProduct,
  extractRows,
  firstValue,
  monthRange,
  paginationFromQuery,
  toDateOnly,
  toNumber,
} = require('./helpers');

function normalizedText(value, maxLength) {
  return String(value || '').trim().slice(0, maxLength);
}

function createOperationalRouter(config) {
  const router = express.Router();
  router.use(requireAuth);

  function normalizeRows(rows) {
    const valid = [];
    const rejected = [];
    for (const [index, row] of rows.entries()) {
      try {
        if (row.monthlyTotal) throw new Error('es un total mensual, no un registro diario');
        const productoCodigo = normalizedText(
          firstValue(row, ['producto_codigo', 'codigo_producto', 'codigo', 'sku', 'producto']),
          120
        );
        if (!productoCodigo) throw new Error('falta producto');
        const productoNombre = normalizedText(
          firstValue(row, ['producto_nombre', 'nombre_producto', 'nombre', 'descripcion', 'producto']) || productoCodigo,
          255
        );
        const fecha = toDateOnly(firstValue(row, ['fecha', 'dia', 'date']), 'fecha', index);
        const cantidad = toNumber(firstValue(row, config.quantityAliases), 'cantidad', index);
        const dimensions = Object.fromEntries(
          config.dimensions.map((dimension) => [
            dimension.name,
            normalizedText(firstValue(row, dimension.aliases), dimension.maxLength),
          ])
        );
        valid.push({
          index,
          fecha,
          producto_codigo: productoCodigo,
          producto_nombre: productoNombre,
          cantidad,
          ...dimensions,
        });
      } catch (error) {
        rejected.push({ fila: index + 1, motivo: error.message });
      }
    }
    return { valid, rejected };
  }

  function consolidateRows(rows) {
    const map = new Map();
    for (const row of rows) {
      const key = [
        row.fecha,
        row.producto_codigo.toUpperCase(),
        ...config.dimensions.map((dimension) => row[dimension.name].toUpperCase()),
      ].join('|');
      const current = map.get(key);
      if (current) current.cantidad += row.cantidad;
      else map.set(key, { ...row });
    }
    return [...map.values()];
  }

  async function importRows(req, res, next) {
    try {
      const rows = extractRows(req.body, config.bodyKeys);
      const archivo = normalizedText(req.body?.archivo, 255);
      const { valid, rejected } = normalizeRows(rows);
      const consolidated = consolidateRows(valid);
      if (!consolidated.length) {
        return res.status(400).json({
          error: `No se encontraron registros diarios válidos de ${config.label}`,
          received: rows.length,
          rejected: rejected.length,
          issues: rejected.slice(0, 25),
        });
      }

      const result = await transaction(async (connection) => {
        const representativeByCode = new Map();
        for (const row of consolidated) {
          const key = row.producto_codigo.toUpperCase();
          if (!representativeByCode.has(key)) representativeByCode.set(key, row);
        }
        const productIds = new Map();
        for (const [key, row] of representativeByCode.entries()) {
          productIds.set(key, await ensureProduct(connection, row, row.index));
        }

        const dates = consolidated.map((row) => row.fecha).sort();
        const fechaMin = dates[0];
        const fechaMax = dates[dates.length - 1];
        const ids = [...new Set(productIds.values())];
        const dimensionSelect = config.dimensions.map((dimension) => `t.${dimension.name}`).join(', ');
        const [existingRows] = await connection.execute(
          `SELECT DATE_FORMAT(t.fecha, '%Y-%m-%d') AS fecha, t.producto_id${dimensionSelect ? `, ${dimensionSelect}` : ''}
           FROM ${config.table} t
           WHERE t.fecha >= ? AND t.fecha <= ? AND t.producto_id IN (${ids.map(() => '?').join(',')})`,
          [fechaMin, fechaMax, ...ids]
        );
        const existingKeys = new Set(
          existingRows.map((row) => [
            row.fecha,
            row.producto_id,
            ...config.dimensions.map((dimension) => String(row[dimension.name] || '').toUpperCase()),
          ].join('|'))
        );
        let inserted = 0;
        let updated = 0;
        for (const row of consolidated) {
          row.producto_id = productIds.get(row.producto_codigo.toUpperCase());
          const key = [
            row.fecha,
            row.producto_id,
            ...config.dimensions.map((dimension) => row[dimension.name].toUpperCase()),
          ].join('|');
          if (existingKeys.has(key)) updated += 1;
          else inserted += 1;
        }

        const columns = ['fecha', 'producto_id', 'cantidad', ...config.dimensions.map((dimension) => dimension.name)];
        for (let offset = 0; offset < consolidated.length; offset += 500) {
          const chunk = consolidated.slice(offset, offset + 500);
          const placeholders = chunk.map(() => `(${columns.map(() => '?').join(',')})`).join(',');
          const values = chunk.flatMap((row) => columns.map((column) => row[column]));
          await connection.query(
            `INSERT INTO ${config.table} (${columns.join(', ')}) VALUES ${placeholders}
             ON DUPLICATE KEY UPDATE cantidad = VALUES(cantidad), updated_at = CURRENT_TIMESTAMP`,
            values
          );
        }

        const detail = {
          duplicadasEnArchivo: valid.length - consolidated.length,
          errores: rejected.slice(0, 25),
        };
        const [importRecord] = await connection.execute(
          `INSERT INTO importaciones_operativas
             (tipo, usuario_id, archivo, filas_recibidas, filas_validas, filas_rechazadas,
              filas_consolidadas, registros_insertados, registros_actualizados, fecha_min, fecha_max, detalle)
           VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)`,
          [
            config.type,
            req.user.id,
            archivo,
            rows.length,
            valid.length,
            rejected.length,
            consolidated.length,
            inserted,
            updated,
            fechaMin,
            fechaMax,
            JSON.stringify(detail),
          ]
        );
        await writeAudit(connection, req.user.id, 'importar', config.type, String(importRecord.insertId), {
          archivo,
          recibidas: rows.length,
          insertadas: inserted,
          actualizadas: updated,
          rechazadas: rejected.length,
        });
        return {
          importId: importRecord.insertId,
          valid: valid.length,
          rejected: rejected.length,
          consolidated: consolidated.length,
          duplicatesInFile: valid.length - consolidated.length,
          inserted,
          updated,
          fechaMin,
          fechaMax,
        };
      });

      res.status(201).json({ ok: true, received: rows.length, ...result, issues: rejected.slice(0, 25) });
    } catch (error) {
      next(error);
    }
  }

  router.post('/importar', requireRole('admin', 'operador'), importRows);
  router.post('/bulk', requireRole('admin', 'operador'), importRows);

  router.get('/importaciones', async (_req, res, next) => {
    try {
      const rows = await query(
        `SELECT i.id, i.tipo, i.archivo, i.filas_recibidas, i.filas_validas, i.filas_rechazadas,
                i.filas_consolidadas, i.registros_insertados, i.registros_actualizados,
                i.fecha_min, i.fecha_max, i.detalle, i.created_at, u.usuario, u.nombre
         FROM importaciones_operativas i
         INNER JOIN usuarios u ON u.id = i.usuario_id
         WHERE i.tipo = ? ORDER BY i.id DESC LIMIT 50`,
        [config.type]
      );
      res.json({ ok: true, rows });
    } catch (error) {
      next(error);
    }
  });

  async function getRows(req, res, next, syncAll = false) {
    try {
      let where = '';
      let params = [];
      let limitClause = '';
      let pagination = null;
      if (syncAll) {
        pagination = paginationFromQuery(req.query);
        where = 'WHERE t.id > ?';
        params = [pagination.cursor];
        limitClause = 'LIMIT ?';
        params.push(pagination.limit + 1);
      } else {
        const { start, next: nextMonth } = monthRange(req.query.mes);
        where = 'WHERE t.fecha >= ? AND t.fecha < ?';
        params = [start, nextMonth];
      }
      const dimensionSelect = config.dimensions.map((dimension) => `t.${dimension.name}`).join(', ');
      const rows = await query(
        `SELECT t.id, DATE_FORMAT(t.fecha, '%Y-%m-%d') AS fecha,
                p.codigo AS producto_codigo, p.nombre AS producto_nombre, t.cantidad
                ${dimensionSelect ? `, ${dimensionSelect}` : ''}
         FROM ${config.table} t
         INNER JOIN productos p ON p.id = t.producto_id
         ${where}
         ORDER BY ${syncAll ? 't.id' : 't.fecha, p.codigo'}
         ${limitClause}`,
        params
      );
      if (!syncAll) return res.json({ ok: true, rows });
      const hasMore = rows.length > pagination.limit;
      const visibleRows = hasMore ? rows.slice(0, pagination.limit) : rows;
      res.json({
        ok: true,
        rows: visibleRows,
        hasMore,
        nextCursor: visibleRows.length ? Number(visibleRows[visibleRows.length - 1].id) : pagination.cursor,
      });
    } catch (error) {
      next(error);
    }
  }

  router.get('/sync', (req, res, next) => getRows(req, res, next, true));
  router.get('/', (req, res, next) => getRows(req, res, next, false));
  return router;
}

module.exports = createOperationalRouter;
