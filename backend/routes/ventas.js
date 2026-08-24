const express = require('express');
const { query, transaction } = require('../db');
const { requireAuth, requireRole, writeAudit } = require('../auth');
const {
  ensureProduct,
  extractRows,
  firstValue,
  monthRange,
  paginationFromQuery,
  parseImportMeta,
  toDateOnly,
  toNumber,
} = require('./helpers');

const router = express.Router();
router.use(requireAuth);

function normalizedText(value, maxLength) {
  return String(value || '').trim().slice(0, maxLength);
}

function normalizeImportRows(rows) {
  const valid = [];
  const rejected = [];
  for (const [index, row] of rows.entries()) {
    try {
      if (row.monthlyTotal) throw new Error('es un total mensual, no una venta diaria');
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
      const cantidad = toNumber(firstValue(row, ['cantidad', 'venta', 'ventas', 'unidades']), 'cantidad', index);
      const importeValue = firstValue(row, ['importe', 'monto', 'total']);
      const importe = importeValue === undefined ? null : toNumber(importeValue, 'importe', index, null);
      const canal = normalizedText(firstValue(row, ['canal', 'zona', 'sucursal']), 120);
      const cliente = normalizedText(firstValue(row, ['cliente', 'customer']), 180);
      valid.push({
        index,
        fecha,
        producto_codigo: productoCodigo,
        producto_nombre: productoNombre,
        cantidad,
        importe,
        canal,
        cliente,
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
    const key = [row.fecha, row.producto_codigo.toUpperCase(), row.canal.toUpperCase(), row.cliente.toUpperCase()].join('|');
    const current = map.get(key);
    if (!current) {
      map.set(key, { ...row });
      continue;
    }
    current.cantidad += row.cantidad;
    if (row.importe !== null) current.importe = (current.importe || 0) + row.importe;
  }
  return [...map.values()];
}

async function importSales(req, res, next) {
  try {
    const rows = extractRows(req.body, ['ventas', 'rows', 'data']);
    const archivo = normalizedText(req.body?.archivo, 255);
    const batchMeta = parseImportMeta(req.body, rows.length);
    const { valid, rejected } = normalizeImportRows(rows);
    const consolidated = consolidateRows(valid);
    if (!consolidated.length) {
      return res.status(400).json({
        error: 'No se encontraron ventas válidas para importar',
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
      const [existingRows] = await connection.execute(
        `SELECT DATE_FORMAT(fecha, '%Y-%m-%d') AS fecha, producto_id, canal, cliente
         FROM ventas_diarias
         WHERE fecha >= ? AND fecha <= ? AND producto_id IN (${ids.map(() => '?').join(',')})`,
        [fechaMin, fechaMax, ...ids]
      );
      const existingKeys = new Set(
        existingRows.map((row) => [row.fecha, row.producto_id, row.canal.toUpperCase(), row.cliente.toUpperCase()].join('|'))
      );
      let inserted = 0;
      let updated = 0;
      for (const row of consolidated) {
        const productId = productIds.get(row.producto_codigo.toUpperCase());
        const key = [row.fecha, productId, row.canal.toUpperCase(), row.cliente.toUpperCase()].join('|');
        if (existingKeys.has(key)) updated += 1;
        else inserted += 1;
        row.producto_id = productId;
      }

      for (let offset = 0; offset < consolidated.length; offset += 500) {
        const chunk = consolidated.slice(offset, offset + 500);
        const placeholders = chunk.map(() => '(?, ?, ?, ?, ?, ?)').join(',');
        const values = chunk.flatMap((row) => [
          row.fecha,
          row.producto_id,
          row.cantidad,
          row.importe,
          row.canal,
          row.cliente,
        ]);
        await connection.query(
          `INSERT INTO ventas_diarias (fecha, producto_id, cantidad, importe, canal, cliente)
           VALUES ${placeholders}
           ON DUPLICATE KEY UPDATE
             cantidad = VALUES(cantidad),
             importe = VALUES(importe),
             updated_at = CURRENT_TIMESTAMP`,
          values
        );
      }

      const detail = {
        duplicadasEnArchivo: valid.length - consolidated.length,
        errores: rejected.slice(0, 25),
        importRunId: batchMeta.importRunId,
        batchIndex: batchMeta.batchIndex,
        batchTotal: batchMeta.batchTotal,
      };
      const [importRecord] = await connection.execute(
        `INSERT INTO importaciones_ventas
           (usuario_id, archivo, filas_recibidas, filas_validas, filas_rechazadas,
            filas_consolidadas, registros_insertados, registros_actualizados, fecha_min, fecha_max, detalle)
         VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)`,
        [
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
      await writeAudit(connection, req.user.id, 'importar', 'ventas', String(importRecord.insertId), {
        archivo,
        recibidas: rows.length,
        insertadas: inserted,
        actualizadas: updated,
        rechazadas: rejected.length,
        importRunId: batchMeta.importRunId,
        batchIndex: batchMeta.batchIndex,
        batchTotal: batchMeta.batchTotal,
      });
      return {
        importId: importRecord.insertId,
        importRunId: batchMeta.importRunId,
        batchIndex: batchMeta.batchIndex,
        batchTotal: batchMeta.batchTotal,
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

    res.status(201).json({
      ok: true,
      received: rows.length,
      ...result,
      issues: rejected.slice(0, 25),
    });
  } catch (error) {
    next(error);
  }
}

router.post('/importar', requireRole('admin', 'operador'), importSales);
router.post('/bulk', requireRole('admin', 'operador'), importSales);

router.get('/importaciones', async (_req, res, next) => {
  try {
    const rows = await query(
      `SELECT i.id, i.archivo, i.filas_recibidas, i.filas_validas, i.filas_rechazadas,
              i.filas_consolidadas, i.registros_insertados, i.registros_actualizados,
              i.fecha_min, i.fecha_max, i.detalle, i.created_at, u.usuario, u.nombre
       FROM importaciones_ventas i
       INNER JOIN usuarios u ON u.id = i.usuario_id
       ORDER BY i.id DESC LIMIT 50`
    );
    res.json({ ok: true, rows });
  } catch (error) {
    next(error);
  }
});

router.get('/sync', async (req, res, next) => {
  try {
    const { cursor, limit } = paginationFromQuery(req.query);
    const rows = await query(
      `SELECT v.id, DATE_FORMAT(v.fecha, '%Y-%m-%d') AS fecha,
               p.codigo AS producto_codigo, p.nombre AS producto_nombre,
               v.cantidad, v.importe, v.canal, v.cliente
        FROM ventas_diarias v
        INNER JOIN productos p ON p.id = v.producto_id
        WHERE v.id > ?
        ORDER BY v.id
        LIMIT ?`,
      [cursor, limit + 1]
    );
    const hasMore = rows.length > limit;
    const visibleRows = hasMore ? rows.slice(0, limit) : rows;
    res.json({
      ok: true,
      rows: visibleRows,
      hasMore,
      nextCursor: visibleRows.length ? Number(visibleRows[visibleRows.length - 1].id) : cursor,
    });
  } catch (error) {
    next(error);
  }
});

router.get('/', async (req, res, next) => {
  try {
    const { mes } = req.query;
    const { start, next: nextMonth } = monthRange(mes);
    const rows = await query(
      `SELECT v.id, v.fecha, DATE_FORMAT(v.fecha, '%Y-%m') AS mes,
              p.codigo AS producto_codigo, p.nombre AS producto_nombre,
              v.cantidad, v.importe, v.canal, v.cliente
       FROM ventas_diarias v
       INNER JOIN productos p ON p.id = v.producto_id
       WHERE v.fecha >= ? AND v.fecha < ?
       ORDER BY v.fecha, p.codigo, v.canal, v.cliente`,
      [start, nextMonth]
    );
    res.json({ ok: true, mes, rows });
  } catch (error) {
    next(error);
  }
});

module.exports = router;
