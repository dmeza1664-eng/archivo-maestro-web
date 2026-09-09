const express = require('express');
const { query, transaction } = require('../db');
const {
  addMonths,
  daysInMonth,
  ensureProduct,
  extractRows,
  firstValue,
  getProduccionSugerida,
  monthRange,
  toDateOnly,
  toNumber,
  validateMes,
} = require('./helpers');

const router = express.Router();

router.post('/bulk', async (req, res, next) => {
  try {
    const rows = extractRows(req.body, ['pronostico', 'rows', 'data']);
    const metodoDefault = req.body?.metodo || 'weekday_colchon_regla';

    const result = await transaction(async (connection) => {
      let insertedOrUpdated = 0;

      for (const [index, row] of rows.entries()) {
        const productoId = await ensureProduct(connection, row, index);
        const fecha = toDateOnly(firstValue(row, ['fecha', 'dia', 'date']), 'fecha', index);
        const cantidad = toNumber(
          firstValue(row, ['cantidad_pronosticada', 'produccion_sugerida', 'cantidad', 'unidades']),
          'cantidad_pronosticada',
          index
        );
        const metodo = String(firstValue(row, ['metodo']) || metodoDefault);

        await connection.execute(
          `INSERT INTO pronostico_diario (fecha, producto_id, cantidad_pronosticada, metodo)
           VALUES (?, ?, ?, ?)
           ON DUPLICATE KEY UPDATE
             cantidad_pronosticada = VALUES(cantidad_pronosticada),
             updated_at = CURRENT_TIMESTAMP`,
          [fecha, productoId, cantidad, metodo]
        );
        insertedOrUpdated += 1;
      }

      return { insertedOrUpdated };
    });

    res.status(201).json({
      ok: true,
      received: rows.length,
      ...result,
    });
  } catch (error) {
    next(error);
  }
});

router.post('/calcular', async (req, res, next) => {
  try {
    const mes = req.body?.mes;
    const metodo = req.body?.metodo || 'weekday_colchon_regla';
    const mesesHistoricos = Number(req.body?.mesesHistoricos || 3);
    const dailyBufferPct = Number(req.body?.dailyBufferPct || 0);
    const weekendBoost = Number(req.body?.weekendBoost || 1);
    validateMes(mes);

    const { start } = monthRange(mes);
    const historyStart = addMonths(start, -mesesHistoricos);
    const days = daysInMonth(mes);
    const productCodes = Array.isArray(req.body?.productos)
      ? req.body.productos.map(String).filter(Boolean)
      : [];

    const result = await transaction(async (connection) => {
      let productFilter = '';
      const params = [historyStart, start];

      if (productCodes.length > 0) {
        productFilter = ` AND p.codigo IN (${productCodes.map(() => '?').join(', ')})`;
        params.push(...productCodes);
      }

      const [sales] = await connection.execute(
        `SELECT
           p.id AS producto_id,
           p.codigo AS producto_codigo,
           p.nombre AS producto_nombre,
           v.fecha,
           v.cantidad
         FROM productos p
         LEFT JOIN ventas_diarias v
           ON v.producto_id = p.id
          AND v.fecha >= ?
          AND v.fecha < ?
         WHERE p.activo = 1${productFilter}
         ORDER BY p.codigo, v.fecha`,
        params
      );

      const byProduct = new Map();
      for (const row of sales) {
        const current = byProduct.get(row.producto_id) || {
          producto_id: row.producto_id,
          producto_codigo: row.producto_codigo,
          producto_nombre: row.producto_nombre,
          weekdays: new Map(),
          total: 0,
          days: 0,
        };
        if (row.fecha) {
          const weekday = new Date(`${row.fecha}T00:00:00`).getDay();
          const bucket = current.weekdays.get(weekday) || { total: 0, count: 0 };
          const cantidad = toNumber(row.cantidad, 'cantidad', 0, 0);
          bucket.total += cantidad;
          bucket.count += 1;
          current.weekdays.set(weekday, bucket);
          current.total += cantidad;
          current.days += 1;
        }
        byProduct.set(row.producto_id, current);
      }

      let insertedOrUpdated = 0;
      for (const product of byProduct.values()) {
        const weekdayAverage = (weekday) => {
          const bucket = product.weekdays.get(weekday);
          return bucket && bucket.count ? bucket.total / bucket.count : 0;
        };
        const flatAverage = product.days ? product.total / product.days : 0;

        for (let day = 1; day <= days; day += 1) {
          const fecha = `${mes}-${String(day).padStart(2, '0')}`;
          const weekday = new Date(`${fecha}T00:00:00`).getDay();
          let dailyForecast = metodo === 'promedio_ventas'
            ? flatAverage
            : weekdayAverage(weekday);
          if (weekday === 0 || weekday === 6) dailyForecast *= weekendBoost || 1;
          dailyForecast += dailyForecast * (Math.max(0, dailyBufferPct) / 100);
          const suggested = getProduccionSugerida(product.producto_nombre || product.producto_codigo, dailyForecast);

          await connection.execute(
            `INSERT INTO pronostico_diario (fecha, producto_id, cantidad_pronosticada, metodo)
             VALUES (?, ?, ?, ?)
             ON DUPLICATE KEY UPDATE
               cantidad_pronosticada = VALUES(cantidad_pronosticada),
               updated_at = CURRENT_TIMESTAMP`,
            [fecha, product.producto_id, suggested, metodo]
          );
          insertedOrUpdated += 1;
        }
      }

      return {
        productosCalculados: byProduct.size,
        diasPorProducto: days,
        insertedOrUpdated,
      };
    });

    res.status(201).json({
      ok: true,
      mes,
      metodo,
      mesesHistoricos,
      dailyBufferPct,
      weekendBoost,
      ...result,
    });
  } catch (error) {
    next(error);
  }
});

router.get('/', async (req, res, next) => {
  try {
    const { mes } = req.query;
    if (!mes) {
      const rows = await query(
        `SELECT
           pd.id,
           pd.fecha,
           DATE_FORMAT(pd.fecha, '%Y-%m-%d') AS fecha_iso,
           DATE_FORMAT(pd.fecha, '%Y-%m') AS mes,
           p.codigo AS producto_codigo,
           p.nombre AS producto_nombre,
           pd.cantidad_pronosticada,
           pd.metodo
         FROM pronostico_diario pd
         INNER JOIN productos p ON p.id = pd.producto_id
         ORDER BY pd.fecha, p.codigo, pd.metodo`
      );
      return res.json({ ok: true, mes: null, rows });
    }

    const { start, next: nextMonth } = monthRange(mes);

    const rows = await query(
      `SELECT
         pd.id,
         pd.fecha,
         DATE_FORMAT(pd.fecha, '%Y-%m') AS mes,
         p.codigo AS producto_codigo,
         p.nombre AS producto_nombre,
         pd.cantidad_pronosticada,
         pd.metodo
       FROM pronostico_diario pd
       INNER JOIN productos p ON p.id = pd.producto_id
       WHERE pd.fecha >= ? AND pd.fecha < ?
       ORDER BY pd.fecha, p.codigo, pd.metodo`,
      [start, nextMonth]
    );

    const mensual = await query(
      `SELECT
         DATE_FORMAT(pd.fecha, '%Y-%m') AS mes,
         p.codigo AS producto_codigo,
         p.nombre AS producto_nombre,
         pd.metodo,
         SUM(pd.cantidad_pronosticada) AS cantidad_pronosticada_mensual
       FROM pronostico_diario pd
       INNER JOIN productos p ON p.id = pd.producto_id
       WHERE pd.fecha >= ? AND pd.fecha < ?
       GROUP BY DATE_FORMAT(pd.fecha, '%Y-%m'), p.codigo, p.nombre, pd.metodo
       ORDER BY p.codigo, pd.metodo`,
      [start, nextMonth]
    );

    res.json({ ok: true, mes, rows, mensual });
  } catch (error) {
    next(error);
  }
});

module.exports = router;
