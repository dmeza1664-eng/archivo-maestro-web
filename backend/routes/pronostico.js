const express = require('express');
const { query, transaction } = require('../db');
const { requireAuth, requireRole, writeAudit } = require('../auth');
const {
  addMonths,
  daysInMonth,
  firstValue,
  monthRange,
  toNumber,
  validateMes,
} = require('./helpers');

const router = express.Router();
router.use(requireAuth);

router.post('/calcular', requireRole('admin', 'operador'), async (req, res, next) => {
  try {
    const mes = req.body?.mes;
    const metodo = req.body?.metodo || 'promedio_ventas';
    const mesesHistoricos = Number(req.body?.mesesHistoricos || 3);
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

      const [averages] = await connection.execute(
        `SELECT
           p.id AS producto_id,
           p.codigo AS producto_codigo,
           p.nombre AS producto_nombre,
           COALESCE(SUM(v.cantidad) / NULLIF(COUNT(DISTINCT v.fecha), 0), 0) AS promedio_diario
         FROM productos p
         LEFT JOIN ventas_diarias v
           ON v.producto_id = p.id
          AND v.fecha >= ?
          AND v.fecha < ?
         WHERE p.activo = 1${productFilter}
         GROUP BY p.id, p.codigo, p.nombre
         ORDER BY p.codigo`,
        params
      );

      let insertedOrUpdated = 0;
      for (const product of averages) {
        const dailyForecast = toNumber(
          firstValue(product, ['promedio_diario']),
          'promedio_diario',
          0,
          0
        );

        for (let day = 1; day <= days; day += 1) {
          const fecha = `${mes}-${String(day).padStart(2, '0')}`;
          await connection.execute(
            `INSERT INTO pronostico_diario (fecha, producto_id, cantidad_pronosticada, metodo)
             VALUES (?, ?, ?, ?)
             ON DUPLICATE KEY UPDATE
               cantidad_pronosticada = VALUES(cantidad_pronosticada),
               updated_at = CURRENT_TIMESTAMP`,
            [fecha, product.producto_id, dailyForecast, metodo]
          );
          insertedOrUpdated += 1;
        }
      }

      await writeAudit(connection, req.user.id, 'calcular', 'pronostico_diario', mes, {
        metodo,
        mesesHistoricos,
        productosCalculados: averages.length,
      });

      return {
        productosCalculados: averages.length,
        diasPorProducto: days,
        insertedOrUpdated,
      };
    });

    res.status(201).json({
      ok: true,
      mes,
      metodo,
      mesesHistoricos,
      ...result,
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
