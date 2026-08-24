const express = require('express');
const { query } = require('../db');
const { requireAuth, requireRole } = require('../auth');
const { monthRange } = require('./helpers');

const router = express.Router();
router.use(requireAuth);

router.post('/calcular', requireRole('admin', 'operador'), async (req, res) => {
  const mes = req.body?.mes || '';
  return res.status(409).json({
    error: 'Este cálculo plano ya no es la fuente oficial del pronóstico.',
    oficial: 'El pronóstico vigente es el snapshot forecast-frozen del mes, generado en la aplicación con el modelo categorySeasonal.',
    mes,
  });
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
