const express = require('express');
const { query, transaction } = require('../db');
const {
  ensureProduct,
  extractRows,
  firstValue,
  monthRange,
  toDateOnly,
  toNumber,
} = require('./helpers');

const router = express.Router();

router.post('/bulk', async (req, res, next) => {
  try {
    const rows = extractRows(req.body, ['produccion', 'produccionReal', 'rows', 'data']);

    const result = await transaction(async (connection) => {
      let insertedOrUpdated = 0;

      for (const [index, row] of rows.entries()) {
        const productoId = await ensureProduct(connection, row, index);
        const fecha = toDateOnly(firstValue(row, ['fecha', 'dia', 'date']), 'fecha', index);
        const cantidad = toNumber(firstValue(row, ['cantidad', 'produccion', 'produccion_real', 'unidades']), 'cantidad', index);
        const turno = String(firstValue(row, ['turno', 'shift']) || '').trim();

        await connection.execute(
          `INSERT INTO produccion_real (fecha, producto_id, cantidad, turno)
           VALUES (?, ?, ?, ?)
           ON DUPLICATE KEY UPDATE
             cantidad = VALUES(cantidad),
             updated_at = CURRENT_TIMESTAMP`,
          [fecha, productoId, cantidad, turno]
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

router.get('/', async (req, res, next) => {
  try {
    const { mes } = req.query;
    const { start, next: nextMonth } = monthRange(mes);

    const rows = await query(
      `SELECT
         pr.id,
         pr.fecha,
         DATE_FORMAT(pr.fecha, '%Y-%m') AS mes,
         p.codigo AS producto_codigo,
         p.nombre AS producto_nombre,
         pr.cantidad,
         pr.turno
       FROM produccion_real pr
       INNER JOIN productos p ON p.id = pr.producto_id
       WHERE pr.fecha >= ? AND pr.fecha < ?
       ORDER BY pr.fecha, p.codigo, pr.turno`,
      [start, nextMonth]
    );

    res.json({ ok: true, mes, rows });
  } catch (error) {
    next(error);
  }
});

module.exports = router;
