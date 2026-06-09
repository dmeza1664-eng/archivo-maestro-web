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
    const rows = extractRows(req.body, ['ventas', 'rows', 'data']);

    const result = await transaction(async (connection) => {
      let insertedOrUpdated = 0;

      for (const [index, row] of rows.entries()) {
        const productoId = await ensureProduct(connection, row, index);
        const fecha = toDateOnly(firstValue(row, ['fecha', 'dia', 'date']), 'fecha', index);
        const cantidad = toNumber(firstValue(row, ['cantidad', 'venta', 'ventas', 'unidades']), 'cantidad', index);
        const importeValue = firstValue(row, ['importe', 'monto', 'total']);
        const importe = importeValue === undefined ? null : toNumber(importeValue, 'importe', index, null);
        const canal = String(firstValue(row, ['canal', 'zona', 'sucursal']) || '').trim();
        const cliente = String(firstValue(row, ['cliente', 'customer']) || '').trim();

        await connection.execute(
          `INSERT INTO ventas_diarias (fecha, producto_id, cantidad, importe, canal, cliente)
           VALUES (?, ?, ?, ?, ?, ?)
           ON DUPLICATE KEY UPDATE
             cantidad = VALUES(cantidad),
             importe = VALUES(importe),
             updated_at = CURRENT_TIMESTAMP`,
          [fecha, productoId, cantidad, importe, canal, cliente]
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
         v.id,
         v.fecha,
         DATE_FORMAT(v.fecha, '%Y-%m') AS mes,
         p.codigo AS producto_codigo,
         p.nombre AS producto_nombre,
         v.cantidad,
         v.importe,
         v.canal,
         v.cliente
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
