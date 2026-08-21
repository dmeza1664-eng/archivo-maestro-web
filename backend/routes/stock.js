const express = require('express');
const { query, transaction } = require('../db');
const { requireAuth, requireRole, writeAudit } = require('../auth');
const {
  ensureProduct,
  extractRows,
  firstValue,
  mesFromDate,
  toDateOnly,
  toNumber,
  validateMes,
} = require('./helpers');

const router = express.Router();
router.use(requireAuth);

router.post('/bulk', requireRole('admin', 'operador'), async (req, res, next) => {
  try {
    const rows = extractRows(req.body, ['stock', 'rows', 'data']);

    const result = await transaction(async (connection) => {
      let insertedOrUpdated = 0;

      for (const [index, row] of rows.entries()) {
        const productoId = await ensureProduct(connection, row, index);
        const explicitMes = firstValue(row, ['mes', 'month']);
        const fechaValue = firstValue(row, ['fecha', 'date']);
        const mes = explicitMes
          ? String(explicitMes).slice(0, 7)
          : mesFromDate(toDateOnly(fechaValue, 'fecha o mes', index));
        validateMes(mes);

        const cantidad = toNumber(firstValue(row, ['cantidad', 'stock', 'stock_fijo', 'unidades']), 'cantidad', index);

        await connection.execute(
          `INSERT INTO stock_fijo (mes, producto_id, cantidad)
           VALUES (?, ?, ?)
           ON DUPLICATE KEY UPDATE
             cantidad = VALUES(cantidad),
             updated_at = CURRENT_TIMESTAMP`,
          [mes, productoId, cantidad]
        );
        insertedOrUpdated += 1;
      }

      await writeAudit(connection, req.user.id, 'importar', 'stock_fijo', '', {
        recibidas: rows.length,
        actualizadas: insertedOrUpdated,
      });

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
    validateMes(mes);

    const rows = await query(
      `SELECT
         s.id,
         s.mes,
         p.codigo AS producto_codigo,
         p.nombre AS producto_nombre,
         s.cantidad
       FROM stock_fijo s
       INNER JOIN productos p ON p.id = s.producto_id
       WHERE s.mes = ?
       ORDER BY p.codigo`,
      [mes]
    );

    res.json({ ok: true, mes, rows });
  } catch (error) {
    next(error);
  }
});

module.exports = router;
