const express = require('express');
const { listProduccion, listStock, listVentas } = require('./lists');

const router = express.Router();

router.get('/', async (_req, res, next) => {
  try {
    const [ventas, stock, produccion] = await Promise.all([
      listVentas(),
      listStock(),
      listProduccion(),
    ]);

    res.json({
      ok: true,
      ventas: ventas.map((row) => ({ ...row, fecha: row.fecha_iso || row.fecha })),
      stock,
      produccion: produccion.map((row) => ({ ...row, fecha: row.fecha_iso || row.fecha })),
    });
  } catch (error) {
    next(error);
  }
});

module.exports = router;
