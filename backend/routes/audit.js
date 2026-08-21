const express = require('express');
const { query } = require('../db');
const { requireAuth, requireRole } = require('../auth');

const router = express.Router();
router.use(requireAuth, requireRole('admin'));

router.get('/', async (req, res, next) => {
  try {
    const limit = Math.min(200, Math.max(1, Number(req.query.limit || 100)));
    const rows = await query(
      `SELECT b.id, b.accion, b.entidad, b.clave_entidad, b.detalle, b.created_at,
              u.usuario, u.nombre
       FROM bitacora b
       LEFT JOIN usuarios u ON u.id = b.usuario_id
       ORDER BY b.id DESC LIMIT ${limit}`
    );
    res.json({ ok: true, rows });
  } catch (error) {
    next(error);
  }
});

module.exports = router;
