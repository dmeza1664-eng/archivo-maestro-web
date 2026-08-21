const express = require('express');
const { query, transaction } = require('../db');
const { requireAuth, requireRole, writeAudit } = require('../auth');

const router = express.Router();
router.use(requireAuth);

function validateType(value) {
  const type = String(value || '').trim().toLowerCase();
  if (!/^[a-z0-9_-]{2,80}$/.test(type)) {
    const error = new Error('Tipo de respaldo inválido');
    error.status = 400;
    throw error;
  }
  return type;
}

function parseJson(value, fallback) {
  if (value === null || value === undefined) return fallback;
  if (typeof value !== 'string') return value;
  try {
    return JSON.parse(value);
  } catch {
    return fallback;
  }
}

router.get('/', async (_req, res, next) => {
  try {
    const rows = await query(
      `SELECT r.id, r.tipo, r.periodo, r.version, r.archivos, r.created_at,
              u.usuario, u.nombre
       FROM respaldos_datos r
       INNER JOIN usuarios u ON u.id = r.usuario_id
       INNER JOIN (
         SELECT tipo, periodo, MAX(version) AS version
         FROM respaldos_datos GROUP BY tipo, periodo
       ) latest ON latest.tipo = r.tipo AND latest.periodo = r.periodo AND latest.version = r.version
       ORDER BY r.tipo, r.periodo`
    );
    res.json({ ok: true, rows: rows.map((row) => ({ ...row, archivos: parseJson(row.archivos, {}) })) });
  } catch (error) {
    next(error);
  }
});

router.get('/:type/history', async (req, res, next) => {
  try {
    const type = validateType(req.params.type);
    const periodo = String(req.query.periodo || 'global').slice(0, 20);
    const rows = await query(
      `SELECT r.id, r.tipo, r.periodo, r.version, r.archivos, r.created_at,
              u.usuario, u.nombre
       FROM respaldos_datos r
       INNER JOIN usuarios u ON u.id = r.usuario_id
       WHERE r.tipo = ? AND r.periodo = ?
       ORDER BY r.version DESC LIMIT 50`,
      [type, periodo]
    );
    res.json({ ok: true, rows: rows.map((row) => ({ ...row, archivos: parseJson(row.archivos, {}) })) });
  } catch (error) {
    next(error);
  }
});

router.get('/:type', async (req, res, next) => {
  try {
    const type = validateType(req.params.type);
    const periodo = String(req.query.periodo || 'global').slice(0, 20);
    const rows = await query(
      `SELECT r.id, r.tipo, r.periodo, r.version, r.archivos, r.contenido, r.created_at,
              u.usuario, u.nombre
       FROM respaldos_datos r
       INNER JOIN usuarios u ON u.id = r.usuario_id
       WHERE r.tipo = ? AND r.periodo = ?
       ORDER BY r.version DESC LIMIT 1`,
      [type, periodo]
    );
    if (!rows.length) return res.status(404).json({ error: 'No existe un respaldo para este tipo' });
    const row = rows[0];
    res.json({
      ok: true,
      snapshot: {
        ...row,
        archivos: parseJson(row.archivos, {}),
        contenido: parseJson(row.contenido, {}),
      },
    });
  } catch (error) {
    next(error);
  }
});

router.post('/:type', requireRole('admin', 'operador'), async (req, res, next) => {
  try {
    const type = validateType(req.params.type);
    const periodo = String(req.body?.periodo || 'global').slice(0, 20);
    const contenido = req.body?.contenido;
    if (!contenido || typeof contenido !== 'object' || Array.isArray(contenido)) {
      return res.status(400).json({ error: 'El respaldo debe incluir un objeto contenido' });
    }
    const serialized = JSON.stringify(contenido);
    if (Buffer.byteLength(serialized, 'utf8') > 3.5 * 1024 * 1024) {
      return res.status(413).json({ error: 'El respaldo excede el límite de 3.5 MB' });
    }
    const archivos = req.body?.archivos && typeof req.body.archivos === 'object' ? req.body.archivos : {};
    const result = await transaction(async (connection) => {
      const [previous] = await connection.execute(
        `SELECT version FROM respaldos_datos
         WHERE tipo = ? AND periodo = ? ORDER BY version DESC LIMIT 1 FOR UPDATE`,
        [type, periodo]
      );
      const version = Number(previous[0]?.version || 0) + 1;
      const [created] = await connection.execute(
        `INSERT INTO respaldos_datos (tipo, version, periodo, archivos, contenido, usuario_id)
         VALUES (?, ?, ?, ?, ?, ?)`,
        [type, version, periodo, JSON.stringify(archivos), serialized, req.user.id]
      );
      await writeAudit(connection, req.user.id, 'respaldar', 'datos', `${type}:${periodo}`, {
        version,
        bytes: Buffer.byteLength(serialized, 'utf8'),
      });
      return { id: created.insertId, version };
    });
    res.status(201).json({ ok: true, ...result, tipo: type, periodo });
  } catch (error) {
    next(error);
  }
});

module.exports = router;
