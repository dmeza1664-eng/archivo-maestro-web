const express = require('express');
const { query, transaction } = require('../db');
const { requireAuth, requireRole, writeAudit } = require('../auth');

const router = express.Router();
router.use(requireAuth);

const ALLOWED_SNAPSHOT_TYPES = new Set(['workspace', 'forecast-frozen', 'monthly-review']);
const SNAPSHOT_WRITE_ROLES = {
  workspace: ['admin', 'operador'],
  'forecast-frozen': ['admin', 'operador'],
  'monthly-review': ['admin', 'operador'],
};

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

function httpError(status, message) {
  const error = new Error(message);
  error.status = status;
  return error;
}

function assertFiniteNumber(value, field) {
  const number = Number(value);
  if (!Number.isFinite(number)) throw httpError(400, `${field} debe ser numérico`);
  return number;
}

function validateWorkspaceContent(contenido) {
  if (!contenido || typeof contenido !== 'object' || Array.isArray(contenido)) {
    throw httpError(400, 'El respaldo de workspace debe ser un objeto');
  }
}

function validateFrozenContent(contenido, periodo) {
  if (!contenido || typeof contenido !== 'object' || Array.isArray(contenido)) {
    throw httpError(400, 'El pronóstico congelado debe incluir un objeto contenido');
  }
  const selectedMonth = String(contenido.selectedMonth || contenido.period || '').slice(0, 20);
  if (selectedMonth && selectedMonth !== periodo) {
    throw httpError(400, 'El mes del pronóstico congelado no coincide con el periodo');
  }
  if (!Array.isArray(contenido.rows) || !contenido.rows.length) {
    throw httpError(400, 'El pronóstico congelado debe incluir filas por producto');
  }
  for (const row of contenido.rows) {
    if (!row || typeof row !== 'object') throw httpError(400, 'Hay una fila de congelado inválida');
    assertFiniteNumber(row.pronosticoBase ?? row.pronosticoVenta ?? 0, `Pronóstico de ${row.producto || 'producto'}`);
  }
}

function validateMonthlyReviewContent(contenido, periodo, user) {
  if (!contenido || typeof contenido !== 'object' || Array.isArray(contenido)) {
    throw httpError(400, 'La revisión mensual debe incluir un objeto contenido');
  }
  const period = String(contenido.period || contenido.selectedMonth || '').slice(0, 20);
  if (period !== periodo) throw httpError(400, 'El mes de la revisión no coincide con el periodo');
  const state = String(contenido.state || 'draft');
  if (!['draft', 'approved'].includes(state)) {
    throw httpError(400, 'El estado de la revisión debe ser draft o approved');
  }
  if (state === 'approved' && user.rol !== 'admin') {
    throw httpError(403, 'Solo un administrador puede aprobar la revisión mensual');
  }
  const frozenVersion = Number(contenido.sourceFrozenVersion);
  if (!Number.isInteger(frozenVersion) || frozenVersion < 1) {
    throw httpError(400, 'La revisión debe referir una versión del pronóstico congelado');
  }
  if (!contenido.inputs || typeof contenido.inputs !== 'object' || Array.isArray(contenido.inputs)) {
    throw httpError(400, 'La revisión debe incluir las decisiones por producto');
  }
  if (!Array.isArray(contenido.rows)) {
    throw httpError(400, 'La revisión debe incluir las filas propuestas');
  }
  if (state === 'approved') {
    const pending = contenido.rows.filter((row) => row?.decision === 'pending').length;
    if (pending > 0) throw httpError(400, 'No se puede aprobar una revisión con propuestas pendientes');
  }
  for (const row of contenido.rows) {
    if (!row || typeof row !== 'object') throw httpError(400, 'Hay una fila de revisión inválida');
    assertFiniteNumber(row.baseForecast ?? 0, `Pronóstico de ${row.producto || 'producto'}`);
    assertFiniteNumber(row.proposed ?? 0, `Propuesta de ${row.producto || 'producto'}`);
    if (Number(row.proposed) < 0 || Number(row.baseForecast) < 0) {
      throw httpError(400, `La revisión de ${row.producto || 'un producto'} contiene cantidades negativas`);
    }
  }
  return { state, frozenVersion };
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
    if (!ALLOWED_SNAPSHOT_TYPES.has(type)) {
      return res.status(400).json({ error: `Tipo de respaldo no permitido: ${type}` });
    }
    if (!SNAPSHOT_WRITE_ROLES[type].includes(req.user.rol)) {
      return res.status(403).json({ error: 'No tienes permiso para guardar este respaldo' });
    }
    const periodo = String(req.body?.periodo || 'global').slice(0, 20);
    const contenido = req.body?.contenido;
    if (!contenido || typeof contenido !== 'object' || Array.isArray(contenido)) {
      return res.status(400).json({ error: 'El respaldo debe incluir un objeto contenido' });
    }

    if (type === 'workspace') validateWorkspaceContent(contenido);
    if (type === 'forecast-frozen') validateFrozenContent(contenido, periodo);
    const reviewMeta = type === 'monthly-review'
      ? validateMonthlyReviewContent(contenido, periodo, req.user)
      : null;

    const serialized = JSON.stringify(contenido);
    if (Buffer.byteLength(serialized, 'utf8') > 3.5 * 1024 * 1024) {
      return res.status(413).json({ error: 'El respaldo excede el límite de 3.5 MB' });
    }
    const archivos = req.body?.archivos && typeof req.body.archivos === 'object' ? req.body.archivos : {};
    const result = await transaction(async (connection) => {
      const [previous] = await connection.execute(
        `SELECT id, version FROM respaldos_datos
         WHERE tipo = ? AND periodo = ? ORDER BY version DESC LIMIT 1 FOR UPDATE`,
        [type, periodo]
      );
      if (type === 'forecast-frozen' && previous.length && req.user.rol !== 'admin') {
        throw httpError(403, 'El pronóstico congelado de este mes ya existe. Solo un administrador puede emitir otra versión.');
      }
      if (reviewMeta) {
        const [frozen] = await connection.execute(
          `SELECT version FROM respaldos_datos
           WHERE tipo = 'forecast-frozen' AND periodo = ? AND version = ? LIMIT 1`,
          [periodo, reviewMeta.frozenVersion]
        );
        if (!frozen.length) {
          throw httpError(400, 'La revisión debe apuntar a un pronóstico congelado existente de ese mes');
        }
      }
      const version = Number(previous[0]?.version || 0) + 1;
      const [created] = await connection.execute(
        `INSERT INTO respaldos_datos (tipo, version, periodo, archivos, contenido, usuario_id)
         VALUES (?, ?, ?, ?, ?, ?)`,
        [type, version, periodo, JSON.stringify(archivos), serialized, req.user.id]
      );
      await writeAudit(connection, req.user.id, 'respaldar', 'datos', `${type}:${periodo}`, {
        version,
        bytes: Buffer.byteLength(serialized, 'utf8'),
        state: contenido.state || null,
      });
      return { id: created.insertId, version };
    });
    res.status(201).json({ ok: true, ...result, tipo: type, periodo });
  } catch (error) {
    next(error);
  }
});

module.exports = router;
