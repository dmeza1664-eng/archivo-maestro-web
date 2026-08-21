const express = require('express');
const { query, transaction } = require('../db');
const {
  createSession,
  hashPassword,
  publicUser,
  requireAuth,
  requireRole,
  tokenHash,
  verifyPassword,
  writeAudit,
} = require('../auth');

const router = express.Router();

function validateCredentials(usuario, password) {
  if (!/^[A-Za-z0-9._-]{3,80}$/.test(usuario || '')) {
    const error = new Error('El usuario debe tener al menos 3 caracteres y usar letras, números, punto, guion o guion bajo');
    error.status = 400;
    throw error;
  }
  if (String(password || '').length < 10) {
    const error = new Error('La contraseña debe tener al menos 10 caracteres');
    error.status = 400;
    throw error;
  }
}

router.get('/status', async (_req, res, next) => {
  try {
    const rows = await query('SELECT COUNT(*) AS total FROM usuarios WHERE activo = 1');
    res.json({ ok: true, needsSetup: Number(rows[0]?.total || 0) === 0 });
  } catch (error) {
    next(error);
  }
});

router.post('/setup', async (req, res, next) => {
  try {
    const setupKey = String(req.body?.setupKey || '');
    if (!process.env.APP_SETUP_KEY || setupKey !== process.env.APP_SETUP_KEY) {
      return res.status(403).json({ error: 'Clave de instalación incorrecta' });
    }
    const usuario = String(req.body?.usuario || '').trim().toLowerCase();
    const nombre = String(req.body?.nombre || '').trim() || usuario;
    const password = String(req.body?.password || '');
    validateCredentials(usuario, password);
    const result = await transaction(async (connection) => {
      const [existing] = await connection.execute('SELECT id FROM usuarios LIMIT 1');
      if (existing.length) {
        const error = new Error('El administrador inicial ya fue creado');
        error.status = 409;
        throw error;
      }
      const credentials = await hashPassword(password);
      const [created] = await connection.execute(
        `INSERT INTO usuarios (usuario, nombre, password_hash, password_salt, rol)
         VALUES (?, ?, ?, ?, 'admin')`,
        [usuario, nombre, credentials.hash, credentials.salt]
      );
      const user = { id: created.insertId, usuario, nombre, rol: 'admin' };
      const token = await createSession(connection, user.id);
      await writeAudit(connection, user.id, 'crear', 'usuario', String(user.id), { rol: 'admin', inicial: true });
      return { token, user };
    });
    res.status(201).json({ ok: true, token: result.token, user: publicUser(result.user) });
  } catch (error) {
    next(error);
  }
});

router.post('/login', async (req, res, next) => {
  try {
    const usuario = String(req.body?.usuario || '').trim().toLowerCase();
    const password = String(req.body?.password || '');
    const result = await transaction(async (connection) => {
      const [rows] = await connection.execute(
        `SELECT id, usuario, nombre, rol, password_hash, password_salt
         FROM usuarios WHERE usuario = ? AND activo = 1 LIMIT 1`,
        [usuario]
      );
      const user = rows[0];
      if (!user || !(await verifyPassword(password, user.password_salt, user.password_hash))) {
        const error = new Error('Usuario o contraseña incorrectos');
        error.status = 401;
        throw error;
      }
      await connection.execute('UPDATE usuarios SET ultimo_acceso = CURRENT_TIMESTAMP WHERE id = ?', [user.id]);
      const token = await createSession(connection, user.id);
      await writeAudit(connection, user.id, 'iniciar_sesion', 'sesion');
      return { token, user };
    });
    res.json({ ok: true, token: result.token, user: publicUser(result.user) });
  } catch (error) {
    next(error);
  }
});

router.get('/me', requireAuth, (req, res) => {
  res.json({ ok: true, user: publicUser(req.user) });
});

router.post('/logout', requireAuth, async (req, res, next) => {
  try {
    await transaction(async (connection) => {
      await connection.execute('DELETE FROM sesiones WHERE token_hash = ?', [tokenHash(req.authToken)]);
      await writeAudit(connection, req.user.id, 'cerrar_sesion', 'sesion');
    });
    res.json({ ok: true });
  } catch (error) {
    next(error);
  }
});

router.get('/users', requireAuth, requireRole('admin'), async (_req, res, next) => {
  try {
    const rows = await query(
      'SELECT id, usuario, nombre, rol, activo, ultimo_acceso, created_at FROM usuarios ORDER BY nombre, usuario'
    );
    res.json({ ok: true, rows });
  } catch (error) {
    next(error);
  }
});

router.post('/users', requireAuth, requireRole('admin'), async (req, res, next) => {
  try {
    const usuario = String(req.body?.usuario || '').trim().toLowerCase();
    const nombre = String(req.body?.nombre || '').trim() || usuario;
    const password = String(req.body?.password || '');
    const rol = ['admin', 'operador', 'consulta'].includes(req.body?.rol) ? req.body.rol : 'operador';
    validateCredentials(usuario, password);
    const result = await transaction(async (connection) => {
      const credentials = await hashPassword(password);
      const [created] = await connection.execute(
        `INSERT INTO usuarios (usuario, nombre, password_hash, password_salt, rol)
         VALUES (?, ?, ?, ?, ?)`,
        [usuario, nombre, credentials.hash, credentials.salt, rol]
      );
      await writeAudit(connection, req.user.id, 'crear', 'usuario', String(created.insertId), { usuario, rol });
      return created.insertId;
    });
    res.status(201).json({ ok: true, id: result });
  } catch (error) {
    if (error.code === 'ER_DUP_ENTRY') error.status = 409;
    next(error);
  }
});

module.exports = router;
