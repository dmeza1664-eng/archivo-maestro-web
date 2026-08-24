const crypto = require('crypto');
const { promisify } = require('util');
const { query } = require('./db');

const scrypt = promisify(crypto.scrypt);
const SESSION_DAYS = Number(process.env.SESSION_DAYS || 7);
const SESSION_COOKIE = 'am_session';

function tokenHash(token) {
  return crypto.createHash('sha256').update(token).digest('hex');
}

function isSecureRequest(req) {
  const proto = String(req.headers['x-forwarded-proto'] || req.protocol || '')
    .split(',')[0]
    .trim()
    .toLowerCase();
  return proto === 'https';
}

function serializeSessionCookie(token, secure) {
  const parts = [
    `${SESSION_COOKIE}=${encodeURIComponent(token)}`,
    'Path=/',
    'HttpOnly',
    'SameSite=Lax',
    `Max-Age=${SESSION_DAYS * 24 * 60 * 60}`,
  ];
  if (secure) parts.push('Secure');
  return parts.join('; ');
}

function clearSessionCookie(secure) {
  const parts = [`${SESSION_COOKIE}=`, 'Path=/', 'HttpOnly', 'SameSite=Lax', 'Max-Age=0'];
  if (secure) parts.push('Secure');
  return parts.join('; ');
}

function tokenFromCookie(req) {
  const header = String(req.headers.cookie || '');
  const match = header
    .split(';')
    .map((part) => part.trim())
    .find((part) => part.startsWith(`${SESSION_COOKIE}=`));
  return match ? decodeURIComponent(match.slice(SESSION_COOKIE.length + 1)) : '';
}

function attachSessionCookie(req, res, token) {
  res.setHeader('Set-Cookie', serializeSessionCookie(token, isSecureRequest(req)));
}

function expireSessionCookie(req, res) {
  res.setHeader('Set-Cookie', clearSessionCookie(isSecureRequest(req)));
}

async function hashPassword(password, salt = crypto.randomBytes(16).toString('hex')) {
  const derived = await scrypt(String(password), salt, 64);
  return { hash: derived.toString('hex'), salt };
}

async function verifyPassword(password, salt, expectedHash) {
  const { hash } = await hashPassword(password, salt);
  const actual = Buffer.from(hash, 'hex');
  const expected = Buffer.from(expectedHash, 'hex');
  return actual.length === expected.length && crypto.timingSafeEqual(actual, expected);
}

function publicUser(user) {
  return {
    id: user.id,
    usuario: user.usuario,
    nombre: user.nombre,
    rol: user.rol,
  };
}

async function createSession(connection, usuarioId) {
  const token = crypto.randomBytes(32).toString('base64url');
  const expiration = new Date(Date.now() + SESSION_DAYS * 24 * 60 * 60 * 1000);
  await connection.execute(
    'INSERT INTO sesiones (usuario_id, token_hash, expira_at) VALUES (?, ?, ?)',
    [usuarioId, tokenHash(token), expiration]
  );
  return token;
}

function bearerToken(req) {
  const value = String(req.headers.authorization || '');
  return value.startsWith('Bearer ') ? value.slice(7).trim() : '';
}

async function requireAuth(req, res, next) {
  try {
    const token = bearerToken(req) || tokenFromCookie(req);
    if (!token) return res.status(401).json({ error: 'Inicia sesión para continuar' });
    const rows = await query(
      `SELECT u.id, u.usuario, u.nombre, u.rol
       FROM sesiones s
       INNER JOIN usuarios u ON u.id = s.usuario_id
       WHERE s.token_hash = ? AND s.expira_at > CURRENT_TIMESTAMP AND u.activo = 1
       LIMIT 1`,
      [tokenHash(token)]
    );
    if (!rows.length) return res.status(401).json({ error: 'La sesión venció o no es válida' });
    req.authToken = token;
    req.user = rows[0];
    next();
  } catch (error) {
    next(error);
  }
}

function requireRole(...roles) {
  return (req, res, next) => {
    if (!req.user || !roles.includes(req.user.rol)) {
      return res.status(403).json({ error: 'No tienes permiso para realizar esta acción' });
    }
    next();
  };
}

async function writeAudit(connection, usuarioId, accion, entidad, claveEntidad = '', detalle = null) {
  await connection.execute(
    `INSERT INTO bitacora (usuario_id, accion, entidad, clave_entidad, detalle)
     VALUES (?, ?, ?, ?, ?)`,
    [usuarioId || null, accion, entidad, claveEntidad, detalle ? JSON.stringify(detalle) : null]
  );
}

module.exports = {
  attachSessionCookie,
  bearerToken,
  createSession,
  expireSessionCookie,
  hashPassword,
  publicUser,
  requireAuth,
  requireRole,
  tokenHash,
  verifyPassword,
  writeAudit,
};
