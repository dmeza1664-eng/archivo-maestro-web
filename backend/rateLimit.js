function clientKey(req) {
  const forwarded = String(req.headers['x-forwarded-for'] || '')
    .split(',')[0]
    .trim();
  return forwarded || req.ip || req.socket?.remoteAddress || 'unknown';
}

function createMemoryRateLimiter({ windowMs, maxAttempts, message }) {
  const hits = new Map();
  return function rateLimit(req, res, next) {
    const key = clientKey(req);
    const now = Date.now();
    const recent = (hits.get(key) || []).filter((stamp) => now - stamp < windowMs);
    if (recent.length >= maxAttempts) {
      const error = new Error(message);
      error.status = 429;
      return next(error);
    }
    recent.push(now);
    hits.set(key, recent);
    if (hits.size > 2000) {
      for (const [storedKey, stamps] of hits.entries()) {
        if (!stamps.some((stamp) => now - stamp < windowMs)) hits.delete(storedKey);
      }
    }
    next();
  };
}

function createPersistLoginRateLimit(query, {
  windowMinutes = 15,
  maxAttempts = 8,
  message = 'Demasiados intentos. Espera unos minutos e inténtalo de nuevo.',
} = {}) {
  const minutes = Math.max(1, Math.min(120, Number(windowMinutes) || 15));
  const limit = Math.max(1, Number(maxAttempts) || 8);
  return async function persistLoginRateLimit(req, res, next) {
    try {
      const key = clientKey(req).slice(0, 160);
      const rows = await query(
        `SELECT COUNT(*) AS total FROM bitacora
         WHERE accion = 'login_fallido' AND entidad = 'sesion' AND clave_entidad = ?
           AND created_at > DATE_SUB(CURRENT_TIMESTAMP, INTERVAL ${minutes} MINUTE)`,
        [key]
      );
      if (Number(rows[0]?.total || 0) >= limit) {
        const error = new Error(message);
        error.status = 429;
        return next(error);
      }
    } catch (error) {
      if (error.status === 429) return next(error);
      console.error(error);
    }
    next();
  };
}

async function recordLoginFailure(query, req, usuario = '') {
  await query(
    `INSERT INTO bitacora (usuario_id, accion, entidad, clave_entidad, detalle)
     VALUES (NULL, 'login_fallido', 'sesion', ?, ?)`,
    [clientKey(req).slice(0, 160), JSON.stringify({ usuario: String(usuario || '').slice(0, 80) })]
  );
}

module.exports = {
  clientKey,
  createMemoryRateLimiter,
  createPersistLoginRateLimit,
  recordLoginFailure,
};

