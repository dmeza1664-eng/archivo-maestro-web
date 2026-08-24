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

module.exports = {
  clientKey,
  createMemoryRateLimiter,
};
