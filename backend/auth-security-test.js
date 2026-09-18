process.env.DB_HOST = process.env.DB_HOST || '127.0.0.1';
process.env.DB_USER = process.env.DB_USER || 'test';
process.env.DB_NAME = process.env.DB_NAME || 'test';

const { createMemoryRateLimiter, createPersistLoginRateLimit, recordLoginFailure } = require('./rateLimit');
const { validateCredentials } = require('./routes/auth');

function mockReq(ip = '203.0.113.10') {
  return { headers: {}, ip, socket: { remoteAddress: ip } };
}

function runLimiter(limiter, req) {
  return new Promise((resolve) => {
    limiter(req, {}, (error) => resolve(error || null));
  });
}

async function main() {
  const limiter = createMemoryRateLimiter({
    windowMs: 60_000,
    maxAttempts: 8,
    message: 'Demasiados intentos. Espera unos minutos e inténtalo de nuevo.',
  });
  for (let attempt = 1; attempt <= 8; attempt += 1) {
    const error = await runLimiter(limiter, mockReq());
    if (error) throw new Error(`el intento ${attempt} no debía bloquearse`);
  }
  const blocked = await runLimiter(limiter, mockReq());
  if (!blocked || blocked.status !== 429) throw new Error('el noveno intento debe devolver 429');
  const otherIp = await runLimiter(limiter, mockReq('198.51.100.20'));
  if (otherIp) throw new Error('otra IP no debe heredar el bloqueo');

  const persist = createPersistLoginRateLimit(async () => [{ total: 8 }], { maxAttempts: 8 });
  const persistBlocked = await runLimiter(persist, mockReq());
  if (!persistBlocked || persistBlocked.status !== 429) {
    throw new Error('el tope persistente debe bloquear con 8 fallos');
  }

  const persistOpen = createPersistLoginRateLimit(async () => [{ total: 2 }], { maxAttempts: 8 });
  const persistAllowed = await runLimiter(persistOpen, mockReq());
  if (persistAllowed) throw new Error('2 fallos no deben bloquear el login');

  const inserts = [];
  await recordLoginFailure(async (_sql, params) => inserts.push(params), mockReq(), 'operador');
  if (inserts[0][0] !== '203.0.113.10' || !String(inserts[0][1]).includes('operador')) {
    throw new Error('el fallo de login debe quedar en bitácora por IP');
  }

  const shortPassword = (() => {
    try {
      validateCredentials('admin', '123');
      return null;
    } catch (error) {
      return error;
    }
  })();
  if (!shortPassword || shortPassword.status !== 400 || !String(shortPassword.message).includes('4 caracteres')) {
    throw new Error('una contraseña de 3 caracteres debe rechazarse');
  }

  validateCredentials('admin', '1210');
  validateCredentials('admin', '12345');

  console.log('auth-security-test ok');
}

main().catch((error) => {
  console.error(error);
  process.exit(1);
});
