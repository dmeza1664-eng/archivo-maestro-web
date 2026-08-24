const { createMemoryRateLimiter } = require('./rateLimit');

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
  console.log('auth-security-test ok');
}

main().catch((error) => {
  console.error(error);
  process.exit(1);
});
