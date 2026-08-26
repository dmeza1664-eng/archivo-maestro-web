const mysql = require('mysql2/promise');

const apiUrl = String(process.argv[2] || '').replace(/\/$/, '');
const setupKey = String(process.argv[3] || '');
const databaseUrl = process.env.MYSQL_PUBLIC_URL || process.env.DATABASE_PUBLIC_URL;
const username = `integration_${Date.now()}`;
const password = `Integration-${Date.now()}-Secure`;

async function request(path, options = {}) {
  const response = await fetch(`${apiUrl}${path}`, {
    ...options,
    headers: {
      ...(options.body ? { 'Content-Type': 'application/json' } : {}),
      ...(options.token ? { Authorization: `Bearer ${options.token}` } : {}),
    },
  });
  const payload = await response.json().catch(() => ({}));
  if (!response.ok) throw new Error(`${path}: ${response.status} ${payload.error || ''}`);
  return payload;
}

async function cleanup() {
  if (!databaseUrl) throw new Error('Falta MYSQL_PUBLIC_URL para limpiar la prueba');
  const connection = await mysql.createConnection(databaseUrl);
  const [users] = await connection.execute('SELECT id FROM usuarios WHERE usuario = ?', [username]);
  if (users.length) {
    const id = users[0].id;
    await connection.execute('DELETE FROM bitacora WHERE usuario_id = ?', [id]);
    await connection.execute('DELETE FROM respaldos_datos WHERE usuario_id = ?', [id]);
    await connection.execute('DELETE FROM sesiones WHERE usuario_id = ?', [id]);
    await connection.execute('DELETE FROM usuarios WHERE id = ?', [id]);
  }
  await connection.end();
}

async function main() {
  if (!apiUrl || !setupKey) throw new Error('Uso: node integration-test.js API_URL SETUP_KEY');
  try {
    const status = await request('/api/auth/status');
    if (!status.needsSetup) throw new Error('La instalación ya tiene usuarios; no se ejecutó la prueba destructiva');
    const setup = await request('/api/auth/setup', {
      method: 'POST',
      body: JSON.stringify({ nombre: 'Integración temporal', usuario: username, password, setupKey }),
    });
    const saved = await request('/api/snapshots/workspace', {
      method: 'POST',
      token: setup.token,
      body: JSON.stringify({
        periodo: 'global',
        archivos: { ventas: 'prueba.xlsx' },
        contenido: { test: true, rows: [{ producto: 'PINA GDE', cantidad: 10 }] },
      }),
    });
    const restored = await request('/api/snapshots/workspace', { token: setup.token });
    const frozen = await request('/api/snapshots/forecast-frozen', {
      method: 'POST',
      token: setup.token,
      body: JSON.stringify({
        periodo: '2026-08',
        archivos: { ventas: 'ventas-julio.xlsx' },
        contenido: {
          frozenAt: new Date().toISOString(),
          modelVersion: 'categorySeasonal',
          operationalMarginPct: 12,
          rows: [{ producto: 'PINA GDE', pronosticoBase: 100, pronosticoOperativo: 112 }],
        },
      }),
    });
    const frozenRestored = await request('/api/snapshots/forecast-frozen?periodo=2026-08', { token: setup.token });
    const audit = await request('/api/audit?limit=10', { token: setup.token });
    if (
      saved.version !== 1
      || restored.snapshot.contenido?.test !== true
      || frozen.version !== 1
      || frozenRestored.snapshot.contenido?.rows?.[0]?.pronosticoOperativo !== 112
      || audit.rows.length < 3
    ) {
      throw new Error('La respuesta integral no contiene los datos esperados');
    }
    console.log('Prueba integral: autenticación, respaldo, pronóstico congelado y bitácora correctos');
  } finally {
    await cleanup();
  }
}

main().catch((error) => {
  console.error(error.message);
  process.exit(1);
});
