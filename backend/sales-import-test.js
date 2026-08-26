const crypto = require('crypto');
const { promisify } = require('util');
const mysql = require('mysql2/promise');

const scrypt = promisify(crypto.scrypt);
const apiUrl = String(process.argv[2] || '').replace(/\/$/, '');
const databaseUrl = process.env.MYSQL_PUBLIC_URL || process.env.DATABASE_PUBLIC_URL;
const suffix = Date.now();
const username = `sales_test_${suffix}`;
const password = `Sales-Test-${suffix}-Secure`;
const productA = `TEST PRODUCT A ${suffix}`;
const productB = `TEST PRODUCT B ${suffix}`;

async function request(path, options = {}) {
  const response = await fetch(`${apiUrl}${path}`, {
    method: options.method || 'GET',
    headers: {
      ...(options.body ? { 'Content-Type': 'application/json' } : {}),
      ...(options.token ? { Authorization: `Bearer ${options.token}` } : {}),
    },
    ...(options.body ? { body: JSON.stringify(options.body) } : {}),
  });
  const payload = await response.json().catch(() => ({}));
  if (!response.ok) throw new Error(`${path}: ${response.status} ${payload.error || ''}`);
  return payload;
}

async function createTestUser(connection) {
  const salt = crypto.randomBytes(16).toString('hex');
  const hash = (await scrypt(password, salt, 64)).toString('hex');
  const [result] = await connection.execute(
    `INSERT INTO usuarios (usuario, nombre, password_hash, password_salt, rol)
     VALUES (?, 'Prueba importación', ?, ?, 'operador')`,
    [username, hash, salt]
  );
  return result.insertId;
}

async function cleanup(connection, userId) {
  const [products] = await connection.execute(
    'SELECT id FROM productos WHERE codigo IN (?, ?)',
    [productA, productB]
  );
  const productIds = products.map((row) => row.id);
  if (productIds.length) {
    await connection.execute(
      `DELETE FROM bajas_diarias WHERE producto_id IN (${productIds.map(() => '?').join(',')})`,
      productIds
    );
    await connection.execute(
      `DELETE FROM produccion_real WHERE producto_id IN (${productIds.map(() => '?').join(',')})`,
      productIds
    );
    await connection.execute(
      `DELETE FROM ventas_diarias WHERE producto_id IN (${productIds.map(() => '?').join(',')})`,
      productIds
    );
  }
  if (userId) {
    await connection.execute('DELETE FROM importaciones_operativas WHERE usuario_id = ?', [userId]);
    await connection.execute('DELETE FROM importaciones_ventas WHERE usuario_id = ?', [userId]);
    await connection.execute('DELETE FROM bitacora WHERE usuario_id = ?', [userId]);
    await connection.execute('DELETE FROM sesiones WHERE usuario_id = ?', [userId]);
    await connection.execute('DELETE FROM usuarios WHERE id = ?', [userId]);
  }
  if (productIds.length) {
    await connection.execute(
      `DELETE FROM productos WHERE id IN (${productIds.map(() => '?').join(',')})`,
      productIds
    );
  }
}

async function main() {
  if (!apiUrl || !databaseUrl) throw new Error('Falta API_URL o MYSQL_PUBLIC_URL');
  const connection = await mysql.createConnection(databaseUrl);
  let userId;
  try {
    userId = await createTestUser(connection);
    const login = await request('/api/auth/login', {
      method: 'POST',
      body: { usuario: username, password },
    });
    const rows = [
      { fecha: '2026-07-21', producto: productA, cantidad: 10, sucursal: 'CENTRO' },
      { fecha: '2026-07-21', producto: productA, cantidad: 5, sucursal: 'CENTRO' },
      { fecha: '2026-07-21', producto: productB, cantidad: 7, sucursal: 'NORTE' },
      { fecha: '2026-07-01', producto: productB, cantidad: 100, monthlyTotal: true },
    ];
    const first = await request('/api/ventas/importar', {
      method: 'POST',
      token: login.token,
      body: { archivo: 'prueba.xlsx', ventas: rows },
    });
    const second = await request('/api/ventas/importar', {
      method: 'POST',
      token: login.token,
      body: { archivo: 'prueba.xlsx', ventas: rows },
    });
    const sales = await request('/api/ventas?mes=2026-07', { token: login.token });
    const imported = sales.rows.filter((row) => [productA, productB].includes(row.producto_codigo));
    const quantityA = imported.find((row) => row.producto_codigo === productA)?.cantidad;
    const salesValid =
      first.inserted === 2 && first.updated === 0 && first.rejected === 1 && first.duplicatesInFile === 1 &&
      second.inserted === 0 && second.updated === 2 && imported.length === 2 && quantityA === 15;
    if (!salesValid) throw new Error(JSON.stringify({ first, second, imported }));

    const productionRows = [
      { fecha: '2026-07-21', producto: productA, cantidad: 10, turno: 'MATUTINO' },
      { fecha: '2026-07-21', producto: productA, cantidad: 5, turno: 'MATUTINO' },
      { fecha: '2026-07-01', producto: productA, cantidad: 100, monthlyTotal: true },
    ];
    const productionFirst = await request('/api/produccion-real/importar', {
      method: 'POST',
      token: login.token,
      body: { archivo: 'produccion.xlsx', produccion: productionRows },
    });
    const productionSecond = await request('/api/produccion-real/importar', {
      method: 'POST',
      token: login.token,
      body: { archivo: 'produccion.xlsx', produccion: productionRows },
    });
    const productionSync = await request('/api/produccion-real/sync', { token: login.token });
    const syncedProduction = productionSync.rows.filter((row) => row.producto_codigo === productA);
    const productionValid =
      productionFirst.inserted === 1 && productionFirst.rejected === 1 && productionFirst.duplicatesInFile === 1 &&
      productionSecond.inserted === 0 && productionSecond.updated === 1 &&
      syncedProduction.length === 1 && syncedProduction[0].cantidad === 15;
    if (!productionValid) throw new Error(JSON.stringify({ productionFirst, productionSecond, syncedProduction }));

    const wasteRows = [
      { fecha: '2026-07-21', producto: productB, cantidad: 2, sucursal: 'CENTRO', motivo: 'CADUCIDAD' },
      { fecha: '2026-07-21', producto: productB, cantidad: 1, sucursal: 'CENTRO', motivo: 'CADUCIDAD' },
      { fecha: '2026-07-21', producto: productB, cantidad: 1, sucursal: 'NORTE', motivo: 'DAÑO' },
      { fecha: '2026-07-01', producto: productB, cantidad: 50, monthlyTotal: true },
    ];
    const wasteFirst = await request('/api/bajas/importar', {
      method: 'POST',
      token: login.token,
      body: { archivo: 'bajas.xlsx', bajas: wasteRows },
    });
    const wasteSecond = await request('/api/bajas/importar', {
      method: 'POST',
      token: login.token,
      body: { archivo: 'bajas.xlsx', bajas: wasteRows },
    });
    const wasteSync = await request('/api/bajas/sync', { token: login.token });
    const syncedWaste = wasteSync.rows.filter((row) => row.producto_codigo === productB);
    const wasteValid =
      wasteFirst.inserted === 2 && wasteFirst.rejected === 1 && wasteFirst.duplicatesInFile === 1 &&
      wasteSecond.inserted === 0 && wasteSecond.updated === 2 &&
      syncedWaste.length === 2 && syncedWaste.reduce((sum, row) => sum + row.cantidad, 0) === 4;
    if (!wasteValid) throw new Error(JSON.stringify({ wasteFirst, wasteSecond, syncedWaste }));

    const salesSync = await request('/api/ventas/sync', { token: login.token });
    if (!salesSync.rows.some((row) => row.producto_codigo === productA)) throw new Error('La sincronización de ventas no devolvió la prueba');
    console.log('Prueba integral: ventas, producción y bajas consolidan, actualizan y sincronizan correctamente');
  } finally {
    await cleanup(connection, userId);
    await connection.end();
  }
}

main().catch((error) => {
  console.error(error.message);
  process.exit(1);
});
