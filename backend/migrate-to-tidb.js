const fs = require('fs');
const path = require('path');
const mysql = require('mysql2/promise');
require('dotenv').config();

const TABLES = [
  'productos',
  'usuarios',
  'homologaciones_productos',
  'ventas_diarias',
  'stock_fijo',
  'produccion_real',
  'pronostico_diario',
  'sesiones',
  'respaldos_datos',
  'bitacora',
  'importaciones_ventas',
  'bajas_diarias',
  'importaciones_operativas',
];
const COPY_BATCH_SIZE = 500;
const SQL_BATCH_BYTES = 2 * 1024 * 1024;
const SNAPSHOT_LIMIT_BYTES = 3.5 * 1024 * 1024;

function targetConfig() {
  if (process.env.TIDB_HOST) {
    return {
      host: process.env.TIDB_HOST,
      user: process.env.TIDB_USER,
      password: process.env.TIDB_PASSWORD || '',
      database: process.env.TIDB_DATABASE,
      port: Number(process.env.TIDB_PORT || 4000),
      ssl: { minVersion: 'TLSv1.2', rejectUnauthorized: true },
      dateStrings: true,
      multipleStatements: true,
    };
  }
  if (!process.env.TARGET_DATABASE_URL) throw new Error('Falta TARGET_DATABASE_URL o la configuración TIDB_*');
  return { uri: process.env.TARGET_DATABASE_URL, dateStrings: true, multipleStatements: true };
}

function sourceConfig() {
  if (!process.env.SOURCE_DATABASE_URL) throw new Error('Falta SOURCE_DATABASE_URL');
  return { uri: process.env.SOURCE_DATABASE_URL, dateStrings: true };
}

function sqlValue(value) {
  if (value === undefined || value === null) return null;
  if (Buffer.isBuffer(value) || value instanceof Date || typeof value !== 'object') return value;
  return JSON.stringify(value);
}

function splitRowsBySize(rows, columns) {
  const batches = [];
  let batch = [];
  let bytes = 0;
  for (const row of rows) {
    const rowBytes = Buffer.byteLength(JSON.stringify(columns.map((column) => row[column])), 'utf8');
    if (batch.length && bytes + rowBytes > SQL_BATCH_BYTES) {
      batches.push(batch);
      batch = [];
      bytes = 0;
    }
    batch.push(row);
    bytes += rowBytes;
  }
  if (batch.length) batches.push(batch);
  return batches;
}

async function applySchema(target) {
  const schema = fs.readFileSync(path.join(__dirname, 'schema.sql'), 'utf8');
  await target.query(schema);
}

async function assertEmptyTarget(target) {
  const populated = [];
  for (const table of TABLES) {
    const [[row]] = await target.query(`SELECT COUNT(*) AS total FROM \`${table}\``);
    if (Number(row.total) > 0) populated.push(`${table}:${row.total}`);
  }
  if (populated.length) {
    throw new Error(`La base destino no está vacía (${populated.join(', ')}). No se modificó para evitar mezclar datos.`);
  }
}

async function copyTable(source, target, table) {
  const [columnRows] = await source.query(`SHOW COLUMNS FROM \`${table}\``);
  const columns = columnRows.map((row) => row.Field);
  const escapedColumns = columns.map((column) => `\`${column}\``).join(', ');
  let cursor = 0;
  let copied = 0;
  while (true) {
    const [rows] = await source.query(
      `SELECT ${escapedColumns} FROM \`${table}\` WHERE id > ? ORDER BY id LIMIT ?`,
      [cursor, COPY_BATCH_SIZE]
    );
    if (!rows.length) break;
    for (const batch of splitRowsBySize(rows, columns)) {
      const placeholders = batch.map(() => `(${columns.map(() => '?').join(',')})`).join(',');
      const values = batch.flatMap((row) => columns.map((column) => sqlValue(row[column])));
      await target.query(`INSERT INTO \`${table}\` (${escapedColumns}) VALUES ${placeholders}`, values);
    }
    cursor = Number(rows[rows.length - 1].id);
    copied += rows.length;
  }
  const [[sourceCount]] = await source.query(`SELECT COUNT(*) AS total FROM \`${table}\``);
  const [[targetCount]] = await target.query(`SELECT COUNT(*) AS total FROM \`${table}\``);
  if (Number(sourceCount.total) !== Number(targetCount.total)) {
    throw new Error(`Conteo inconsistente en ${table}: origen ${sourceCount.total}, destino ${targetCount.total}`);
  }
  const [[maxRow]] = await target.query(`SELECT COALESCE(MAX(id), 0) AS maxId FROM \`${table}\``);
  const nextId = Number(maxRow.maxId) + 1;
  await target.query(`ALTER TABLE \`${table}\` AUTO_INCREMENT = ${nextId}`);
  return copied;
}

function normalizeProduct(value) {
  const normalized = String(value || '')
    .trim()
    .toUpperCase()
    .normalize('NFD')
    .replace(/[\u0300-\u036f]/g, '')
    .replace(/[.,/\\_-]+/g, ' ')
    .replace(/\s+/g, ' ')
    .trim()
    .replace(/CHE{1,2}S{1,2}ECAKE/g, 'CHESSECAKE');
  const compact = normalized.replace(/[^A-Z0-9]/g, '');
  return compact === 'PINAGDE' || compact === 'PINAGRANDE' ? 'PINA GDE' : normalized;
}

function dateOnly(value) {
  if (!value) return '';
  if (typeof value === 'string') {
    const match = value.match(/^\d{4}-\d{2}-\d{2}/);
    if (match) return match[0];
  }
  const date = new Date(value);
  return Number.isNaN(date.getTime()) ? '' : date.toISOString().slice(0, 10);
}

function resolvedProduct(row, aliases) {
  const original = normalizeProduct(row.producto || row.producto_codigo);
  return normalizeProduct(aliases[original] || original);
}

function consolidateWorkspaceRows(rows, type, aliases) {
  const map = new Map();
  for (const row of rows || []) {
    const fecha = dateOnly(row.fecha || row.fechaKey);
    const producto = resolvedProduct(row, aliases);
    const cantidad = Number(row.cantidad || 0);
    if (row.monthlyTotal || !fecha || !producto || !Number.isFinite(cantidad)) continue;
    const dimensions = type === 'ventas'
      ? [String(row.sucursal || row.canal || '').trim(), String(row.cliente || '').trim()]
      : type === 'produccion'
        ? [String(row.turno || '').trim()]
        : [String(row.sucursal || row.canal || '').trim(), String(row.motivo || '').trim()];
    const key = [fecha, producto, ...dimensions.map((value) => value.toUpperCase())].join('|');
    const current = map.get(key);
    const importeValue = Number(row.importe);
    const importe = Number.isFinite(importeValue) ? importeValue : null;
    if (!current) {
      map.set(key, { fecha, producto, cantidad, importe, dimensions });
      continue;
    }
    current.cantidad += cantidad;
    if (type === 'ventas' && importe !== null) current.importe = Number(current.importe || 0) + importe;
  }
  return [...map.values()];
}

async function productIdsForRows(connection, rows) {
  const ids = new Map();
  for (const producto of [...new Set(rows.map((row) => row.producto))]) {
    const [result] = await connection.execute(
      `INSERT INTO productos (codigo, nombre) VALUES (?, ?)
       ON DUPLICATE KEY UPDATE nombre = VALUES(nombre), id = LAST_INSERT_ID(id)`,
      [producto, producto]
    );
    ids.set(producto, Number(result.insertId));
  }
  return ids;
}

async function upsertWorkspaceRows(connection, rows, type) {
  if (!rows.length) return 0;
  const productIds = await productIdsForRows(connection, rows);
  const config = type === 'ventas'
    ? {
        table: 'ventas_diarias',
        columns: ['fecha', 'producto_id', 'cantidad', 'importe', 'canal', 'cliente'],
        values: (row) => [row.fecha, productIds.get(row.producto), row.cantidad, row.importe, ...row.dimensions],
        update: 'cantidad = VALUES(cantidad), importe = VALUES(importe), updated_at = CURRENT_TIMESTAMP',
      }
    : type === 'produccion'
      ? {
          table: 'produccion_real',
          columns: ['fecha', 'producto_id', 'cantidad', 'turno'],
          values: (row) => [row.fecha, productIds.get(row.producto), row.cantidad, ...row.dimensions],
          update: 'cantidad = VALUES(cantidad), updated_at = CURRENT_TIMESTAMP',
        }
      : {
          table: 'bajas_diarias',
          columns: ['fecha', 'producto_id', 'cantidad', 'sucursal', 'motivo'],
          values: (row) => [row.fecha, productIds.get(row.producto), row.cantidad, ...row.dimensions],
          update: 'cantidad = VALUES(cantidad), updated_at = CURRENT_TIMESTAMP',
        };
  for (let offset = 0; offset < rows.length; offset += COPY_BATCH_SIZE) {
    const batch = rows.slice(offset, offset + COPY_BATCH_SIZE);
    const placeholders = batch.map(() => `(${config.columns.map(() => '?').join(',')})`).join(',');
    await connection.query(
      `INSERT INTO ${config.table} (${config.columns.join(',')}) VALUES ${placeholders}
       ON DUPLICATE KEY UPDATE ${config.update}`,
      batch.flatMap(config.values)
    );
  }
  return rows.length;
}

async function createServerlessWorkspace(target) {
  const [rows] = await target.query(
    `SELECT id, version, archivos, contenido, usuario_id
     FROM respaldos_datos WHERE tipo = 'workspace' AND periodo = 'global'
     ORDER BY version DESC LIMIT 1`
  );
  if (!rows.length) return { created: false, reason: 'sin workspace' };
  const source = rows[0];
  const content = typeof source.contenido === 'string' ? JSON.parse(source.contenido) : source.contenido;
  const aliases = content.productAliases && typeof content.productAliases === 'object' ? content.productAliases : {};
  const sales = consolidateWorkspaceRows(content.ventas, 'ventas', aliases);
  const production = consolidateWorkspaceRows(content.realProduction, 'produccion', aliases);
  const waste = consolidateWorkspaceRows(content.bajas, 'bajas', aliases);
  await target.beginTransaction();
  try {
    await upsertWorkspaceRows(target, sales, 'ventas');
    await upsertWorkspaceRows(target, production, 'produccion');
    await upsertWorkspaceRows(target, waste, 'bajas');
    const slimContent = {
      ...content,
      schemaVersion: 2,
      migratedAt: new Date().toISOString(),
      dataStorage: 'database',
      ventas: (content.ventas || []).filter((row) => row.monthlyTotal || !dateOnly(row.fecha)),
      realProduction: (content.realProduction || []).filter((row) => row.monthlyTotal || !dateOnly(row.fecha || row.fechaKey)),
      bajas: (content.bajas || []).filter((row) => row.monthlyTotal || !dateOnly(row.fecha)),
    };
    const serialized = JSON.stringify(slimContent);
    const bytes = Buffer.byteLength(serialized, 'utf8');
    if (bytes > SNAPSHOT_LIMIT_BYTES) {
      throw new Error(`El workspace reducido todavía pesa ${bytes} bytes; requiere una migración adicional antes de Vercel`);
    }
    const version = Number(source.version) + 1;
    await target.execute(
      `INSERT INTO respaldos_datos (tipo, version, periodo, archivos, contenido, usuario_id)
       VALUES ('workspace', ?, 'global', ?, ?, ?)`,
      [version, typeof source.archivos === 'string' ? source.archivos : JSON.stringify(source.archivos || {}), serialized, source.usuario_id]
    );
    await target.commit();
    return { created: true, version, bytes, sales: sales.length, production: production.length, waste: waste.length };
  } catch (error) {
    await target.rollback();
    throw error;
  }
}

async function main() {
  const source = await mysql.createConnection(sourceConfig());
  const target = await mysql.createConnection(targetConfig());
  try {
    await source.query('SELECT 1');
    await target.query('SELECT 1');
    await applySchema(target);
    await assertEmptyTarget(target);
    const copied = {};
    for (const table of TABLES) {
      copied[table] = await copyTable(source, target, table);
    }
    const workspace = await createServerlessWorkspace(target);
    console.log(JSON.stringify({ ok: true, copied, workspace }, null, 2));
  } finally {
    await source.end();
    await target.end();
  }
}

main().catch((error) => {
  console.error(error.message);
  process.exit(1);
});
