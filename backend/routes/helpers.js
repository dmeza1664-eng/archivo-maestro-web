function httpError(status, message) {
  const error = new Error(message);
  error.status = status;
  return error;
}

function extractRows(body, keys) {
  if (Array.isArray(body)) return body;

  for (const key of keys) {
    if (Array.isArray(body?.[key])) return body[key];
  }

  throw httpError(400, `El cuerpo debe ser un arreglo o contener uno de estos campos: ${keys.join(', ')}`);
}

function firstValue(row, keys) {
  for (const key of keys) {
    if (row[key] !== undefined && row[key] !== null && row[key] !== '') {
      return row[key];
    }
  }

  return undefined;
}

function toNumber(value, fieldName, rowIndex, defaultValue = undefined) {
  if (value === undefined || value === null || value === '') {
    if (defaultValue !== undefined) return defaultValue;
    throw httpError(400, `Fila ${rowIndex + 1}: falta ${fieldName}`);
  }

  const normalized = typeof value === 'string'
    ? value.trim().replace(/\s/g, '').replace(',', '.')
    : value;
  const number = Number(normalized);

  if (!Number.isFinite(number)) {
    throw httpError(400, `Fila ${rowIndex + 1}: ${fieldName} no es numérico`);
  }

  return number;
}

function toDateOnly(value, fieldName, rowIndex) {
  if (value === undefined || value === null || value === '') {
    throw httpError(400, `Fila ${rowIndex + 1}: falta ${fieldName}`);
  }

  if (typeof value === 'string' && /^\d{4}-\d{2}-\d{2}$/.test(value)) {
    return value;
  }

  if (typeof value === 'number') {
    const excelEpoch = Date.UTC(1899, 11, 30);
    const date = new Date(excelEpoch + value * 24 * 60 * 60 * 1000);
    return date.toISOString().slice(0, 10);
  }

  const date = new Date(value);
  if (Number.isNaN(date.getTime())) {
    throw httpError(400, `Fila ${rowIndex + 1}: ${fieldName} no tiene formato de fecha válido`);
  }

  return date.toISOString().slice(0, 10);
}

function validateMes(mes) {
  if (!/^\d{4}-(0[1-9]|1[0-2])$/.test(mes || '')) {
    throw httpError(400, 'El parámetro mes es requerido con formato YYYY-MM, por ejemplo 2026-04');
  }
}

function mesFromDate(date) {
  return date.slice(0, 7);
}

function monthRange(mes) {
  validateMes(mes);
  const [year, month] = mes.split('-').map(Number);
  const start = `${mes}-01`;
  const next = month === 12
    ? `${year + 1}-01-01`
    : `${year}-${String(month + 1).padStart(2, '0')}-01`;

  return { start, next };
}

function addMonths(dateString, amount) {
  const [year, month, day] = dateString.split('-').map(Number);
  const date = new Date(Date.UTC(year, month - 1 + amount, day));
  return date.toISOString().slice(0, 10);
}

function daysInMonth(mes) {
  validateMes(mes);
  const [year, month] = mes.split('-').map(Number);
  return new Date(Date.UTC(year, month, 0)).getUTCDate();
}

function paginationFromQuery(query, defaultLimit = 4000, maxLimit = 10000) {
  const cursor = Number(query?.cursor || 0);
  const requestedLimit = Number(query?.limit || defaultLimit);
  if (!Number.isSafeInteger(cursor) || cursor < 0) {
    throw httpError(400, 'El cursor de sincronización no es válido');
  }
  if (!Number.isSafeInteger(requestedLimit) || requestedLimit < 1) {
    throw httpError(400, 'El límite de sincronización no es válido');
  }
  return { cursor, limit: Math.min(requestedLimit, maxLimit) };
}

// TiDB rechaza los marcadores dentro de LIMIT en sentencias preparadas y
// responde "Incorrect arguments to LIMIT". El valor se interpola, y por eso
// aqui solo se acepta un entero acotado: nada que venga del cliente sin validar
// puede llegar a la consulta.
function limitClause(limit, maxLimit = 10001) {
  if (!Number.isSafeInteger(limit) || limit < 1 || limit > maxLimit) {
    throw httpError(400, 'El límite de sincronización no es válido');
  }
  return `LIMIT ${limit}`;
}

const MAX_IMPORT_BATCH_ROWS = 1500;

function parseImportMeta(body, rowCount) {
  if (rowCount > MAX_IMPORT_BATCH_ROWS) {
    throw httpError(
      400,
      `Cada lote admite máximo ${MAX_IMPORT_BATCH_ROWS} filas. Usa la carga por lotes de la aplicación.`
    );
  }
  const importRunId = String(body?.importRunId || '').trim().slice(0, 80) || null;
  const batchIndex = Number(body?.batchIndex);
  const batchTotal = Number(body?.batchTotal);
  const hasBatch = Number.isInteger(batchIndex) && Number.isInteger(batchTotal);
  if (hasBatch && (batchIndex < 1 || batchTotal < 1 || batchIndex > batchTotal)) {
    throw httpError(400, 'El lote de importación es inválido');
  }
  return {
    importRunId,
    batchIndex: hasBatch ? batchIndex : 1,
    batchTotal: hasBatch ? batchTotal : 1,
  };
}

function productInput(row, rowIndex) {
  const codigo = firstValue(row, [
    'producto_codigo',
    'codigo_producto',
    'codigo',
    'sku',
    'producto',
  ]);
  const nombre = firstValue(row, [
    'producto_nombre',
    'nombre_producto',
    'nombre',
    'descripcion',
    'producto',
  ]);

  if (!codigo) {
    throw httpError(400, `Fila ${rowIndex + 1}: falta codigo de producto`);
  }

  return {
    codigo: String(codigo).trim(),
    nombre: String(nombre || codigo).trim(),
  };
}

async function ensureProduct(connection, row, rowIndex) {
  const product = productInput(row, rowIndex);

  const [result] = await connection.execute(
    `INSERT INTO productos (codigo, nombre)
     VALUES (?, ?)
     ON DUPLICATE KEY UPDATE
       nombre = VALUES(nombre),
       id = LAST_INSERT_ID(id)`,
    [product.codigo, product.nombre]
  );

  await saveHomologacionIfPresent(connection, row, result.insertId);
  return result.insertId;
}

async function saveHomologacionIfPresent(connection, row, productoId) {
  const codigoOrigen = firstValue(row, [
    'codigo_origen',
    'producto_origen_codigo',
    'codigo_excel',
    'codigo_original',
  ]);

  if (!codigoOrigen) return;

  const nombreOrigen = firstValue(row, [
    'nombre_origen',
    'producto_origen_nombre',
    'nombre_excel',
    'nombre_original',
  ]);

  await connection.execute(
    `INSERT INTO homologaciones_productos (codigo_origen, nombre_origen, producto_id)
     VALUES (?, ?, ?)
     ON DUPLICATE KEY UPDATE
       nombre_origen = VALUES(nombre_origen),
       producto_id = VALUES(producto_id)`,
    [String(codigoOrigen).trim(), nombreOrigen ? String(nombreOrigen).trim() : null, productoId]
  );
}

module.exports = {
  addMonths,
  daysInMonth,
  ensureProduct,
  extractRows,
  firstValue,
  httpError,
  limitClause,
  mesFromDate,
  monthRange,
  paginationFromQuery,
  parseImportMeta,
  MAX_IMPORT_BATCH_ROWS,
  toDateOnly,
  toNumber,
  validateMes,
};
