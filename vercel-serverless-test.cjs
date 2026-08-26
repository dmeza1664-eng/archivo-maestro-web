const path = require('path');
const Module = require('module');
const esbuild = require('esbuild');
const XLSX = require('xlsx');
const { paginationFromQuery } = require('./backend/routes/helpers');

async function loadAppFunctions() {
  const built = await esbuild.build({
    entryPoints: [path.join(__dirname, 'App.jsx')],
    bundle: true,
    platform: 'node',
    format: 'cjs',
    write: false,
    loader: { '.css': 'text' },
    define: {
      'import.meta.env.VITE_API_URL': 'undefined',
      'import.meta.env.DEV': 'false',
    },
    logLevel: 'silent',
  });
  const appModule = new Module('vercel-serverless-test');
  appModule.filename = path.join(__dirname, 'vercel-serverless-test.bundle.cjs');
  appModule.paths = module.paths;
  appModule._compile(built.outputFiles[0].text, appModule.filename);
  return appModule.exports;
}

function assert(condition, message) {
  if (!condition) throw new Error(message);
}

async function main() {
  const app = await loadAppFunctions();
  const stockWorkbook = XLSX.readFile('C:/Users/X13/Downloads/STOCK IDEAL SUCURSALES.xlsx', { cellDates: true });
  const stock = app.parseStock(stockWorkbook);
  const existencias = app.parseExistencias(stockWorkbook);
  const stockTotal = stock.reduce((sum, row) => sum + row.stock, 0);
  console.log(JSON.stringify({ stockRows: stock.length, stockTotal, existenceRows: existencias.length }));
  assert(stock.length === 113, `El stock ideal debe reconocer 113 productos; reconoció ${stock.length}`);
  assert(stockTotal === 9195, `El stock ideal debe usar la columna STOCK; sumó ${stockTotal}`);
  assert(existencias.length === 113, `Las existencias deben leer 113 productos; reconoció ${existencias.length}`);

  const juneWasteWorkbook = XLSX.readFile('C:/Users/X13/Documents/DEVOLUCION Y BAJAS/JUNIO BAJAS.xlsx', { cellDates: true });
  const juneWaste = app.parseSalesOrReturns(juneWasteWorkbook, 'bajas', 'JUNIO BAJAS.xlsx');
  const juneWasteSummary = app.parseBajasSummaryWorkbook(juneWasteWorkbook);
  assert(juneWaste.length === 3098, `Las bajas de junio deben leer 3098 eventos; reconoció ${juneWaste.length}`);
  assert(juneWasteSummary.reduce((sum, row) => sum + row.cantidad, 0) === 812, 'El resumen de bajas no debe duplicar el subtotal de Erick');
  const julyWasteWorkbook = XLSX.readFile('C:/Users/X13/Documents/DEVOLUCION Y BAJAS/JULIO BAJAS.xlsx', { cellDates: true });
  const julyWaste = app.parseSalesOrReturns(julyWasteWorkbook, 'bajas', 'JULIO BAJAS.xlsx');
  const julyWasteSummary = app.parseBajasSummaryWorkbook(julyWasteWorkbook);
  assert(julyWaste.length === 630, 'Las bajas de julio deben leer los 630 eventos del corte');
  assert(julyWasteSummary.reduce((sum, row) => sum + row.cantidad, 0) === 276, 'El resumen de julio no debe duplicar el subtotal de Erick');

  const sales = app.consolidateSalesRowsForUpload([
    { fecha: '2026-08-01', producto_codigo: 'A', cantidad: 2, importe: 20, sucursal: 'Norte', cliente: '' },
    { fecha: '2026-08-01', producto_codigo: 'A', cantidad: 3, importe: 30, sucursal: 'Norte', cliente: '' },
    { fecha: '2026-08-01', producto_codigo: 'TOTAL', cantidad: 5, importe: null, monthlyTotal: true },
  ]);
  assert(sales.length === 1, 'Las ventas repetidas deben consolidarse antes de dividir lotes');
  assert(sales[0].cantidad === 5 && sales[0].importe === 50, 'La consolidación de ventas es incorrecta');

  const production = app.consolidateOperationalRowsForUpload([
    { fecha: '2026-08-01', producto_codigo: 'A', cantidad: 2, turno: 'Mañana' },
    { fecha: '2026-08-01', producto_codigo: 'A', cantidad: 4, turno: 'Mañana' },
  ], true);
  assert(production.length === 1 && production[0].cantidad === 6, 'La consolidación operativa es incorrecta');

  const page = paginationFromQuery({ cursor: '10', limit: '25000' });
  assert(page.cursor === 10 && page.limit === 10000, 'La paginación debe validar el cursor y limitar el tamaño');
  let invalidCursorRejected = false;
  try {
    paginationFromQuery({ cursor: '-1' });
  } catch (error) {
    invalidCursorRejected = error.status === 400;
  }
  assert(invalidCursorRejected, 'Debe rechazar cursores inválidos');

  process.env.DB_HOST = 'localhost';
  process.env.DB_USER = 'test';
  process.env.DB_NAME = 'test';
  const expressApp = require('./backend/server');
  assert(typeof expressApp === 'function', 'El backend debe exportar la aplicación Express para Vercel');

  console.log(JSON.stringify({ ok: true, stockRows: stock.length, existenceRows: existencias.length, juneWasteRows: juneWaste.length, julyWasteRows: julyWaste.length, salesRows: sales.length, productionRows: production.length, page }));
}

main().catch((error) => {
  console.error(error.message);
  process.exit(1);
});
