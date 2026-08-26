const path = require('path');
const Module = require('module');
const esbuild = require('esbuild');
const XLSX = require('xlsx');

async function loadAppFunctions() {
  const built = await esbuild.build({
    entryPoints: [path.join(__dirname, 'App.jsx')],
    bundle: true,
    platform: 'node',
    format: 'cjs',
    write: false,
    loader: { '.css': 'text' },
    define: { 'import.meta.env.VITE_API_URL': JSON.stringify('') },
    logLevel: 'silent',
  });
  const appModule = new Module('monthly-close-test');
  appModule.filename = path.join(__dirname, 'monthly-close-test.bundle.cjs');
  appModule.paths = module.paths;
  appModule._compile(built.outputFiles[0].text, appModule.filename);
  return appModule.exports;
}

function assert(condition, message) {
  if (!condition) throw new Error(message);
}

async function main() {
  const app = await loadAppFunctions();
  const close = app.buildMonthlyCloseSummary({
    forecastRows: [
      { producto: 'PASTEL A', pronosticoVenta: 100 },
      { producto: 'PASTEL B', pronosticoVenta: 200 },
    ],
    salesRows: [
      { producto: 'PASTEL A', cantidad: 70 },
      { producto: 'PASTEL A', cantidad: 40 },
      { producto: 'PASTEL B', cantidad: 180 },
      { producto: 'ACCESORIO EXTRA', cantidad: 50 },
    ],
    productionRows: [
      { producto: 'PASTEL A', cantidad: 115 },
      { producto: 'PASTEL B', cantidad: 190 },
      { producto: 'ACCESORIO EXTRA', cantidad: 55 },
    ],
  });
  assert(close.summary.pronostico === 300, 'El pronóstico total es incorrecto');
  assert(close.summary.ventaReal === 290, 'La venta real debe consolidar productos repetidos');
  assert(close.summary.producido === 305, 'La producción total es incorrecta');
  assert(Number(close.summary.wape.toFixed(2)) === 10.34, 'El WAPE es incorrecto');
  assert(close.summary.mae === 15, 'El MAE es incorrecto');
  assert(close.summary.diferenciaProduccion === 15, 'El balance de producción es incorrecto');
  assert(close.unmatchedSalesTotal === 50, 'La venta fuera de catálogo debe separarse');
  assert(close.unmatchedProductionTotal === 55, 'La producción fuera de catálogo debe separarse');

  const salesPath = 'C:/Users/X13/Documents/ventas por mes/ventas julio.xlsx';
  const productionPath = 'C:/Users/X13/Documents/PRODUCIDO/PRODUCIDO JUlIO.xlsx';
  const julySales = app.parseSalesOrReturns(
    XLSX.readFile(salesPath, { cellDates: true }),
    'ventas',
    path.basename(salesPath)
  );
  const julyProduction = app.parseSalesOrReturns(
    XLSX.readFile(productionPath, { cellDates: true }),
    'ventas',
    path.basename(productionPath)
  );
  assert(julySales.length === 923, 'El resumen de ventas de julio cambió de estructura');
  assert(julySales.every((row) => row.monthlyTotal), 'Ventas debe reconocerse como resumen mensual');
  assert(!julySales.some((row) => String(row.producto).startsWith('TOTAL')), 'Total Productos no debe importarse');
  assert(julyProduction.length > 100, 'No se reconoció el resumen mensual de producción');
  assert(julyProduction.every((row) => row.monthlyTotal), 'Producción debe reconocerse como resumen mensual');

  console.log(JSON.stringify({
    ok: true,
    julySalesProducts: julySales.length,
    julyProductionProducts: julyProduction.length,
    syntheticWape: Number(close.summary.wape.toFixed(2)),
  }));
}

main().catch((error) => {
  console.error(error.message);
  process.exit(1);
});
