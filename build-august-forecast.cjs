const fs = require('fs');
const path = require('path');
const Module = require('module');
const esbuild = require('esbuild');
const XLSX = require('xlsx');

const ROOT = __dirname;
const DOWNLOADS = path.resolve(ROOT, '..');
const JULY_SALES_PATH = 'C:/Users/X13/Documents/ventas por mes/ventas julio.xlsx';
const JUNE_SALES_PATH = 'C:/Users/X13/Documents/ventas por mes/ventas junio.xlsx';
const OUTPUT_PATH = 'C:/Users/X13/Documents/PRONOSTICO_AGOSTO_2026_BASE_Y_ESCENARIO_12.xlsx';
const TARGET_MONTH = '2026-08';

function normalized(value) {
  return String(value || '')
    .toUpperCase()
    .normalize('NFD')
    .replace(/[\u0300-\u036f]/g, '')
    .replace(/\s+/g, ' ')
    .trim();
}

function findDownload(search) {
  const target = normalized(search);
  const match = fs.readdirSync(DOWNLOADS).find((name) => normalized(name) === target);
  if (!match) throw new Error(`No se encontró ${search}`);
  return path.join(DOWNLOADS, match);
}

async function loadAppFunctions() {
  const built = await esbuild.build({
    entryPoints: [path.join(ROOT, 'App.jsx')],
    bundle: true,
    platform: 'node',
    format: 'cjs',
    write: false,
    loader: { '.css': 'text' },
    define: { 'import.meta.env.VITE_API_URL': JSON.stringify('') },
    logLevel: 'silent',
  });
  const appModule = new Module('august-forecast');
  appModule.filename = path.join(ROOT, 'august-forecast.bundle.cjs');
  appModule.paths = module.paths;
  appModule._compile(built.outputFiles[0].text, appModule.filename);
  return appModule.exports;
}

function appendSheet(workbook, rows, name, widths) {
  const sheet = XLSX.utils.json_to_sheet(rows);
  sheet['!cols'] = widths.map((wch) => ({ wch }));
  sheet['!autofilter'] = { ref: sheet['!ref'] };
  XLSX.utils.book_append_sheet(workbook, sheet, name);
}

async function main() {
  const app = await loadAppFunctions();
  const stockRows = app.parseStock(XLSX.readFile(findDownload('STOCK IDEAL SUCURSALES.xlsx'), { cellDates: true }));
  const historicalSourceNames = [
    'Venta de diciembre 2024.xlsx',
    'VENTA AÑO 2025 - ANGEL.xlsx',
    'VENTA ENERO 2026.xlsx',
    'VENTAS FEEEBRERO 2026.xlsx',
    'MARZO VENTAS.xlsx',
    'venta abril 2026.xlsx',
    'VENTAS DE MAYO Y JUNIO - ANGEL.xlsx',
  ];
  const sourceSales = historicalSourceNames.flatMap((sourceName) => {
    const filePath = findDownload(sourceName);
    return app.parseSalesOrReturns(
      XLSX.readFile(filePath, { cellDates: true }),
      'ventas',
      path.basename(filePath)
    );
  });
  const throughMay = app.filterVentasBeforeMonth(sourceSales, '2026-06');
  const juneSales = app.parseSalesOrReturns(
    XLSX.readFile(JUNE_SALES_PATH, { cellDates: true }),
    'ventas',
    path.basename(JUNE_SALES_PATH)
  );
  const julySales = app.parseSalesOrReturns(
    XLSX.readFile(JULY_SALES_PATH, { cellDates: true }),
    'ventas',
    path.basename(JULY_SALES_PATH)
  );
  const historicalSales = [...throughMay, ...juneSales, ...julySales];
  const forecastRows = app.calculateForecast({
    stockRows,
    historicalVentas: app.filterVentasBeforeMonth(historicalSales, TARGET_MONTH),
    bajas: [],
    existencias: [],
    realProduction: [],
    selectedMonth: TARGET_MONTH,
    dailyBufferPct: 0,
  });
  const scenarioRows = app.buildOperationalForecastScenario(forecastRows);
  const baseTotal = scenarioRows.reduce((sum, row) => sum + row.pronosticoBase, 0);
  const marginTotal = scenarioRows.reduce((sum, row) => sum + row.margenOperativoPiezas, 0);
  const operationalTotal = scenarioRows.reduce((sum, row) => sum + row.pronosticoOperativo, 0);
  const generatedAt = new Date().toISOString();
  const summary = [
    { Indicador: 'Mes pronosticado', Valor: TARGET_MONTH },
    { Indicador: 'Fecha de generación', Valor: new Date(generatedAt).toLocaleString('es-MX') },
    { Indicador: 'Productos', Valor: scenarioRows.length },
    { Indicador: 'Pronóstico estadístico', Valor: Number(baseTotal.toFixed(2)) },
    { Indicador: 'Margen operativo 12%', Valor: Number(marginTotal.toFixed(2)) },
    { Indicador: 'Escenario operativo +12%', Valor: Number(operationalTotal.toFixed(2)) },
    { Indicador: 'Último mes histórico', Valor: '2026-07' },
    { Indicador: 'Control', Valor: 'El escenario +12% no modifica el pronóstico estadístico.' },
  ];
  const products = scenarioRows.map((row) => ({
    Producto: row.producto,
    Categoría: row.categoria,
    'Pronóstico estadístico': Number(row.pronosticoBase.toFixed(2)),
    'Margen 12% piezas': Number(row.margenOperativoPiezas.toFixed(2)),
    'Escenario operativo +12%': Number(row.pronosticoOperativo.toFixed(2)),
    'Método seleccionado': row.metodoPronostico,
    'Meses usados': row.mesesUsados,
  }));
  const categories = [...scenarioRows.reduce((map, row) => {
    const current = map.get(row.categoria) || { products: 0, base: 0, margin: 0, operational: 0 };
    current.products += 1;
    current.base += row.pronosticoBase;
    current.margin += row.margenOperativoPiezas;
    current.operational += row.pronosticoOperativo;
    map.set(row.categoria, current);
    return map;
  }, new Map()).entries()].map(([category, values]) => ({
    Categoría: category,
    Productos: values.products,
    'Pronóstico estadístico': Number(values.base.toFixed(2)),
    'Margen 12% piezas': Number(values.margin.toFixed(2)),
    'Escenario operativo +12%': Number(values.operational.toFixed(2)),
  }));
  const methodology = [
    { Concepto: 'Corte histórico', Detalle: 'Incluye ventas hasta julio de 2026 y excluye cualquier venta de agosto.' },
    { Concepto: 'Junio y julio', Detalle: 'Usa los resúmenes mensuales finales como meses completos.' },
    { Concepto: 'Pronóstico estadístico', Detalle: 'Modelo validado por categoría sin margen operativo.' },
    { Concepto: 'Escenario +12%', Detalle: 'Pronóstico estadístico multiplicado por 1.12; debe evaluarse por separado al cierre.' },
    { Concepto: 'Congelamiento', Detalle: 'Este archivo conserva la referencia generada; la versión oficial también debe congelarse desde la página.' },
  ];
  const workbook = XLSX.utils.book_new();
  appendSheet(workbook, summary, 'Resumen', [38, 90]);
  appendSheet(workbook, products, 'Por producto', [38, 24, 24, 20, 28, 30, 34]);
  appendSheet(workbook, categories, 'Por categoría', [24, 12, 24, 20, 28]);
  appendSheet(workbook, methodology, 'Metodología', [28, 105]);
  XLSX.writeFile(workbook, OUTPUT_PATH);

  console.log(JSON.stringify({
    ok: true,
    outputPath: OUTPUT_PATH,
    targetMonth: TARGET_MONTH,
    products: scenarioRows.length,
    historyRows: historicalSales.length,
    juneRows: juneSales.length,
    julyRows: julySales.length,
    baseForecast: Number(baseTotal.toFixed(2)),
    margin12Pct: Number(marginTotal.toFixed(2)),
    operationalScenario: Number(operationalTotal.toFixed(2)),
  }, null, 2));
}

main().catch((error) => {
  console.error(error.message);
  process.exit(1);
});
