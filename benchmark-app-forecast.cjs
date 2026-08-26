const fs = require('fs');
const path = require('path');
const Module = require('module');
const esbuild = require('esbuild');
const XLSX = require('xlsx');

const ROOT = __dirname;
const DOWNLOADS = path.resolve(ROOT, '..');

function normalized(value) {
  return String(value)
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

function monthKey(value) {
  const date = value instanceof Date ? value : new Date(value);
  if (Number.isNaN(date.getTime())) return '';
  return `${date.getFullYear()}-${String(date.getMonth() + 1).padStart(2, '0')}`;
}

function category(product) {
  const value = normalized(product);
  if (value.includes('GELATINA')) return 'Gelatinas';
  if (value.includes('GALLETA')) return 'Galletas';
  if (/\b(GDE|GRANDE)\b/.test(value)) return 'Pasteles grandes';
  if (/\b(MED|MEDIANO)\b/.test(value) && !value.includes('MINI')) return 'Pasteles medianos';
  if (/\b(CH|CHICO)\b/.test(value)) return 'Pasteles chicos';
  if (value.includes('MINI')) return 'Mini medianos';
  if (value.includes('BOLLO') || value.includes('PAN')) return 'Pan';
  return 'Otros';
}

function metrics(rows) {
  const actual = rows.reduce((sum, row) => sum + row.actual, 0);
  const forecast = rows.reduce((sum, row) => sum + row.forecast, 0);
  const absoluteError = rows.reduce((sum, row) => sum + Math.abs(row.actual - row.forecast), 0);
  return {
    products: rows.length,
    actual: Number(actual.toFixed(2)),
    forecast: Number(forecast.toFixed(2)),
    bias: Number((actual - forecast).toFixed(2)),
    wape: actual > 0 ? Number((absoluteError / actual * 100).toFixed(2)) : null,
    mae: rows.length ? Number((absoluteError / rows.length).toFixed(2)) : null,
    inside15: rows.filter((row) => Math.abs(row.actual - row.forecast) <= 15).length,
  };
}

function productMonthTotals(records, product) {
  return Object.fromEntries(
    [...records.reduce((map, row) => {
      if (normalized(row.producto) !== product) return map;
      const key = monthKey(row.fecha);
      map.set(key, (map.get(key) || 0) + Number(row.cantidad || 0));
      return map;
    }, new Map()).entries()].sort().map(([key, value]) => [key, Number(value.toFixed(2))])
  );
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
  const appModule = new Module('forecast-benchmark');
  appModule.filename = path.join(ROOT, 'forecast-benchmark.cjs');
  appModule.paths = module.paths;
  appModule._compile(built.outputFiles[0].text, appModule.filename);
  return appModule.exports;
}

async function main() {
  const app = await loadAppFunctions();
  const stockPath = findDownload('STOCK IDEAL SUCURSALES.xlsx');
  const stockRows = app.parseStock(XLSX.readFile(stockPath, { cellDates: true }));
  const sourceNames = [
    'Venta de diciembre 2024.xlsx',
    'VENTA AÑO 2025 - ANGEL.xlsx',
    'VENTA ENERO 2026.xlsx',
    'VENTAS FEEEBRERO 2026.xlsx',
    'MARZO VENTAS.xlsx',
    'venta abril 2026.xlsx',
    'VENTAS DE MAYO Y JUNIO - ANGEL.xlsx',
  ];
  const parsedSources = sourceNames.map((sourceName) => {
    const filePath = findDownload(sourceName);
    const fileName = path.basename(filePath);
    const workbook = XLSX.readFile(filePath, { cellDates: true });
    const rows = app.parseSalesOrReturns(workbook, 'ventas', fileName);
    return { fileName, sheetNames: workbook.SheetNames, rows };
  });
  const allSales = parsedSources.flatMap((source) => source.rows);
  const includeProductHistories = process.env.METRICS_ONLY !== '1' && process.env.CATEGORY_ONLY !== '1';
  const results = {};
  for (const targetMonth of ['2026-04', '2026-05', '2026-06']) {
    const historical = app.filterVentasBeforeMonth(allSales, targetMonth);
    const actualMap = new Map();
    for (const row of allSales.filter((sale) => monthKey(sale.fecha) === targetMonth)) {
      const product = normalized(row.producto);
      actualMap.set(product, (actualMap.get(product) || 0) + Number(row.cantidad || 0));
    }
    results[targetMonth] = { historyRows: historical.length };
    for (const modelVersion of ['default', 'operational12', 'defaultStringDates', 'legacy', 'rolling', 'categorySeasonal', 'seasonal25', 'seasonal50', 'seasonal75', 'seasonal100']) {
      const isOperationalScenario = modelVersion === 'operational12';
      const historicalInput = modelVersion === 'defaultStringDates'
        ? historical.map((row) => ({ ...row, fecha: row.fecha instanceof Date ? row.fecha.toISOString().slice(0, 10) : row.fecha }))
        : historical;
      const forecastRows = app.calculateForecast({
        stockRows,
        historicalVentas: historicalInput,
        bajas: [],
        existencias: [],
        realProduction: [],
        selectedMonth: targetMonth,
        dailyBufferPct: 0,
        ...(['default', 'operational12', 'defaultStringDates'].includes(modelVersion) ? {} : { modelVersion }),
      });
      const comparison = forecastRows.map((row) => ({
        product: normalized(row.producto),
        category: category(row.producto),
        actual: actualMap.get(normalized(row.producto)) || 0,
        forecast: Number(row.pronosticoVenta || 0) * (isOperationalScenario ? 1.12 : 1),
        method: row.metodoPronostico,
        ...(includeProductHistories ? { history: productMonthTotals(historical, normalized(row.producto)) } : {}),
      }));
      const categories = {};
      for (const categoryName of [...new Set(comparison.map((row) => row.category))]) {
        categories[categoryName] = metrics(comparison.filter((row) => row.category === categoryName));
      }
      results[targetMonth][modelVersion] = {
        overall: metrics(comparison),
        categories,
        methods: Object.fromEntries(
          [...comparison.reduce((map, row) => map.set(row.method, (map.get(row.method) || 0) + 1), new Map()).entries()]
            .sort((a, b) => b[1] - a[1])
        ),
        topErrors: comparison
          .map((row) => ({ ...row, absoluteError: Math.abs(row.actual - row.forecast) }))
          .sort((a, b) => b.absoluteError - a.absoluteError)
          .slice(0, 10),
      };
    }
    results[targetMonth].wapeDeltas = Object.fromEntries(
      Object.entries(results[targetMonth])
        .filter(([, value]) => value?.overall)
        .map(([version, value]) => [version, Number((value.overall.wape - results[targetMonth].legacy.overall.wape).toFixed(2))])
    );
  }
  const sourceSummary = parsedSources.map((source) => ({
      file: source.fileName,
      sheets: source.sheetNames,
      rows: source.rows.length,
      quantity: Number(source.rows.reduce((sum, row) => sum + Number(row.cantidad || 0), 0).toFixed(2)),
      dateCoverage: (() => {
        const dates = [...new Set(source.rows.map((row) => row.fecha instanceof Date ? row.fecha.toISOString().slice(0, 10) : ''))]
          .filter(Boolean).sort();
        return { count: dates.length, first: dates[0] || null, last: dates.at(-1) || null };
      })(),
      months: Object.fromEntries(
        [...source.rows.reduce((map, row) => {
          const key = monthKey(row.fecha) || 'sin-fecha';
          map.set(key, (map.get(key) || 0) + 1);
          return map;
        }, new Map()).entries()].sort()
      ),
      monthQuantities: Object.fromEntries(
        [...source.rows.reduce((map, row) => {
          const key = monthKey(row.fecha) || 'sin-fecha';
          map.set(key, (map.get(key) || 0) + Number(row.cantidad || 0));
          return map;
        }, new Map()).entries()].sort().map(([key, value]) => [key, Number(value.toFixed(2))])
      ),
    }));
  if (process.env.SOURCE_ONLY === '1') {
    const operationalProducts = new Set(
      stockRows
        .map((row) => normalized(row.producto))
        .filter((product) => !/\b(PROMO|PROMOCION|PROMOCIONAL)\b/.test(product))
    );
    const operationalMonthTotals = Object.fromEntries(
      [...allSales.reduce((map, row) => {
        const product = normalized(row.producto);
        if (!operationalProducts.has(product)) return map;
        const key = monthKey(row.fecha) || 'sin-fecha';
        map.set(key, (map.get(key) || 0) + Number(row.cantidad || 0));
        return map;
      }, new Map()).entries()].sort().map(([key, value]) => [key, Number(value.toFixed(2))])
    );
    console.log(JSON.stringify({ sources: sourceSummary, operationalMonthTotals }, null, 2));
    return;
  }
  if (process.env.METRICS_ONLY === '1') {
    console.log(JSON.stringify(Object.fromEntries(
      Object.entries(results).map(([month, monthResult]) => [month, Object.fromEntries(
        Object.entries(monthResult).map(([version, value]) => [version, value?.overall || value])
      )])
    ), null, 2));
    return;
  }
  if (process.env.CATEGORY_ONLY === '1') {
    console.log(JSON.stringify(Object.fromEntries(
      Object.entries(results).map(([month, monthResult]) => [month, Object.fromEntries(
        ['legacy', 'seasonal25', 'seasonal50', 'seasonal75', 'seasonal100'].map((version) => [version, monthResult[version].categories])
      )])
    ), null, 2));
    return;
  }
  if (process.env.CATEGORY_WAPE_ONLY === '1') {
    const versions = ['legacy', 'seasonal25', 'seasonal50', 'seasonal75', 'seasonal100'];
    console.log(JSON.stringify(Object.fromEntries(
      Object.entries(results).map(([month, monthResult]) => [month, Object.fromEntries(
        [...new Set(versions.flatMap((version) => Object.keys(monthResult[version].categories)))].map((categoryName) => [
          categoryName,
          Object.fromEntries(versions.map((version) => [version, monthResult[version].categories[categoryName]?.wape ?? null])),
        ])
      )])
    ), null, 2));
    return;
  }
  if (process.env.TOP_ERRORS_ONLY === '1') {
    console.log(JSON.stringify(Object.fromEntries(
      Object.entries(results).map(([month, monthResult]) => [month, {
        overall: monthResult.categorySeasonal.overall,
        topErrors: monthResult.categorySeasonal.topErrors,
      }])
    ), null, 2));
    return;
  }
  console.log(JSON.stringify({
    sources: sourceSummary,
    stockProducts: stockRows.length,
    results,
  }, null, 2));
}

main().catch((error) => {
  console.error(error);
  process.exit(1);
});
