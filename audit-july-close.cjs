const fs = require('fs');
const path = require('path');
const Module = require('module');
const esbuild = require('esbuild');
const XLSX = require('xlsx');

const ROOT = __dirname;
const DOWNLOADS = path.resolve(ROOT, '..');
const SALES_PATH = 'C:/Users/X13/Documents/ventas por mes/ventas julio.xlsx';
const PRODUCTION_PATH = 'C:/Users/X13/Documents/PRODUCIDO/PRODUCIDO JUlIO.xlsx';
const FINAL_JUNE_SALES_PATH = 'C:/Users/X13/Documents/ventas por mes/ventas junio.xlsx';
const DISPLAYED_JULY_FORECAST = 26061;
const FORECAST_MARGIN = 0.12;

function normalized(value) {
  return String(value || '')
    .toUpperCase()
    .normalize('NFD')
    .replace(/[\u0300-\u036f]/g, '')
    .replace(/[.,/\\_-]+/g, ' ')
    .replace(/\s+/g, ' ')
    .trim();
}

function matchKey(value) {
  return normalized(value)
    .replace(/\b(PASTEL|TARTA|PANQUE|PAY|DE|DEL|LA|EL)\b/g, ' ')
    .replace(/\bGRANDE\b/g, 'GDE')
    .replace(/\bMEDIANO\b/g, 'MED')
    .replace(/\bCHICO\b/g, 'CH')
    .replace(/\s+/g, ' ')
    .trim();
}

function findDownload(search) {
  const target = normalized(search);
  const match = fs.readdirSync(DOWNLOADS).find((name) => normalized(name) === target);
  if (!match) throw new Error(`No se encontró ${search}`);
  return path.join(DOWNLOADS, match);
}

function category(product) {
  const value = normalized(product);
  if (value.includes('GELATINA')) return 'Gelatinas';
  if (value.includes('GALLETA')) return 'Galletas';
  if (/\b(GDE|GRANDE)\b/.test(value)) return 'Pasteles grandes';
  if (/\b(MED|MEDIANO)\b/.test(value) && !value.includes('MINI')) return 'Pasteles medianos';
  if (/\b(CH|CHICO)\b/.test(value)) return 'Pasteles chicos';
  if (value.includes('MINI')) return 'Mini medianos';
  if (value.includes('BOLLO') || /\bPAN\b/.test(value)) return 'Pan';
  return 'Otros';
}

function isExcluded(product) {
  const value = normalized(product);
  return /\b(REBANADA|REBANADAS|REB|RBN)\b/.test(value) || /\b(PROMO|PROMOCION|PROMOCIONAL)\b/.test(value);
}

function metrics(rows) {
  const actual = rows.reduce((sum, row) => sum + row.actual, 0);
  const forecast = rows.reduce((sum, row) => sum + row.forecast, 0);
  const produced = rows.reduce((sum, row) => sum + row.produced, 0);
  const absoluteError = rows.reduce((sum, row) => sum + Math.abs(row.actual - row.forecast), 0);
  return {
    products: rows.length,
    actual: Number(actual.toFixed(2)),
    forecast: Number(forecast.toFixed(2)),
    produced: Number(produced.toFixed(2)),
    forecastBias: Number((actual - forecast).toFixed(2)),
    productionBalance: Number((produced - actual).toFixed(2)),
    wape: actual > 0 ? Number((absoluteError / actual * 100).toFixed(2)) : null,
    mae: rows.length ? Number((absoluteError / rows.length).toFixed(2)) : null,
    inside15: rows.filter((row) => Math.abs(row.actual - row.forecast) <= 15).length,
  };
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
  const appModule = new Module('july-close-audit');
  appModule.filename = path.join(ROOT, 'july-close-audit.bundle.cjs');
  appModule.paths = module.paths;
  appModule._compile(built.outputFiles[0].text, appModule.filename);
  return appModule.exports;
}

function aggregateToCatalog(records, officialProducts) {
  const direct = new Set(officialProducts);
  const byMatchKey = new Map();
  for (const product of officialProducts) {
    const key = matchKey(product);
    if (!byMatchKey.has(key)) byMatchKey.set(key, []);
    byMatchKey.get(key).push(product);
  }
  const matched = new Map();
  const unmatched = new Map();
  for (const row of records) {
    const sourceProduct = normalized(row.producto);
    const quantity = Number(row.cantidad || 0);
    if (!sourceProduct || quantity === 0 || isExcluded(sourceProduct)) continue;
    const candidates = byMatchKey.get(matchKey(sourceProduct)) || [];
    const official = direct.has(sourceProduct) ? sourceProduct : candidates.length === 1 ? candidates[0] : '';
    const target = official ? matched : unmatched;
    const key = official || sourceProduct;
    target.set(key, (target.get(key) || 0) + quantity);
  }
  return { matched, unmatched };
}

async function main() {
  const app = await loadAppFunctions();
  const stockRows = app.parseStock(XLSX.readFile(findDownload('STOCK IDEAL SUCURSALES.xlsx'), { cellDates: true }));
  const officialProducts = stockRows
    .map((row) => normalized(row.producto))
    .filter((product) => !isExcluded(product));
  const sourceNames = [
    'Venta de diciembre 2024.xlsx',
    'VENTA AÑO 2025 - ANGEL.xlsx',
    'VENTA ENERO 2026.xlsx',
    'VENTAS FEEEBRERO 2026.xlsx',
    'MARZO VENTAS.xlsx',
    'venta abril 2026.xlsx',
    'VENTAS DE MAYO Y JUNIO - ANGEL.xlsx',
  ];
  const historicalSales = sourceNames.flatMap((sourceName) => {
    const filePath = findDownload(sourceName);
    return app.parseSalesOrReturns(
      XLSX.readFile(filePath, { cellDates: true }),
      'ventas',
      path.basename(filePath)
    );
  });
  historicalSales.push(...app.parseSalesOrReturns(
    XLSX.readFile(FINAL_JUNE_SALES_PATH, { cellDates: true }),
    'ventas',
    path.basename(FINAL_JUNE_SALES_PATH)
  ));
  const forecastInput = {
    stockRows,
    historicalVentas: app.filterVentasBeforeMonth(historicalSales, '2026-07'),
    bajas: [],
    existencias: [],
    realProduction: [],
    selectedMonth: '2026-07',
    dailyBufferPct: 0,
  };
  const forecastRows = app.calculateForecast(forecastInput);
  const julySales = app.parseSalesOrReturns(
    XLSX.readFile(SALES_PATH, { cellDates: true }),
    'ventas',
    path.basename(SALES_PATH)
  );
  const julyProduction = app.parseProductionReal(XLSX.readFile(PRODUCTION_PATH, { cellDates: true }));
  const sales = aggregateToCatalog(julySales, officialProducts);
  const production = aggregateToCatalog(julyProduction, officialProducts);
  const buildComparison = (rows) => rows.map((row) => {
    const product = normalized(row.producto);
    return {
      product,
      category: category(product),
      forecast: Number(row.pronosticoVenta || 0),
      actual: sales.matched.get(product) || 0,
      produced: production.matched.get(product) || 0,
    };
  });
  const comparison = buildComparison(forecastRows);
  const categories = Object.fromEntries(
    [...new Set(comparison.map((row) => row.category))].map((name) => [
      name,
      metrics(comparison.filter((row) => row.category === name)),
    ])
  );
  const rawSalesWorkbook = XLSX.readFile(SALES_PATH);
  const rawSalesRows = XLSX.utils.sheet_to_json(
    rawSalesWorkbook.Sheets[rawSalesWorkbook.SheetNames[0]],
    { header: 1, defval: '' }
  ).slice(1);
  const rawSalesTotal = rawSalesRows
    .filter((row) => !normalized(row[0]).startsWith('TOTAL'))
    .reduce((sum, row) => sum + Number(row[1] || 0), 0);
  const rawProductionWorkbook = XLSX.readFile(PRODUCTION_PATH);
  const rawProductionRows = XLSX.utils.sheet_to_json(
    rawProductionWorkbook.Sheets[rawProductionWorkbook.SheetNames[0]],
    { header: 1, defval: '' }
  ).slice(1);
  const rawProductionTotal = rawProductionRows.reduce((sum, row) => sum + Number(row[0] || 0), 0);
  const overall = metrics(comparison);
  const adjustedComparison = comparison.map((row) => ({
    ...row,
    forecast: row.forecast * (1 + FORECAST_MARGIN),
  }));
  const adjustedOverall = metrics(adjustedComparison);
  const adjustedCategories = Object.fromEntries(
    [...new Set(adjustedComparison.map((row) => row.category))].map((name) => [
      name,
      metrics(adjustedComparison.filter((row) => row.category === name)),
    ])
  );

  const result = {
    files: {
      sales: {
        parsedProducts: julySales.length,
        monthlyTotals: julySales.filter((row) => row.monthlyTotal).length,
        dailyRows: julySales.filter((row) => !row.monthlyTotal).length,
        rawTotal: rawSalesTotal,
      },
      production: {
        parsedRows: julyProduction.length,
        rowsWithDate: julyProduction.filter((row) => row.fecha).length,
        rawTotal: rawProductionTotal,
      },
    },
    catalog: {
      forecastProducts: comparison.length,
      matchedSalesProducts: sales.matched.size,
      matchedProductionProducts: production.matched.size,
      unmatchedSalesTotal: Number([...sales.unmatched.values()].reduce((sum, value) => sum + value, 0).toFixed(2)),
      unmatchedProductionTotal: Number([...production.unmatched.values()].reduce((sum, value) => sum + value, 0).toFixed(2)),
    },
    displayedForecastAggregate: {
      actualRegularCatalog: overall.actual,
      displayedForecast: DISPLAYED_JULY_FORECAST,
      difference: Number((overall.actual - DISPLAYED_JULY_FORECAST).toFixed(2)),
      differencePct: Number(((overall.actual - DISPLAYED_JULY_FORECAST) / DISPLAYED_JULY_FORECAST * 100).toFixed(2)),
      producedRegularCatalog: overall.produced,
      productionBalance: overall.productionBalance,
    },
    reconstructedForecast: overall,
    adjustedForecast12Pct: adjustedOverall,
    categories,
    adjustedCategories12Pct: adjustedCategories,
    topForecastErrors: comparison
      .map((row) => ({ ...row, absoluteError: Math.abs(row.actual - row.forecast) }))
      .sort((a, b) => b.absoluteError - a.absoluteError)
      .slice(0, 15),
    topProductionDifferences: comparison
      .map((row) => ({ ...row, productionDifference: row.produced - row.actual }))
      .sort((a, b) => Math.abs(b.productionDifference) - Math.abs(a.productionDifference))
      .slice(0, 15),
    unmatchedSales: [...sales.unmatched.entries()]
      .map(([product, quantity]) => ({ product, quantity }))
      .sort((a, b) => b.quantity - a.quantity)
      .slice(0, 20),
    unmatchedProduction: [...production.unmatched.entries()]
      .map(([product, quantity]) => ({ product, quantity }))
      .sort((a, b) => b.quantity - a.quantity)
      .slice(0, 20),
  };
  if (process.env.WRITE_REPORT === '1') {
    const reportPath = 'C:/Users/X13/Documents/COMPARATIVO_JULIO_REAL_VS_PRONOSTICADO.xlsx';
    const summarySheet = [
      { Indicador: 'Periodo', Valor: 'Julio 2026' },
      { Indicador: 'Estado', Valor: 'Comparativo con archivos mensuales sin fecha diaria' },
      { Indicador: 'Pronostico agregado mostrado en pagina', Valor: DISPLAYED_JULY_FORECAST },
      { Indicador: 'Venta regular comparable', Valor: overall.actual },
      { Indicador: 'Diferencia agregada real - pagina', Valor: Number((overall.actual - DISPLAYED_JULY_FORECAST).toFixed(2)) },
      { Indicador: 'Desviacion agregada contra pagina', Valor: `${(Math.abs(overall.actual - DISPLAYED_JULY_FORECAST) / DISPLAYED_JULY_FORECAST * 100).toFixed(2)}%` },
      { Indicador: 'Cumplimiento agregado contra pagina', Valor: `${(overall.actual / DISPLAYED_JULY_FORECAST * 100).toFixed(2)}%` },
      { Indicador: 'Pronostico reconstruido del modelo actual', Valor: overall.forecast },
      { Indicador: 'Margen operativo agregado 12%', Valor: Number((overall.forecast * FORECAST_MARGIN).toFixed(2)) },
      { Indicador: 'Pronostico ajustado con margen 12%', Valor: adjustedOverall.forecast },
      { Indicador: 'Diferencia real - pronostico ajustado', Valor: adjustedOverall.forecastBias },
      { Indicador: 'Desviacion total del pronostico ajustado', Valor: `${(Math.abs(adjustedOverall.forecastBias) / adjustedOverall.forecast * 100).toFixed(2)}%` },
      { Indicador: 'Cumplimiento del pronostico ajustado', Valor: `${(adjustedOverall.actual / adjustedOverall.forecast * 100).toFixed(2)}%` },
      { Indicador: 'WAPE por producto sin margen', Valor: `${overall.wape}%` },
      { Indicador: 'WAPE por producto con margen 12%', Valor: `${adjustedOverall.wape}%` },
      { Indicador: 'MAE sin margen', Valor: overall.mae },
      { Indicador: 'MAE con margen 12%', Valor: adjustedOverall.mae },
      { Indicador: 'Productos dentro de +/-15 sin margen', Valor: `${overall.inside15} de ${overall.products}` },
      { Indicador: 'Productos dentro de +/-15 con margen', Valor: `${adjustedOverall.inside15} de ${adjustedOverall.products}` },
      { Indicador: 'Venta total del archivo', Valor: rawSalesTotal },
      { Indicador: 'Nota', Valor: 'El margen 12% es un escenario posterior y no forma parte del pronostico original congelado.' },
    ];
    const categorySheet = Object.entries(categories).map(([name, row]) => {
      const adjusted = adjustedCategories[name];
      return {
        Categoria: name,
        Productos: row.products,
        'Venta real': row.actual,
        'Pronostico sin margen': row.forecast,
        'Pronostico con margen 12%': adjusted.forecast,
        'Diferencia real - ajustado': adjusted.forecastBias,
        'WAPE sin margen': row.wape === null ? '' : `${row.wape}%`,
        'WAPE con margen 12%': adjusted.wape === null ? '' : `${adjusted.wape}%`,
        'MAE con margen 12%': adjusted.mae,
        'Dentro de +/-15 con margen': adjusted.inside15,
      };
    });
    const productSheet = comparison
      .map((row) => {
        const adjustedForecast = row.forecast * (1 + FORECAST_MARGIN);
        const adjustedError = Math.abs(row.actual - adjustedForecast);
        return {
          Producto: row.product,
          Categoria: row.category,
          'Venta real': row.actual,
          'Pronostico sin margen': Number(row.forecast.toFixed(2)),
          'Margen 12% piezas': Number((row.forecast * FORECAST_MARGIN).toFixed(2)),
          'Pronostico con margen 12%': Number(adjustedForecast.toFixed(2)),
          'Diferencia real - ajustado': Number((row.actual - adjustedForecast).toFixed(2)),
          'Error absoluto sin margen': Number(Math.abs(row.actual - row.forecast).toFixed(2)),
          'Error absoluto con margen': Number(adjustedError.toFixed(2)),
          'Cumplimiento ajustado %': adjustedForecast > 0 ? `${(row.actual / adjustedForecast * 100).toFixed(2)}%` : '',
          'Error ajustado % sobre venta real': row.actual > 0 ? `${(adjustedError / row.actual * 100).toFixed(2)}%` : '',
          'Dentro de +/-15 ajustado': adjustedError <= 15 ? 'Si' : 'No',
          Estado: row.actual - adjustedForecast > 15
          ? 'Venta superior al pronostico'
          : row.actual - adjustedForecast < -15
            ? 'Venta inferior al pronostico'
            : 'En rango',
        };
      })
      .sort((a, b) => b['Error absoluto con margen'] - a['Error absoluto con margen']);
    const methodologySheet = [
      { Concepto: 'Universo comparable', Detalle: 'Productos del catalogo de pronostico, sin rebanadas ni promociones.' },
      { Concepto: 'Referencia agregada', Detalle: 'El total 26,061 es el pronostico mostrado en la pagina; solo puede compararse contra el total real porque no existe su exportacion por producto.' },
      { Concepto: 'Detalle por producto', Detalle: 'Se reconstruyo con el modelo actual usando exclusivamente historia anterior a julio.' },
      { Concepto: 'Margen 12%', Detalle: 'Escenario operativo calculado como pronostico base multiplicado por 1.12.' },
      { Concepto: 'Comparacion honesta', Detalle: 'El margen se agrego despues de conocer el resultado de julio; no debe presentarse como parte del pronostico original.' },
      { Concepto: 'WAPE', Detalle: 'Suma de errores absolutos por producto dividida entre la venta real comparable.' },
      { Concepto: 'MAE', Detalle: 'Promedio de piezas de error absoluto por producto.' },
      { Concepto: 'Signo de diferencia', Detalle: 'Positivo: se vendio mas que el pronostico. Negativo: se vendio menos.' },
      { Concepto: 'Limitacion', Detalle: 'Los archivos mensuales no contienen fechas y no permiten evaluar el avance por semana.' },
    ];
    const workbook = XLSX.utils.book_new();
    const appendSheet = (rows, name, widths) => {
      const sheet = XLSX.utils.json_to_sheet(rows);
      sheet['!cols'] = widths.map((wch) => ({ wch }));
      sheet['!autofilter'] = { ref: sheet['!ref'] };
      XLSX.utils.book_append_sheet(workbook, sheet, name);
    };
    appendSheet(summarySheet, 'Resumen', [42, 92]);
    appendSheet(productSheet, 'Por producto', [38, 24, 14, 22, 19, 27, 25, 25, 25, 24, 32, 26, 32]);
    appendSheet(categorySheet, 'Por categoria', [24, 12, 14, 24, 28, 26, 20, 24, 22, 30]);
    appendSheet(methodologySheet, 'Metodologia', [28, 110]);
    XLSX.writeFile(workbook, reportPath);
    result.reportPath = reportPath;
  }
  const output = process.env.SUMMARY_ONLY === '1'
    ? {
        files: result.files,
        catalog: result.catalog,
        displayedForecastAggregate: result.displayedForecastAggregate,
        reconstructedForecast: result.reconstructedForecast,
        adjustedForecast12Pct: result.adjustedForecast12Pct,
        categories: result.categories,
        reportPath: result.reportPath,
      }
    : result;
  console.log(JSON.stringify(output, null, 2));
}

main().catch((error) => {
  console.error(error);
  process.exit(1);
});
