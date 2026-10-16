// Vista «Producción diaria vs pronóstico»: los números salen solo de los Excel
// diarios (hojas Resumen del día y Por producto). Re-subir un día reemplaza.
const assert = require("assert");
const fs = require("fs");
const path = require("path");
const Module = require("module");
const esbuild = require("esbuild");
const XLSX = require("xlsx");

const ROOT = path.join(__dirname, "..");
const FIXTURE_12 = path.join(ROOT, "scripts/fixtures/Reporte_diario_Produccion_vs_Pronostico_2026-10-12.xlsx");
const FIXTURE_13 = path.join(ROOT, "scripts/fixtures/Reporte_diario_Produccion_vs_Pronostico_2026-10-13.xlsx");

async function loadApp() {
  const built = await esbuild.build({
    entryPoints: [path.join(ROOT, "App.jsx")],
    bundle: true,
    platform: "node",
    format: "cjs",
    write: false,
    loader: { ".css": "text" },
    define: {
      "import.meta.env.VITE_API_URL": '""',
      "import.meta.env.DEV": "false",
      "import.meta.env.PROD": "true",
    },
    logLevel: "silent",
  });
  const appModule = new Module("daily-production-vs-forecast-test");
  appModule.filename = path.join(__dirname, "daily-production-vs-forecast-test.bundle.cjs");
  appModule.paths = module.paths;
  appModule._compile(built.outputFiles[0].text, appModule.filename);
  return appModule.exports;
}

function workbookFromSheets(sheets) {
  const workbook = XLSX.utils.book_new();
  for (const [name, rows] of Object.entries(sheets)) {
    XLSX.utils.book_append_sheet(workbook, XLSX.utils.aoa_to_sheet(rows), name);
  }
  return workbook;
}

function almostEqual(actual, expected, tolerance = 0.51) {
  assert(
    Math.abs(Number(actual) - Number(expected)) <= tolerance,
    `esperado ${expected}, obtuvo ${actual}`
  );
}

async function main() {
  assert(fs.existsSync(FIXTURE_12), "Falta el fixture del 12-oct-2026.");
  assert(fs.existsSync(FIXTURE_13), "Falta el fixture del 13-oct-2026.");

  const appSource = fs.readFileSync(path.join(ROOT, "App.jsx"), "utf8");
  const viewStart = appSource.indexOf("function DailyVsForecastPage(");
  const dashboardStart = appSource.indexOf("function Dashboard(");
  assert(viewStart > 0 && dashboardStart > viewStart, "Existe la vista Producción diaria vs pronóstico.");
  const viewSource = appSource.slice(viewStart, dashboardStart);
  assert.match(viewSource, /Producción diaria vs pronóstico/, "El título de la vista está en español.");
  assert.match(viewSource, /margen de seguridad/, "La vista usa la frase margen de seguridad.");
  assert.doesNotMatch(viewSource, /colch[oó]n/i, "La vista no usa la palabra colchón.");
  assert.match(viewSource, /Distribución no muestra porcentaje de error/, "Distribución no lleva % de error.");
  assert.match(viewSource, /solo de los Excel diarios/, "La vista declara que los datos salen de los archivos.");
  assert.match(appSource, /function detectDailyVsForecastDate/, "Se detecta la fecha del Día o del título.");
  assert.match(appSource, /se reemplaza; no se duplica/, "Re-subir un día reemplaza.");
  const css = fs.readFileSync(path.join(ROOT, "style.css"), "utf8");
  assert.match(
    css,
    /\.main\.daily-vs-forecast-page > \* \{\s*order: 0;/,
    "En esta vista la carga no se manda al fondo por el order del dashboard."
  );

  const {
    DAILY_VS_FORECAST_MARGIN_PCT,
    detectDailyVsForecastDate,
    parseDailyProductionVsForecastWorkbook,
    upsertDailyVsForecastReports,
    sanitizeDailyVsForecastReports,
    buildDailyVsForecastMetrics,
    buildDailyVsForecastViewModel,
    recommendedFromForecast,
    formatDailyVsForecastPct,
    isDistribucionChannel,
  } = await loadApp();

  assert.strictEqual(DAILY_VS_FORECAST_MARGIN_PCT, 10, "El margen de seguridad de esta vista es 10%.");
  assert.strictEqual(
    detectDailyVsForecastDate({
      title: "Reporte diario · Planta León · Mandado a producir vs. recomendado · martes 13 de octubre de 2026",
      diaCell: "Día 2026-10-13",
    }),
    "2026-10-13"
  );
  assert.strictEqual(
    detectDailyVsForecastDate({
      title: "Reporte diario · Planta León · lunes 12 de octubre de 2026",
      diaCell: "Día",
    }),
    "2026-10-12",
    "Si la celda Día no trae fecha, se lee del título."
  );

  const oct13 = parseDailyProductionVsForecastWorkbook(XLSX.readFile(FIXTURE_13, { cellDates: true }), {
    fileName: path.basename(FIXTURE_13),
  });
  const oct12 = parseDailyProductionVsForecastWorkbook(XLSX.readFile(FIXTURE_12, { cellDates: true }), {
    fileName: path.basename(FIXTURE_12),
  });

  assert.strictEqual(oct13.fecha, "2026-10-13");
  assert.strictEqual(oct12.fecha, "2026-10-12");
  assert.match(oct13.title, /martes 13 de octubre de 2026/);
  assert.match(oct13.note, /margen de seguridad 10%/);
  assert.doesNotMatch(oct13.note, /colch[oó]n/i);
  almostEqual(oct13.day.recomendado, 1207);
  almostEqual(oct13.day.mandado, 1361);
  almostEqual(oct13.day.difPiezas, 154);
  almostEqual(oct13.day.difPct, 0.1276824923357359, 1e-9);
  almostEqual(oct13.day.producido, 1330);
  almostEqual(oct13.day.pronosticoVenta, 1097);
  almostEqual(oct13.monthToDate.recomendado, 11682);
  assert.strictEqual(oct13.monthToDate.fechaCorte, "2026-10-13");
  assert.match(oct13.monthToDate.concepto, /Acumulado mes al 2026-10-13/);
  almostEqual(oct12.day.recomendado, 1149);
  almostEqual(oct12.day.mandado, 1077);
  almostEqual(recommendedFromForecast(oct13.day.pronosticoVenta), oct13.day.recomendado, 1);
  almostEqual(recommendedFromForecast(oct12.day.pronosticoVenta), oct12.day.recomendado, 1);
  assert(oct13.products.length >= 70, `El 13-oct trae el detalle por producto (${oct13.products.length}).`);
  assert(oct13.areas.length >= 8, "El 13-oct trae el desglose por área.");
  const pina = oct13.products.find((row) => row.producto === "PINA GDE");
  assert(pina, "Se lee PINA GDE.");
  assert.strictEqual(pina.area, "3 Leches Grandes");
  almostEqual(pina.recomendado, 10.8, 0.05);
  almostEqual(pina.mandado, 15);
  almostEqual(pina.pronosticoVenta, 9.8, 0.05);

  const first = upsertDailyVsForecastReports([], [oct12]);
  assert.strictEqual(first.reports.length, 1);
  assert.deepStrictEqual(first.replaced, []);
  const added = upsertDailyVsForecastReports(first.reports, [oct13]);
  assert.strictEqual(added.reports.length, 2, "Dos días distintos no se pisan.");
  assert.deepStrictEqual(added.replaced, []);
  const replacement = parseDailyProductionVsForecastWorkbook(XLSX.readFile(FIXTURE_13, { cellDates: true }), {
    fileName: "reenvio-13.xlsx",
  });
  replacement.day = { ...replacement.day, mandado: 9999 };
  const replaced = upsertDailyVsForecastReports(added.reports, [replacement]);
  assert.strictEqual(replaced.reports.length, 2, "Re-subir el 13 no duplica el día.");
  assert.deepStrictEqual(replaced.replaced, ["2026-10-13"]);
  const kept = replaced.reports.find((row) => row.fecha === "2026-10-13");
  almostEqual(kept.day.mandado, 9999);
  assert.strictEqual(kept.sourceFile, "reenvio-13.xlsx");
  assert.strictEqual(replaced.reports.filter((row) => row.fecha === "2026-10-13").length, 1);

  const sanitized = sanitizeDailyVsForecastReports([oct13, oct13, oct12]);
  assert.strictEqual(sanitized.length, 2, "sanitize tampoco duplica por fecha.");

  const dist = buildDailyVsForecastMetrics({
    recomendado: 10,
    mandado: 12,
    producido: 12,
    pronosticoVenta: 9,
    difPiezas: 2,
    difPct: 0.2,
    area: "Distribución",
    producto: "SURTIDO PLANTA",
  });
  assert(isDistribucionChannel("Distribución"));
  assert.strictEqual(dist.difPct, null, "Distribución no lleva Dif. %.");
  assert.strictEqual(formatDailyVsForecastPct(dist.difPct), "—");
  const regular = buildDailyVsForecastMetrics({
    recomendado: 10,
    mandado: 12,
    difPct: 0.2,
    producto: "PINA GDE",
  });
  assert.strictEqual(regular.difPct, 0.2);
  assert.strictEqual(formatDailyVsForecastPct(regular.difPct), "20.0%");

  const distWorkbook = workbookFromSheets({
    "Resumen del día": [
      ["Reporte diario · Planta León · Mandado a producir vs. recomendado · miércoles 14 de octubre de 2026"],
      ["Recomendado = pronóstico de venta de sucursales + margen de seguridad 10%."],
      [],
      ["Concepto", "Recomendado", "Mandado a producir", "Dif. piezas", "Dif. %", "Producido", "Pronóstico de venta (sin margen de seguridad)"],
      ["Día 2026-10-14", 100, 110, 10, 0.1, 108, 91],
      ["Acumulado mes al 2026-10-14", 200, 210, 10, 0.05, 208, 182],
      [],
      ["Área", "Recomendado", "Mandado a producir", "Dif. piezas", "Dif. %", "Producido"],
      ["Distribución", 40, 44, 4, 0.1, 44],
      ["Vainillas", 60, 66, 6, 0.1, 64],
    ],
    "Por producto": [
      ["Área", "Producto", "Recomendado", "Mandado a producir", "Dif. piezas", "Dif. %", "Producido", "Pronóstico de venta"],
      ["Distribución", "SURTIDO PLANTA", 40, 44, 4, 0.1, 44, 36],
      ["Vainillas", "VAINILLA GDE", 60, 66, 6, 0.1, 64, 55],
    ],
  });
  const distReport = parseDailyProductionVsForecastWorkbook(distWorkbook, { fileName: "dist.xlsx" });
  assert.strictEqual(distReport.fecha, "2026-10-14");
  const distArea = distReport.areas.find((row) => row.area === "Distribución");
  const distProduct = distReport.products.find((row) => row.producto === "SURTIDO PLANTA");
  assert.strictEqual(distArea.difPct, null, "El área Distribución no muestra error %.");
  assert.strictEqual(distProduct.difPct, null, "El producto de Distribución no muestra error %.");
  assert.strictEqual(distReport.areas.find((row) => row.area === "Vainillas").difPct, 0.1);

  const model = buildDailyVsForecastViewModel({
    reports: [oct12, oct13],
    month: "2026-10",
    fecha: "2026-10-13",
  });
  assert.strictEqual(model.selected.fecha, "2026-10-13");
  assert.strictEqual(model.days.length, 2);
  assert(model.productsForDay.some((row) => row.producto === "PINA GDE"));
  assert(model.areasForDay.some((row) => row.area === "Gelatinas"));
  almostEqual(model.officialMonthToDate.recomendado, 11682);
  almostEqual(model.uploadedMonthToDate.recomendado, oct12.day.recomendado + oct13.day.recomendado, 0.6);
  const monthModel = buildDailyVsForecastViewModel({
    reports: [oct12, oct13],
    month: "2026-10",
    productQuery: "PINA GDE",
  });
  assert(
    monthModel.productsMonthToDate.every((row) => row.producto === "PINA GDE"),
    "El filtro de producto aplica al acumulado."
  );
  almostEqual(
    monthModel.productsMonthToDate[0].mandado,
    oct12.products.find((row) => row.producto === "PINA GDE").mandado
      + oct13.products.find((row) => row.producto === "PINA GDE").mandado
  );

  const titleOnly = parseDailyProductionVsForecastWorkbook(workbookFromSheets({
    "Resumen del día": [
      ["Reporte diario · Planta León · Mandado a producir vs. recomendado · martes 13 de octubre de 2026"],
      ["Nota de prueba. Recomendado = pronóstico + margen de seguridad 10%."],
      [],
      ["Concepto", "Recomendado", "Mandado a producir", "Dif. piezas", "Dif. %", "Producido", "Pronóstico de venta (sin margen de seguridad)"],
      ["Día", 10, 11, 1, 0.1, 11, 9],
      ["Acumulado mes", 20, 21, 1, 0.05, 21, 18],
    ],
    "Por producto": [
      ["Área", "Producto", "Recomendado", "Mandado a producir", "Dif. piezas", "Dif. %", "Producido", "Pronóstico de venta"],
      ["Vainillas", "VAINILLA GDE", 10, 11, 1, 0.1, 11, 9],
    ],
  }));
  assert.strictEqual(titleOnly.fecha, "2026-10-13", "La fecha sale del título si la fila Día no la trae.");

  console.log("daily-production-vs-forecast-test: fixtures 12 y 13-oct, replace-by-day y Distribución sin % de error");
}

main().catch((error) => {
  console.error(error);
  process.exit(1);
});
