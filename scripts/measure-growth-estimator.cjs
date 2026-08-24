const fs = require("fs");
const path = require("path");
const Module = require("module");
const esbuild = require("esbuild");
const XLSX = require("xlsx");

const ROOT = path.join(__dirname, "..");
const DOWNLOADS = path.resolve(ROOT, "..");
const JULY_SALES = "C:/Users/X13/Documents/ventas por mes/ventas julio.xlsx";
const JUNE_CLOSE = "C:/Users/X13/Documents/ventas por mes/ventas junio.xlsx";
const SOURCES = [
  "Venta de diciembre 2024.xlsx",
  "VENTA AÑO 2025 - ANGEL.xlsx",
  "VENTA ENERO 2026.xlsx",
  "VENTAS FEEEBRERO 2026.xlsx",
  "MARZO VENTAS.xlsx",
  "venta abril 2026.xlsx",
  "VENTAS DE MAYO Y JUNIO - ANGEL.xlsx",
];
const TARGET_MONTHS = ["2026-04", "2026-05", "2026-06", "2026-07"];
const VERSIONS = ["categorySeasonal", "csG3S50", "csG6S50", "csG1S25", "csG3S25", "csG3S20"];

function normalized(value) {
  return String(value || "")
    .toUpperCase()
    .normalize("NFD")
    .replace(/[\u0300-\u036f]/g, "")
    .replace(/\s+/g, " ")
    .trim();
}

function findDownload(search) {
  const target = normalized(search);
  const match = fs.readdirSync(DOWNLOADS).find((name) => normalized(name) === target);
  if (!match) throw new Error(`No se encontro ${search}`);
  return path.join(DOWNLOADS, match);
}

function monthKey(value) {
  const date = value instanceof Date ? value : new Date(value);
  if (Number.isNaN(date.getTime())) return "";
  return `${date.getFullYear()}-${String(date.getMonth() + 1).padStart(2, "0")}`;
}

async function loadAppFunctions() {
  const built = await esbuild.build({
    entryPoints: [path.join(ROOT, "App.jsx")],
    bundle: true,
    platform: "node",
    format: "cjs",
    write: false,
    loader: { ".css": "text" },
    define: {
      "import.meta.env.VITE_API_URL": JSON.stringify(""),
      "import.meta.env.DEV": "false",
      "import.meta.env.PROD": "true",
    },
    logLevel: "silent",
  });
  const appModule = new Module("growth-estimator");
  appModule.filename = path.join(ROOT, "growth-estimator.bundle.cjs");
  appModule.paths = module.paths;
  appModule._compile(built.outputFiles[0].text, appModule.filename);
  return appModule.exports;
}

async function main() {
  const app = await loadAppFunctions();
  const stockRows = app.parseStock(XLSX.readFile(findDownload("STOCK IDEAL SUCURSALES.xlsx"), { cellDates: true }));
  const read = (filePath) =>
    app.parseSalesOrReturns(XLSX.readFile(filePath, { cellDates: true }), "ventas", path.basename(filePath));
  const measuredSales = [...SOURCES.flatMap((name) => read(findDownload(name))), ...read(JULY_SALES)];
  const allSales = [...measuredSales, ...read(JUNE_CLOSE)];

  const catalog = new Set(stockRows.map((row) => normalized(row.producto)));
  const monthTotals = new Map();
  for (const row of measuredSales) {
    const product = normalized(row.producto);
    if (!catalog.has(product)) continue;
    const key = monthKey(row.fecha);
    if (!key) continue;
    monthTotals.set(key, (monthTotals.get(key) || 0) + Number(row.cantidad || 0));
  }
  console.log("=== Totales del catalogo (archivo diario / julio mensual) ===");
  for (const key of [...monthTotals.keys()].sort()) {
    if (key >= "2025-06" && key <= "2026-07") console.log(`${key}: ${Math.round(monthTotals.get(key))}`);
  }

  const results = new Map();
  for (const targetMonth of TARGET_MONTHS) {
    const historical = app.filterVentasBeforeMonth(allSales, targetMonth);
    const actualMap = new Map();
    for (const row of measuredSales.filter((sale) => monthKey(sale.fecha) === targetMonth)) {
      const product = normalized(row.producto);
      actualMap.set(product, (actualMap.get(product) || 0) + Number(row.cantidad || 0));
    }
    for (const version of VERSIONS) {
      const forecastRows = app.calculateForecast({
        stockRows,
        historicalVentas: historical,
        bajas: [],
        existencias: [],
        realProduction: [],
        selectedMonth: targetMonth,
        dailyBufferPct: 0,
        modelVersion: version,
      });
      const rows = forecastRows.map((row) => ({
        actual: actualMap.get(normalized(row.producto)) || 0,
        forecast: Number(row.pronosticoVenta || 0),
      }));
      const actual = rows.reduce((sum, row) => sum + row.actual, 0);
      const absoluteError = rows.reduce((sum, row) => sum + Math.abs(row.actual - row.forecast), 0);
      if (!results.has(version)) results.set(version, new Map());
      results.get(version).set(targetMonth, {
        actual,
        forecast: rows.reduce((sum, row) => sum + row.forecast, 0),
        absoluteError,
        wape: actual > 0 ? (absoluteError / actual) * 100 : null,
        inside: rows.filter((row) => Math.abs(row.actual - row.forecast) <= 15).length,
      });
    }
  }

  console.log("\n=== WAPE fuera de muestra (menor es mejor) ===");
  console.log(`${"Version".padEnd(20)}${TARGET_MONTHS.map((month) => month.slice(5).padStart(8)).join("")}${"Pond".padStart(8)}`);
  for (const version of VERSIONS) {
    let error = 0;
    let actual = 0;
    const cells = TARGET_MONTHS.map((month) => {
      const row = results.get(version).get(month);
      error += row.absoluteError;
      actual += row.actual;
      return row.wape.toFixed(2).padStart(8);
    });
    console.log(`${version.padEnd(20)}${cells.join("")}${((error / actual) * 100).toFixed(2).padStart(8)}`);
  }
}

main().catch((error) => {
  console.error(error);
  process.exit(1);
});
