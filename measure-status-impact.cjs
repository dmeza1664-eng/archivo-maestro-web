const fs = require("fs");
const path = require("path");
const Module = require("module");
const esbuild = require("esbuild");
const XLSX = require("xlsx");

const ROOT = __dirname;
const DOWNLOADS = path.resolve(ROOT, "..");
const STATUS_FILE = path.join(DOWNLOADS, "ESTATUS_PRODUCTOS_PARA_LLENAR.xlsx");
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

function normalized(value) {
  return String(value || "")
    .toUpperCase().normalize("NFD").replace(/[\u0300-\u036f]/g, "")
    .replace(/[.,/\\_-]+/g, " ").replace(/\s+/g, " ").trim();
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

function loadStatus() {
  const workbook = XLSX.readFile(STATUS_FILE);
  const rows = XLSX.utils.sheet_to_json(workbook.Sheets["Estatus productos"], { defval: "" });
  const status = new Map();
  for (const row of rows) {
    const product = normalized(row.Producto);
    if (!product) continue;
    status.set(product, {
      estatus: String(row.Estatus || "").trim().toUpperCase(),
      desde: String(row.Desde || "").trim(),
      ultimaVenta: String(row["Ultima venta"] || "").trim(),
    });
  }
  return status;
}

async function loadAppFunctions() {
  const built = await esbuild.build({
    entryPoints: [path.join(ROOT, "App.jsx")],
    bundle: true, platform: "node", format: "cjs", write: false,
    loader: { ".css": "text" },
    define: { "import.meta.env.VITE_API_URL": JSON.stringify("") },
    logLevel: "silent",
  });
  const appModule = new Module("status-impact");
  appModule.filename = path.join(ROOT, "status-impact.bundle.cjs");
  appModule.paths = module.paths;
  appModule._compile(built.outputFiles[0].text, appModule.filename);
  return appModule.exports;
}

async function main() {
  const app = await loadAppFunctions();
  const status = loadStatus();
  const stockRows = app.parseStock(XLSX.readFile(findDownload("STOCK IDEAL SUCURSALES.xlsx"), { cellDates: true }));
  const read = (p) => app.parseSalesOrReturns(XLSX.readFile(p, { cellDates: true }), "ventas", path.basename(p));
  const measuredSales = [...SOURCES.flatMap((n) => read(findDownload(n))), ...read(JULY_SALES)];
  const allSales = [...measuredSales, ...read(JUNE_CLOSE)];

  const pad = (v, w) => String(v).padStart(w);
  console.log(`${"Mes".padEnd(9)}${pad("Error base", 12)}${pad("Baja", 9)}${pad("Estacional", 12)}${pad("Error nuevo", 13)}${pad("WAPE antes", 12)}${pad("WAPE despues", 14)}`);

  const totals = { before: 0, after: 0, actual: 0 };
  const detail = [];
  for (const month of TARGET_MONTHS) {
    const historical = app.filterVentasBeforeMonth(allSales, month);
    const actualMap = new Map();
    for (const row of measuredSales.filter((s) => monthKey(s.fecha) === month)) {
      const p = normalized(row.producto);
      actualMap.set(p, (actualMap.get(p) || 0) + Number(row.cantidad || 0));
    }
    const forecastRows = app.calculateForecast({
      stockRows, historicalVentas: historical, bajas: [], existencias: [], realProduction: [],
      selectedMonth: month, dailyBufferPct: 0,
    });

    let errorBefore = 0;
    let errorAfter = 0;
    let savedBaja = 0;
    let savedSeasonal = 0;
    let actualTotal = 0;
    const missed = [];
    for (const row of forecastRows) {
      const product = normalized(row.producto);
      const actual = actualMap.get(product) || 0;
      const forecast = Number(row.pronosticoVenta || 0);
      const info = status.get(product);
      const estatus = info?.estatus || "ACTIVO";
      const suppressed = estatus === "BAJA" || estatus === "ESTACIONAL";
      const newForecast = suppressed ? 0 : forecast;

      errorBefore += Math.abs(actual - forecast);
      errorAfter += Math.abs(actual - newForecast);
      actualTotal += actual;
      if (suppressed) {
        const saved = Math.abs(actual - forecast) - Math.abs(actual - newForecast);
        if (estatus === "BAJA") savedBaja += saved;
        else savedSeasonal += saved;
        if (actual > 0) missed.push({ product, actual, forecast, estatus });
      }
    }

    totals.before += errorBefore;
    totals.after += errorAfter;
    totals.actual += actualTotal;
    detail.push({ month, missed });
    console.log(
      `${month.padEnd(9)}${pad(errorBefore.toFixed(0), 12)}${pad(savedBaja.toFixed(0), 9)}${pad(savedSeasonal.toFixed(0), 12)}${pad(errorAfter.toFixed(0), 13)}${pad((errorBefore / actualTotal * 100).toFixed(2) + "%", 12)}${pad((errorAfter / actualTotal * 100).toFixed(2) + "%", 14)}`
    );
  }

  console.log(`\nPonderado antes:  ${(totals.before / totals.actual * 100).toFixed(2)}%`);
  console.log(`Ponderado despues: ${(totals.after / totals.actual * 100).toFixed(2)}%`);
  console.log(`Ganancia: ${((totals.before - totals.after) / totals.actual * 100).toFixed(2)} puntos, ${(totals.before - totals.after).toFixed(0)} piezas de error`);

  console.log("\n=== ALERTA: productos suprimidos que SI vendieron ===");
  let anyMissed = false;
  for (const entry of detail) {
    for (const row of entry.missed) {
      anyMissed = true;
      console.log(`  ${entry.month}  ${row.product.padEnd(34)}${row.estatus.padEnd(12)}vendio ${row.actual}, pronostico era ${row.forecast.toFixed(0)}`);
    }
  }
  if (!anyMissed) console.log("  ninguno");
}

main().catch((error) => { console.error(error.message); process.exit(1); });
