const fs = require("fs");
const path = require("path");
const Module = require("module");
const esbuild = require("esbuild");
const XLSX = require("xlsx");

const ROOT = __dirname;
const DOWNLOADS = path.resolve(ROOT, "..");
const JULY_SALES = "C:/Users/X13/Documents/ventas por mes/ventas julio.xlsx";
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
  return String(value || "").toUpperCase().normalize("NFD").replace(/[\u0300-\u036f]/g, "").replace(/\s+/g, " ").trim();
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
    bundle: true, platform: "node", format: "cjs", write: false,
    loader: { ".css": "text" },
    define: { "import.meta.env.VITE_API_URL": JSON.stringify("") },
    logLevel: "silent",
  });
  const appModule = new Module("dormancy");
  appModule.filename = path.join(ROOT, "dormancy.bundle.cjs");
  appModule.paths = module.paths;
  appModule._compile(built.outputFiles[0].text, appModule.filename);
  return appModule.exports;
}

async function main() {
  const app = await loadAppFunctions();
  const stockRows = app.parseStock(XLSX.readFile(findDownload("STOCK IDEAL SUCURSALES.xlsx"), { cellDates: true }));
  const read = (p) => app.parseSalesOrReturns(XLSX.readFile(p, { cellDates: true }), "ventas", path.basename(p));
  const allSales = [...SOURCES.flatMap((n) => read(findDownload(n))), ...read(JULY_SALES)];

  for (const targetMonth of TARGET_MONTHS) {
    const historical = app.filterVentasBeforeMonth(allSales, targetMonth);
    const actualMap = new Map();
    for (const row of allSales.filter((s) => monthKey(s.fecha) === targetMonth)) {
      const p = normalized(row.producto);
      actualMap.set(p, (actualMap.get(p) || 0) + Number(row.cantidad || 0));
    }
    // Ultimo mes con venta por producto, usando solo historia previa al mes objetivo.
    const lastSaleMonth = new Map();
    const monthsWithSales = new Map();
    for (const row of historical) {
      const p = normalized(row.producto);
      const m = monthKey(row.fecha);
      if (!m || Number(row.cantidad || 0) <= 0) continue;
      if (!lastSaleMonth.has(p) || m > lastSaleMonth.get(p)) lastSaleMonth.set(p, m);
      if (!monthsWithSales.has(p)) monthsWithSales.set(p, new Set());
      monthsWithSales.get(p).add(m);
    }
    const historicalMonths = [...new Set(historical.map((r) => monthKey(r.fecha)).filter(Boolean))].sort();

    const forecastRows = app.calculateForecast({
      stockRows, historicalVentas: historical, bajas: [], existencias: [], realProduction: [],
      selectedMonth: targetMonth, dailyBufferPct: 0,
    });

    const zeroActual = forecastRows
      .map((row) => {
        const p = normalized(row.producto);
        const actual = actualMap.get(p) || 0;
        const forecast = Number(row.pronosticoVenta || 0);
        const last = lastSaleMonth.get(p) || "";
        const gap = last ? historicalMonths.filter((m) => m > last).length : historicalMonths.length;
        return { product: row.producto, actual, forecast, last, gap };
      })
      .filter((row) => row.actual === 0 && row.forecast > 0)
      .sort((a, b) => b.forecast - a.forecast);

    console.log(`\n=== ${targetMonth}: productos con venta real 0 pero pronostico > 0 ===`);
    console.log(`Total: ${zeroActual.length} productos, ${zeroActual.reduce((s, r) => s + r.forecast, 0).toFixed(0)} piezas de error`);
    console.log(`${"Producto".padEnd(32)}${"Pronost".padStart(9)}${"Ult venta".padStart(11)}${"Meses inactivo".padStart(16)}`);
    for (const row of zeroActual.slice(0, 12)) {
      console.log(`${row.product.slice(0, 31).padEnd(32)}${row.forecast.toFixed(0).padStart(9)}${(row.last || "nunca").padStart(11)}${String(row.gap).padStart(16)}`);
    }
    const catchable = zeroActual.filter((r) => r.gap >= 1);
    console.log(`Detectables con regla de inactividad de 1 mes o mas: ${catchable.length} productos, ${catchable.reduce((s, r) => s + r.forecast, 0).toFixed(0)} piezas`);
    const catchable2 = zeroActual.filter((r) => r.gap >= 2);
    console.log(`Detectables con 2 meses o mas: ${catchable2.length} productos, ${catchable2.reduce((s, r) => s + r.forecast, 0).toFixed(0)} piezas`);
  }
}

main().catch((error) => { console.error(error.message); process.exit(1); });
