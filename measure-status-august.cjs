const fs = require("fs");
const path = require("path");
const Module = require("module");
const esbuild = require("esbuild");
const XLSX = require("xlsx");

const ROOT = __dirname;
const DOWNLOADS = path.resolve(ROOT, "..");
const TARGET = "2026-08";
const SOURCES = [
  "Venta de diciembre 2024.xlsx",
  "VENTA AÑO 2025 - ANGEL.xlsx",
  "VENTA ENERO 2026.xlsx",
  "VENTAS FEEEBRERO 2026.xlsx",
  "MARZO VENTAS.xlsx",
  "venta abril 2026.xlsx",
  "VENTAS DE MAYO Y JUNIO - ANGEL.xlsx",
];

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

async function loadAppFunctions() {
  const built = await esbuild.build({
    entryPoints: [path.join(ROOT, "App.jsx")],
    bundle: true, platform: "node", format: "cjs", write: false,
    loader: { ".css": "text" },
    define: { "import.meta.env.VITE_API_URL": JSON.stringify("") },
    logLevel: "silent",
  });
  const appModule = new Module("august-status");
  appModule.filename = path.join(ROOT, "august-status.bundle.cjs");
  appModule.paths = module.paths;
  appModule._compile(built.outputFiles[0].text, appModule.filename);
  return appModule.exports;
}

async function main() {
  const app = await loadAppFunctions();
  const statusSheet = XLSX.readFile(path.join(DOWNLOADS, "ESTATUS_PRODUCTOS_PARA_LLENAR.xlsx"))
    .Sheets["Estatus productos"];
  const status = new Map(
    XLSX.utils.sheet_to_json(statusSheet, { defval: "" })
      .map((row) => [normalized(row.Producto), String(row.Estatus || "").trim().toUpperCase()])
  );

  const stockRows = app.parseStock(XLSX.readFile(findDownload("STOCK IDEAL SUCURSALES.xlsx"), { cellDates: true }));
  const read = (filePath) =>
    app.parseSalesOrReturns(XLSX.readFile(filePath, { cellDates: true }), "ventas", path.basename(filePath));
  const sales = [
    ...SOURCES.flatMap((name) => read(findDownload(name))),
    ...read("C:/Users/X13/Documents/ventas por mes/ventas julio.xlsx"),
    ...read("C:/Users/X13/Documents/ventas por mes/ventas junio.xlsx"),
  ];

  const forecastRows = app.calculateForecast({
    stockRows,
    historicalVentas: app.filterVentasBeforeMonth(sales, TARGET),
    bajas: [], existencias: [], realProduction: [],
    selectedMonth: TARGET,
    dailyBufferPct: 0,
  });

  let total = 0;
  let suppressedBaja = 0;
  let suppressedSeasonal = 0;
  const suppressed = [];
  for (const row of forecastRows) {
    const product = normalized(row.producto);
    const forecast = Number(row.pronosticoVenta || 0);
    const estatus = status.get(product) || "ACTIVO";
    total += forecast;
    if (estatus === "BAJA") {
      suppressedBaja += forecast;
      if (forecast > 0.5) suppressed.push({ product, forecast, estatus });
    }
    if (estatus === "ESTACIONAL") {
      suppressedSeasonal += forecast;
      if (forecast > 0.5) suppressed.push({ product, forecast, estatus });
    }
  }

  console.log(`Pronostico de agosto sin aplicar estatus: ${total.toFixed(0)} piezas`);
  console.log(`Piezas asignadas a productos marcados BAJA: ${suppressedBaja.toFixed(0)}`);
  console.log(`Piezas asignadas a productos marcados ESTACIONAL: ${suppressedSeasonal.toFixed(0)}`);
  console.log(`Pronostico de agosto con estatus aplicado: ${(total - suppressedBaja - suppressedSeasonal).toFixed(0)} piezas`);

  console.log("\nProductos suprimidos que todavia recibian pronostico:");
  if (!suppressed.length) {
    console.log("  ninguno: el modelo ya les asignaba cero");
  } else {
    suppressed.sort((a, b) => b.forecast - a.forecast)
      .forEach((row) => console.log(`  ${row.product.padEnd(36)}${row.forecast.toFixed(1).padStart(8)}  ${row.estatus}`));
  }
}

main().catch((error) => { console.error(error.message); process.exit(1); });
