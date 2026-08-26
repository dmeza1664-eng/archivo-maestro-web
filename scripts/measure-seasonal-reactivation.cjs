// Mide si los productos de temporada (pan de muerto, calabaza) pueden entrar al
// pronostico. Responde una pregunta concreta antes de congelar septiembre y
// planear octubre: el modelo puede proponer piezas para ellos, o no existen.
//
// Uso: node scripts/measure-seasonal-reactivation.cjs

const fs = require("fs");
const path = require("path");
const Module = require("module");
const esbuild = require("esbuild");
const XLSX = require("xlsx");

const ROOT = path.resolve(__dirname, "..");
const DOWNLOADS = path.resolve(ROOT, "..");
const MONTHLY_SALES = "C:/Users/X13/Documents/ventas por mes";
const SEASON_MONTHS = ["2025-09", "2025-10", "2025-11"];
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
  const appModule = new Module("seasonal-reactivation");
  appModule.filename = path.join(ROOT, "seasonal-reactivation.bundle.cjs");
  appModule.paths = module.paths;
  appModule._compile(built.outputFiles[0].text, appModule.filename);
  return appModule.exports;
}

async function main() {
  const app = await loadAppFunctions();

  const statusRows = XLSX.utils.sheet_to_json(
    XLSX.readFile(findDownload("ESTATUS_PRODUCTOS_PARA_LLENAR.xlsx")).Sheets["Estatus productos"],
    { defval: "" }
  );
  const status = new Map(
    statusRows.map((row) => [normalized(row.Producto), String(row.Estatus || "").trim().toUpperCase()])
  );

  const stockWorkbook = XLSX.readFile(findDownload("STOCK IDEAL SUCURSALES.xlsx"), { cellDates: true });
  const stockRows = app.parseStock(stockWorkbook);
  const catalog = stockRows.map((row) => normalized(row.producto));
  const inCatalog = new Set(catalog);

  const sheetNotice = app.assessStockSheetSelection(stockWorkbook);
  console.log("=== Hoja de la que sale el catalogo ===");
  console.log(sheetNotice.message || `Hoja "${sheetNotice.chosenSheet}": ninguna otra hoja candidata aporta productos distintos.`);
  for (const alternative of sheetNotice.alternatives) {
    console.log(`  "${alternative.sheet}" tiene ${alternative.missing.length} productos que la elegida no incluye.`);
  }
  console.log("");

  const read = (filePath) =>
    app.parseSalesOrReturns(XLSX.readFile(filePath, { cellDates: true }), "ventas", path.basename(filePath));
  const sales = [
    ...SOURCES.flatMap((name) => read(findDownload(name))),
    ...read(path.join(MONTHLY_SALES, "ventas junio.xlsx")),
    ...read(path.join(MONTHLY_SALES, "ventas julio.xlsx")),
  ];

  const sold = new Map();
  for (const row of sales) {
    const key = `${normalized(row.producto)}|${monthKey(row.fecha)}`;
    sold.set(key, (sold.get(key) || 0) + Number(row.cantidad || 0));
  }

  console.log("=== Cruce entre catalogo y captura de estatus ===");
  console.log(`Catalogo (stock ideal): ${catalog.length} productos`);
  console.log(`Archivo de estatus: ${status.size} productos`);
  console.log(`Cruzan por nombre: ${catalog.filter((p) => status.has(p)).length}`);
  console.log(`Del catalogo sin estatus: ${catalog.filter((p) => !status.has(p)).length}`);
  console.log(`Con estatus pero fuera del catalogo: ${[...status.keys()].filter((p) => !inCatalog.has(p)).length}`);

  const seasonal = [...status.entries()].filter(([, value]) => value === "ESTACIONAL").map(([product]) => product);
  console.log(`\n=== Los ${seasonal.length} productos marcados ESTACIONAL ===`);
  console.log(`  ${"Producto".padEnd(42)}${SEASON_MONTHS.map((m) => m.padStart(10)).join("")}${"total".padStart(10)}   en catalogo`);

  let seasonTotal = 0;
  let orphanTotal = 0;
  for (const product of seasonal) {
    const values = SEASON_MONTHS.map((month) => sold.get(`${product}|${month}`) || 0);
    const total = values.reduce((sum, value) => sum + value, 0);
    seasonTotal += total;
    if (!inCatalog.has(product)) orphanTotal += total;
    console.log(`  ${product.padEnd(42)}${values.map((v) => v.toFixed(0).padStart(10)).join("")}${total.toFixed(0).padStart(10)}   ${inCatalog.has(product)}`);
  }
  console.log(`  ${"TOTAL".padEnd(42)}${" ".repeat(10 * SEASON_MONTHS.length)}${seasonTotal.toFixed(0).padStart(10)}`);

  console.log(`\nPiezas de la temporada pasada en productos que NO existen en el stock ideal: ${orphanTotal.toFixed(0)}`);
  if (orphanTotal > 0) {
    console.log("El modelo no puede pronosticarlos: calculateForecast solo emite renglones");
    console.log("para productos del stock ideal. No es un problema de pesos ni de estatus.");
  }
}

main().catch((error) => { console.error(error); process.exit(1); });
