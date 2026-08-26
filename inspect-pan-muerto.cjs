const fs = require("fs");
const path = require("path");
const Module = require("module");
const esbuild = require("esbuild");
const XLSX = require("xlsx");

const ROOT = __dirname;
const DOWNLOADS = path.resolve(ROOT, "..");
const SOURCES = ["VENTA AÑO 2025 - ANGEL.xlsx", "Venta de diciembre 2024.xlsx"];

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
  const appModule = new Module("pan-muerto");
  appModule.filename = path.join(ROOT, "pan-muerto.bundle.cjs");
  appModule.paths = module.paths;
  appModule._compile(built.outputFiles[0].text, appModule.filename);
  return appModule.exports;
}

async function main() {
  const app = await loadAppFunctions();
  const stockRows = app.parseStock(XLSX.readFile(findDownload("STOCK IDEAL SUCURSALES.xlsx"), { cellDates: true }));
  const catalogPanMuerto = stockRows
    .map((row) => normalized(row.producto))
    .filter((product) => /MUERTO|HUESITOS|CALAVERA|CALABAZA|BRUJA/.test(product));

  const read = (filePath) =>
    app.parseSalesOrReturns(XLSX.readFile(filePath, { cellDates: true }), "ventas", path.basename(filePath));
  const sales = SOURCES.flatMap((name) => read(findDownload(name)));

  const season = new Map();
  for (const row of sales) {
    const month = monthKey(row.fecha);
    if (!/^2025-(09|10|11)$/.test(month) && month !== "2024-12") continue;
    const product = normalized(row.producto);
    if (!/MUERTO|HUESITOS|CALAVERA|CALABAZA|BRUJA|CHOCOLATE/.test(product)) continue;
    const key = `${product}|${month}`;
    season.set(key, (season.get(key) || 0) + Number(row.cantidad || 0));
  }

  console.log("=== CATALOGO: productos de temporada de muertos ===");
  catalogPanMuerto.forEach((product) => console.log(`  ${product}`));

  console.log("\n=== VENTAS de sep-nov 2025 con MUERTO, HUESITOS, CALAVERA, CALABAZA, BRUJA ===");
  const grouped = new Map();
  for (const [key, quantity] of season.entries()) {
    const [product, month] = key.split("|");
    if (!/MUERTO|HUESITOS|CALAVERA|CALABAZA|BRUJA/.test(product)) continue;
    if (!grouped.has(product)) grouped.set(product, new Map());
    grouped.get(product).set(month, quantity);
  }
  const inCatalog = new Set(catalogPanMuerto);
  [...grouped.entries()]
    .sort((a, b) => {
      const totalA = [...a[1].values()].reduce((s, v) => s + v, 0);
      const totalB = [...b[1].values()].reduce((s, v) => s + v, 0);
      return totalB - totalA;
    })
    .forEach(([product, months]) => {
      const total = [...months.values()].reduce((sum, value) => sum + value, 0);
      const detail = [...months.entries()].sort().map(([m, v]) => `${m}:${Math.round(v)}`).join(" ");
      const flag = inCatalog.has(product) ? "en catalogo" : "NO ESTA EN CATALOGO";
      console.log(`  ${product.padEnd(42)}${String(Math.round(total)).padStart(6)}  ${detail.padEnd(24)}  ${flag}`);
    });

  console.log("\n=== Ventas de temporada que mencionan CHOCOLATE ===");
  const chocolate = [...season.entries()].filter(([key]) => /CHOCOLATE/.test(key.split("|")[0]));
  if (!chocolate.length) {
    console.log("  ninguna");
  } else {
    const byProduct = new Map();
    for (const [key, quantity] of chocolate) {
      const [product, month] = key.split("|");
      if (!byProduct.has(product)) byProduct.set(product, new Map());
      byProduct.get(product).set(month, quantity);
    }
    [...byProduct.entries()].forEach(([product, months]) => {
      const detail = [...months.entries()].sort().map(([m, v]) => `${m}:${Math.round(v)}`).join(" ");
      console.log(`  ${product.padEnd(42)}${detail}`);
    });
  }
}

main().catch((error) => { console.error(error.message); process.exit(1); });
