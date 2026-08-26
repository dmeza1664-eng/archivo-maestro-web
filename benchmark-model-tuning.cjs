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
const VERSIONS = ["legacy", "rolling", "seasonal25", "seasonal50", "categorySeasonal", "seasonal75", "seasonal100", "seasonalAdaptive"];

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
    define: { "import.meta.env.VITE_API_URL": JSON.stringify("") },
    logLevel: "silent",
  });
  const appModule = new Module("model-tuning");
  appModule.filename = path.join(ROOT, "model-tuning.bundle.cjs");
  appModule.paths = module.paths;
  appModule._compile(built.outputFiles[0].text, appModule.filename);
  return appModule.exports;
}

async function main() {
  const app = await loadAppFunctions();
  const stockRows = app.parseStock(XLSX.readFile(findDownload("STOCK IDEAL SUCURSALES.xlsx"), { cellDates: true }));

  const read = (filePath) =>
    app.parseSalesOrReturns(XLSX.readFile(filePath, { cellDates: true }), "ventas", path.basename(filePath));

  // Se usa el archivo diario de junio para conservar la forma por dia de semana.
  // Con INCLUIR_CIERRE_JUNIO=1 se agrega tambien el cierre mensual de junio, que trae
  // total declarado sin fechas, para probar la reconciliacion de forma y nivel.
  // La venta real siempre se mide con una sola fuente por mes para no duplicarla.
  const measuredSales = [...SOURCES.flatMap((name) => read(findDownload(name))), ...read(JULY_SALES)];
  const allSales = [
    ...measuredSales,
    ...(process.env.INCLUIR_CIERRE_JUNIO === "1"
      ? read("C:/Users/X13/Documents/ventas por mes/ventas junio.xlsx")
      : []),
  ];

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
      const forecast = rows.reduce((sum, row) => sum + row.forecast, 0);
      const absoluteError = rows.reduce((sum, row) => sum + Math.abs(row.actual - row.forecast), 0);
      if (!results.has(version)) results.set(version, new Map());
      results.get(version).set(targetMonth, {
        actual,
        forecast,
        absoluteError,
        wape: actual > 0 ? (absoluteError / actual) * 100 : null,
        inside: rows.filter((row) => Math.abs(row.actual - row.forecast) <= 15).length,
        products: rows.length,
      });
    }
  }

  const pad = (value, width) => String(value).padStart(width);
  console.log("=== WAPE POR VERSION DEL MODELO (menor es mejor) ===\n");
  console.log(`${"Version".padEnd(18)}${TARGET_MONTHS.map((m) => pad(m.slice(5) + "/26", 9)).join("")}${pad("Promedio", 11)}${pad("Ponderado", 11)}`);
  console.log("-".repeat(18 + 9 * TARGET_MONTHS.length + 22));

  const summary = [];
  for (const version of VERSIONS) {
    const perMonth = results.get(version);
    const wapes = TARGET_MONTHS.map((m) => perMonth.get(m).wape);
    const average = wapes.reduce((sum, value) => sum + value, 0) / wapes.length;
    const totalActual = TARGET_MONTHS.reduce((sum, m) => sum + perMonth.get(m).actual, 0);
    const totalError = TARGET_MONTHS.reduce((sum, m) => sum + perMonth.get(m).absoluteError, 0);
    const weighted = (totalError / totalActual) * 100;
    summary.push({ version, average, weighted, wapes });
    const label = version === "categorySeasonal" ? "categorySeasonal*" : version;
    console.log(
      `${label.padEnd(18)}${wapes.map((w) => pad(w.toFixed(2), 9)).join("")}${pad(average.toFixed(2), 11)}${pad(weighted.toFixed(2), 11)}`
    );
  }
  console.log("\n* configuracion actual en produccion (75% estacional, 50% en Otros y Mini medianos)");

  const best = [...summary].sort((a, b) => a.weighted - b.weighted)[0];
  const current = summary.find((row) => row.version === "categorySeasonal");
  console.log(`\nMejor ponderado: ${best.version} con ${best.weighted.toFixed(2)}%`);
  console.log(`Actual: categorySeasonal con ${current.weighted.toFixed(2)}%`);
  console.log(`Diferencia: ${(current.weighted - best.weighted).toFixed(2)} puntos`);

  console.log("\n=== JULIO EN DETALLE ===");
  for (const version of VERSIONS) {
    const july = results.get(version).get("2026-07");
    console.log(
      `${version.padEnd(18)} pronostico ${pad(july.forecast.toFixed(0), 7)}  real ${july.actual}  dif ${pad((july.actual - july.forecast).toFixed(0), 7)}  WAPE ${july.wape.toFixed(2)}%  dentro15 ${july.inside}`
    );
  }

  console.log("\n=== DIRECCION DEL SESGO POR MES (real - pronostico, configuracion actual) ===");
  for (const month of TARGET_MONTHS) {
    const row = results.get("categorySeasonal").get(month);
    const bias = row.actual - row.forecast;
    console.log(`${month}: ${bias > 0 ? "quedo corto" : "se paso"} por ${Math.abs(bias).toFixed(0)} piezas (${(bias / row.actual * 100).toFixed(1)}%)`);
  }

  // Separa el error de estatus de producto del error de estimacion de cantidad.
  console.log("\n=== DE DONDE VIENE EL ERROR (configuracion actual) ===");
  console.log(`${"Mes".padEnd(10)}${pad("Error total", 13)}${pad("Descontinuado", 15)}${pad("Nuevo", 9)}${pad("Estimacion", 13)}${pad("% evitable", 12)}`);
  for (const month of TARGET_MONTHS) {
    const historical = app.filterVentasBeforeMonth(allSales, month);
    const actualMap = new Map();
    for (const row of measuredSales.filter((sale) => monthKey(sale.fecha) === month)) {
      const product = normalized(row.producto);
      actualMap.set(product, (actualMap.get(product) || 0) + Number(row.cantidad || 0));
    }
    const forecastRows = app.calculateForecast({
      stockRows,
      historicalVentas: historical,
      bajas: [],
      existencias: [],
      realProduction: [],
      selectedMonth: month,
      dailyBufferPct: 0,
      modelVersion: "categorySeasonal",
    });

    let discontinued = 0;
    let brandNew = 0;
    let estimation = 0;
    for (const row of forecastRows) {
      const actual = actualMap.get(normalized(row.producto)) || 0;
      const forecast = Number(row.pronosticoVenta || 0);
      const error = Math.abs(actual - forecast);
      if (actual === 0 && forecast > 0) discontinued += error;
      else if (actual > 0 && forecast < 1) brandNew += error;
      else estimation += error;
    }
    const total = discontinued + brandNew + estimation;
    console.log(
      `${month.padEnd(10)}${pad(total.toFixed(0), 13)}${pad(discontinued.toFixed(0), 15)}${pad(brandNew.toFixed(0), 9)}${pad(estimation.toFixed(0), 13)}${pad(((discontinued + brandNew) / total * 100).toFixed(1) + "%", 12)}`
    );
  }
  console.log("\nDescontinuado: el producto no vendio nada pero el modelo si lo pronostico.");
  console.log("Nuevo: el producto vendio pero no tenia historico para pronosticarlo.");
  console.log("Evitable con un estatus operativo capturado antes del mes.");
}

main().catch((error) => {
  console.error(error.message);
  process.exit(1);
});
