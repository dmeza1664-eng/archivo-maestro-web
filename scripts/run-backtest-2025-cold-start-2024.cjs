/**
 * Walk-forward 2025 con historia 2024. Compara contra post22 (main a050fa9).
 * Catálogo = SKUs mapeados + extras de stock en mapeo-sin-match (sin stock_ideal.xlsx).
 */
const fs = require("fs");
const path = require("path");
const Module = require("module");
const esbuild = require("esbuild");

const ROOT = path.join(__dirname, "..");
const UPLOADS = "/home/ubuntu/.cursor/projects/workspace/uploads";
const SALES_JSON = path.join(UPLOADS, "ventas-2024-2025-pepes-para-forecast_01b9.json");
const POST22_JSON = path.join(UPLOADS, "comparacion-vs-post22_645c.json");
const MAPEO_CSV = path.join(UPLOADS, "mapeo-sin-match_effd.csv");

const TARGET_MONTHS = [
  "2025-01", "2025-02", "2025-03", "2025-04", "2025-05", "2025-06",
  "2025-07", "2025-08", "2025-09", "2025-10", "2025-11", "2025-12",
];

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
  const appModule = new Module("backtest-2025-cold-start");
  appModule.filename = path.join(ROOT, "scripts", "backtest-2025-cold-start.bundle.cjs");
  appModule.paths = Module._nodeModulePaths(ROOT);
  appModule._compile(built.outputFiles[0].text, appModule.filename);
  return appModule.exports;
}

function monthKeyOf(row) {
  if (row.mes) return row.mes;
  const d = row.fecha instanceof Date ? row.fecha : new Date(row.fecha);
  if (Number.isNaN(d.getTime())) return null;
  return `${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, "0")}`;
}

function hydrateSalesRows(payload) {
  return (payload.rows || []).map((row) => ({
    ...row,
    fecha: new Date(row.fecha),
    cantidad: Number(row.cantidad) || 0,
    monthlyTotal: true,
    monthDays: row.monthDays || (() => {
      const [y, m] = String(row.mes).split("-").map(Number);
      return new Date(y, m, 0).getDate();
    })(),
  }));
}

function buildStockRows(salesRows) {
  const products = new Set(salesRows.map((row) => row.producto).filter(Boolean));
  if (fs.existsSync(MAPEO_CSV)) {
    const lines = fs.readFileSync(MAPEO_CSV, "utf8").trim().split("\n").slice(1);
    for (const line of lines) {
      const [side, producto] = line.split(",");
      if (side === "stock" && producto) products.add(producto);
    }
  }
  return [...products].sort().map((producto, index) => ({ producto, stock: 20, orden: index + 1 }));
}

function weightedOverall(monthResults, filterFn) {
  let abs = 0;
  let act = 0;
  let fcst = 0;
  const included = [];
  for (const m of TARGET_MONTHS) {
    const x = monthResults[m];
    if (!x || !filterFn(m, x)) continue;
    if (!(x.actualTotal > 0)) continue;
    abs += x.absoluteErrorTotal;
    act += x.actualTotal;
    fcst += x.forecastTotal;
    included.push(m);
  }
  return {
    weightedWapePct: act > 0 ? Number(((abs / act) * 100).toFixed(2)) : null,
    sumAbsoluteError: Number(abs.toFixed(2)),
    sumActual: Number(act.toFixed(2)),
    sumForecast: Number(fcst.toFixed(2)),
    monthsIncluded: included,
  };
}

async function main() {
  const app = await loadAppFunctions();
  const { filterVentasBeforeMonth, calculateForecast, analyzeForecastProductErrors } = app;
  const salesPayload = JSON.parse(fs.readFileSync(SALES_JSON, "utf8"));
  const post22 = JSON.parse(fs.readFileSync(POST22_JSON, "utf8"));
  const ventas = hydrateSalesRows(salesPayload);
  const stockRows = buildStockRows(ventas);

  const ventas2025 = ventas.filter((row) => String(monthKeyOf(row) || "").startsWith("2025"));

  function runWalkForward(historicalVentas) {
    const monthResults = {};
    for (const hideMonth of TARGET_MONTHS) {
      const historical = filterVentasBeforeMonth(historicalVentas, hideMonth);
      const forecastRows = calculateForecast({
        stockRows,
        historicalVentas: historical,
        bajas: [],
        existencias: [],
        realProduction: [],
        selectedMonth: hideMonth,
        dailyBufferPct: 10,
      });
      const actualMap = new Map();
      for (const row of ventas.filter((r) => monthKeyOf(r) === hideMonth)) {
        actualMap.set(row.producto, (actualMap.get(row.producto) || 0) + (Number(row.cantidad) || 0));
      }
      const analysis = analyzeForecastProductErrors(forecastRows, actualMap, { topN: 8 });
      const absErr = analysis.rows.reduce((s, r) => s + r.absoluteError, 0);
      const post = post22.monthly[hideMonth];
      monthResults[hideMonth] = {
        wape: analysis.wape != null ? Number(Number(analysis.wape).toFixed(2)) : null,
        actualTotal: Number(Number(analysis.actual).toFixed(2)),
        forecastTotal: Number(Number(analysis.forecast).toFixed(2)),
        absoluteErrorTotal: Number(Number(absErr).toFixed(2)),
        post22: post ? post.post22 : null,
        deltaPts: post && analysis.wape != null ? Number((analysis.wape - post.post22).toFixed(2)) : null,
      };
    }
    return monthResults;
  }

  const monthResults = runWalkForward(ventas);
  const sameCatalogNo2024 = runWalkForward(ventas2025);
  for (const hideMonth of TARGET_MONTHS) {
    const with2024 = monthResults[hideMonth];
    const without = sameCatalogNo2024[hideMonth];
    with2024.sameCatalogNo2024 = without ? without.wape : null;
    with2024.sameCatalogDeltaPts = without && with2024.wape != null
      ? Number((with2024.wape - without.wape).toFixed(2))
      : null;
  }

  const cuts = {
    "Ene–Dic": weightedOverall(monthResults, () => true),
    "Sin enero": weightedOverall(monthResults, (m) => m !== "2025-01"),
    "Mar–Nov": weightedOverall(monthResults, (m) => m >= "2025-03" && m <= "2025-11"),
    "Sep–Nov": weightedOverall(monthResults, (m) => ["2025-09", "2025-10", "2025-11"].includes(m)),
  };
  const sameCatalogCuts = {
    "Ene–Dic": weightedOverall(sameCatalogNo2024, () => true),
    "Sin enero": weightedOverall(sameCatalogNo2024, (m) => m !== "2025-01"),
    "Mar–Nov": weightedOverall(sameCatalogNo2024, (m) => m >= "2025-03" && m <= "2025-11"),
    "Sep–Nov": weightedOverall(sameCatalogNo2024, (m) => ["2025-09", "2025-10", "2025-11"].includes(m)),
  };

  const report = {
    stockProducts: stockRows.length,
    salesRows: ventas.length,
    cuts,
    sameCatalogNo2024Cuts: sameCatalogCuts,
    monthly: monthResults,
  };
  const outPath = path.join(ROOT, "scripts", "backtest-2025-cold-start-2024.json");
  fs.writeFileSync(outPath, JSON.stringify(report, null, 2), "utf8");
  console.log(JSON.stringify(report, null, 2));
}

main().catch((err) => {
  console.error(err);
  process.exit(1);
});
