/**
 * Walk-forward 2025. Compara contra post22 (main a050fa9).
 * Catálogo = SKUs de ventas + extras de stock en mapeo-sin-match.
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

const POST22_EXPECTED = {
  cuts: {
    "Ene–Dic": 20.58,
    "Sin enero": 13.96,
    "Mar–Nov": 12.75,
    "Sep–Nov": 9.17,
  },
  monthly: {
    "2025-01": 100,
    "2025-02": 13.39,
    "2025-03": 20.14,
    "2025-04": 12.89,
    "2025-05": 17.16,
    "2025-06": 8.47,
    "2025-07": 18.45,
    "2025-08": 10.43,
    "2025-09": 8.08,
    "2025-10": 10.55,
    "2025-11": 8.82,
    "2025-12": 23.09,
  },
};

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

function salesInputsExist() {
  return fs.existsSync(SALES_JSON) && fs.existsSync(POST22_JSON);
}

async function runPepes2025Backtest() {
  const app = await loadAppFunctions();
  const { filterVentasBeforeMonth, calculateForecast, analyzeForecastProductErrors } = app;
  const salesPayload = JSON.parse(fs.readFileSync(SALES_JSON, "utf8"));
  const post22 = JSON.parse(fs.readFileSync(POST22_JSON, "utf8"));
  const ventas = hydrateSalesRows(salesPayload);
  const stockRows = buildStockRows(ventas);
  const ventas2025 = ventas.filter((row) => String(monthKeyOf(row) || "").startsWith("2025"));

  function runWalkForward(historicalVentas) {
    const monthResults = {};
    const reactivations = [];
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
      const gapRows = analysis.rows.filter((row) => /reactivación estacional/i.test(row.metodo || ""));
      for (const row of gapRows) {
        reactivations.push({
          month: hideMonth,
          producto: row.producto,
          forecast: Number(row.forecast.toFixed(2)),
          actual: Number(row.actual.toFixed(2)),
          method: row.metodo,
        });
      }
      monthResults[hideMonth] = {
        wape: analysis.wape != null ? Number(Number(analysis.wape).toFixed(2)) : null,
        actualTotal: Number(Number(analysis.actual).toFixed(2)),
        forecastTotal: Number(Number(analysis.forecast).toFixed(2)),
        absoluteErrorTotal: Number(Number(absErr).toFixed(2)),
        post22: post ? post.post22 : POST22_EXPECTED.monthly[hideMonth],
        deltaPts: post && analysis.wape != null ? Number((analysis.wape - post.post22).toFixed(2)) : null,
        gapReactivations: gapRows.length,
      };
    }
    return { monthResults, reactivations };
  }

  const with2024 = runWalkForward(ventas);
  const without2024 = runWalkForward(ventas2025);
  for (const hideMonth of TARGET_MONTHS) {
    const row = with2024.monthResults[hideMonth];
    const baseline = without2024.monthResults[hideMonth];
    row.sin2024 = baseline ? baseline.wape : null;
    row.deltaVsSin2024 = baseline && row.wape != null
      ? Number((row.wape - baseline.wape).toFixed(2))
      : null;
  }

  const cutsOf = (monthResults) => ({
    "Ene–Dic": weightedOverall(monthResults, () => true),
    "Sin enero": weightedOverall(monthResults, (m) => m !== "2025-01"),
    "Mar–Nov": weightedOverall(monthResults, (m) => m >= "2025-03" && m <= "2025-11"),
    "Sep–Nov": weightedOverall(monthResults, (m) => ["2025-09", "2025-10", "2025-11"].includes(m)),
  });

  return {
    stockProducts: stockRows.length,
    salesRows: ventas.length,
    cutsSin2024: cutsOf(without2024.monthResults),
    cutsCon2024: cutsOf(with2024.monthResults),
    monthlySin2024: without2024.monthResults,
    monthlyCon2024: with2024.monthResults,
    gapReactivationCon2024: with2024.reactivations,
    gapReactivationSin2024: without2024.reactivations,
  };
}

async function main() {
  const report = await runPepes2025Backtest();
  const outPath = path.join(ROOT, "scripts", "backtest-2025-cold-start-2024.json");
  fs.writeFileSync(outPath, JSON.stringify(report, null, 2), "utf8");
  console.log(JSON.stringify(report, null, 2));
}

module.exports = {
  POST22_EXPECTED,
  TARGET_MONTHS,
  salesInputsExist,
  runPepes2025Backtest,
};

if (require.main === module) {
  main().catch((err) => {
    console.error(err);
    process.exit(1);
  });
}
