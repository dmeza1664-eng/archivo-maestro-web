/**
 * Real WAPE backtest — Pastelería Pepes.
 * Prefers keeping daily + monthly close together when both exist
 * (buildMonthlyForecastData scales daily shape to close total).
 */
const fs = require("fs");
const path = require("path");
const Module = require("module");
const esbuild = require("esbuild");
const XLSX = require("xlsx");

const EXCEL_DIR = process.env.WAPE_EXCEL_DIR || path.join(__dirname, "..", "wape-excel");
const APP_ROOT = path.join(__dirname, "..");

async function loadAppFunctions() {
  const built = await esbuild.build({
    entryPoints: [path.join(APP_ROOT, "App.jsx")],
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
  const appModule = new Module("real-wape");
  appModule.filename = path.join(EXCEL_DIR, "real-wape.bundle.cjs");
  appModule.paths = Module._nodeModulePaths(APP_ROOT);
  appModule._compile(built.outputFiles[0].text, appModule.filename);
  return appModule.exports;
}

function monthKeyOf(row) {
  const d = row.fecha instanceof Date ? row.fecha : new Date(row.fecha);
  if (Number.isNaN(d.getTime())) return null;
  return `${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, "0")}`;
}

function readWorkbook(fileName) {
  return XLSX.readFile(path.join(EXCEL_DIR, fileName), { cellDates: true });
}

function sourceKind(rows, month) {
  const subset = rows.filter((r) => monthKeyOf(r) === month);
  const hasDaily = subset.some((r) => !r.monthlyTotal);
  const hasClose = subset.some((r) => r.monthlyTotal);
  if (hasDaily && hasClose) return "both";
  if (hasDaily) return "daily";
  if (hasClose) return "close";
  return "empty";
}

function resolvePreferDailyPlusClose(entries) {
  const names = entries.map((e) => e.name);
  const byName = new Map(entries.map((e) => [e.name, { name: e.name, rows: [...(e.rows || [])], file: e.file }]));
  const months = new Set();
  for (const entry of byName.values()) {
    for (const row of entry.rows) {
      const key = monthKeyOf(row);
      if (key) months.add(key);
    }
  }
  const decisions = [];
  for (const month of [...months].sort()) {
    const candidates = names.filter((name) => byName.get(name).rows.some((row) => monthKeyOf(row) === month));
    if (candidates.length < 2) continue;
    const kinds = Object.fromEntries(candidates.map((name) => [name, sourceKind(byName.get(name).rows, month)]));
    const hasDailySrc = candidates.some((n) => kinds[n] === "daily" || kinds[n] === "both");
    const hasCloseSrc = candidates.some((n) => kinds[n] === "close" || kinds[n] === "both");

    if (hasDailySrc && hasCloseSrc) {
      const dailyKeep = candidates.find((n) => kinds[n] === "daily" || kinds[n] === "both");
      const closeKeep = candidates
        .filter((n) => kinds[n] === "close" || kinds[n] === "both")
        .sort((a, b) => {
          const score = (name) =>
            (/ventas (mayo|junio|julio|agosto|octubre|noviembre)/i.test(name) ? 2 : 0) +
            (/año|anio|2025\.xlsx/i.test(name) ? -1 : 0);
          return score(b) - score(a);
        })[0];
      const keep = new Set([dailyKeep, closeKeep].filter(Boolean));
      const omitted = [];
      for (const name of candidates) {
        if (keep.has(name)) continue;
        const entry = byName.get(name);
        entry.rows = entry.rows.filter((row) => monthKeyOf(row) !== month);
        omitted.push(name);
      }
      decisions.push({ month, strategy: "keep-daily-and-close", kept: [...keep], omitted, kinds });
      continue;
    }

    const winner = [...candidates].sort((a, b) => {
      const dedicated = (n) => (/^ventas (mayo|junio|julio|agosto|octubre|noviembre)/i.test(n) ? 1 : 0);
      return dedicated(b) - dedicated(a) || names.indexOf(b) - names.indexOf(a);
    })[0];
    const omitted = candidates.filter((n) => n !== winner);
    for (const name of omitted) {
      const entry = byName.get(name);
      entry.rows = entry.rows.filter((row) => monthKeyOf(row) !== month);
    }
    decisions.push({ month, strategy: "same-kind-winner", winner, omitted, kinds });
  }
  return { entries: names.map((n) => byName.get(n)), decisions };
}

function summarizeSource(name, rows) {
  const byMonth = {};
  let daily = 0;
  let monthly = 0;
  let qty = 0;
  for (const row of rows) {
    const key = monthKeyOf(row);
    if (!key) continue;
    if (!byMonth[key]) byMonth[key] = { rows: 0, qty: 0, daily: 0, monthly: 0 };
    byMonth[key].rows += 1;
    byMonth[key].qty += Number(row.cantidad) || 0;
    if (row.monthlyTotal) {
      monthly += 1;
      byMonth[key].monthly += 1;
    } else {
      daily += 1;
      byMonth[key].daily += 1;
    }
    qty += Number(row.cantidad) || 0;
  }
  return { name, rowCount: rows.length, daily, monthly, totalQty: Math.round(qty), byMonth };
}

function compactTopError(row) {
  return {
    producto: row.producto,
    forecast: Number(Number(row.forecast).toFixed(2)),
    actual: Number(Number(row.actual).toFixed(2)),
    error: Number(Number(row.error).toFixed(2)),
    absoluteError: Number(Number(row.absoluteError).toFixed(2)),
    errorShare: row.errorShare != null ? Number(Number(row.errorShare).toFixed(4)) : null,
    apePct: row.actual > 0 ? Number(((row.absoluteError / row.actual) * 100).toFixed(2)) : null,
  };
}

function gdeTop(rows, n = 15) {
  return rows
    .filter((r) => /GDE$/i.test(String(r.producto || "")))
    .sort((a, b) => b.absoluteError - a.absoluteError)
    .slice(0, n)
    .map(compactTopError);
}

function gdeFocus(rows, names) {
  const wanted = new Set(names);
  return rows
    .filter((r) => wanted.has(r.producto))
    .map(compactTopError);
}

async function main() {
  const app = await loadAppFunctions();
  const {
    parseStock,
    parseSalesOrReturns,
    resolveCanonicalMonthSources,
    filterVentasBeforeMonth,
    calculateForecast,
    analyzeForecastProductErrors,
    buildForecastHealth,
    buildSalesMonthCoverage,
    monthsNamedInFileName,
    computeAnnualGrowthFactor,
    buildMonthlyForecastData,
  } = app;

  const stockRows = parseStock(readWorkbook("stock_ideal.xlsx"));

  const salesSources = [
    { file: "venta_anio_2025.xlsx", uploadName: "VENTA AÑO 2025.xlsx" },
    { file: "ventas_octubre_2025.xlsx", uploadName: "ventas octubre 2025.xlsx" },
    { file: "ventas_noviembre_2025.xlsx", uploadName: "ventas noviembre 2025.xlsx" },
    { file: "ventas_mayo_junio_angel.xlsx", uploadName: "VENTAS DE MAYO Y JUNIO 2026.xlsx" },
    { file: "ventas_mayo.xlsx", uploadName: "ventas mayo.xlsx" },
    { file: "ventas_junio.xlsx", uploadName: "ventas junio.xlsx" },
    { file: "ventas_julio.xlsx", uploadName: "ventas julio.xlsx" },
    { file: "ventas_agosto.xlsx", uploadName: "ventas agosto.xlsx" },
  ];

  const parseNotes = [];
  const entries = [];
  for (const src of salesSources) {
    const full = path.join(EXCEL_DIR, src.file);
    if (!fs.existsSync(full)) {
      parseNotes.push({ file: src.file, ok: false, reason: "archivo no encontrado" });
      continue;
    }
    try {
      const wb = readWorkbook(src.file);
      const rows = parseSalesOrReturns(wb, "ventas", src.uploadName);
      const namedMonths = monthsNamedInFileName(src.uploadName);
      if (!rows.length) {
        parseNotes.push({ file: src.file, uploadName: src.uploadName, ok: false, reason: "0 filas", namedMonths });
        continue;
      }
      parseNotes.push({ file: src.file, uploadName: src.uploadName, ok: true, namedMonths, ...summarizeSource(src.uploadName, rows) });
      entries.push({ name: src.uploadName, rows, file: src.file });
    } catch (err) {
      parseNotes.push({ file: src.file, uploadName: src.uploadName, ok: false, reason: String(err.message || err) });
    }
  }

  const appDefault = resolveCanonicalMonthSources(entries.map((e) => ({ name: e.name, rows: e.rows })));
  const preferred = resolvePreferDailyPlusClose(entries);
  const ventas = preferred.entries.flatMap((e) => e.rows);
  const coverage = buildSalesMonthCoverage(ventas);
  const ventasAppDefault = appDefault.entries.flatMap((e) => e.rows);

  const focusProducts = ["FRUTAS GDE", "MOKA GDE", "M & M GDE", "CHOCOLATE GDE", "NUTELA GDE", "DURAZNO GDE", "DUO DURAZNO GDE", "PAY DE FRESA GDE"];

  function runBacktests(ventasSet) {
    const hideMonths = ["2026-06", "2026-07", "2026-08"];
    const out = {};
    for (const hideMonth of hideMonths) {
      const historical = filterVentasBeforeMonth(ventasSet, hideMonth);
      const forecastRows = calculateForecast({
        stockRows,
        historicalVentas: historical,
        bajas: [],
        existencias: [],
        realProduction: [],
        selectedMonth: hideMonth,
        dailyBufferPct: 10,
      });
      const monthRows = ventasSet.filter((row) => monthKeyOf(row) === hideMonth);
      const hasClose = monthRows.some((row) => row.monthlyTotal);
      const actualSource = hasClose ? monthRows.filter((row) => row.monthlyTotal) : monthRows.filter((row) => !row.monthlyTotal);
      const actualMap = new Map();
      for (const row of actualSource) {
        const p = row.producto;
        actualMap.set(p, (actualMap.get(p) || 0) + (Number(row.cantidad) || 0));
      }
      const analysis = analyzeForecastProductErrors(forecastRows, actualMap, { topN: 15 });
      const histMonths = [...new Set(historical.map(monthKeyOf).filter(Boolean))].sort();
      const gdeRows = analysis.rows.filter((r) => /GDE$/i.test(String(r.producto || "")));
      const gdeAbs = gdeRows.reduce((s, r) => s + r.absoluteError, 0);
      const gdeActual = gdeRows.reduce((s, r) => s + r.actual, 0);
      out[hideMonth] = {
        wape: analysis.wape != null ? Number(Number(analysis.wape).toFixed(2)) : null,
        mae: analysis.mae != null ? Number(Number(analysis.mae).toFixed(2)) : null,
        inside15: analysis.inside15,
        productCount: analysis.products ?? analysis.rows.length,
        actualTotal: Number(Number(analysis.actual).toFixed(2)),
        forecastTotal: Number(Number(analysis.forecast).toFixed(2)),
        absoluteErrorTotal: Number(Number(analysis.rows.reduce((s, r) => s + r.absoluteError, 0)).toFixed(2)),
        gdeWape: gdeActual > 0 ? Number(((gdeAbs / gdeActual) * 100).toFixed(2)) : null,
        gdeAbsError: Number(gdeAbs.toFixed(2)),
        gdeForecast: Number(gdeRows.reduce((s, r) => s + r.forecast, 0).toFixed(2)),
        gdeActual: Number(gdeActual.toFixed(2)),
        actualProductsWithSales: [...actualMap.values()].filter((q) => q > 0).length,
        historicalMonthCount: histMonths.length,
        historicalMonths: histMonths,
        topErrors: (analysis.topErrors || []).map(compactTopError),
        topGdeErrors: gdeTop(analysis.rows, 15),
        focusGde: gdeFocus(analysis.rows, focusProducts),
      };
    }
    return out;
  }

  const monthResults = runBacktests(ventas);
  const monthResultsAppDefault = runBacktests(ventasAppDefault);

  const diagnosis = {};
  for (const product of focusProducts) {
    const productRows = ventas.filter((row) => row.producto === product);
    const monthly = {};
    for (const row of productRows) {
      const key = monthKeyOf(row);
      if (!key) continue;
      if (!monthly[key]) monthly[key] = { daily: 0, close: 0 };
      if (row.monthlyTotal) monthly[key].close += Number(row.cantidad) || 0;
      else monthly[key].daily += Number(row.cantidad) || 0;
    }
    const closeOrDaily = {};
    for (const [key, val] of Object.entries(monthly)) {
      closeOrDaily[key] = val.close > 0 ? val.close : val.daily;
    }
    const mapForGrowth = new Map();
    for (const [key, total] of Object.entries(closeOrDaily)) {
      mapForGrowth.set(key, { total });
    }
    diagnosis[product] = {
      monthly: closeOrDaily,
      growthJuly: Number(computeAnnualGrowthFactor(mapForGrowth, "2026-07", 3).toFixed(4)),
      growthAugust: Number(computeAnnualGrowthFactor(mapForGrowth, "2026-08", 3).toFixed(4)),
      yoy: {
        "2026-05": closeOrDaily["2025-05"] ? Number((closeOrDaily["2026-05"] / closeOrDaily["2025-05"]).toFixed(3)) : null,
        "2026-06": closeOrDaily["2025-06"] ? Number((closeOrDaily["2026-06"] / closeOrDaily["2025-06"]).toFixed(3)) : null,
        "2026-07": closeOrDaily["2025-07"] ? Number((closeOrDaily["2026-07"] / closeOrDaily["2025-07"]).toFixed(3)) : null,
        "2026-08": closeOrDaily["2025-08"] ? Number((closeOrDaily["2026-08"] / closeOrDaily["2025-08"]).toFixed(3)) : null,
      },
    };
  }

  let health = null;
  try {
    const h = buildForecastHealth({
      stockRows,
      ventas: ventasAppDefault,
      selectedMonth: "2026-09",
      dailyBufferPct: 10,
    });
    health = {
      ready: h.ready,
      currentTotal: h.currentTotal != null ? Number(Number(h.currentTotal).toFixed(2)) : null,
      checks: (h.checks || []).map((c) => ({ code: c.code, level: c.level, message: c.message })),
      backtests: (h.backtests || []).map((b) => ({
        month: b.month,
        wape: b.wape != null ? Number(Number(b.wape).toFixed(2)) : null,
        topErrors: (b.topErrors || []).slice(0, 8).map(compactTopError),
      })),
    };
  } catch (err) {
    health = { error: String(err.message || err) };
  }

  const presentMonths = new Set((coverage || []).map((c) => c.monthKey).filter(Boolean));
  const expected = [];
  for (let y = 2025; y <= 2026; y++) {
    for (let m = 1; m <= (y === 2026 ? 8 : 12); m++) expected.push(`${y}-${String(m).padStart(2, "0")}`);
  }
  const missingMonths = expected.filter((m) => !presentMonths.has(m));

  const report = {
    generatedAt: new Date().toISOString(),
    timezoneNote: "America/Mazatlan (UTC-7) on box",
    model: { dailyBufferPct: 10, bajas: [], existencias: [], realProduction: [] },
    stock: {
      products: stockRows.length,
      sample: stockRows.slice(0, 5).map((r) => ({ producto: r.producto, stock: r.stock })),
    },
    sources: parseNotes,
    resolution: {
      preferredStrategy: "keep complementary daily + monthly close",
      preferredDecisions: preferred.decisions,
      appDefaultDecisions: appDefault.decisions,
    },
    coverage,
    dataGaps: { missingMonths },
    diagnosis,
    backtests: monthResults,
    backtestsAppDefaultCanonical: monthResultsAppDefault,
    forecastHealth2026_09: health,
  };

  if (process.env.WRITE_REPORT === "1") {
    fs.writeFileSync(path.join(EXCEL_DIR, "wape_report.json"), JSON.stringify(report, null, 2), "utf8");
  }

  const hideMonths = ["2026-06", "2026-07", "2026-08"];
  console.log(
    JSON.stringify(
      {
        preferred: Object.fromEntries(
          hideMonths.map((m) => [
            m,
            {
              wape: monthResults[m].wape,
              gdeWape: monthResults[m].gdeWape,
              gdeAbs: monthResults[m].gdeAbsError,
              forecast: monthResults[m].forecastTotal,
              actual: monthResults[m].actualTotal,
            },
          ])
        ),
        appDefault: Object.fromEntries(
          hideMonths.map((m) => [
            m,
            { wape: monthResultsAppDefault[m].wape, gdeWape: monthResultsAppDefault[m].gdeWape },
          ])
        ),
        julyFocus: monthResults["2026-07"].focusGde,
        juneFocus: monthResults["2026-06"].focusGde,
        augustFocus: monthResults["2026-08"].focusGde,
        diagnosis,
        missingMonths,
      },
      null,
      2
    )
  );
}

main().catch((err) => {
  console.error(err);
  process.exit(1);
});
