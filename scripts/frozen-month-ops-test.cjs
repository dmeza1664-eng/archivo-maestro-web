const path = require("path");
const Module = require("module");
const esbuild = require("esbuild");

async function loadAppFunctions() {
  const built = await esbuild.build({
    entryPoints: [path.join(__dirname, "..", "App.jsx")],
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
  const appModule = new Module("frozen-month-ops-test");
  appModule.filename = path.join(__dirname, "frozen-month-ops-test.bundle.cjs");
  appModule.paths = module.paths;
  appModule._compile(built.outputFiles[0].text, appModule.filename);
  return appModule.exports;
}

function assert(condition, message) {
  if (!condition) throw new Error(message);
}

function almostEqual(actual, expected, tolerance, message) {
  if (Math.abs(actual - expected) > tolerance) {
    throw new Error(`${message} (${actual} vs ${expected})`);
  }
}

function monthClose(monthKey, product, quantity) {
  const [year, month] = monthKey.split("-").map(Number);
  return {
    fecha: new Date(year, month - 1, 1),
    producto: product,
    cantidad: quantity,
    monthlyTotal: true,
    monthDays: new Date(year, month, 0).getDate(),
  };
}

function forecastTotal(rows) {
  return rows.reduce((sum, row) => sum + Number(row.pronosticoVenta || 0), 0);
}

function weekdayShape(row) {
  return [
    row.promedioLunes,
    row.promedioMartes,
    row.promedioMiercoles,
    row.promedioJueves,
    row.promedioViernes,
    row.promedioSabado,
    row.promedioDomingo,
  ].map((value) => Number(value || 0));
}

async function main() {
  const app = await loadAppFunctions();
  const {
    buildMonthlyCloseSummary,
    calculateDailyForecast,
    calculateForecast,
    hydrateForecastFromOperationalRows,
    resolveEffectiveForecast,
    snapshotForecastRowsForFreeze,
    summarizeForecastAccuracy,
    weightedWapeFromBacktests,
    applyPromoUpliftToQuantity,
    applyInventoryToProductionSuggestion,
    normalizeActivePromo,
  } = app;

  const stockRows = [
    { producto: "FRUTAS GDE", stock: 40, orden: 1 },
    { producto: "MOKA GDE", stock: 40, orden: 2 },
  ];
  const history = [
    monthClose("2025-08", "FRUTAS GDE", 500),
    monthClose("2025-08", "MOKA GDE", 620),
    monthClose("2026-06", "FRUTAS GDE", 510),
    monthClose("2026-06", "MOKA GDE", 630),
    monthClose("2026-07", "FRUTAS GDE", 520),
    monthClose("2026-07", "MOKA GDE", 640),
  ];

  const liveBefore = calculateForecast({
    stockRows,
    historicalVentas: history,
    bajas: [],
    existencias: [],
    realProduction: [],
    selectedMonth: "2026-08",
    dailyBufferPct: 10,
  });
  const frozenSnapshot = {
    version: 3,
    periodo: "2026-08",
    contenido: {
      schemaVersion: 2,
      selectedMonth: "2026-08",
      rows: liveBefore.map((row) => ({
        producto: row.producto,
        pronosticoBase: row.pronosticoVenta,
        pronosticoOperativo: row.pronosticoVenta * 1.12,
      })),
      forecastRows: snapshotForecastRowsForFreeze(liveBefore),
    },
  };

  const laterHistory = history.map((row) => (
    row.producto === "FRUTAS GDE" && row.cantidad === 520
      ? { ...row, cantidad: 820 }
      : row
  ));
  const liveAfter = calculateForecast({
    stockRows,
    historicalVentas: laterHistory,
    bajas: [],
    existencias: [],
    realProduction: [],
    selectedMonth: "2026-08",
    dailyBufferPct: 10,
  });
  assert(
    Math.abs(forecastTotal(liveAfter) - forecastTotal(liveBefore)) > 20,
    "la mutación de ventas debe mover el modelo vivo; si no, el test no demuestra el freeze"
  );

  const pinned = resolveEffectiveForecast({
    liveForecast: liveAfter,
    frozenSnapshot,
    selectedMonth: "2026-08",
  });
  assert(pinned.source === "frozen", "con forecastRows el mes debe quedar anclado al congelado");
  assert(pinned.label === "pronóstico congelado v3", `etiqueta de operador incorrecta: ${pinned.label}`);
  almostEqual(forecastTotal(pinned.rows), forecastTotal(liveBefore), 0.001, "el total congelado no debe derivar");

  const frutasBefore = liveBefore.find((row) => row.producto === "FRUTAS GDE");
  const frutasPinned = pinned.rows.find((row) => row.producto === "FRUTAS GDE");
  weekdayShape(frutasBefore).forEach((value, index) => {
    almostEqual(weekdayShape(frutasPinned)[index], value, 0.001, `el promedio semanal ${index} no debe derivar`);
  });
  almostEqual(frutasPinned.produccionSugerida, frutasBefore.produccionSugerida, 0.001, "la producción sugerida congelada no debe derivar");

  const dailyBefore = calculateDailyForecast({
    monthlyRows: liveBefore,
    ventasReales: [],
    realProduction: [],
    selectedMonth: "2026-08",
    dailyBufferPct: 10,
  });
  const dailyPinned = calculateDailyForecast({
    monthlyRows: pinned.rows,
    ventasReales: [],
    realProduction: [],
    selectedMonth: "2026-08",
    dailyBufferPct: 10,
    activePromos: [],
    dailyBranchStock: [],
    dailyColdRoom: [],
  });
  const dailyLiveAfter = calculateDailyForecast({
    monthlyRows: liveAfter,
    ventasReales: [],
    realProduction: [],
    selectedMonth: "2026-08",
    dailyBufferPct: 10,
  });
  const plantTotal = (rows) => rows.reduce((sum, row) => sum + row.pronosticoVentaDia, 0);
  almostEqual(plantTotal(dailyPinned), plantTotal(dailyBefore), 0.05, "planta no debe derivar el pronóstico diario de un mes congelado");
  assert(
    Math.abs(plantTotal(dailyLiveAfter) - plantTotal(dailyBefore)) > 20,
    "sin freeze, planta sí cambiaría al recargar ventas"
  );

  const closeBefore = buildMonthlyCloseSummary({
    forecastRows: pinned.rows,
    salesRows: [
      { producto: "FRUTAS GDE", cantidad: 530 },
      { producto: "MOKA GDE", cantidad: 650 },
    ],
    productionRows: [],
  });
  const closeIfLive = buildMonthlyCloseSummary({
    forecastRows: liveAfter,
    salesRows: [
      { producto: "FRUTAS GDE", cantidad: 530 },
      { producto: "MOKA GDE", cantidad: 650 },
    ],
    productionRows: [],
  });
  almostEqual(closeBefore.summary.pronostico, forecastTotal(liveBefore), 0.05, "el cierre debe medir contra el congelado");
  assert(
    Math.abs(closeIfLive.summary.pronostico - closeBefore.summary.pronostico) > 20,
    "el cierre contra el modelo vivo sí se movería"
  );

  const otherMonth = resolveEffectiveForecast({
    liveForecast: liveAfter,
    frozenSnapshot,
    selectedMonth: "2026-09",
  });
  assert(otherMonth.source === "live", "otro mes no hereda el freeze");

  const legacy = resolveEffectiveForecast({
    liveForecast: liveAfter,
    frozenSnapshot: {
      version: 1,
      periodo: "2026-08",
      contenido: {
        rows: frozenSnapshot.contenido.rows,
      },
    },
    selectedMonth: "2026-08",
  });
  assert(legacy.source === "frozen-operational", "un congelado viejo sigue anclando el total");
  almostEqual(forecastTotal(legacy.rows), forecastTotal(liveBefore), 0.05, "el total de un snapshot viejo no deriva");

  const hydrated = hydrateForecastFromOperationalRows(frozenSnapshot.contenido.rows, "2026-08");
  almostEqual(forecastTotal(hydrated), forecastTotal(liveBefore), 0.05, "hidratar el operativo debe conservar el total");

  const promo = normalizeActivePromo({
    producto: "FRUTAS GDE",
    startDate: "2026-08-03",
    endDate: "2026-08-07",
    multiplier: 1.5,
    extraPiecesPerDay: 0,
  });
  const dailyWithPromo = calculateDailyForecast({
    monthlyRows: pinned.rows,
    ventasReales: [],
    realProduction: [],
    selectedMonth: "2026-08",
    dailyBufferPct: 10,
    activePromos: [promo],
  });
  const promoDay = dailyWithPromo.find((row) => row.producto === "FRUTAS GDE" && row.fecha === "2026-08-04");
  const baseDay = dailyPinned.find((row) => row.producto === "FRUTAS GDE" && row.fecha === "2026-08-04");
  const expectedPromo = applyPromoUpliftToQuantity("FRUTAS GDE", baseDay.produccionBrutaDia, promo);
  almostEqual(promoDay.produccionBrutaDia, expectedPromo, 0.05, "la promo sigue empujando encima del mes congelado");

  const withInventory = applyInventoryToProductionSuggestion(baseDay.produccionBrutaDia, 8, 5);
  assert(withInventory === Math.max(0, baseDay.produccionBrutaDia - 13), "sucursales + CF siguen restando sobre el bruto congelado");

  const summary = summarizeForecastAccuracy([
    { actual: 100, forecast: 90 },
    { actual: 50, forecast: 80 },
  ]);
  almostEqual(summary.wape, (40 / 150) * 100, 1e-9, "WAPE = suma |error| / suma real");
  almostEqual(summary.absoluteError, 40, 1e-9, "la suma de error absoluto debe exportarse para el ponderado");
  almostEqual(
    weightedWapeFromBacktests([
      { actual: 100, absoluteError: 10 },
      { actual: 50, absoluteError: 30 },
    ]),
    (40 / 150) * 100,
    1e-9,
    "WAPE ponderado = suma de errores / suma de reales"
  );
  assert(weightedWapeFromBacktests([]) == null, "sin meses cerrados no hay WAPE ponderado");

  console.log("frozen-month-ops-test: freeze ancla planta/cierre y el WAPE operador usa la misma cuenta");
}

main().catch((error) => {
  console.error(error.message);
  process.exit(1);
});
