const assert = require("assert");
const path = require("path");
const Module = require("module");
const esbuild = require("esbuild");

async function loadApp() {
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
  const appModule = new Module("data-freshness-test");
  appModule.filename = path.join(__dirname, "data-freshness-test.bundle.cjs");
  appModule.paths = module.paths;
  appModule._compile(built.outputFiles[0].text, appModule.filename);
  return appModule.exports;
}

function daily(fecha, producto = "FRUTAS GDE", cantidad = 10) {
  return { fecha, producto, cantidad };
}

function monthly(month, producto = "FRUTAS GDE", cantidad = 200) {
  const [year, monthNumber] = month.split("-").map(Number);
  return {
    fecha: new Date(year, monthNumber - 1, 1),
    producto,
    cantidad,
    monthlyTotal: true,
    monthDays: new Date(year, monthNumber, 0).getDate(),
  };
}

(async () => {
  const app = await loadApp();
  const {
    assessDataFreshness,
    expectedDataFreshnessDate,
    isProductionWorkingDay,
    lastCompletedWorkingDay,
    latestLoadedDailyDate,
    normalizeDataFreshnessLag,
    DEFAULT_DATA_FRESHNESS_MAX_LAG_WORKING_DAYS,
    DATA_FRESHNESS_MAX_LAG_WORKING_DAYS,
    buildProductValidationSummary,
    isDistribucionChannel,
  } = app;

  assert.strictEqual(DEFAULT_DATA_FRESHNESS_MAX_LAG_WORKING_DAYS, 0, "el umbral por defecto es 0 días hábiles de atraso");
  assert.strictEqual(DATA_FRESHNESS_MAX_LAG_WORKING_DAYS, 0, "la constante exportada coincide con el default");
  assert.strictEqual(normalizeDataFreshnessLag(""), 0);
  assert.strictEqual(normalizeDataFreshnessLag(-3), 0);
  assert.strictEqual(normalizeDataFreshnessLag(2.4), 2);
  assert.strictEqual(normalizeDataFreshnessLag(99), 14);

  // 2026-10-14 miércoles, 11 domingo, 12 lunes, 10 sábado.
  assert.strictEqual(isProductionWorkingDay("2026-10-10"), true, "sábado es día de planta");
  assert.strictEqual(isProductionWorkingDay("2026-10-11"), false, "domingo no hay producción");
  assert.strictEqual(isProductionWorkingDay("2026-10-12"), true, "lunes es día de planta");

  assert.strictEqual(lastCompletedWorkingDay("2026-10-14"), "2026-10-13", "miércoles → martes");
  assert.strictEqual(lastCompletedWorkingDay("2026-10-12"), "2026-10-10", "lunes → sábado");
  assert.strictEqual(lastCompletedWorkingDay("2026-10-11"), "2026-10-10", "domingo → sábado (el pedido del domingo es para el lunes)");
  assert.strictEqual(lastCompletedWorkingDay("2026-10-10"), "2026-10-09", "sábado → viernes");

  assert.strictEqual(expectedDataFreshnessDate("2026-10-14", 0), "2026-10-13");
  assert.strictEqual(expectedDataFreshnessDate("2026-10-14", 1), "2026-10-12", "un día hábil de atraso: lunes");
  assert.strictEqual(expectedDataFreshnessDate("2026-10-12", 0), "2026-10-10");
  assert.strictEqual(expectedDataFreshnessDate("2026-10-11", 0), "2026-10-10");
  assert.strictEqual(expectedDataFreshnessDate("2026-10-12", 1), "2026-10-09");

  assert.strictEqual(latestLoadedDailyDate([monthly("2026-09"), daily("2026-10-10"), daily("2026-10-08")]), "2026-10-10");
  assert.strictEqual(latestLoadedDailyDate([monthly("2026-09")]), "", "el cierre mensual no cuenta como fecha diaria");

  const empty = assessDataFreshness({ asOf: "2026-10-14" });
  assert.strictEqual(empty.checked, false, "sin datos cargados no se avisa");
  assert.strictEqual(empty.fresh, true);
  assert.strictEqual(empty.message, "");

  const salesFresh = assessDataFreshness({
    ventas: [daily("2026-10-13")],
    asOf: "2026-10-14",
  });
  assert.strictEqual(salesFresh.fresh, true, "venta del último día hábil no avisa");
  assert.strictEqual(salesFresh.checked, true);
  assert.strictEqual(salesFresh.sales.latest, "2026-10-13");

  const salesStale = assessDataFreshness({
    ventas: [daily("2026-10-10")],
    asOf: "2026-10-14",
  });
  assert.strictEqual(salesStale.fresh, false, "venta anterior al último día hábil avisa");
  assert.strictEqual(salesStale.expectedDate, "2026-10-13");
  assert.match(salesStale.message, /ventas llegan al 10\/10\/2026/);
  assert.match(salesStale.message, /último día hábil \(13\/10\/2026\)/);
  assert.match(salesStale.detail, /lunes a sábado/);
  assert.match(salesStale.detail, /pedido del domingo es para el lunes/);
  assert.strictEqual(salesStale.title, "Datos atrasados");

  const mondayNeedsSaturday = assessDataFreshness({
    ventas: [daily("2026-10-09")],
    realProduction: [daily("2026-10-10", "FRUTAS GDE", 20)],
    asOf: "2026-10-12",
  });
  assert.strictEqual(mondayNeedsSaturday.fresh, false, "el lunes exige el sábado");
  assert.strictEqual(mondayNeedsSaturday.sales.fresh, false);
  assert.strictEqual(mondayNeedsSaturday.production.fresh, true);
  assert.match(mondayNeedsSaturday.message, /ventas llegan al 09\/10\/2026/);
  assert.doesNotMatch(mondayNeedsSaturday.message, /producción/, "solo avisa la fuente atrasada");

  const sundayUsesSaturday = assessDataFreshness({
    ventas: [daily("2026-10-10")],
    realProduction: [daily("2026-10-10")],
    asOf: "2026-10-11",
  });
  assert.strictEqual(sundayUsesSaturday.fresh, true, "el domingo se cubre con el sábado: el pedido es para el lunes");

  const productionOnlyStale = assessDataFreshness({
    ventas: [daily("2026-10-13")],
    realProduction: [daily("2026-10-08")],
    asOf: "2026-10-14",
  });
  assert.strictEqual(productionOnlyStale.fresh, false);
  assert.strictEqual(productionOnlyStale.sales.fresh, true);
  assert.strictEqual(productionOnlyStale.production.fresh, false);
  assert.match(productionOnlyStale.message, /producción llega al 08\/10\/2026/);

  const monthlyOnly = assessDataFreshness({
    ventas: [monthly("2026-09")],
    asOf: "2026-10-14",
  });
  assert.strictEqual(monthlyOnly.fresh, false, "un cierre mensual no sustituye el día hábil");
  assert.match(monthlyOnly.message, /no tienen fecha diaria/);

  const configurableLag = assessDataFreshness({
    ventas: [daily("2026-10-12")],
    asOf: "2026-10-14",
    maxLagWorkingDays: 1,
  });
  assert.strictEqual(configurableLag.fresh, true, "con umbral 1 el lunes cubre el miércoles");
  assert.strictEqual(configurableLag.expectedDate, "2026-10-12");

  const configurableLagStillStale = assessDataFreshness({
    ventas: [daily("2026-10-09")],
    asOf: "2026-10-14",
    maxLagWorkingDays: 1,
  });
  assert.strictEqual(configurableLagStillStale.fresh, false);
  assert.match(configurableLagStillStale.message, /1 día hábil de atraso permitido/);

  const bothStale = assessDataFreshness({
    ventas: [daily("2026-10-08")],
    realProduction: [daily("2026-10-09")],
    asOf: "2026-10-14",
  });
  assert.strictEqual(bothStale.fresh, false);
  assert.match(bothStale.message, /ventas llegan al 08\/10\/2026 y la producción llega al 09\/10\/2026/);

  const newerThanExpected = assessDataFreshness({
    ventas: [daily("2026-10-14")],
    asOf: "2026-10-14",
  });
  assert.strictEqual(newerThanExpected.fresh, true, "si ya cargaron el día de hoy, no se avisa");

  const validation = buildProductValidationSummary([
    {
      producto: "DISTRIBUCIÓN EN SUCURSALES",
      pronosticoVentaDia: 10,
      ventaRealDia: 4,
      hasVentaReal: true,
      precisionVenta: 40,
    },
    {
      producto: "BOLILLO",
      pronosticoVentaDia: 10,
      ventaRealDia: 8,
      hasVentaReal: true,
      precisionVenta: 80,
    },
  ], []);
  const dist = validation.find((row) => isDistribucionChannel(row.producto));
  const bolillo = validation.find((row) => row.producto === "BOLILLO");
  assert(dist && dist.errorPct === null, "Distribución sigue sin porcentaje de error (solo apoyo a la decisión)");
  assert(bolillo && bolillo.errorPct !== null, "un producto de sucursal sí puede tener error %");

  console.log("data-freshness-test: aviso de datos atrasados (último día hábil, umbral configurable)");
})().catch((error) => {
  console.error(error);
  process.exit(1);
});
