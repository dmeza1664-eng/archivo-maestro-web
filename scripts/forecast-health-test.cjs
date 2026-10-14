/**
 * Salud del pronóstico (PR #42): buildForecastHealth / WAPE y
 * actualSalesByProductForMonth. Datos sintéticos.
 *
 * El cierre (monthlyTotal) es el total del producto/mes; el diario solo
 * aporta forma. Sumar ambas cuentas infló el WAPE de septiembre a ~50%
 * cuando la cuenta única ronda ~9.6%. Estas pruebas fallan si se vuelve
 * a contar dos veces: diario + cierre, o el mismo mes subido dos veces.
 *
 * Distribución es apoyo a la decisión: nunca lleva porcentaje de error.
 */
const fs = require("fs");
const path = require("path");
const Module = require("module");
const esbuild = require("esbuild");

const ROOT = path.join(__dirname, "..");

async function loadApp() {
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
  const appModule = new Module("forecast-health-test");
  appModule.filename = path.join(__dirname, "forecast-health-test.bundle.cjs");
  appModule.paths = module.paths;
  appModule._compile(built.outputFiles[0].text, appModule.filename);
  return appModule.exports;
}

function assert(condition, message) {
  if (!condition) throw new Error(`[salud WAPE] ${message}`);
}

function almostEqual(actual, expected, tolerance, message) {
  if (Math.abs(actual - expected) > tolerance) {
    throw new Error(`[salud WAPE] ${message} (${actual} vs ${expected})`);
  }
}

function monthClose(monthKey, producto, cantidad, extra = {}) {
  const [year, month] = monthKey.split("-").map(Number);
  return {
    fecha: new Date(year, month - 1, 1),
    producto,
    cantidad,
    monthlyTotal: true,
    monthDays: new Date(year, month, 0).getDate(),
    ...extra,
  };
}

function monthDaily(monthKey, producto, total, sourceFile) {
  const [year, month] = monthKey.split("-").map(Number);
  const days = new Date(year, month, 0).getDate();
  const rows = [];
  for (let day = 1; day <= days; day += 1) {
    rows.push({
      fecha: new Date(year, month - 1, day),
      producto,
      cantidad: total / days,
      sourceFile,
    });
  }
  return rows;
}

function flattenResolved(resolved) {
  return resolved.entries.flatMap((entry) => entry.rows);
}

function pickMonth(health, month) {
  return (health.backtests || []).find((row) => row.month === month);
}

function naiveMonthSum(records, monthKey) {
  return (records || []).reduce((sum, row) => {
    const key = row.fecha instanceof Date
      ? `${row.fecha.getFullYear()}-${String(row.fecha.getMonth() + 1).padStart(2, "0")}`
      : String(row.fecha || "").slice(0, 7);
    return key === monthKey ? sum + Number(row.cantidad || 0) : sum;
  }, 0);
}

(async () => {
  const app = await loadApp();
  const {
    actualSalesByProductForMonth,
    buildForecastHealth,
    buildProductValidationSummary,
    isDistribucionChannel,
    resolveCanonicalMonthSources,
    summarizeForecastAccuracy,
    weightedWapeFromBacktests,
  } = app;

  const stockRows = [
    { producto: "FRUTAS GDE", stock: 40, orden: 1 },
    { producto: "MOKA GDE", stock: 40, orden: 2 },
  ];
  const sepCloseQty = { "FRUTAS GDE": 510, "MOKA GDE": 630 };
  const sepCloseTotal = 510 + 630;
  const history = {
    "FRUTAS GDE": {
      "2025-07": 480, "2025-08": 500, "2025-09": 492,
      "2026-06": 503, "2026-07": 525, "2026-08": 518,
    },
    "MOKA GDE": {
      "2025-07": 600, "2025-08": 620, "2025-09": 604,
      "2026-06": 624, "2026-07": 650, "2026-08": 640,
    },
  };
  const historyRows = [];
  for (const [producto, months] of Object.entries(history)) {
    for (const [month, qty] of Object.entries(months)) {
      historyRows.push(monthClose(month, producto, qty, { sourceFile: "historia.xlsx" }));
    }
  }
  const sepCloseRows = Object.entries(sepCloseQty).map(([producto, qty]) =>
    monthClose("2026-09", producto, qty, { sourceFile: "ventas septiembre 2026.xlsx" })
  );
  const sepDailyRows = Object.entries(sepCloseQty).flatMap(([producto, qty]) =>
    monthDaily("2026-09", producto, qty, "VENTAS DIARIO SEPTIEMBRE 2026.xlsx")
  );
  const closeOnly = [...historyRows, ...sepCloseRows];
  const dailyOnly = [...historyRows, ...sepDailyRows];
  const dailyAndClose = [...historyRows, ...sepCloseRows, ...sepDailyRows];
  const uploadedTwiceDaily = [
    ...historyRows,
    ...sepCloseRows,
    ...sepDailyRows,
    ...sepDailyRows.map((row) => ({ ...row, sourceFile: "Base de datos", databaseSynced: true })),
  ];

  // --- actualSalesByProductForMonth: una sola cuenta ---
  const fromClose = actualSalesByProductForMonth(sepCloseRows, "2026-09");
  const fromDaily = actualSalesByProductForMonth(sepDailyRows, "2026-09");
  const fromBoth = actualSalesByProductForMonth(dailyAndClose, "2026-09");
  const fromTwiceDaily = actualSalesByProductForMonth(uploadedTwiceDaily, "2026-09");
  assert(fromClose.get("FRUTAS GDE") === 510, "solo cierre: FRUTAS = 510");
  assert(fromClose.get("MOKA GDE") === 630, "solo cierre: MOKA = 630");
  almostEqual(fromDaily.get("FRUTAS GDE"), 510, 1e-6, "solo diario suma el mes una vez");
  almostEqual(fromDaily.get("MOKA GDE"), 630, 1e-6, "solo diario: MOKA = 630");
  assert(fromBoth.get("FRUTAS GDE") === 510, "diario + cierre no suma FRUTAS");
  assert(fromBoth.get("MOKA GDE") === 630, "diario + cierre no suma MOKA");
  assert(fromTwiceDaily.get("FRUTAS GDE") === 510, "diario de sesión + diario de base no cuenta si hay cierre");
  assert(fromTwiceDaily.get("MOKA GDE") === 630, "diario duplicado + cierre: MOKA sigue siendo el cierre");

  const naiveBoth = naiveMonthSum(dailyAndClose, "2026-09");
  assert(naiveBoth > sepCloseTotal * 1.5, "la suma cruda de diario + cierre sí dobla septiembre");
  assert(
    fromBoth.get("FRUTAS GDE") + fromBoth.get("MOKA GDE") === sepCloseTotal,
    "la cuenta canónica de septiembre es 1140, no 2280"
  );

  const incompleteDaily = [
    ...sepCloseRows,
    ...monthDaily("2026-09", "FRUTAS GDE", 400, "diario incompleto.xlsx"),
    ...monthDaily("2026-09", "MOKA GDE", 200, "diario incompleto.xlsx"),
  ];
  const fromIncomplete = actualSalesByProductForMonth(incompleteDaily, "2026-09");
  assert(fromIncomplete.get("FRUTAS GDE") === 510, "si el diario no empató, el cierre sigue ganando el total");
  assert(fromIncomplete.get("MOKA GDE") === 630, "el cierre gana aunque el diario sea otra cifra");

  const otherMonth = actualSalesByProductForMonth(
    [...sepCloseRows, monthClose("2026-08", "FRUTAS GDE", 999)],
    "2026-09"
  );
  assert(otherMonth.get("FRUTAS GDE") === 510, "no mezcla el mes anterior");
  assert(!otherMonth.has("PINA GDE"), "un producto ausente no aparece");

  // --- El WAPE ~9.6% se vuelve ~50% si se dobla la venta real ---
  // Caso reportado: pronóstico cerca del cierre; |1030-1140|/1140 ≈ 9.65%.
  // Si se suman diario y cierre, actual = 2280 y |1030-2280|/2280 ≈ 54.8%.
  const reportedSingle = summarizeForecastAccuracy([
    { actual: 510, forecast: 461 },
    { actual: 630, forecast: 569 },
  ]);
  const reportedDoubled = summarizeForecastAccuracy([
    { actual: 1020, forecast: 461 },
    { actual: 1260, forecast: 569 },
  ]);
  almostEqual(reportedSingle.wape, (110 / 1140) * 100, 1e-9, "WAPE de una sola cuenta ≈ 9.65%");
  almostEqual(reportedDoubled.wape, (1250 / 2280) * 100, 1e-9, "WAPE con doble conteo ≈ 54.8%");
  assert(reportedSingle.wape < 12, "la cuenta correcta queda cerca del 9.6% reportado");
  assert(reportedDoubled.wape > 45, "el doble conteo dispara el WAPE hacia ~50%");
  assert(
    reportedDoubled.wape - reportedSingle.wape > 30,
    "la diferencia entre una cuenta y dos es el síntoma de septiembre"
  );

  assert(summarizeForecastAccuracy([]).wape == null, "sin filas no hay WAPE");
  assert(weightedWapeFromBacktests([]) == null, "sin meses cerrados no hay WAPE ponderado");
  almostEqual(
    weightedWapeFromBacktests([
      { actual: 1140, absoluteError: 110 },
      { actual: 1000, absoluteError: 80 },
    ]),
    (190 / 2140) * 100,
    1e-9,
    "WAPE ponderado = suma |error| / suma real"
  );

  // --- buildForecastHealth: diario + cierre no dobla WAPE ---
  function healthOf(ventas) {
    return buildForecastHealth({
      stockRows,
      ventas,
      selectedMonth: "2026-10",
      dailyBufferPct: 10,
    });
  }
  const healthClose = healthOf(closeOnly);
  const healthDaily = healthOf(dailyOnly);
  const healthBoth = healthOf(dailyAndClose);
  const healthTwiceDaily = healthOf(uploadedTwiceDaily);
  const sepClose = pickMonth(healthClose, "2026-09");
  const sepDaily = pickMonth(healthDaily, "2026-09");
  const sepBoth = pickMonth(healthBoth, "2026-09");
  const sepTwiceDaily = pickMonth(healthTwiceDaily, "2026-09");
  assert(sepClose && sepDaily && sepBoth && sepTwiceDaily, "octubre backtestea septiembre");
  almostEqual(sepBoth.actual, sepClose.actual, 1e-6, "diario + cierre no dobla la venta real");
  almostEqual(sepTwiceDaily.actual, sepClose.actual, 1e-6, "diario subido dos veces + cierre no dobla");
  almostEqual(sepDaily.actual, sepClose.actual, 1e-6, "solo diario empató el cierre");
  almostEqual(sepBoth.forecast, sepClose.forecast, 1e-6, "el pronóstico no cambia al cargar diario + cierre");
  almostEqual(sepBoth.wape, sepClose.wape, 0.05, "WAPE con ambos archivos = WAPE de una sola cuenta");
  almostEqual(sepTwiceDaily.wape, sepClose.wape, 0.05, "WAPE con diario duplicado = WAPE del cierre");
  assert(sepBoth.wape < 20, `WAPE sep no debe parecer ~50% por doble conteo (salió ${sepBoth.wape.toFixed(2)}%)`);
  assert(sepTwiceDaily.wape < 20, `WAPE sep con diario duplicado no debe parecer ~50% (salió ${sepTwiceDaily.wape.toFixed(2)}%)`);

  const inflatedIfDoubled = summarizeForecastAccuracy([
    { actual: sepClose.actual * 2, forecast: sepClose.forecast },
  ]);
  assert(
    inflatedIfDoubled.wape > 35,
    `si se doblara la venta real el WAPE saltaría (saldría ${inflatedIfDoubled.wape.toFixed(2)}%), y la prueba lo detecta`
  );
  assert(
    Math.abs(sepBoth.wape - inflatedIfDoubled.wape) > 15,
    "el WAPE real no es el de la venta doblada"
  );

  // --- Mes subido dos veces (dos cierres o dos diarios) ---
  const resolvedTwoCloses = resolveCanonicalMonthSources([
    { name: "cierre septiembre extra.xlsx", rows: sepCloseRows.map((row) => ({ ...row, sourceFile: "cierre septiembre extra.xlsx" })) },
    { name: "ventas septiembre 2026.xlsx", rows: sepCloseRows },
    { name: "historia.xlsx", rows: historyRows },
  ]);
  assert(
    resolvedTwoCloses.decisions.find((decision) => decision.month === "2026-09")?.strategy === "same-kind-winner",
    "dos cierres del mismo mes son excluyentes"
  );
  const twoCloseRows = flattenResolved(resolvedTwoCloses);
  const twoCloseActual = actualSalesByProductForMonth(twoCloseRows, "2026-09");
  assert(twoCloseActual.get("FRUTAS GDE") === 510, "dos Excel de cierre no suman FRUTAS");
  assert(twoCloseActual.get("MOKA GDE") === 630, "dos Excel de cierre no suman MOKA");
  const naiveTwoCloses = naiveMonthSum(
    [...sepCloseRows, ...sepCloseRows.map((row) => ({ ...row, sourceFile: "copia.xlsx" }))],
    "2026-09"
  );
  assert(naiveTwoCloses === sepCloseTotal * 2, "sin resolver, dos cierres concatenados sí doblan");
  const sepTwoCloses = pickMonth(healthOf(twoCloseRows), "2026-09");
  almostEqual(sepTwoCloses.actual, sepClose.actual, 1e-6, "dos cierres resueltos no doblan el actual de salud");
  almostEqual(sepTwoCloses.wape, sepClose.wape, 0.05, "WAPE con el mes subido dos veces como cierre = una cuenta");
  assert(sepTwoCloses.wape < 20, `dos cierres no deben inflar el WAPE a ~50% (salió ${sepTwoCloses.wape.toFixed(2)}%)`);

  const otherDailyQty = { "FRUTAS GDE": 800, "MOKA GDE": 900 };
  const otherDailyRows = Object.entries(otherDailyQty).flatMap(([producto, qty]) =>
    monthDaily("2026-09", producto, qty, "VENTAS DIARIO SEPTIEMBRE combinado.xlsx")
  );
  const resolvedTwoDailies = resolveCanonicalMonthSources([
    { name: "VENTAS DIARIO SEPTIEMBRE combinado.xlsx", rows: otherDailyRows },
    { name: "VENTAS DIARIO SEPTIEMBRE 2026.xlsx", rows: sepDailyRows },
    { name: "historia.xlsx", rows: historyRows },
  ]);
  assert(
    resolvedTwoDailies.decisions.find((decision) => decision.month === "2026-09")?.strategy === "same-kind-winner",
    "dos diarios del mismo mes son excluyentes"
  );
  const twoDailyRows = flattenResolved(resolvedTwoDailies);
  const twoDailyActual = actualSalesByProductForMonth(twoDailyRows, "2026-09");
  almostEqual(twoDailyActual.get("FRUTAS GDE"), 510, 1e-6, "el diario dedicado gana; no se suman 800+510");
  almostEqual(twoDailyActual.get("MOKA GDE"), 630, 1e-6, "el diario dedicado gana en MOKA");
  assert(
    twoDailyActual.get("FRUTAS GDE") + twoDailyActual.get("MOKA GDE") < 1400,
    "dos diarios no deben dejar la venta en 1700"
  );
  const sepTwoDailies = pickMonth(healthOf(twoDailyRows), "2026-09");
  almostEqual(sepTwoDailies.actual, sepClose.actual, 1e-6, "dos diarios resueltos no doblan septiembre");
  almostEqual(sepTwoDailies.wape, sepDaily.wape, 0.05, "WAPE con el mes subido dos veces como diario = una cuenta");

  const resolvedDailyAndClose = resolveCanonicalMonthSources([
    { name: "VENTAS DIARIO SEPTIEMBRE 2026.xlsx", rows: sepDailyRows },
    { name: "ventas septiembre 2026.xlsx", rows: sepCloseRows },
    { name: "historia.xlsx", rows: historyRows },
  ]);
  assert(
    resolvedDailyAndClose.decisions.find((decision) => decision.month === "2026-09")?.strategy === "keep-daily-and-close",
    "septiembre diario + cierre se conserva"
  );
  const resolvedBothRows = flattenResolved(resolvedDailyAndClose);
  const resolvedBothActual = actualSalesByProductForMonth(resolvedBothRows, "2026-09");
  assert(resolvedBothActual.get("FRUTAS GDE") === 510, "tras resolveCanonical el actual es el cierre");
  const sepResolved = pickMonth(healthOf(resolvedBothRows), "2026-09");
  almostEqual(sepResolved.wape, sepClose.wape, 0.05, "WAPE por resolveCanonical no dobla el cierre");
  assert(sepResolved.wape < 20, "WAPE resuelto no parece ~50%");

  // --- Mensajes de salud en español; sin stock / sin ventas ---
  const missingStock = buildForecastHealth({ ventas: closeOnly, selectedMonth: "2026-10" });
  assert(missingStock.checks.some((item) => item.code === "missing-stock"), "sin stock marca error");
  assert(/Carga el stock ideal/.test(missingStock.checks.find((item) => item.code === "missing-stock").message), "aviso de stock en español");
  assert(!missingStock.ready, "sin stock no está listo");

  const missingSales = buildForecastHealth({ stockRows, selectedMonth: "2026-10" });
  assert(missingSales.checks.some((item) => item.code === "missing-sales"), "sin ventas marca error");
  assert(/Carga ventas históricas/.test(missingSales.checks.find((item) => item.code === "missing-sales").message), "aviso de ventas en español");

  assert(healthClose.checks.some((item) => item.code === "forecast-ready"), "con datos el control marca el modelo activo");
  assert(
    /Pronóstico/.test(healthClose.checks.find((item) => item.code === "forecast-ready").message),
    "el aviso de listo habla de Pronóstico, en español"
  );
  assert(healthClose.backtests.every((row) => row.wape != null), "cada mes oculto tiene WAPE");
  assert(healthClose.backtests.every((row) => Array.isArray(row.topErrors)), "cada mes listo tiene productos con más error");

  // --- Distribución: apoyo a la decisión, nunca error % ---
  assert(isDistribucionChannel("Distribución"), "se reconoce el canal Distribución");
  assert(isDistribucionChannel("DISTRIBUCIÓN EN SUCURSALES"), "DISTRIBUCIÓN EN SUCURSALES es Distribución");
  assert(!isDistribucionChannel("Suc. León"), "una sucursal no es Distribución");
  const validation = buildProductValidationSummary(
    [
      {
        producto: "DISTRIBUCIÓN EN SUCURSALES",
        pronosticoVentaDia: 10,
        ventaRealDia: 4,
        hasVentaReal: true,
        precisionVenta: 40,
      },
      {
        producto: "FRUTAS GDE",
        pronosticoVentaDia: 10,
        ventaRealDia: 8,
        hasVentaReal: true,
        precisionVenta: 80,
      },
    ],
    []
  );
  const dist = validation.find((row) => isDistribucionChannel(row.producto));
  const sucursal = validation.find((row) => row.producto === "FRUTAS GDE");
  assert(dist && dist.errorPct === null, "Distribución nunca muestra porcentaje de error");
  assert(sucursal && sucursal.errorPct !== null, "un producto de sucursal sí puede tener error %");

  const appSource = fs.readFileSync(path.join(ROOT, "App.jsx"), "utf8");
  assert(
    /isDistribucionChannel\(row\.producto\)\s*\n\s*\? null/.test(appSource),
    "la validación sigue anulando error % en Distribución"
  );
  assert(/WAPE de meses cerrados/.test(appSource), "la tarjeta de WAPE sigue en español");
  assert(/Carga ventas de meses anteriores para ver el WAPE automático/.test(appSource), "el vacío de WAPE sigue en español");
  assert(/Ver productos con más error/.test(appSource), "el detalle de salud sigue en español");
  assert(
    !/Distribuci[oó]n[\s\S]{0,180}error\s*%/i.test(appSource.match(/function ForecastHealthStrip[\s\S]+?\nfunction /)[0]),
    "la tira de salud no agrega error % a Distribución"
  );

  console.log("forecast-health-test ok");
  console.log(JSON.stringify({
    sepCloseWape: Number(sepClose.wape.toFixed(2)),
    sepBothWape: Number(sepBoth.wape.toFixed(2)),
    sepTwoClosesWape: Number(sepTwoCloses.wape.toFixed(2)),
    sepTwoDailiesWape: Number(sepTwoDailies.wape.toFixed(2)),
    reportedSingleWape: Number(reportedSingle.wape.toFixed(2)),
    reportedDoubledWape: Number(reportedDoubled.wape.toFixed(2)),
    sepActual: sepBoth.actual,
  }));
})().catch((error) => {
  console.error(error);
  process.exit(1);
});
