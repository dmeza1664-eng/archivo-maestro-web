const path = require("path");
const Module = require("module");
const esbuild = require("esbuild");
const XLSX = require("xlsx");

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
  const appModule = new Module("parser-test");
  appModule.filename = path.join(__dirname, "parser-test.bundle.cjs");
  appModule.paths = module.paths;
  appModule._compile(built.outputFiles[0].text, appModule.filename);
  return appModule.exports;
}

function workbookFromSheets(sheets) {
  const workbook = XLSX.utils.book_new();
  for (const [name, rows] of Object.entries(sheets)) {
    XLSX.utils.book_append_sheet(workbook, XLSX.utils.aoa_to_sheet(rows), name);
  }
  return workbook;
}

function assert(condition, message) {
  if (!condition) throw new Error(message);
}

function monthKey(row) {
  if (row.fecha instanceof Date) {
    return `${row.fecha.getFullYear()}-${String(row.fecha.getMonth() + 1).padStart(2, "0")}`;
  }
  return String(row.fecha || "").slice(0, 7);
}

async function main() {
  const {
    monthsNamedInFileName,
    parseBajasReport,
    parseBajasSummaryWorkbook,
    parseExistencias,
    parseMonthlySummaryWorkbook,
    parseSalesOrReturns,
    parseStock,
    resolveCanonicalMonthSources,
    sourceKindForMonth,
    describeSourceDecision,
    computeAnnualGrowthFactor,
    resolvePriorYearSeasonal,
    computeRecentMomentumFactor,
    isPriceTaggedProduct,
    applyCatalogOutlierCleanup,
    assessForecastFreezeReadiness,
    assessStockSheetSelection,
    buildSalesMonthCoverage,
    buildForecastHealth,
    analyzeForecastProductErrors,
    calculateForecast,
    buildMonthlyForecastData,
    getProduccionSugerida,
    countCapturedProductStatuses,
  } = await loadAppFunctions();

  assert(monthsNamedInFileName("VENTAS DE MAYO Y JUNIO 2026.xlsx").join(",") === "2026-05,2026-06", "el combinado debe nombrar mayo y junio");
  assert(monthsNamedInFileName("ventas junio.xlsx")[0] === "2026-06", "el cierre de junio debe ser un mes");

  const combined = { name: "VENTAS DE MAYO Y JUNIO.xlsx", rows: [
    { fecha: "2026-05-03", producto: "BOLILLO", cantidad: 10 },
    { fecha: "2026-06-04", producto: "BOLILLO", cantidad: 20 },
  ] };
  const juneClose = { name: "ventas junio.xlsx", rows: [
    { fecha: "2026-06-01", producto: "BOLILLO", cantidad: 18, monthlyTotal: true, monthDays: 30 },
  ] };
  assert(sourceKindForMonth(combined.rows, "2026-06") === "daily", "el combinado de junio es diario");
  assert(sourceKindForMonth(juneClose.rows, "2026-06") === "close", "ventas junio.xlsx es cierre");
  const juneFirst = resolveCanonicalMonthSources([juneClose, combined]);
  const combinedAfterJuneFirst = juneFirst.entries.find((entry) => entry.name === combined.name);
  const juneAfterJuneFirst = juneFirst.entries.find((entry) => entry.name === juneClose.name);
  assert(combinedAfterJuneFirst.rows.some((row) => monthKey(row) === "2026-06"), "el diario de junio se conserva junto al cierre aunque el cierre se cargue primero");
  assert(combinedAfterJuneFirst.rows.some((row) => monthKey(row) === "2026-05"), "mayo del combinado no se toca");
  assert(juneAfterJuneFirst.rows.length === 1 && juneAfterJuneFirst.rows[0].monthlyTotal, "el cierre dedicado de junio se conserva");
  assert(juneFirst.decisions[0].winner === "ventas junio.xlsx", "el cierre dedicado gana el total del mes");
  assert(juneFirst.decisions[0].strategy === "keep-daily-and-close", "diario y cierre del mismo mes son complementarios");
  assert(juneFirst.decisions[0].kept.includes(combined.name), "el combinado diario queda en kept");
  assert(!juneFirst.decisions[0].omitted.includes(combined.name), "no se omite el diario complementario");
  assert(
    /cierre de ventas junio/.test(describeSourceDecision(juneFirst.decisions[0])) &&
      /diario de VENTAS DE MAYO Y JUNIO/.test(describeSourceDecision(juneFirst.decisions[0])),
    "el mensaje debe nombrar cierre + diario"
  );

  const marchClose = resolveCanonicalMonthSources([
    { name: "VENTAS FEBRERO Y MARZO.xlsx", rows: [
      { fecha: "2026-02-01", producto: "BOLILLO", cantidad: 5 },
      { fecha: "2026-03-01", producto: "BOLILLO", cantidad: 9 },
    ] },
    { name: "cierre marzo.xlsx", rows: [{ fecha: "2026-03-01", producto: "BOLILLO", cantidad: 7, monthlyTotal: true, monthDays: 31 }] },
  ]);
  assert(marchClose.decisions[0].month === "2026-03", "el conflicto genérico no es solo junio");
  assert(marchClose.decisions[0].winner === "cierre marzo.xlsx", "marzo dedicado gana el total");
  assert(marchClose.decisions[0].strategy === "keep-daily-and-close", "marzo diario + cierre también se conservan");
  assert(
    marchClose.entries.find((entry) => entry.name === "VENTAS FEBRERO Y MARZO.xlsx").rows.some((row) => monthKey(row) === "2026-03"),
    "el diario de marzo no se tira"
  );

  const twoDailies = resolveCanonicalMonthSources([
    { name: "VENTAS DE MAYO Y JUNIO.xlsx", rows: [
      { fecha: "2026-05-03", producto: "BOLILLO", cantidad: 10 },
      { fecha: "2026-06-04", producto: "BOLILLO", cantidad: 20 },
    ] },
    { name: "ventas junio angel.xlsx", rows: [
      { fecha: "2026-06-05", producto: "BOLILLO", cantidad: 22 },
    ] },
  ]);
  assert(twoDailies.decisions[0].strategy === "same-kind-winner", "dos diarios del mismo mes siguen siendo excluyentes");
  assert(twoDailies.decisions[0].winner === "ventas junio angel.xlsx", "el diario dedicado gana al combinado");
  assert(
    twoDailies.entries.find((entry) => entry.name === "VENTAS DE MAYO Y JUNIO.xlsx").rows.every((row) => monthKey(row) !== "2026-06"),
    "el diario redundante de junio sí se omite"
  );

  const twoCloses = resolveCanonicalMonthSources([
    { name: "cierre junio extra.xlsx", rows: [{ fecha: "2026-06-01", producto: "BOLILLO", cantidad: 11, monthlyTotal: true, monthDays: 30 }] },
    { name: "ventas junio.xlsx", rows: [{ fecha: "2026-06-01", producto: "BOLILLO", cantidad: 18, monthlyTotal: true, monthDays: 30 }] },
  ]);
  assert(twoCloses.decisions[0].strategy === "same-kind-winner", "dos cierres del mismo mes son excluyentes");
  assert(twoCloses.decisions[0].winner === "ventas junio.xlsx", "el cierre con el mes en el nombre gana");

  const overridden = resolveCanonicalMonthSources([combined, juneClose], { "2026-06": combined.name });
  const combinedKept = overridden.entries.find((entry) => entry.name === combined.name);
  const closeAfterOverride = overridden.entries.find((entry) => entry.name === juneClose.name);
  assert(combinedKept.rows.some((row) => monthKey(row) === "2026-06"), "el override debe conservar junio del combinado");
  assert(closeAfterOverride.rows.every((row) => monthKey(row) !== "2026-06"), "el override exclusivo sí quita el cierre");
  assert(overridden.decisions[0].strategy === "override", "el override queda marcado como exclusivo");

  const stockWorkbook = workbookFromSheets({
    "OTRA HOJA": [
      ["PRODUCTO", "TOTAL GRAL SUC"],
      ["PAN MAL", 9999],
    ],
    "TOTAL A TENER SUC.(EXIST.+DIST)": [
      ["PRODUCTO", "SUC", "STOCK"],
      ["BOLILLO", 1, 113],
    ],
  });
  const stockRows = parseStock(stockWorkbook);
  assert(stockRows.length === 1 && stockRows[0].stock === 113, "stock debe leer la hoja TOTAL A TENER y la columna STOCK");

  const existenciasWorkbook = workbookFromSheets({
    "EXISTENCIA EN SUCURSALES": [
      ["PRODUCTO", "TOTAL GRAL SUC", "C.F.", "SUMA SUC+CF"],
      ["BOLILLO", 10, 2, 12],
    ],
  });
  const existenciasRows = parseExistencias(existenciasWorkbook);
  assert(existenciasRows.length === 1 && existenciasRows[0].sumaSucCf === 12, "existencias no debe confundirse con stock");
  assert(parseStock(existenciasWorkbook)[0]?.stock !== 12, "parseStock no debe tomar existencias como stock objetivo");

  // El 2026-08-21 un cambio de orden en las hojas candidatas saco del catalogo
  // los 16 productos de temporada sin avisar. La advertencia evita repetirlo.
  const catalogDriftWorkbook = workbookFromSheets({
    "TOTAL A TENER SUC.(EXIST.+DIST)": [
      ["PRODUCTO", "STOCK"],
      ["BOLILLO", 100],
    ],
    "EXIST. SUCURSALES Y RESTANTE CF": [
      ["PRODUCTO", "STOCK"],
      ["BOLILLO", 80],
      ["PAN MUERTO IND AZUCAR 50GR", 0],
    ],
  });
  const drift = assessStockSheetSelection(catalogDriftWorkbook);
  assert(drift.chosenSheet === "TOTAL A TENER SUC.(EXIST.+DIST)", "el catalogo debe seguir saliendo de la hoja prioritaria");
  assert(drift.missingTotal === 1, "debe contar los productos que la hoja elegida no trae");
  assert(/PAN MUERTO IND AZUCAR 50GR/.test(drift.message), "la advertencia debe nombrar el producto ausente");
  assert(drift.alternatives[0]?.sheet === "EXIST. SUCURSALES Y RESTANTE CF", "debe decir en que hoja si estaba");

  const catalogAgreesWorkbook = workbookFromSheets({
    "TOTAL A TENER SUC.(EXIST.+DIST)": [
      ["PRODUCTO", "STOCK"],
      ["BOLILLO", 100],
    ],
    "STOCK DE SUCURSALES": [
      ["PRODUCTO", "STOCK"],
      ["BOLILLO", 90],
    ],
  });
  const agrees = assessStockSheetSelection(catalogAgreesWorkbook);
  assert(agrees.missingTotal === 0, "hojas con el mismo universo no deben advertir");
  assert(agrees.message === "", "sin diferencias no hay mensaje que mostrar");

  const bajasWorkbook = workbookFromSheets({
    Resumen: [
      ["PRODUCTO", "CANTIDAD"],
      ["BOLILLO", 500],
    ],
    Reporte: [
      ["PRODUCTO", "CANTIDAD", "FECHA", "SUCURSAL", "MOTIVO"],
      ["BOLILLO", 3, "2026-07-02", "CENTRO", "Merma"],
      ["ERICK", 80, "2026-07-02", "CENTRO", "Subtotal"],
    ],
  });
  const bajasRows = parseSalesOrReturns(bajasWorkbook, "bajas", "bajas julio.xlsx");
  assert(bajasRows.length === 1, "bajas diarias deben salir solo de la hoja Reporte");
  assert(bajasRows[0].cantidad === 3, "no debe colarse el total mensual de Resumen");
  assert(parseBajasReport(bajasWorkbook).every((row) => !/^ERICK/.test(row.producto)), "Erick no es un producto de bajas");

  const summaryWorkbook = workbookFromSheets({
    Totales: [
      ["PRODUCTO", "CANTIDAD"],
      ["BOLILLO", 42],
    ],
  });
  const summaryRows = parseMonthlySummaryWorkbook(summaryWorkbook, { year: 2026, monthIndex: 5 });
  assert(summaryRows.length === 1 && summaryRows[0].monthlyTotal === true, "un resumen sin fechas diarias es total mensual");
  const salesSummary = parseSalesOrReturns(summaryWorkbook, "ventas", "ventas junio.xlsx");
  assert(salesSummary.every((row) => row.monthlyTotal), "parseSalesOrReturns no debe inventar días a partir del total mensual");

  const summaryBajas = parseBajasSummaryWorkbook(workbookFromSheets({
    "BAJAS ERICK": [
      ["ETIQUETAS", "CANTIDAD"],
      ["BOLILLO", 11],
      ["ERICK", 80],
    ],
  }));
  assert(summaryBajas.length === 1 && summaryBajas[0].producto.includes("BOLILLO"), "el subtotal Erick no entra al resumen de bajas");

  const wideWorkbook = workbookFromSheets({
    "JULIO 2026": [
      ["", "LUNES", "MARTES"],
      ["", 1, 2],
      ["BOLILLO", 10, 20],
      ["CONCHA", 4, 5],
    ],
  });
  const remapped = parseSalesOrReturns(wideWorkbook, "ventas", "VENTAS DE MAYO Y JUNIO.xlsx");
  assert(remapped.length >= 2, "la hoja ancha debe producir ventas diarias");
  assert(remapped.every((row) => monthKey(row) === "2026-06"), "hoja JULIO en archivo mayo-junio se interpreta como junio");

  const growthData = new Map([
    ["2025-03", { total: 100 }],
    ["2025-04", { total: 80 }],
    ["2025-05", { total: 100 }],
    ["2026-03", { total: 110 }],
    ["2026-04", { total: 100 }],
    ["2026-05", { total: 90 }],
  ]);
  assert(Math.abs(computeAnnualGrowthFactor(growthData, "2026-06", 1) - 0.9) < 1e-9, "un mes de crecimiento usa solo el mes previo");
  assert(Math.abs(computeAnnualGrowthFactor(growthData, "2026-06", 3) - 1.1) < 1e-9, "tres meses usan la mediana anual");
  const strongGrowth = new Map([
    ["2025-03", { total: 100 }],
    ["2025-04", { total: 100 }],
    ["2025-05", { total: 100 }],
    ["2026-03", { total: 140 }],
    ["2026-04", { total: 130 }],
    ["2026-05", { total: 90 }],
  ]);
  assert(Math.abs(computeAnnualGrowthFactor(strongGrowth, "2026-06", 3) - 1.28) < 1e-9, "dos meses fuertes permiten un tope de 1.28");
  assert(
    Math.abs(computeAnnualGrowthFactor(growthData, "2026-06", 1, { dampenDecline: 0.4 }) - 0.96) < 1e-9,
    "el recorte YoY de GDE se amortigua: 0.90 → 0.96"
  );

  const frutasDip = new Map([
    ["2025-05", { total: 658 }],
    ["2025-06", { total: 519 }],
    ["2025-07", { total: 475 }],
    ["2025-08", { total: 524 }],
    ["2025-09", { total: 492 }],
  ]);
  const frutasRef = resolvePriorYearSeasonal(frutasDip, "2026-07");
  assert(!frutasRef.usedDipCorrection, "sin dipRatio suave, FRUTAS 8.5% no corrige");
  const frutasSoft = resolvePriorYearSeasonal(frutasDip, "2026-07", { dipRatio: 0.92 });
  assert(frutasSoft.usedDipCorrection, "con umbral 8% FRUTAS sí corrige la caída de 475");
  assert(frutasSoft.levelFactor > 1.08, "la corrección debe subir al nivel de los vecinos (~522)");

  function monthWithHalves(monthKey, firstHalf, secondHalf, days = 30) {
    const [year, month] = monthKey.split("-").map(Number);
    const valuesByDate = new Map();
    for (let day = 1; day <= days; day += 1) {
      const key = `${year}-${String(month).padStart(2, "0")}-${String(day).padStart(2, "0")}`;
      valuesByDate.set(key, day <= 15 ? firstHalf / 15 : secondHalf / (days - 15));
    }
    return { total: firstHalf + secondHalf, valuesByDate, filledFromMonthlyTotal: false };
  }
  const momentumData = new Map([
    ["2025-06", { total: 519 }],
    ["2025-07", { total: 475 }],
    ["2025-08", { total: 524 }],
    ["2026-06", monthWithHalves("2026-06", 223, 291)],
  ]);
  const momentum = computeRecentMomentumFactor(momentumData, "2026-07");
  assert(momentum > 1.08 && momentum <= 1.12, `junio acelerado debe impulsar julio (factor ${momentum})`);
  const mildAccel = new Map([
    ["2025-06", { total: 274 }],
    ["2025-07", { total: 274 }],
    ["2025-08", { total: 322 }],
    ["2026-06", monthWithHalves("2026-06", 142, 159)],
  ]);
  assert(computeRecentMomentumFactor(mildAccel, "2026-07") === 1, "1.12× justo no dispara el default");
  assert(
    computeRecentMomentumFactor(mildAccel, "2026-07", { momentumTrigger: 1.1 }) > 1,
    "con umbral 1.10 el DUO de 1.12× sí impulsa"
  );
  const maySource = new Map([
    ["2026-05", monthWithHalves("2026-05", 200, 400, 31)],
  ]);
  assert(computeRecentMomentumFactor(maySource, "2026-06") === 1, "mayo (Día de las Madres) no impulsa junio");
  const payFade = new Map([
    ["2025-06", { total: 264 }],
    ["2025-07", { total: 199 }],
    ["2025-08", { total: 222 }],
    ["2026-06", monthWithHalves("2026-06", 89, 159)],
  ]);
  assert(computeRecentMomentumFactor(payFade, "2026-07") === 1, "PAY DE FRESA es baja de temporada: sin impulso");
  const julyCloseOnlyMonth = new Map([
    ["2026-07", { total: 580, valuesByDate: new Map([["2026-07-01", 18.7]]), filledFromMonthlyTotal: true, syntheticDays: 31 }],
  ]);
  assert(computeRecentMomentumFactor(julyCloseOnlyMonth, "2026-08") === 1, "un cierre sin diario no impulsa agosto");

  function dailyRowsForMonth(monthKey, dayCount, quantity = 10) {
    const [year, month] = monthKey.split("-").map(Number);
    return Array.from({ length: dayCount }, (_, index) => ({
      fecha: new Date(year, month - 1, index + 1),
      producto: "BOLILLO",
      cantidad: quantity,
    }));
  }

  function closeRowForMonth(monthKey, quantity = 300) {
    const [year, month] = monthKey.split("-").map(Number);
    return {
      fecha: new Date(year, month - 1, 1),
      producto: "BOLILLO",
      cantidad: quantity,
      monthlyTotal: true,
      monthDays: new Date(year, month, 0).getDate(),
    };
  }

  const completeSync = {
    sales: { ok: true, count: 1 },
    production: { ok: true, count: 1 },
    waste: { ok: true, count: 1 },
  };

  const juneDailyOnly = buildSalesMonthCoverage(dailyRowsForMonth("2026-06", 30));
  assert(juneDailyOnly[0].status === "daily-only", "30 días de junio sin cierre deben marcar solo diario");

  const juneCloseOnly = buildSalesMonthCoverage([closeRowForMonth("2026-06")]);
  assert(juneCloseOnly[0].status === "close-only", "cierre de junio sin diario debe marcar solo cierre");

  const juneComplete = buildSalesMonthCoverage([...dailyRowsForMonth("2026-06", 30), closeRowForMonth("2026-06", 25220)]);
  assert(juneComplete[0].status === "complete", "diario y cierre juntos marcan el mes completo");

  const junePartial = buildSalesMonthCoverage([...dailyRowsForMonth("2026-06", 10), closeRowForMonth("2026-06")]);
  assert(junePartial[0].status === "close-and-partial-daily", "10 días más cierre no alcanzan cobertura diaria");

  const inferredOnly = buildSalesMonthCoverage([{
    fecha: new Date(2026, 5, 1),
    producto: "BOLILLO",
    cantidad: 0,
    monthlyTotal: true,
    monthDays: 30,
    inferredZeroMonth: true,
  }]);
  assert(!inferredOnly.length || inferredOnly[0].status === "empty", "ceros inferidos no cuentan como cierre");

  const ready = assessForecastFreezeReadiness({
    selectedMonth: "2026-08",
    coverageRows: buildSalesMonthCoverage([...dailyRowsForMonth("2026-07", 31), closeRowForMonth("2026-07")]),
    databaseSync: completeSync,
    catalogCount: 111,
    capturedStatuses: 111,
  });
  assert(ready.canFreeze, "julio completo y agosto sin ventas debe permitir congelar");

  const withAugustSales = assessForecastFreezeReadiness({
    selectedMonth: "2026-08",
    coverageRows: buildSalesMonthCoverage([
      ...dailyRowsForMonth("2026-07", 31),
      closeRowForMonth("2026-07"),
      ...dailyRowsForMonth("2026-08", 24),
    ]),
    databaseSync: completeSync,
  });
  assert(withAugustSales.blockers.some((item) => item.code === "target-has-sales"), "ventas de agosto bloquean congelar agosto");

  const julyDailyOnly = assessForecastFreezeReadiness({
    selectedMonth: "2026-08",
    coverageRows: buildSalesMonthCoverage(dailyRowsForMonth("2026-07", 31)),
    databaseSync: completeSync,
  });
  assert(julyDailyOnly.blockers.some((item) => item.code === "previous-missing-close"), "julio solo diario bloquea por falta de cierre");

  const julyCloseOnly = assessForecastFreezeReadiness({
    selectedMonth: "2026-08",
    coverageRows: buildSalesMonthCoverage([closeRowForMonth("2026-07")]),
    databaseSync: completeSync,
  });
  assert(julyCloseOnly.canFreeze, "un cierre mensual debe permitir congelar aunque no haya diario");
  assert(
    julyCloseOnly.warnings.some((item) => item.code === "previous-missing-daily"),
    "julio solo cierre avisa que no hay forma por día de semana"
  );
  assert(!julyCloseOnly.blockers.some((item) => item.code === "previous-missing-daily"), "el cierre solo no debe bloquear el mes");

  const incompleteSync = assessForecastFreezeReadiness({
    selectedMonth: "2026-08",
    coverageRows: buildSalesMonthCoverage([...dailyRowsForMonth("2026-07", 31), closeRowForMonth("2026-07")]),
    databaseSync: { sales: { ok: false, error: "timeout" }, production: { ok: true, count: 1 }, waste: { ok: true, count: 1 } },
  });
  assert(incompleteSync.blockers.some((item) => item.code === "sync-incomplete"), "sync incompleto debe bloquear");

  const missingStatus = assessForecastFreezeReadiness({
    selectedMonth: "2026-08",
    coverageRows: buildSalesMonthCoverage([...dailyRowsForMonth("2026-07", 31), closeRowForMonth("2026-07")]),
    databaseSync: completeSync,
    catalogCount: 111,
    capturedStatuses: 0,
  });
  assert(missingStatus.canFreeze && missingStatus.warnings.some((item) => item.code === "missing-status"), "falta de estatus avisa pero no bloquea");
  assert(countCapturedProductStatuses({ "PAN DE MUERTO": { status: "ESTACIONAL" }, OTRO: {} }) === 1, "estatus capturado cuenta solo valores válidos");

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

  const healthStock = [
    { producto: "PINA GDE", stock: 20, orden: 1 },
    { producto: "FRUTAS GDE", stock: 20, orden: 2 },
  ];
  const healthVentas = [
    monthClose("2026-05", "PINA GDE", 300),
    monthClose("2026-05", "FRUTAS GDE", 500),
    monthClose("2026-06", "PINA GDE", 300),
    monthClose("2026-06", "FRUTAS GDE", 500),
    monthClose("2026-07", "PINA GDE", 310),
    monthClose("2026-07", "FRUTAS GDE", 520),
  ];

  const missingStock = buildForecastHealth({ ventas: healthVentas, selectedMonth: "2026-08" });
  assert(missingStock.checks.some((item) => item.code === "missing-stock"), "sin stock el control debe marcar error");
  assert(!missingStock.ready, "sin stock no está listo");

  const missingSales = buildForecastHealth({ stockRows: healthStock, selectedMonth: "2026-08" });
  assert(missingSales.checks.some((item) => item.code === "missing-sales"), "sin ventas el control debe marcar error");

  const health = buildForecastHealth({
    stockRows: healthStock,
    ventas: healthVentas,
    selectedMonth: "2026-08",
    dailyBufferPct: 10,
  });
  assert(health.currentTotal > 0, "con stock y cierres el pronóstico de agosto no puede ser cero");
  assert(health.backtests.length >= 1, "debe comparar al menos un mes oculto contra su venta real");
  assert(health.backtests.every((row) => row.wape !== null), "cada mes oculto debe tener WAPE");
  assert(health.backtests.every((row) => Array.isArray(row.topErrors)), "cada mes oculto debe listar productos con más error absoluto");
  assert(health.checks.some((item) => item.code === "missing-year-ago"), "sin 2025 debe avisar que no hay estacionalidad");
  assert(health.checks.some((item) => item.code === "forecast-ready"), "con total > 0 el control marca el modelo activo");

  const juneOnlyForecast = calculateForecast({
    stockRows: healthStock,
    historicalVentas: healthVentas.filter((row) => monthKey(row) === "2026-06"),
    bajas: [],
    existencias: [],
    realProduction: [],
    selectedMonth: "2026-07",
    dailyBufferPct: 10,
  });
  assert(juneOnlyForecast[0].pronosticoVenta > 0, "un solo cierre mensual debe producir pronóstico");
  assert(juneOnlyForecast[0].metodoPronostico !== "Sin histórico", "un cierre de junio no debe etiquetarse como sin histórico");

  const cheesecakeForecast = calculateForecast({
    stockRows: [{ producto: "CHEESECAKE GDE", stock: 20, orden: 1 }],
    historicalVentas: [monthClose("2026-06", "CHEESECAKE GDE", 180)],
    bajas: [],
    existencias: [],
    realProduction: [],
    selectedMonth: "2026-07",
    dailyBufferPct: 10,
  });
  assert(cheesecakeForecast[0].pronosticoVenta > 0, "CHEESECAKE del catálogo debe empatar con las ventas homologadas a CHESSECAKE");

  const rankedErrors = analyzeForecastProductErrors([
    { producto: "FRUTAS GDE", forecast: 400, actual: 520 },
    { producto: "GELATINA IND FRESA", forecast: 200, actual: 205 },
    { producto: "GALLETA NUEZ", forecast: 90, actual: 88 },
  ]);
  assert(rankedErrors.topErrors[0].producto === "FRUTAS GDE", "el análisis debe poner primero al producto con más error absoluto");
  assert(rankedErrors.topErrors[0].errorShare > 0.8, "FRUTAS GDE debe concentrar la mayor parte del error absoluto");
  assert(rankedErrors.wape > 10, "el WAPE del ejemplo de FRUTAS debe quedar por encima de 10");

  const juneDaily = [];
  const [juneYear, juneMonth] = [2026, 6];
  for (let day = 1; day <= 30; day += 1) {
    const fecha = new Date(juneYear, juneMonth - 1, day);
    const weekday = fecha.getDay();
    juneDaily.push({
      fecha,
      producto: "PINA GDE",
      cantidad: weekday === 6 ? 24 : weekday === 2 ? 8 : 12,
    });
  }
  const julyCloseRows = monthClose("2026-07", "PINA GDE", 400);
  const monthlyData = buildMonthlyForecastData([...juneDaily, julyCloseRows]);
  const julyData = monthlyData.get("2026-07");
  const saturday = julyData.weekdays.get(6);
  const tuesday = julyData.weekdays.get(2);
  assert(julyData.inheritedWeekdayShapeFrom === "2026-06", "el cierre de julio debe heredar la forma diaria de junio");
  assert(saturday.total / saturday.count > tuesday.total / tuesday.count, "el sábado heredado debe quedar por encima del martes");

  const accelDaily = [];
  for (let day = 1; day <= 30; day += 1) {
    accelDaily.push({
      fecha: new Date(2026, 5, day),
      producto: "FRUTAS GDE",
      cantidad: day <= 15 ? 223 / 15 : 296 / 15,
    });
  }
  const accelClose = monthClose("2026-06", "FRUTAS GDE", 519);
  const resolvedAccel = resolveCanonicalMonthSources([
    { name: "VENTAS DE MAYO Y JUNIO 2026.xlsx", rows: accelDaily },
    { name: "ventas junio.xlsx", rows: [accelClose] },
  ]);
  const resolvedAccelRows = resolvedAccel.entries.flatMap((entry) => entry.rows);
  assert(resolvedAccel.decisions[0].strategy === "keep-daily-and-close", "junio acelerado debe quedar diario+cierre");
  const accelMonthly = buildMonthlyForecastData(resolvedAccelRows);
  const juneAccel = accelMonthly.get("2026-06");
  assert(Math.abs(juneAccel.total - 519) < 1e-6, "el total de junio debe ser el del cierre, no la suma diaria");
  assert(!juneAccel.filledFromMonthlyTotal, "con diario complementario junio no se rellena en uniforme");
  const closeOnlyMonthly = buildMonthlyForecastData([accelClose]);
  assert(closeOnlyMonthly.get("2026-06").filledFromMonthlyTotal, "cierre solo sí rellena días sintéticos");
  assert(computeRecentMomentumFactor(closeOnlyMonthly, "2026-07") === 1, "sin diario el impulso de julio no corre");
  const accelFactor = computeRecentMomentumFactor(accelMonthly, "2026-07", { momentumTrigger: 1.1 });
  assert(accelFactor > 1.05, `diario+cierre debe disparar el impulso GDE de julio (factor ${accelFactor})`);

  assert(getProduccionSugerida("PINA GDE", 7.9) === 0, "menos de 8 no se produce");
  assert(getProduccionSugerida("PINA GDE", 8) === 10, "de 8 a 12 se hace lote 10");
  assert(getProduccionSugerida("DURAZNO GDE", 12.9) === 10, "12.9 se hace 10");
  assert(getProduccionSugerida("MOKA GDE", 13) === 15, "13 se hace 15");
  assert(getProduccionSugerida("FRUTAS GDE", 17.9) === 15, "17.9 se hace 15");
  assert(getProduccionSugerida("MOKA GDE", 18) === 20, "18 se hace 20");
  assert(getProduccionSugerida("GELATINA IND FRESA", 12.1) === 13, "lo que no es pastel se redondea hacia arriba");

  assert(isPriceTaggedProduct("PALETA GALLETA $35"), "PALETA GALLETA $35 debe detectarse como precio etiquetado");
  assert(isPriceTaggedProduct("JERICALLA UN CUARTO $45"), "JERICALLA $45 es precio etiquetado");
  assert(!isPriceTaggedProduct("GELATINA IND FRESA"), "GELATINA IND FRESA no trae etiqueta de precio");

  const paletaHistory = [
    monthClose("2026-01", "PALETA GALLETA $35", 30),
    monthClose("2026-02", "PALETA GALLETA $35", 184),
    monthClose("2026-03", "PALETA GALLETA $35", 2),
    monthClose("2026-04", "PALETA GALLETA $35", 0),
    monthClose("2026-05", "PALETA GALLETA $35", 79),
    monthClose("2026-06", "PALETA GALLETA $35", 296),
  ];
  const paletaRaw = calculateForecast({
    stockRows: [{ producto: "PALETA GALLETA $35", stock: 10, orden: 1 }],
    historicalVentas: paletaHistory,
    bajas: [],
    existencias: [],
    realProduction: [],
    selectedMonth: "2026-07",
    dailyBufferPct: 0,
  });
  assert(paletaRaw[0].pronosticoVenta < 180, `PALETA $35 no debe extrapolar el pico de junio (fc ${paletaRaw[0].pronosticoVenta.toFixed(1)})`);
  assert(/limpieza catálogo/i.test(paletaRaw[0].metodoPronostico || ""), "PALETA $35 debe anotar limpieza de catálogo");

  const cajitaHistory = [
    monthClose("2025-04", "CAJITA FELIZ", 599),
    monthClose("2026-01", "CAJITA FELIZ", 1),
    monthClose("2026-02", "CAJITA FELIZ", 8),
    monthClose("2026-03", "CAJITA FELIZ", 0),
    monthClose("2026-04", "CAJITA FELIZ", 773),
    monthClose("2026-05", "CAJITA FELIZ", 75),
    monthClose("2026-06", "CAJITA FELIZ", 0),
  ];
  const cajitaJuly = calculateForecast({
    stockRows: [{ producto: "CAJITA FELIZ", stock: 5, orden: 2 }],
    historicalVentas: cajitaHistory,
    bajas: [],
    existencias: [],
    realProduction: [],
    selectedMonth: "2026-07",
    dailyBufferPct: 0,
  });
  assert(cajitaJuly[0].pronosticoVenta <= 5, `CAJITA FELIZ con junio en cero no debe inventar volumen (fc ${cajitaJuly[0].pronosticoVenta.toFixed(1)})`);

  const dormantHistory = [
    monthClose("2026-02", "RAMO FLORAL 12CM 3 CAPAS", 84),
    monthClose("2026-03", "RAMO FLORAL 12CM 3 CAPAS", 0),
    monthClose("2026-04", "RAMO FLORAL 12CM 3 CAPAS", 0),
    monthClose("2026-05", "RAMO FLORAL 12CM 3 CAPAS", 0),
    monthClose("2026-06", "RAMO FLORAL 12CM 3 CAPAS", 0),
  ];
  const ramoJuly = calculateForecast({
    stockRows: [{ producto: "RAMO FLORAL 12CM 3 CAPAS", stock: 5, orden: 3 }],
    historicalVentas: dormantHistory,
    bajas: [],
    existencias: [],
    realProduction: [],
    selectedMonth: "2026-07",
    dailyBufferPct: 0,
  });
  assert(ramoJuly[0].pronosticoVenta === 0, `RAMO dormido 2+ meses debe quedar en 0 (fc ${ramoJuly[0].pronosticoVenta})`);

  const stableCake = calculateForecast({
    stockRows: [{ producto: "FRUTAS GDE", stock: 40, orden: 4 }],
    historicalVentas: [
      monthClose("2025-07", "FRUTAS GDE", 510),
      monthClose("2026-04", "FRUTAS GDE", 480),
      monthClose("2026-05", "FRUTAS GDE", 500),
      monthClose("2026-06", "FRUTAS GDE", 520),
    ],
    bajas: [],
    existencias: [],
    realProduction: [],
    selectedMonth: "2026-07",
    dailyBufferPct: 0,
  });
  assert(stableCake[0].pronosticoVenta > 450, `pasteles regulares no deben verse afectados por la limpieza (fc ${stableCake[0].pronosticoVenta.toFixed(1)})`);
  assert(!/limpieza catálogo/i.test(stableCake[0].metodoPronostico || ""), "FRUTAS GDE no debe activar limpieza de catálogo");


  console.log("parser-test ok");
}

main().catch((error) => {
  console.error(error);
  process.exit(1);
});
