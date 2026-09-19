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
  const appModule = new Module("daily-branch-stock-test");
  appModule.filename = path.join(__dirname, "daily-branch-stock-test.bundle.cjs");
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

function monthClose(month, producto, cantidad) {
  const [year, monthNumber] = month.split("-").map(Number);
  return {
    fecha: new Date(year, monthNumber - 1, 1),
    producto,
    cantidad,
    monthlyTotal: true,
    monthDays: new Date(year, monthNumber, 0).getDate(),
  };
}

async function main() {
  const {
    parseDailyBranchStock,
    sanitizeDailyBranchStock,
    upsertDailyBranchStock,
    sumDailyBranchStockByProductDate,
    applyDailyBranchStockToPlantSuggestion,
    applyPromoUpliftToQuantity,
    collectSucursales,
    calculateForecast,
    calculateDailyForecast,
    normalizeActivePromo,
  } = await loadAppFunctions();

  const longBook = workbookFromSheets({
    Inventario: [
      ["Fecha", "Sucursal", "Producto", "Cantidad"],
      ["2026-09-19", "Centro", "FRUTAS GDE", 8],
      ["2026-09-19", "Norte", "FRUTAS GDE", 2],
      ["2026-09-19", "Centro", "GELATINA IND FRESA", 10],
      ["2026-09-19", "Centro", "FRUTAS GDE", 5],
      ["2026-09-20", "Centro", "FRUTAS GDE", 4],
      ["2026-09-19", "Centro", "", 3],
    ],
  });
  const longRows = parseDailyBranchStock(longBook);
  const frutasCentro = longRows.find((row) => row.producto === "FRUTAS GDE" && row.sucursal === "Centro" && row.fecha === "2026-09-19");
  assert(frutasCentro?.cantidad === 5, "la segunda fila del mismo día/sucursal/SKU pisa a la primera");
  assert(longRows.filter((row) => row.fecha === "2026-09-19" && row.producto === "FRUTAS GDE").length === 2, "Centro y Norte conviven el mismo día");
  assert(!longRows.some((row) => !row.producto), "filas sin producto se omiten");

  const wideBook = workbookFromSheets({
    "19-09-2026": [
      ["Producto", "Centro", "Norte", "Sur"],
      ["MOKA GDE", 3, "", 1],
      ["GELATINA IND FRESA", 0, 6, 2],
    ],
  });
  const wideRows = parseDailyBranchStock(wideBook, "2026-09-01");
  assert(wideRows.find((row) => row.producto === "MOKA GDE" && row.sucursal === "Centro")?.cantidad === 3, "formato ancho lee sucursal-columna");
  assert(!wideRows.some((row) => row.producto === "MOKA GDE" && row.sucursal === "Norte"), "celda vacía del cruce no se guarda");
  assert(wideRows.find((row) => row.producto === "GELATINA IND FRESA" && row.sucursal === "Norte")?.cantidad === 6, "0 y positivos del cruce se conservan");
  assert(wideRows.every((row) => row.fecha === "2026-09-19"), "la fecha del nombre de hoja alimenta el formato ancho");

  const merged = upsertDailyBranchStock(longRows, [
    { fecha: "2026-09-19", sucursal: "Centro", producto: "FRUTAS GDE", cantidad: 1 },
    { fecha: "2026-09-19", sucursal: "Sur", producto: "FRUTAS GDE", cantidad: 9 },
  ]);
  assert(merged.find((row) => row.producto === "FRUTAS GDE" && row.sucursal === "Centro" && row.fecha === "2026-09-19")?.cantidad === 1, "upsert pisa el mismo key");
  assert(merged.find((row) => row.sucursal === "Sur")?.cantidad === 9, "upsert agrega sucursal nueva");

  const totals = sumDailyBranchStockByProductDate(merged);
  assert(totals.get("2026-09-19|FRUTAS GDE") === 1 + 2 + 9, "el total del día suma sucursales");
  assert(totals.get("2026-09-20|FRUTAS GDE") === 4, "otro día no se mezcla");
  assert(!totals.has("2026-09-19|MOKA GDE"), "un SKU ausente no entra al mapa");

  assert(applyDailyBranchStockToPlantSuggestion(15, 8) === 7, "pastel 15 con 8 en tienda pide 7, sin rearmar lote");
  assert(applyDailyBranchStockToPlantSuggestion(15, 20) === 0, "si la sucursal cubre el día, el pedido queda en 0");
  assert(applyDailyBranchStockToPlantSuggestion(13, 0) === 13, "stock 0 capturado no cambia el bruto");
  assert(
    applyDailyBranchStockToPlantSuggestion(
      applyPromoUpliftToQuantity("FRUTAS GDE", 10, { multiplier: 1.3, extraPiecesPerDay: 0 }),
      8
    ) === 7,
    "promo 10×1.3 → lote 15, luego 15-8=7; no se vuelve a lote"
  );
  assert(sanitizeDailyBranchStock([{ fecha: "no", producto: "X", sucursal: "Y", cantidad: 2 }]).length === 0, "fecha inválida se descarta");
  assert(
    collectSucursales({
      ventas: [{ sucursal: "Centro" }],
      bajas: [{ canal: "Norte" }],
      dailyBranchStock: [{ sucursal: "Sur" }],
      extra: [" Centro ", "Oriente"],
    }).join(",") === "Centro,Norte,Oriente,Sur",
    "sucursales se unen y ordenan sin duplicar"
  );

  const gelatinaHistory = [
    monthClose("2026-04", "GELATINA IND FRESA", 210),
    monthClose("2026-05", "GELATINA IND FRESA", 220),
    monthClose("2026-06", "GELATINA IND FRESA", 230),
  ];
  const gelatinaRows = calculateForecast({
    stockRows: [{ producto: "GELATINA IND FRESA", stock: 20, orden: 1 }],
    historicalVentas: gelatinaHistory,
    bajas: [],
    existencias: [],
    realProduction: [],
    selectedMonth: "2026-07",
    dailyBufferPct: 0,
  });
  const gelatinaDailyBase = calculateDailyForecast({
    monthlyRows: gelatinaRows,
    ventasReales: [],
    realProduction: [],
    selectedMonth: "2026-07",
    dailyBufferPct: 0,
  });
  const gelatinaPromo = normalizeActivePromo({
    producto: "GELATINA IND FRESA",
    startDate: "2026-07-01",
    durationPreset: "3dias",
    multiplier: 2,
    extraPiecesPerDay: 5,
    active: true,
  });
  const gelatinaDailyPromo = calculateDailyForecast({
    monthlyRows: gelatinaRows,
    ventasReales: [],
    realProduction: [],
    selectedMonth: "2026-07",
    dailyBufferPct: 0,
    activePromos: [gelatinaPromo],
  });
  const july1Base = gelatinaDailyBase.find((row) => row.fecha === "2026-07-01");
  const july1Promo = gelatinaDailyPromo.find((row) => row.fecha === "2026-07-01");
  const expectedPromo = july1Base.produccionSugeridaDia * 2 + 5;
  assert(july1Promo.produccionSugeridaDia === expectedPromo, "la promo sigue aplicando igual si no hay inventario diario");

  const gelatinaDailyStock = calculateDailyForecast({
    monthlyRows: gelatinaRows,
    ventasReales: [],
    realProduction: [],
    selectedMonth: "2026-07",
    dailyBufferPct: 0,
    activePromos: [gelatinaPromo],
    dailyBranchStock: [
      { fecha: "2026-07-01", sucursal: "Centro", producto: "GELATINA IND FRESA", cantidad: 10 },
      { fecha: "2026-07-01", sucursal: "Norte", producto: "GELATINA IND FRESA", cantidad: 4 },
      { fecha: "2026-07-02", sucursal: "Centro", producto: "GELATINA IND FRESA", cantidad: 99 },
    ],
  });
  const july1Stock = gelatinaDailyStock.find((row) => row.fecha === "2026-07-01");
  const july2Stock = gelatinaDailyStock.find((row) => row.fecha === "2026-07-02");
  const july2Promo = gelatinaDailyPromo.find((row) => row.fecha === "2026-07-02");
  const july4Stock = gelatinaDailyStock.find((row) => row.fecha === "2026-07-04");
  const july4Promo = gelatinaDailyPromo.find((row) => row.fecha === "2026-07-04");
  const july5Stock = gelatinaDailyStock.find((row) => row.fecha === "2026-07-05");
  assert(july1Stock.produccionBrutaDia === expectedPromo, "el bruto conserva el uplift de promo antes de restar stock");
  assert(july1Stock.inventarioSucursalesDia === 14, "el inventario del 1 de julio suma sucursales");
  assert(july1Stock.produccionSugeridaDia === expectedPromo - 14, "pedido neto = promo − stock del día");
  assert(july1Stock.hasDailyBranchStock, "el 1 de julio queda marcado como día con inventario");
  assert(
    july2Stock.produccionSugeridaDia === applyDailyBranchStockToPlantSuggestion(july2Promo.produccionSugeridaDia, 99),
    "otro día usa su propio stock y no queda negativo"
  );
  assert(july4Stock.produccionSugeridaDia === july4Promo.produccionSugeridaDia, "sin captura ese día, el pedido no cambia");
  assert(july5Stock.weekday === 0 && july5Stock.produccionSugeridaDia === 0, "domingo sigue en cero aunque haya stock otros días");

  const cakeRows = calculateForecast({
    stockRows: [{ producto: "FRUTAS GDE", stock: 40, orden: 1 }],
    historicalVentas: [
      monthClose("2026-04", "FRUTAS GDE", 300),
      monthClose("2026-05", "FRUTAS GDE", 310),
      monthClose("2026-06", "FRUTAS GDE", 320),
    ],
    bajas: [],
    existencias: [],
    realProduction: [],
    selectedMonth: "2026-07",
    dailyBufferPct: 0,
  });
  const cakeDaily = calculateDailyForecast({
    monthlyRows: cakeRows,
    ventasReales: [],
    realProduction: [],
    selectedMonth: "2026-07",
    dailyBufferPct: 0,
    dailyBranchStock: [
      { fecha: "2026-07-01", sucursal: "Centro", producto: "FRUTAS GDE", cantidad: 8 },
    ],
  });
  const cakeDay = cakeDaily.find((row) => row.fecha === "2026-07-01");
  assert(cakeDay.produccionBrutaDia >= 10, "el bruto de pastel sigue saliendo en lote");
  assert(
    cakeDay.produccionSugeridaDia === applyDailyBranchStockToPlantSuggestion(cakeDay.produccionBrutaDia, 8),
    "después de restar stock no se reaplica el lote de 10/15/20"
  );
  assert(cakeDay.produccionSugeridaDia === cakeDay.produccionBrutaDia - 8, "pedido de pastel = lote − 8 en sucursal");

  const otherSkuStock = calculateDailyForecast({
    monthlyRows: gelatinaRows,
    ventasReales: [],
    realProduction: [],
    selectedMonth: "2026-07",
    dailyBufferPct: 0,
    dailyBranchStock: [
      { fecha: "2026-07-01", sucursal: "Centro", producto: "FRUTAS GDE", cantidad: 50 },
    ],
  });
  assert(
    otherSkuStock.find((row) => row.fecha === "2026-07-01").produccionSugeridaDia === july1Base.produccionSugeridaDia,
    "el stock de otro SKU no mueve el pedido"
  );

  console.log("daily-branch-stock-test ok");
}

main().catch((event) => {
  console.error(event);
  process.exit(1);
});
