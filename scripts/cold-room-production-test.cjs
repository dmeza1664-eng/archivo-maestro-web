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
  const appModule = new Module("cold-room-production-test");
  appModule.filename = path.join(__dirname, "cold-room-production-test.bundle.cjs");
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
    parseDailyInventory,
    parseDailyBranchStock,
    parseDailyColdRoom,
    sanitizeDailyColdRoom,
    upsertDailyColdRoom,
    mapDailyColdRoomByProductDate,
    applyDailyBranchStockToPlantSuggestion,
    applyInventoryToProductionSuggestion,
    applyPromoUpliftToQuantity,
    calculateForecast,
    calculateDailyForecast,
    normalizeActivePromo,
  } = await loadAppFunctions();

  assert(applyInventoryToProductionSuggestion(20, 8, 5) === 7, "20 − 8 suc − 5 CF = 7 a producir");
  assert(applyInventoryToProductionSuggestion(15, 8, 20) === 0, "si sucursales + CF cubren el target, A producir = 0");
  assert(applyInventoryToProductionSuggestion(13, 0, 0) === 13, "sin inventario, A producir = bruto");
  assert(applyInventoryToProductionSuggestion(10, 4) === 6, "CF omitido se trata como 0");
  assert(
    applyDailyBranchStockToPlantSuggestion(20, 8) === applyInventoryToProductionSuggestion(20, 8, 0),
    "pedido planta sigue siendo bruto − sucursales, sin restar CF"
  );
  assert(
    applyInventoryToProductionSuggestion(
      applyPromoUpliftToQuantity("FRUTAS GDE", 10, { multiplier: 1.3, extraPiecesPerDay: 0 }),
      8,
      3
    ) === 4,
    "promo 10×1.3 → lote 15, luego 15 − 8 − 3 = 4; no se vuelve a lote"
  );
  assert(sanitizeDailyColdRoom([{ fecha: "no", producto: "X", cantidad: 2 }]).length === 0, "fecha inválida de CF se descarta");
  assert(
    upsertDailyColdRoom(
      [{ fecha: "2026-09-20", producto: "FRUTAS GDE", cantidad: 4 }],
      [{ fecha: "2026-09-20", producto: "FRUTAS GDE", cantidad: 9 }]
    )[0].cantidad === 9,
    "upsert de CF pisa el mismo día/SKU"
  );

  const raizWide = parseDailyInventory(workbookFromSheets({
    RAIZ: [
      ["Producto", "Centro", "Norte", "Total suc", "Cuarto frío", "Suma suc+CF"],
      ["FRUTAS GDE", 8, 2, 10, 5, 15],
      ["GELATINA IND FRESA", 0, 6, 6, 0, 6],
      ["MOKA GDE", 3, "", 3, "", 3],
    ],
  }), "2026-09-20");
  assert(raizWide.branchStock.find((row) => row.producto === "FRUTAS GDE" && row.sucursal === "Centro")?.cantidad === 8, "RAIZ ancho lee sucursal");
  assert(!raizWide.branchStock.some((row) => /total|cuarto|suma/i.test(row.sucursal)), "total, CF y suma no entran como sucursal");
  assert(raizWide.coldRoom.find((row) => row.producto === "FRUTAS GDE")?.cantidad === 5, "RAIZ ancho lee restante de cuarto frío");
  assert(raizWide.coldRoom.find((row) => row.producto === "GELATINA IND FRESA")?.cantidad === 0, "CF en 0 se conserva");
  assert(!raizWide.coldRoom.some((row) => row.producto === "MOKA GDE"), "celda vacía de CF no se guarda");

  const existenciasLike = parseDailyInventory(workbookFromSheets({
    "EXISTENCIA EN SUCURSALES": [
      ["PRODUCTO", "TOTAL GRAL SUC", "C.F.", "SUMA SUC+CF"],
      ["BOLILLO", 10, 2, 12],
    ],
  }), "2026-09-21");
  assert(existenciasLike.branchStock.find((row) => row.producto === "BOLILLO")?.cantidad === 10, "hoja tipo existencias usa total sucursales");
  assert(existenciasLike.branchStock[0].sucursal === "Sucursales", "sin cruce de tiendas, el total vive como Sucursales");
  assert(existenciasLike.coldRoom.find((row) => row.producto === "BOLILLO")?.cantidad === 2, "C.F. de existencias entra a cuarto frío");

  const longCold = parseDailyColdRoom(workbookFromSheets({
    Inventario: [
      ["Fecha", "Sucursal", "Producto", "Cantidad", "Cuarto frío"],
      ["2026-09-20", "Centro", "FRUTAS GDE", 8, 4],
      ["2026-09-20", "Norte", "FRUTAS GDE", 2, 4],
    ],
  }));
  assert(longCold.find((row) => row.producto === "FRUTAS GDE")?.cantidad === 4, "columna CF del formato largo no se duplica por sucursal");

  const namedCold = parseDailyColdRoom(workbookFromSheets({
    "Cuarto frío": [
      ["Producto", "Cantidad"],
      ["FRUTAS GDE", 7],
    ],
  }), "2026-09-21");
  assert(namedCold.find((row) => row.producto === "FRUTAS GDE")?.cantidad === 7, "hoja llamada Cuarto frío lee restante de planta");
  assert(parseDailyBranchStock(workbookFromSheets({
    "Cuarto frío": [
      ["Producto", "Cantidad"],
      ["FRUTAS GDE", 7],
    ],
  }), "2026-09-21").length === 0, "la hoja de CF no se guarda como sucursal");

  const locationCold = parseDailyInventory(workbookFromSheets({
    Inventario: [
      ["Fecha", "Sucursal", "Producto", "Cantidad"],
      ["2026-09-20", "Cuarto frío", "MOKA GDE", 11],
      ["2026-09-20", "Centro", "MOKA GDE", 3],
    ],
  }));
  assert(locationCold.coldRoom.find((row) => row.producto === "MOKA GDE")?.cantidad === 11, "sucursal Cuarto frío se enruta a CF");
  assert(locationCold.branchStock.find((row) => row.sucursal === "Centro")?.cantidad === 3, "la sucursal real sigue en sucursales");
  assert(!locationCold.branchStock.some((row) => /cuarto/i.test(row.sucursal)), "CF no contamina el pedido de sucursales");

  const fullRaiz = parseDailyInventory(workbookFromSheets({
    "TOTAL A TENER SUC.(EXIST.+DIST)": [
      ["PRODUCTO", "SUC", "STOCK"],
      ["FRUTAS GDE", 2, 40],
      ["GELATINA IND FRESA", 2, 80],
    ],
    "STOCK DE SUCURSALES": [
      ["PRODUCTO", "STOCK"],
      ["FRUTAS GDE", 40],
    ],
    "EXIST. SUCURSALES Y RESTANTE CF": [
      ["PRODUCTO", "Centro", "Norte", "RESTANTE"],
      ["FRUTAS GDE", 8, 2, 5],
      ["GELATINA IND FRESA", 0, 6, 0],
    ],
    "TOTAL LOCALES": [
      ["PRODUCTO", "STOCK"],
      ["FRUTAS GDE", 99],
    ],
    "TOTAL FORANEAS": [
      ["PRODUCTO", "CANTIDAD"],
      ["FRUTAS GDE", 77],
    ],
    "TOTAL GRAL": [
      ["PRODUCTO", "TOTAL GRAL"],
      ["FRUTAS GDE", 120],
    ],
  }), "2026-09-21");
  assert(!fullRaiz.branchStock.some((row) => /total|tener|locales|forane|gral|suma|stock/i.test(row.sucursal)), "TOTAL A TENER, TOTAL GRAL, LOCALES/FORANEAS y sumas no son sucursal");
  assert(fullRaiz.branchStock.find((row) => row.producto === "FRUTAS GDE" && row.sucursal === "Centro")?.cantidad === 8, "el workbook RAIZ sigue leyendo sucursales reales");
  assert(fullRaiz.branchStock.find((row) => row.producto === "FRUTAS GDE" && row.sucursal === "Norte")?.cantidad === 2, "Norte del cruce RAIZ no se pierde");
  assert(fullRaiz.coldRoom.find((row) => row.producto === "FRUTAS GDE")?.cantidad === 5, "RESTANTE sin CF en EXIST. SUCURSALES entra a cuarto frío");
  assert(fullRaiz.coldRoom.find((row) => row.producto === "GELATINA IND FRESA")?.cantidad === 0, "RESTANTE 0 de la hoja CF se conserva");
  assert(!fullRaiz.branchStock.some((row) => row.cantidad === 40 || row.cantidad === 99 || row.cantidad === 77 || row.cantidad === 120), "las metas y rollups no se cuelan como inventario");

  const restanteSheet = parseDailyColdRoom(workbookFromSheets({
    RESTANTE: [
      ["Producto", "Cantidad"],
      ["BOLILLO", 4],
    ],
  }), "2026-09-21");
  assert(restanteSheet.find((row) => row.producto === "BOLILLO")?.cantidad === 4, "hoja RESTANTE sin CF entra a cuarto frío");
  assert(parseDailyBranchStock(workbookFromSheets({
    RESTANTE: [
      ["Producto", "Cantidad"],
      ["BOLILLO", 4],
    ],
  }), "2026-09-21").length === 0, "la hoja RESTANTE no se guarda como sucursal");

  const restanteCfCol = parseDailyInventory(workbookFromSheets({
    "EXIST. SUCURSALES Y RESTANTE CF": [
      ["Producto", "Centro", "Norte", "RESTANTE CF"],
      ["MOKA GDE", 3, 1, 7],
    ],
  }), "2026-09-20");
  assert(restanteCfCol.branchStock.find((row) => row.sucursal === "Centro")?.cantidad === 3, "sucursales de EXIST. SUCURSALES siguen en el cruce");
  assert(restanteCfCol.coldRoom.find((row) => row.producto === "MOKA GDE")?.cantidad === 7, "columna Restante CF de esa hoja va a cuarto frío");
  assert(!restanteCfCol.branchStock.some((row) => /restante|exist/i.test(row.sucursal)), "ni RESTANTE ni el nombre de hoja se vuelven sucursal");

  const totalSucRestante = parseDailyInventory(workbookFromSheets({
    Inventario: [
      ["PRODUCTO", "TOTAL GRAL SUC", "RESTANTE"],
      ["BOLILLO", 10, 2],
    ],
  }), "2026-09-21");
  assert(totalSucRestante.branchStock.find((row) => row.producto === "BOLILLO")?.cantidad === 10, "patrón total suc + RESTANTE lee sucursales");
  assert(totalSucRestante.branchStock[0].sucursal === "Sucursales", "sin cruce de tiendas el total vive como Sucursales");
  assert(totalSucRestante.coldRoom.find((row) => row.producto === "BOLILLO")?.cantidad === 2, "RESTANTE junto a total sucursales es cuarto frío");

  const catalogOnly = parseDailyInventory(workbookFromSheets({
    "TOTAL A TENER SUC.(EXIST.+DIST)": [
      ["PRODUCTO", "STOCK"],
      ["BOLILLO", 113],
    ],
    "EXIST. SUCURSALES Y RESTANTE CF": [
      ["PRODUCTO", "STOCK"],
      ["BOLILLO", 80],
    ],
  }), "2026-09-21");
  assert(catalogOnly.branchStock.length === 0, "fotos de catálogo TOTAL A TENER / STOCK no se importan como inventario del día");
  assert(catalogOnly.coldRoom.length === 0, "STOCK de catálogo no se toma como restante de CF");

  const coldMap = mapDailyColdRoomByProductDate([
    { fecha: "2026-09-20", producto: "FRUTAS GDE", cantidad: 5 },
    { fecha: "2026-09-21", producto: "FRUTAS GDE", cantidad: 1 },
  ]);
  assert(coldMap.get("2026-09-20|FRUTAS GDE") === 5, "el mapa de CF es por día y SKU");
  assert(coldMap.get("2026-09-21|FRUTAS GDE") === 1, "otro día de CF no se mezcla");

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
  const gelatinaPromo = normalizeActivePromo({
    producto: "GELATINA IND FRESA",
    startDate: "2026-07-01",
    durationPreset: "3dias",
    multiplier: 2,
    extraPiecesPerDay: 5,
    active: true,
  });
  const gelatinaDaily = calculateDailyForecast({
    monthlyRows: gelatinaRows,
    ventasReales: [],
    realProduction: [],
    selectedMonth: "2026-07",
    dailyBufferPct: 0,
    activePromos: [gelatinaPromo],
    dailyBranchStock: [
      { fecha: "2026-07-01", sucursal: "Centro", producto: "GELATINA IND FRESA", cantidad: 10 },
      { fecha: "2026-07-01", sucursal: "Norte", producto: "GELATINA IND FRESA", cantidad: 4 },
    ],
    dailyColdRoom: [
      { fecha: "2026-07-01", producto: "GELATINA IND FRESA", cantidad: 6 },
      { fecha: "2026-07-02", producto: "GELATINA IND FRESA", cantidad: 99 },
    ],
  });
  const july1 = gelatinaDaily.find((row) => row.fecha === "2026-07-01");
  const july2 = gelatinaDaily.find((row) => row.fecha === "2026-07-02");
  const july4 = gelatinaDaily.find((row) => row.fecha === "2026-07-04");
  const july5 = gelatinaDaily.find((row) => row.fecha === "2026-07-05");
  assert(july1.inventarioSucursalesDia === 14, "el 1 de julio sigue sumando sucursales");
  assert(july1.cuartoFrioDia === 6, "el 1 de julio registra restante de cuarto frío");
  assert(july1.produccionSugeridaDia === applyDailyBranchStockToPlantSuggestion(july1.produccionBrutaDia, 14), "pedido planta no resta cuarto frío");
  assert(july1.aProducirDia === applyInventoryToProductionSuggestion(july1.produccionBrutaDia, 14, 6), "A producir = max(0, bruto − sucursales − CF)");
  assert(july1.hasDailyColdRoom, "el 1 de julio queda marcado con CF");
  assert(
    july2.produccionSugeridaDia === july2.produccionBrutaDia,
    "sin sucursales ese día, el pedido planta no cambia"
  );
  assert(
    july2.aProducirDia === applyInventoryToProductionSuggestion(july2.produccionBrutaDia, 0, 99),
    "CF solo ese día baja A producir y no el pedido"
  );
  assert(july4.aProducirDia === july4.produccionSugeridaDia, "sin captura, A producir = pedido = bruto");
  assert(july5.weekday === 0 && july5.aProducirDia === 0, "domingo sigue en cero aunque haya CF otros días");

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
    dailyColdRoom: [
      { fecha: "2026-07-01", producto: "FRUTAS GDE", cantidad: 3 },
    ],
  });
  const cakeDay = cakeDaily.find((row) => row.fecha === "2026-07-01");
  assert(cakeDay.produccionBrutaDia >= 10, "el bruto de pastel sigue saliendo en lote");
  assert(
    cakeDay.produccionSugeridaDia === applyDailyBranchStockToPlantSuggestion(cakeDay.produccionBrutaDia, 8),
    "pedido de pastel no reaplica lote al restar sucursales"
  );
  assert(
    cakeDay.aProducirDia === applyInventoryToProductionSuggestion(cakeDay.produccionBrutaDia, 8, 3),
    "A producir de pastel = lote − sucursales − CF, sin rearmar a 10/15/20"
  );

  const otherSkuCold = calculateDailyForecast({
    monthlyRows: gelatinaRows,
    ventasReales: [],
    realProduction: [],
    selectedMonth: "2026-07",
    dailyBufferPct: 0,
    dailyColdRoom: [
      { fecha: "2026-07-01", producto: "FRUTAS GDE", cantidad: 50 },
    ],
  });
  const gelatinaBase = calculateDailyForecast({
    monthlyRows: gelatinaRows,
    ventasReales: [],
    realProduction: [],
    selectedMonth: "2026-07",
    dailyBufferPct: 0,
  }).find((row) => row.fecha === "2026-07-01");
  assert(
    otherSkuCold.find((row) => row.fecha === "2026-07-01").aProducirDia === gelatinaBase.aProducirDia,
    "el CF de otro SKU no mueve A producir"
  );

  console.log("cold-room-production-test ok");
}

main().catch((event) => {
  console.error(event);
  process.exit(1);
});
