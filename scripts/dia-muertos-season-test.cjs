/**
 * Temporada de pan de muerto / calabaza (decisión 2026-10-03, Ángel).
 * Datos sintéticos + el respaldo embebido de sucursales 2025; no usa Planta León.
 */
const path = require("path");
const Module = require("module");
const esbuild = require("esbuild");
const XLSX = require("xlsx");
const temporada = require("../datos/catalogo/temporada-dia-muertos.json");

async function loadApp() {
  const built = await esbuild.build({
    entryPoints: [path.join(__dirname, "..", "App.jsx")],
    bundle: true,
    platform: "node",
    format: "cjs",
    write: false,
    loader: { ".css": "text", ".json": "json" },
    define: {
      "import.meta.env.VITE_API_URL": JSON.stringify(""),
      "import.meta.env.DEV": "false",
      "import.meta.env.PROD": "true",
    },
    logLevel: "silent",
  });
  const appModule = new Module("dia-muertos-season-test");
  appModule.filename = path.join(__dirname, "dia-muertos-season-test.bundle.cjs");
  appModule.paths = module.paths;
  appModule._compile(built.outputFiles[0].text, appModule.filename);
  return appModule.exports;
}

function assert(condition, message) {
  if (!condition) throw new Error(`[temporada muertos] ${message}`);
}

function workbookFromSheets(sheets) {
  const workbook = XLSX.utils.book_new();
  for (const [name, rows] of Object.entries(sheets)) {
    XLSX.utils.book_append_sheet(workbook, XLSX.utils.aoa_to_sheet(rows), name);
  }
  return workbook;
}

function close(monthKey, producto, cantidad, extra = {}) {
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

(async () => {
  const app = await loadApp();
  const names = temporada.productos;
  assert(names.length === 16, "deben ser 16 SKU de temporada");

  for (const name of names) {
    assert(app.isDiaDeMuertosSeasonalProduct(name), `${name} es ESTACIONAL de muertos`);
  }
  assert(app.isDiaDeMuertosSeasonalProduct("PAN DE MUERTO CHOCOLATE GDE"), "chocolate GDE es ESTACIONAL");
  assert(app.isDiaDeMuertosSeasonalProduct("PAN DE MUERTO IND CHOCOLATE"), "chocolate IND es ESTACIONAL");
  assert(app.isDiaDeMuertosForecastMonth("2026-10"), "octubre es mes de pronóstico");
  assert(app.isDiaDeMuertosForecastMonth("2026-11"), "noviembre es mes de pronóstico");
  assert(!app.isDiaDeMuertosForecastMonth("2026-08"), "agosto no es temporada");
  assert(app.isDiaDeMuertosSeasonDate("2026-10-03"), "3 oct está en temporada");
  assert(app.isDiaDeMuertosSeasonDate("2026-11-02"), "2 nov está en temporada");
  assert(!app.isDiaDeMuertosSeasonDate("2026-11-03"), "3 nov ya cerró la temporada");

  const currentSheetOnly = workbookFromSheets({
    "TOTAL A TENER SUC.(EXIST.+DIST)": [
      ["PRODUCTO", "STOCK"],
      ["BOLILLO", 100],
      ["PINA GDE", 20],
    ],
  });
  const parsedCurrent = app.parseStock(currentSheetOnly);
  for (const name of names) {
    const row = parsedCurrent.find((item) => app.normalizeProduct(item.producto) === app.normalizeProduct(name));
    assert(row, `${name} entra aunque la hoja actual no lo traiga`);
    assert(row.estatusOperativo === "ESTACIONAL" || app.isDiaDeMuertosSeasonalProduct(row.producto), `${name} queda ESTACIONAL`);
  }
  assert(parsedCurrent.some((row) => row.producto === "BOLILLO" && row.stock === 100), "el catálogo regular no se toca");

  const bothSheets = workbookFromSheets({
    "TOTAL A TENER SUC.(EXIST.+DIST)": [
      ["PRODUCTO", "STOCK"],
      ["BOLILLO", 100],
    ],
    "EXIST. SUCURSALES Y RESTANTE CF": [
      ["PRODUCTO", "STOCK"],
      ["BOLILLO", 80],
      ["PAN MUERTO IND AZUCAR 50GR", 4],
    ],
  });
  const parsedBoth = app.parseStock(bothSheets);
  const azucarStock = parsedBoth.find((row) => /PAN MUERTO IND AZUCAR/.test(row.producto));
  assert(azucarStock?.stock === 4, "si la hoja vieja trae stock, se usa");
  const drift = app.assessStockSheetSelection(bothSheets);
  assert(drift.missingTotal === 0, "ya no se pide rehacer la hoja por los 16 de muertos");

  const smallCatalog = [{ producto: "FRUTAS GDE", stock: 10, orden: 1 }];
  const august = app.calculateForecast({
    stockRows: smallCatalog,
    historicalVentas: [close("2026-07", "FRUTAS GDE", 400)],
    bajas: [],
    existencias: [],
    realProduction: [],
    selectedMonth: "2026-08",
    dailyBufferPct: 10,
  });
  assert(august.length === 1 && august[0].producto === "FRUTAS GDE", "en agosto el catálogo sintético no se llena de muertos");

  const realish = [
    { producto: "BOLILLO", stock: 10, orden: 1 },
    ...Array.from({ length: 40 }, (_, index) => ({ producto: `SKU REGULAR ${index + 1}`, stock: 1, orden: index + 2 })),
  ];
  const october = app.calculateForecast({
    stockRows: realish,
    historicalVentas: [
      close("2026-09", "PAN MUERTO IND AZUCAR 50GR", 555, { sucursal: "Suc. León" }),
      close("2026-09", "PAN MUERTO IND AZUCAR 50GR", 465, { sucursal: "Planta León" }),
    ],
    bajas: [],
    existencias: [],
    realProduction: [],
    selectedMonth: "2026-10",
    dailyBufferPct: 10,
  });
  const seasonalOct = october.filter((row) => app.isDiaDeMuertosSeasonalProduct(row.producto));
  assert(seasonalOct.length === 16, `octubre emite 16 filas de temporada (obtuvo ${seasonalOct.length})`);
  assert(seasonalOct.every((row) => row.estatusOperativo === "ESTACIONAL"), "las 16 filas van ESTACIONAL");

  const azucar = seasonalOct.find((row) => /PAN MUERTO IND AZUCAR/.test(row.producto));
  assert(azucar, "PAN MUERTO IND AZUCAR 50GR tiene fila de octubre");
  assert(azucar.pronosticoVenta > 2000, `azúcar octubre usa 2025 y el ritmo de sep-2026 (${azucar.pronosticoVenta})`);
  assert(azucar.pronosticoVenta < 2900, `el crecimiento de azúcar queda acotado (${azucar.pronosticoVenta})`);

  const chocolate = seasonalOct.find((row) => app.normalizeProduct(row.producto) === app.normalizeProduct("PAN DE MUERTO CHOCOLATE GDE"));
  assert(chocolate, "el chocolate GDE aparece en octubre");
  assert(chocolate.pronosticoVenta === 0, "chocolate GDE no vendió la temporada pasada: 0");

  const plantaOnly = app.calculateForecast({
    stockRows: realish,
    historicalVentas: [close("2026-09", "PAN MUERTO IND AZUCAR 50GR", 9999, { sucursal: "Planta León · Piso de venta" })],
    bajas: [],
    existencias: [],
    realProduction: [],
    selectedMonth: "2026-10",
    dailyBufferPct: 10,
  });
  const azucarSinPlanta = plantaOnly.find((row) => /PAN MUERTO IND AZUCAR/.test(row.producto));
  assert(azucarSinPlanta.pronosticoValidacionModelo === 555, "la semilla de sep-2026 es sucursal; Planta no sustituye el ritmo");
  assert(azucarSinPlanta.pronosticoVenta === azucar.pronosticoVenta, "Planta León no cambia el pronóstico de muertos");

  const december = app.calculateForecast({
    stockRows: realish,
    historicalVentas: [close("2026-09", "PAN MUERTO IND AZUCAR 50GR", 555, { sucursal: "Suc. León" })],
    bajas: [],
    existencias: [],
    realProduction: [],
    selectedMonth: "2026-12",
    dailyBufferPct: 10,
  });
  assert(december.every((row) => !app.isDiaDeMuertosSeasonalProduct(row.producto)), "en diciembre ya no salen filas de muertos");

  const november = app.calculateForecast({
    stockRows: realish,
    historicalVentas: [close("2026-10", "PAN MUERTO IND AZUCAR 50GR", 2000, { sucursal: "Suc. León" })],
    bajas: [],
    existencias: [],
    realProduction: [],
    selectedMonth: "2026-11",
    dailyBufferPct: 10,
  });
  const azucarNov = november.find((row) => /PAN MUERTO IND AZUCAR/.test(row.producto));
  assert(azucarNov, "noviembre todavía lista el SKU");
  assert(azucarNov.pronosticoVenta === 0, "nov-2025 fue 0: el mensual de noviembre queda en 0");

  const dailyNov = app.calculateDailyForecast({
    monthlyRows: november,
    ventasReales: [],
    realProduction: [],
    selectedMonth: "2026-11",
    dailyBufferPct: 10,
  });
  const afterSeason = dailyNov.filter((row) => /PAN MUERTO IND AZUCAR/.test(row.producto) && row.fecha > "2026-11-02");
  assert(afterSeason.length > 0, "hay días posteriores al 2 nov");
  assert(afterSeason.every((row) => row.aProducirDia === 0 && row.pronosticoVentaDia === 0), "después del 2 nov Mandar a producir es 0");

  const dailyOct = app.calculateDailyForecast({
    monthlyRows: october,
    ventasReales: [],
    realProduction: [],
    selectedMonth: "2026-10",
    dailyBufferPct: 10,
  });
  const azucarDays = dailyOct.filter((row) => /PAN MUERTO IND AZUCAR/.test(row.producto));
  assert(azucarDays.length === 31, "octubre tiene fila diaria del azúcar");
  assert(azucarDays.some((row) => (row.aProducirDia ?? row.produccionSugeridaDia) > 0), "en octubre Mandar a producir tiene piezas");

  const review = app.buildMonthlyReviewRows({
    sourceRows: october.map((row) => ({
      producto: row.producto,
      pronosticoBase: row.pronosticoVenta,
      pronosticoOperativo: row.produccionSugerida,
      orden: row.orden,
    })),
    forecastRows: october,
    historicalVentas: [],
    loadedExistencias: [],
    inputs: {},
    selectedMonth: "2026-10",
  });
  const reviewAzucar = review.find((row) => /PAN MUERTO IND AZUCAR/.test(row.producto));
  assert(reviewAzucar?.status === "ESTACIONAL", "la revisión marca ESTACIONAL sin captura manual");
  assert(reviewAzucar.proposed > 0, "en temporada ESTACIONAL sí propone producir");

  const validation = app.buildProductValidationSummary([
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
  const dist = validation.find((row) => app.isDistribucionChannel(row.producto));
  const bolillo = validation.find((row) => row.producto === "BOLILLO");
  assert(dist && dist.errorPct === null, "Distribución nunca muestra porcentaje de error");
  assert(bolillo && bolillo.errorPct !== null, "un producto de sucursal sí puede tener error %");
  assert(app.isDistribucionChannel("Distribución"), "se reconoce el canal Distribución");
  assert(!app.isDistribucionChannel("Suc. León"), "una sucursal no es Distribución");

  const demand = app.filterDemandSales([
    close("2026-09", "PAN MUERTO IND AZUCAR 50GR", 555, { sucursal: "Suc. León" }),
    close("2026-09", "PAN MUERTO IND AZUCAR 50GR", 465, { sucursal: "Planta León" }),
  ]);
  assert(demand.length === 1 && demand[0].cantidad === 555, "la demanda de muertos es solo sucursal");

  console.log("dia-muertos-season-test ok");
})().catch((error) => {
  console.error(error);
  process.exit(1);
});
