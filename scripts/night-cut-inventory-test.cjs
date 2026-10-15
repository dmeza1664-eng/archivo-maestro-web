const fs = require("fs");
const path = require("path");
const Module = require("module");
const esbuild = require("esbuild");
const XLSX = require("xlsx");

const FIXTURE = path.join(__dirname, "fixtures", "inventario-hoy-2026-10-14.xlsx");
const PLANTILLA = path.join(__dirname, "fixtures", "plantilla-total-a-tener-2026-10-14.xlsx");
const APP_SOURCE = fs.readFileSync(path.join(__dirname, "..", "App.jsx"), "utf8");

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
  const appModule = new Module("night-cut-inventory-test");
  appModule.filename = path.join(__dirname, "night-cut-inventory-test.bundle.cjs");
  appModule.paths = module.paths;
  appModule._compile(built.outputFiles[0].text, appModule.filename);
  return appModule.exports;
}

function assert(condition, message) {
  if (!condition) throw new Error(message);
}

function product(parsed, name) {
  return parsed.products.find((row) => row.producto === name);
}

async function main() {
  assert(fs.existsSync(FIXTURE), "falta el Excel de ejemplo del corte de noche");
  const {
    parseNightCutWorkbook,
    parseDailyInventory,
    isNightCutWorkbook,
    matchNightCutProducts,
    nightCutFileDateLooksStale,
    productionDayForNightCut,
    buildNightCutOrderRows,
    sortNightCutOrderRows,
    summarizeNightCutOrder,
    buildNightCutExportRows,
    buildNightCutExportWorkbook,
    sumDailyBranchExcessByProductDate,
    branchStockKey,
    normalizeProduct,
    calculateDailyForecast,
    getProduccionSugerida,
    NIGHT_CUT_FORMULA,
    NIGHT_CUT_EXPORT_COLUMNS,
  } = await loadAppFunctions();

  assert(NIGHT_CUT_FORMULA.includes("margen de seguridad"), "la fórmula en pantalla dice margen de seguridad");
  assert(!/colch[oó]n/i.test(NIGHT_CUT_FORMULA), "la fórmula no dice colchón");
  assert(!/error\s*%/i.test(NIGHT_CUT_FORMULA), "la fórmula no es un porcentaje de error");
  assert(!/lalo|erick|ángel|angel/i.test(NIGHT_CUT_FORMULA), "la fórmula no lleva nombres de personas");

  const sectionStart = APP_SOURCE.indexOf("function NightCutOrderSection(");
  const dashboardStart = APP_SOURCE.indexOf("function Dashboard(");
  const appStart = APP_SOURCE.indexOf("\nfunction App()");
  const section = APP_SOURCE.slice(sectionStart, dashboardStart);
  const page = APP_SOURCE.slice(dashboardStart, appStart);
  assert(section.includes("Para el encargado de turno"), "el resultado del corte queda en la página");
  assert(section.includes("Calcular mandar a producir"), "el encargado confirma antes de calcular");
  assert(section.includes("Esa fecha se ve atrasada"), "se avisa si la fecha del archivo se ve atrasada");
  assert(section.includes("NIGHT_CUT_FORMULA"), "la fórmula de una línea está en la página");
  assert(section.includes("night-cut-equation"), "el desglose se ve en recuadros");
  assert(section.includes("Pronóstico"), "la tabla muestra el pronóstico");
  assert(section.includes("En sucursales"), "la tabla muestra la existencia en sucursales");
  assert(section.includes("Suma suc. y C.F."), "la tabla muestra la suma del Excel");
  assert(section.includes("Diferencia"), "la tabla muestra la diferencia contra el stock");
  assert(section.includes("sin margen"), "el pronóstico se marca sin margen");
  assert(section.includes("Margen de seguridad"), "el margen de seguridad es un recuadro");
  assert(section.includes('role="switch"'), "ocultar ceros es un interruptor con etiqueta");
  assert(section.includes("Ocultar productos en cero"), "el interruptor dice qué oculta");
  assert(section.includes("night-cut-dates"), "el encabezado muestra inventario y pedido");
  const resultAt = section.indexOf('className="night-cut-result"');
  const datesAt = section.indexOf("night-cut-dates", resultAt);
  const equationAt = section.indexOf("night-cut-equation", resultAt);
  const tableAt = section.indexOf("night-cut-table", resultAt);
  assert(resultAt > 0 && datesAt < equationAt && equationAt < tableAt, "el resultado muestra fechas, luego el desglose y luego la tabla");
  assert(section.includes("Pedido para"), "la fecha del pedido queda a la vista");
  assert(section.includes("night-cut-caption"), "la explicación queda en una línea corta");
  assert(!section.includes('className="night-cut-formula"'), "la fórmula larga ya no es un bloque gris");
  const css = fs.readFileSync(path.join(__dirname, "..", "style.css"), "utf8");
  assert(/\.night-cut-table\s*\{[^}]*min-width:\s*0/.test(css), "la tabla del corte no fuerza el ancho de las demás tablas");
  assert(css.includes(".night-cut-toggle"), "el interruptor de ceros tiene estilo propio");
  assert(/\.night-cut-table th\.produce-col[\s\S]*?text-align:\s*right/.test(css), "Mandar a producir queda alineado a la derecha");
  assert(page.includes("<NightCutOrderSection"), "el tablero muestra el pedido del corte");
  assert(!/Inventario del día/.test(page), "no regresa el panel viejo de inventario del día");
  assert(!/colch[oó]n/i.test(section), "el bloque del encargado no dice colchón");
  assert(!/lalo|erick/i.test(section), "el bloque del encargado no nombra personas");

  const workbook = XLSX.readFile(FIXTURE, { cellDates: true });
  assert(isNightCutWorkbook(workbook), "el archivo de ejemplo es el corte de noche");
  const parsed = parseNightCutWorkbook(workbook);
  assert(parsed.fileDate === "2026-10-06", `A1 se lee pero no se usa como fecha del pedido (${parsed.fileDate})`);
  assert(nightCutFileDateLooksStale(parsed.fileDate, "2026-10-15"), "6 oct se ve atrasada si hoy es 15 oct");
  assert(nightCutFileDateLooksStale(parsed.fileDate, "2026-10-14"), "6 oct también se ve atrasada el 14 oct");
  assert(!nightCutFileDateLooksStale("2026-10-14", "2026-10-15"), "ayer no se marca como atrasada");
  assert(!nightCutFileDateLooksStale("2026-10-15", "2026-10-15"), "hoy no se marca como atrasada");
  assert(productionDayForNightCut("2026-10-14") === "2026-10-15", "el corte del miércoles 14 pide el jueves 15");
  assert(productionDayForNightCut("2026-10-17") === "2026-10-19", "el sábado pide el lunes");
  assert(productionDayForNightCut("2026-10-18") === "2026-10-19", "el domingo pide el lunes");

  const pina = product(parsed, "PINA GDE");
  const frutas = product(parsed, "FRUTAS GDE");
  const moka = product(parsed, "MOKA GDE");
  assert(pina, "PIÑA GDE entra al corte");
  assert(pina.existenciaSucursales === 34, `PIÑA existencia ${pina.existenciaSucursales}`);
  assert(pina.cuartoFrio === 5, `PIÑA cuarto frío restante es 5, no la otra columna (${pina.cuartoFrio})`);
  assert(pina.sobra === 0, `PIÑA no sobra: 34 en sucursales contra stock 50 (${pina.sobra})`);
  assert(pina.piso === 34 && pina.sumaSucCf === 39 && pina.stockAf === 50 && pina.stockAg === 40 && pina.diferenciaStock === -11, "PIÑA coteja el total a tener: suma 39, stock 50 y 40, diferencia -11");
  assert(pina.codigo === "02", `PIÑA se reconoce por código 02 (${pina.codigo})`);
  assert(pina.sucursales.length === 0, "el total a tener no se reparte como piso de cada sucursal");
  assert(frutas.existenciaSucursales === 50 && frutas.cuartoFrio === 20 && frutas.sobra === 0, "FRUTAS GDE: 50 en sucursales, 0 de sobra, 20 de cuarto frío");
  assert(moka.existenciaSucursales === 57 && moka.cuartoFrio === 10 && moka.sobra === 0, "MOKA GDE: 57 en sucursales, sobra 0, cuarto frío 10");

  assert(!parsed.products.some((row) => /3 LECHES GRANDES|3 LECHES MEDIANOS|3 LECHES CHICOS|MINI MEDIANOS 3 LECHES|MINI MINI 3 LECHES|TOTAL AREA|ESPECIALIDAD/.test(row.producto)), "los subtotales del archivo no son productos");
  assert(parsed.products.some((row) => row.producto === "GELATINA IND FRESA"), "un producto real después de un subtotal se conserva");
  assert(parsed.products.some((row) => row.producto === normalizeProduct("ROSCA C/CAJETA")), "ROSCA se conserva");
  const zanahoriaPetit = product(parsed, "ZANAHORIA PETIT");
  assert(zanahoriaPetit && zanahoriaPetit.cuartoFrio === 4 && zanahoriaPetit.existenciaSucursales === 44, "ZANAHORIA PETIT toma el cuarto frío y la existencia del total a tener");
  assert(!parsed.branchRows.some((row) => /lalo|erick/i.test(row.sucursal)), "el nombre de una hoja de persona no se vuelve sucursal");
  assert(!parsed.branchRows.some((row) => /total|cuarto|gral|suma|stock|tener/i.test(row.sucursal)), "totales y metas no entran como sucursal");
  assert(!parsed.coldRows.some((row) => row.producto === "PINA GDE" && row.cantidad !== 5), "el cuarto frío de PIÑA no se toma de ATRASADOS ni de la hoja que va después");

  const dated = parseDailyInventory(workbook, "2026-10-14");
  assert(dated.format === "corte-noche", "el lector diario reconoce este libro");
  assert(dated.fileDate === "2026-10-06", "sigue reportando la fecha de A1");
  assert(dated.branchStock.length === 0, "el total a tener no carga un piso por sucursal");
  assert(dated.coldRoom.length > 0 && dated.coldRoom.every((row) => row.fecha === "2026-10-14"), "la fecha la pone quien confirma, no A1");
  assert(dated.coldRoom.find((row) => row.producto === "PINA GDE")?.cantidad === 5, "el cuarto frío fechado sigue siendo el restante");
  const undated = parseDailyInventory(workbook, "");
  assert(undated.branchStock.length === 0 && undated.coldRoom.length === 0, "sin fecha confirmada no se carga el corte con la fecha vieja de A1");
  assert(undated.products.length === parsed.products.length, "aunque no haya fecha, los productos sí se leen");
  assert(pina.sobra === 0, "la sobra de PIÑA usa el total general contra el stock, no el resumen oculto");

  const codigoMap = require("../datos/catalogo/stock-codigo-map.json");
  const muertos = require("../datos/catalogo/temporada-dia-muertos.json");
  const official = [...new Set([
    ...Object.values(codigoMap).filter((row) => row.inCatalog).map((row) => normalizeProduct(row.producto)),
    ...muertos.productos.map((name) => normalizeProduct(name)),
  ])];
  const { matched, unmatched } = matchNightCutProducts(parsed.products, official);
  assert(matched.some((row) => row.producto === "PINA GDE"), "PIÑA coincide con el catálogo");
  assert(matched.some((row) => row.producto === normalizeProduct("PAN MUERTO AZUCAR GDE 250GR")), "el pan de muerto de temporada coincide por nombre");
  assert(!unmatched.some((row) => /3 LECHES GRANDES|TOTAL AREA/.test(normalizeProduct(row.productoOriginal))), "un subtotal no se reporta como producto sin catálogo");

  const orderRows = buildNightCutOrderRows([
    {
      fecha: "2026-10-15",
      producto: "PINA GDE",
      baseSinLoteDia: 20,
      hasDailyColdRoom: true,
      cuartoFrioDia: 5,
      hasDailyBranchStock: true,
      excedenteSucursalesDia: 4,
      aProducirDia: 11,
    },
    {
      fecha: "2026-10-14",
      producto: "OTRO",
      baseSinLoteDia: 9,
      aProducirDia: 9,
    },
  ], "2026-10-15");
  assert(orderRows.length === 1 && orderRows[0].mandar === 11 && orderRows[0].recomendado === 20, "la tabla del encargado es la del día de producción");
  assert(orderRows[0].cuartoFrio === 5 && orderRows[0].sobra === 4, "la tabla muestra cuarto frío y sobra");
  assert(Math.abs(orderRows[0].pronostico - 20 / 1.1) < 1e-9, "el pronóstico es el recomendado sin el margen de seguridad");
  assert(Math.abs(orderRows[0].pronostico + orderRows[0].margen - orderRows[0].recomendado) < 1e-9, "pronóstico más margen de seguridad es el recomendado");
  const conPronostico = buildNightCutOrderRows([
    {
      fecha: "2026-10-15",
      producto: "PINA GDE",
      pronosticoVentaDia: 20,
      colchonDiario: 2,
      baseConColchonDia: 22,
      baseSinLoteDia: 22,
      aProducirDia: 20,
    },
  ], "2026-10-15", 10);
  assert(conPronostico[0].pronostico === 20 && conPronostico[0].margen === 2, "si el día ya trae pronóstico, se muestra ese número");
  const sabado = buildNightCutOrderRows([
    {
      fecha: "2026-10-17",
      producto: "FRUTAS GDE",
      pronosticoVentaDia: 20,
      colchonDiario: 2,
      baseConColchonDia: 22,
      baseSinLoteDia: 33,
      aProducirDia: 30,
    },
  ], "2026-10-17", 10);
  assert(Math.abs(sabado[0].pronostico - 30) < 1e-9 && Math.abs(sabado[0].margen - 3) < 1e-9, "el sábado incluye el pronóstico del domingo, todavía sin margen");
  const promo = buildNightCutOrderRows([
    {
      fecha: "2026-10-15",
      producto: "GELATINA IND FRESA",
      promoActiva: true,
      pronosticoVentaDia: 10,
      colchonDiario: 1,
      baseSinLoteDia: 15,
      aProducirDia: 12,
    },
  ], "2026-10-15", 10);
  assert(promo[0].pronostico === 10 && promo[0].margen === 1 && promo[0].recomendado === 15, "la promo no se disfraza de margen de seguridad");
  const sorted = sortNightCutOrderRows([
    { producto: "MOKA GDE", mandar: 2, recomendado: 1, pronostico: 0.9, margen: 0.1, cuartoFrio: 0, sobra: 0 },
    { producto: "PINA GDE", mandar: 11, recomendado: 3, pronostico: 2.7, margen: 0.3, cuartoFrio: 0, sobra: 0 },
  ], { key: "pronostico", dir: "desc" });
  assert(sorted[0].producto === "PINA GDE", "se puede ordenar por pronóstico");
  const byMandar = sortNightCutOrderRows(sorted, { key: "mandar", dir: "desc" });
  assert(byMandar[0].producto === "PINA GDE", "se puede ordenar por Mandar a producir");
  const totals = summarizeNightCutOrder(byMandar);
  assert(totals.mandar === 13 && totals.productos === 2 && Math.abs(totals.pronostico - 3.6) < 1e-9, "el total suma pronóstico y lo que se manda");
  const exported = buildNightCutExportRows(byMandar);
  assert(Object.keys(exported[0]).join("|") === NIGHT_CUT_EXPORT_COLUMNS.join("|"), "el Excel incluye Pronóstico, la suma y el stock del corte");
  assert(NIGHT_CUT_EXPORT_COLUMNS.includes("Pronóstico") && NIGHT_CUT_EXPORT_COLUMNS.includes("Suma suc. y C.F.") && NIGHT_CUT_EXPORT_COLUMNS.includes("Diferencia vs stock"), "el Excel deja cotejar suma, stock y diferencia");
  assert(exported.at(-1).Producto === "TOTAL" && exported.at(-1)["Mandar a producir"] === 13, "el Excel cierra con el total");
  assert(exported.at(-1)["Pronóstico"] === 3.6, "el total del Excel incluye el pronóstico");
  assert(!exported.some((row) => Object.keys(row).some((key) => /error|%/i.test(key))), "el Excel no lleva columnas de error");
  assert(!exported.some((row) => Object.keys(row).some((key) => /colch[oó]n/i.test(key))), "el Excel no dice colchón");
  const exportWorkbook = buildNightCutExportWorkbook(byMandar, { productionDate: "2026-10-16", inventoryDate: "2026-10-15" });
  const note = XLSX.utils.sheet_to_json(exportWorkbook.Sheets.Nota, { header: 1 });
  assert(String(note[2][1]).includes("margen de seguridad") && !/colch[oó]n/i.test(String(note[2][1])), "la nota del Excel conserva la fórmula");

  assert(fs.existsSync(PLANTILLA), "falta la plantilla anonimizada del total a tener");
  const plantillaBook = XLSX.readFile(PLANTILLA, { cellDates: true });
  assert(isNightCutWorkbook(plantillaBook), "la plantilla es corte de noche");
  const hiddenNames = (plantillaBook.Workbook?.Sheets || [])
    .map((meta, index) => ({ name: plantillaBook.SheetNames[index], hidden: meta?.Hidden }))
    .filter((meta) => meta.hidden === 1 || meta.hidden === 2);
  assert(hiddenNames.some((meta) => /EXIST/.test(meta.name) && /RESTANTE/.test(meta.name)), "el resumen viejo va oculto después del total");
  const plantilla = parseNightCutWorkbook(plantillaBook);
  assert(plantilla.fileDate === "2026-10-06", "A1 de la plantilla se lee y no fecha el pedido");
  const durazno = product(plantilla, "DURAZNO GDE");
  const pinero = product(plantilla, normalizeProduct("MINI MEDIANO PIÑERO"));
  assert(durazno, "DURAZNO GDE está en el total a tener");
  assert(pinero, "MINI MEDIANO PIÑERO está en el total a tener");
  assert(durazno.existenciaSucursales === 36 && durazno.cuartoFrio === 9 && durazno.sumaSucCf === 45 && durazno.stockAf === 60 && durazno.stockAg === 50 && durazno.diferenciaStock === -15, `DURAZNO GDE debe ser Z 36, AB 9, AD 45, AF 60, AG 50, AH -15 (${JSON.stringify({ z: durazno.existenciaSucursales, ab: durazno.cuartoFrio, ad: durazno.sumaSucCf, af: durazno.stockAf, ag: durazno.stockAg, ah: durazno.diferenciaStock })})`);
  assert(durazno.sobra === 0, "DURAZNO está corto de stock: no hay sobra que restar");
  assert(pinero.existenciaSucursales === 113 && pinero.cuartoFrio === 22 && pinero.sumaSucCf === 135 && pinero.stockAf === 180 && pinero.stockAg === 180 && pinero.diferenciaStock === -45, `MINI MEDIANO PIÑERO debe ser Z 113, AB 22, AD 135, AF 180, AG 180, AH -45 (${JSON.stringify({ z: pinero.existenciaSucursales, ab: pinero.cuartoFrio, ad: pinero.sumaSucCf, af: pinero.stockAf, ag: pinero.stockAg, ah: pinero.diferenciaStock })})`);
  assert(pinero.sobra === 0, "MINI MEDIANO PIÑERO está corto de stock: no hay sobra que restar");
  assert(!plantilla.products.some((row) => /3 LECHES GRANDES|3 LECHES MEDIANOS|3 LECHES CHICOS|MINI MEDIANOS 3 LECHES|MINI MINI 3 LECHES|TOTAL AREA|ESPECIALIDAD/.test(row.producto)), "los subtotales de la plantilla no son productos");
  assert(plantilla.products.every((row) => row.corteTotal), "el corte sale de la hoja de total a tener");
  const conSobra = plantilla.products.find((row) => row.sobra > 0 && row.stockAf !== null);
  assert(conSobra && conSobra.sobra === conSobra.existenciaSucursales - conSobra.stockAf, "la sobra es existencia menos stock, cuando la existencia es mayor");

  const scrambled = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(scrambled, XLSX.utils.aoa_to_sheet([
    ["2020-01-01"],
    ["STOCK", "PRODUCTO", "LO QUE VA A QUEDAR EN CUARTO FRIO", "SUMA SUC. Y C.F.", "TOTAL GRAL.", "STOCK"],
    [60, "DURAZNO GDE", 9, 45, 36, 50],
  ]), "TOTAL A TENER SUC.(EXIST.+DIST)");
  XLSX.utils.book_append_sheet(scrambled, XLSX.utils.aoa_to_sheet([
    ["2020-01-01"],
    ["PRODUCTO", "TOTAL EXISTENCIA EN SUCURSALES", "RESTANTE CUARTO FRIO"],
    ["DURAZNO GDE", 29, 1],
  ]), "EXIST. SUCURSALES Y RESTANTE CF");
  XLSX.utils.book_append_sheet(scrambled, XLSX.utils.aoa_to_sheet([
    ["PRODUCTO", "MATRIZ", "PLAZA", "VISTAS"],
    ["DURAZNO GDE", 1, 1, 1],
  ]), "EXISTENCIA EN SUCURSALES");
  scrambled.Workbook = scrambled.Workbook || {};
  scrambled.Workbook.Sheets = [
    { Hidden: 0 },
    { Hidden: 1 },
    { Hidden: 0 },
  ];
  const moved = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(moved, scrambled.Sheets["EXISTENCIA EN SUCURSALES"], "EXISTENCIA EN SUCURSALES");
  XLSX.utils.book_append_sheet(moved, scrambled.Sheets["EXIST. SUCURSALES Y RESTANTE CF"], "EXIST. SUCURSALES Y RESTANTE CF");
  XLSX.utils.book_append_sheet(moved, scrambled.Sheets["TOTAL A TENER SUC.(EXIST.+DIST)"], "TOTAL A TENER SUC.(EXIST.+DIST)");
  moved.Workbook = { Sheets: [{ Hidden: 0 }, { Hidden: 1 }, { Hidden: 0 }] };
  const byHeader = parseNightCutWorkbook(moved);
  const duraznoMoved = product(byHeader, "DURAZNO GDE");
  assert(duraznoMoved && duraznoMoved.existenciaSucursales === 36 && duraznoMoved.cuartoFrio === 9 && duraznoMoved.stockAf === 60 && duraznoMoved.stockAg === 50 && duraznoMoved.diferenciaStock === -15, "las columnas se encuentran por el texto del encabezado, no por la letra");

  const weekday = {
    promedioLunes: 20,
    promedioMartes: 20,
    promedioMiercoles: 20,
    promedioJueves: 20,
    promedioViernes: 20,
    promedioSabado: 20,
    promedioDomingo: 0,
    produccionSugerida: 400,
  };
  const daily = calculateDailyForecast({
    monthlyRows: [durazno, pinero].map((row) => ({ producto: row.producto, ...weekday })),
    ventasReales: [],
    realProduction: [],
    selectedMonth: "2026-10",
    dailyBufferPct: 10,
    nightCutInventory: [durazno, pinero].map((row) => ({ ...row, fecha: "2026-10-15" })),
  });
  const order = buildNightCutOrderRows(daily, "2026-10-15", 10);
  const duraznoOrder = order.find((row) => row.producto === "DURAZNO GDE");
  const pineroOrder = order.find((row) => row.producto === pinero.producto);
  assert(duraznoOrder.recomendado === 22 && duraznoOrder.cuartoFrio === 9 && duraznoOrder.sobra === 0 && duraznoOrder.existencia === 36 && duraznoOrder.sumaSucCf === 45 && duraznoOrder.stockAf === 60 && duraznoOrder.diferenciaStock === -15, "DURAZNO muestra el Excel y resta solo el cuarto frío");
  assert(duraznoOrder.mandar === getProduccionSugerida("DURAZNO GDE", 22 - 9), "DURAZNO manda el lote de pronóstico + margen − 9");
  assert(Math.abs(duraznoOrder.pronostico - 20) < 1e-9 && Math.abs(duraznoOrder.margen - 2) < 1e-9, "DURAZNO sigue mostrando el pronóstico sin el margen");
  assert(pineroOrder.cuartoFrio === 22 && pineroOrder.existencia === 113 && pineroOrder.diferenciaStock === -45 && pineroOrder.mandar === 0, "MINI MEDIANO PIÑERO resta los 22 de cuarto frío y no manda si el neto no llega a lote");
  const fullDaily = calculateDailyForecast({
    monthlyRows: plantilla.products.map((row) => ({ producto: row.producto, ...weekday })),
    ventasReales: [],
    realProduction: [],
    selectedMonth: "2026-10",
    dailyBufferPct: 10,
    nightCutInventory: plantilla.products.map((row) => ({ ...row, fecha: "2026-10-15" })),
  });
  const fullOrder = buildNightCutOrderRows(fullDaily, "2026-10-15", 10);
  const fullTotals = summarizeNightCutOrder(fullOrder);
  const inOrder = new Set(fullOrder.map((row) => row.producto));
  const forecastable = plantilla.products.filter((row) => inOrder.has(row.producto));
  const cold = forecastable.reduce((sum, row) => sum + row.cuartoFrio, 0);
  const surplus = forecastable.reduce((sum, row) => sum + row.sobra, 0);
  assert(forecastable.length === fullOrder.length && fullTotals.cuartoFrio === cold && fullTotals.sobra === surplus, "el pedido resta el cuarto frío y la sobra de los productos que sí se pronostican");
  assert(fullTotals.mandar === fullOrder.reduce((sum, row) => sum + row.mandar, 0), "el total a mandar suma la tabla");

  console.log("night-cut-inventory-test ok");
  console.log(`DURAZNO GDE Z=${durazno.existenciaSucursales} AB=${durazno.cuartoFrio} AD=${durazno.sumaSucCf} AF=${durazno.stockAf} AG=${durazno.stockAg} AH=${durazno.diferenciaStock} sobra=${durazno.sobra} mandar(base 22)=${duraznoOrder.mandar}`);
  console.log(`MINI MEDIANO PIÑERO (${pinero.producto}) Z=${pinero.existenciaSucursales} AB=${pinero.cuartoFrio} AD=${pinero.sumaSucCf} AF=${pinero.stockAf} AG=${pinero.stockAg} AH=${pinero.diferenciaStock} sobra=${pinero.sobra} mandar(base 22)=${pineroOrder.mandar}`);
  console.log(`plantilla productos=${plantilla.products.length} cuartoFrio=${cold} sobra=${surplus} diferencia=${plantilla.products.reduce((sum, row) => sum + (row.diferenciaStock || 0), 0)} mandar con pronóstico 20=${fullTotals.mandar}`);
  console.log(`ejemplo 2026-10-14 → pedido 2026-10-15: PIÑA existencia ${pina.existenciaSucursales}, sobra ${pina.sobra}, CF ${pina.cuartoFrio}; FRUTAS existencia ${frutas.existenciaSucursales}, sobra ${frutas.sobra}, CF ${frutas.cuartoFrio}; MOKA existencia ${moka.existenciaSucursales}, sobra ${moka.sobra}, CF ${moka.cuartoFrio}; sin catálogo con piezas ${unmatched.filter((row) => row.existenciaSucursales > 0 || row.cuartoFrio > 0).length}`);
}

main().catch((error) => {
  console.error(error);
  process.exit(1);
});
