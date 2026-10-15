const fs = require("fs");
const path = require("path");
const Module = require("module");
const esbuild = require("esbuild");
const XLSX = require("xlsx");

const FIXTURE = path.join(__dirname, "fixtures", "inventario-hoy-2026-10-14.xlsx");
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
    NIGHT_CUT_FORMULA,
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
  assert(pina.existenciaSucursales === 24, `PIÑA existencia ${pina.existenciaSucursales}`);
  assert(pina.cuartoFrio === 5, `PIÑA cuarto frío restante es 5, no la otra columna (${pina.cuartoFrio})`);
  assert(pina.sobra === 4, `PIÑA sobra 4 del stock fijo (${pina.sobra})`);
  assert(pina.piso === 24, "el piso por sucursal de PIÑA suma el total");
  assert(pina.codigo === "02", `PIÑA se reconoce por código 02 (${pina.codigo})`);
  assert(pina.sucursales.find((row) => branchStockKey(row.sucursal) === "MATRIZ")?.cantidad === 1, "MATRIZ de PIÑA es existencia 1, no el stock ni la distribución");
  assert(frutas.existenciaSucursales === 41 && frutas.cuartoFrio === 20 && frutas.sobra === 0, "FRUTAS GDE: 41 en sucursales, 0 de sobra, 20 de cuarto frío");
  assert(moka.existenciaSucursales === 42 && moka.cuartoFrio === 10 && moka.sobra === 2, "MOKA GDE: 42 en sucursales, sobra 2, cuarto frío 10");

  assert(!parsed.products.some((row) => /3 LECHES GRANDES|3 LECHES MEDIANOS|3 LECHES CHICOS|MINI MEDIANOS 3 LECHES|MINI MINI 3 LECHES|TOTAL AREA|ESPECIALIDAD/.test(row.producto)), "los subtotales del archivo no son productos");
  assert(parsed.products.some((row) => row.producto === "GELATINA IND FRESA"), "un producto real después de un subtotal se conserva");
  assert(parsed.products.some((row) => row.producto === normalizeProduct("ROSCA C/CAJETA")), "ROSCA se conserva");
  const zanahoriaPetit = product(parsed, "ZANAHORIA PETIT");
  assert(zanahoriaPetit && zanahoriaPetit.cuartoFrio === 0, "un restante negativo de cuarto frío no suma piezas");
  assert(!parsed.branchRows.some((row) => /lalo|erick/i.test(row.sucursal)), "el nombre de una hoja de persona no se vuelve sucursal");
  assert(!parsed.branchRows.some((row) => /total|cuarto|gral|suma|stock|tener/i.test(row.sucursal)), "totales y metas no entran como sucursal");
  assert(!parsed.coldRows.some((row) => row.producto === "PINA GDE" && row.cantidad !== 5), "el cuarto frío de PIÑA no se toma de ATRASADOS ni de la hoja que va después");

  const dated = parseDailyInventory(workbook, "2026-10-14");
  assert(dated.format === "corte-noche", "el lector diario reconoce este libro");
  assert(dated.fileDate === "2026-10-06", "sigue reportando la fecha de A1");
  assert(dated.branchStock.length > 0 && dated.branchStock.every((row) => row.fecha === "2026-10-14"), "la fecha la pone quien confirma, no A1");
  assert(dated.coldRoom.find((row) => row.producto === "PINA GDE")?.cantidad === 5, "el cuarto frío fechado sigue siendo el restante");
  const undated = parseDailyInventory(workbook, "");
  assert(undated.branchStock.length === 0 && undated.coldRoom.length === 0, "sin fecha confirmada no se carga el corte con la fecha vieja de A1");
  assert(undated.products.length === parsed.products.length, "aunque no haya fecha, los productos sí se leen");

  const targetMap = new Map(pina.targets.map((target) => [`PINA GDE|${branchStockKey(target.sucursal)}`, target.stock]));
  const excess = sumDailyBranchExcessByProductDate(
    dated.branchStock.filter((row) => row.producto === "PINA GDE"),
    targetMap
  );
  assert(excess.get("2026-10-14|PINA GDE")?.excedente === 4, "la sobra de PIÑA usa la misma cuenta que Mandar a producir");

  const codigoMap = require("../datos/catalogo/stock-codigo-map.json");
  const muertos = require("../datos/catalogo/temporada-dia-muertos.json");
  const official = [...new Set([
    ...Object.values(codigoMap).filter((row) => row.inCatalog).map((row) => normalizeProduct(row.producto)),
    ...muertos.productos.map((name) => normalizeProduct(name)),
  ])];
  const { matched, unmatched } = matchNightCutProducts(parsed.products, official);
  assert(matched.some((row) => row.producto === "PINA GDE"), "PIÑA coincide con el catálogo");
  assert(matched.some((row) => row.producto === normalizeProduct("PAN DE MUERTO CHOCOLATE GDE")), "el pan de muerto de temporada coincide por nombre");
  const nuez = unmatched.find((row) => normalizeProduct(row.productoOriginal) === normalizeProduct("NUEZ TRES LECHES MED"));
  assert(nuez && nuez.existenciaSucursales === 40 && nuez.cuartoFrio === 3, "NUEZ TRES LECHES MED queda sin catálogo, con sus piezas");
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
  assert(Object.keys(exported[0]).join("|") === ["Producto", "Pronóstico", "Recomendado", "Cuarto frío restante", "Sobra en sucursales", "Mandar a producir"].join("|"), "el Excel incluye Pronóstico junto a Recomendado");
  assert(exported.at(-1).Producto === "TOTAL" && exported.at(-1)["Mandar a producir"] === 13, "el Excel cierra con el total");
  assert(exported.at(-1)["Pronóstico"] === 3.6, "el total del Excel incluye el pronóstico");
  assert(!exported.some((row) => Object.keys(row).some((key) => /error|%/i.test(key))), "el Excel no lleva columnas de error");
  assert(!exported.some((row) => Object.keys(row).some((key) => /colch[oó]n/i.test(key))), "el Excel no dice colchón");
  const exportWorkbook = buildNightCutExportWorkbook(byMandar, { productionDate: "2026-10-16", inventoryDate: "2026-10-15" });
  const note = XLSX.utils.sheet_to_json(exportWorkbook.Sheets.Nota, { header: 1 });
  assert(String(note[2][1]).includes("margen de seguridad") && !/colch[oó]n/i.test(String(note[2][1])), "la nota del Excel conserva la fórmula");

  console.log("night-cut-inventory-test ok");
  console.log(`ejemplo 2026-10-14 → pedido 2026-10-15: PIÑA existencia ${pina.existenciaSucursales}, sobra ${pina.sobra}, CF ${pina.cuartoFrio}; FRUTAS existencia ${frutas.existenciaSucursales}, sobra ${frutas.sobra}, CF ${frutas.cuartoFrio}; MOKA existencia ${moka.existenciaSucursales}, sobra ${moka.sobra}, CF ${moka.cuartoFrio}; sin catálogo con piezas ${unmatched.filter((row) => row.existenciaSucursales > 0 || row.cuartoFrio > 0).length}`);
}

main().catch((error) => {
  console.error(error);
  process.exit(1);
});
