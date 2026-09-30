// Página simplificada (sep 2026): solo carga de archivos, salud mínima, promo y la
// tabla diaria con «Mandar a producir». Las secciones de seguimiento/auditoría y el
// inventario del día salen de la interfaz; la lógica compartida sigue exportada.
const assert = require("assert");
const fs = require("fs");
const path = require("path");
const Module = require("module");
const esbuild = require("esbuild");

const ROOT = path.join(__dirname, "..");
const app = fs.readFileSync(path.join(ROOT, "App.jsx"), "utf8");
const css = fs.readFileSync(path.join(ROOT, "style.css"), "utf8");

// Solo el JSX de Dashboard (desde su return hasta function App) para no confundir
// textos de exportación a Excel o de la lógica con lo que se ve en la página.
const dashboardStart = app.indexOf("function Dashboard(");
const appStart = app.indexOf("\nfunction App()");
const dashboard = app.slice(dashboardStart, appStart);
const renderStart = dashboard.lastIndexOf('    <div className="app">');
assert(renderStart > 0, "Se encontró el render de Dashboard.");
const page = dashboard.slice(renderStart);

// Lo que se queda
assert.match(app, /function SectionDisclosure/, "El disclosure nativo sigue disponible.");
assert.match(page, /className=\{`promo-section compact-analytics-section/, "Promo activa sigue como accordion.");
assert.match(page, /Más opciones/, "Usuarios y filtros del Excel mensual viven en Más opciones.");
assert.match(page, /Paso 1 · Datos/, "La carga de datos es el primer paso visible.");
assert.match(page, /title="Stock fijo"/, "Se sube el stock fijo.");
assert.match(page, /title="Ventas"/, "Se suben las ventas.");
assert.match(page, /Paso 2 · Salud/, "La salud mínima del pronóstico sigue visible.");
assert.match(page, /<ForecastHealthStrip/, "La tira de salud (meses de venta, WAPE del último cierre) sigue.");
assert.match(page, /<FrozenMonthBanner/, "Un mes congelado se declara en la UI.");
assert.match(app, /no cambian si cargas más datos/, "El freeze avisa que planta no deriva en silencio.");
assert.match(page, /<FreezeReadinessStrip/, "Se explica por qué no se puede congelar.");
assert.match(page, /Paso 3 · Planta/, "La tabla diaria es el tercer paso visible.");
assert.match(page, /Guardar ventas en la base/, "Se pueden guardar las ventas en la base para no depender del respaldo.");
assert.match(page, /Mandar a producir/, "La tabla diaria muestra «Mandar a producir».");
assert.match(page, /produce-col/, "La columna de producir está resaltada.");
assert.match(page, /loadedSalesMonthKeys/, "La tarjeta de datos dice cuántos meses de venta hay cargados.");

// Columnas finales de la tabla diaria
const tableStart = page.indexOf('<table className="daily-table">');
const tableEnd = page.indexOf("</thead>", tableStart);
const headers = [...page.slice(tableStart, tableEnd).matchAll(/<th[^>]*>([^<]+)<\/th>/g)].map((m) => m[1].trim());
assert.deepStrictEqual(
  headers,
  ["Fecha", "Día", "Producto", "Pronóstico de venta", "Margen de seguridad", "Mandar a producir"],
  "Columnas de la tabla diaria"
);

// Lo que se quitó de la página
for (const [pattern, label] of [
  [/daily-stock-panel/, "Inventario del día"],
  [/Inventario del día/, "Inventario del día"],
  [/Vaciar este día/, "Vaciar este día"],
  [/Base con margen/, "Base con margen"],
  [/Bruto planta/, "Bruto planta"],
  [/Pedido planta/, "Pedido planta"],
  [/Inventario sucursales/, "Inventario sucursales"],
  [/weekly-progress-section/, "Avance semanal"],
  [/monthly-review-section/, "Revisión mensual asistida"],
  [/monthly-close-section/, "Cierre mensual"],
  [/homologation-section/, "Homologación de productos"],
  [/historical-validation-section/, "Ventas, producido y bajas"],
  [/notes-section/, "Notas de interpretación"],
  [/validation-section/, "Validación de cálculos"],
  [/secondary-tools-heading/, "Seguimiento y auditoría"],
  [/title="Más archivos"/, "Más archivos"],
  [/<ForecastAccuracyPanel/, "Tabla WAPE de meses cerrados"],
  [/Escenario operativo/, "KPI Escenario operativo"],
  [/daily-more-filters/, "Más filtros de la tabla"],
]) {
  assert.doesNotMatch(page, pattern, `${label} ya no aparece en la página.`);
}

// «Inventario de hoy»: una sola tarjeta de carga (sin cuadrícula de captura). Lo cargado
// (piso de venta de sucursales + cuarto frío) se resta de «Mandar a producir».
assert.match(page, /title="Inventario de hoy"/, "Se sube el inventario de hoy.");
assert.match(page, /onFile=\{handleDailyBranchStockFile\}/, "La tarjeta usa el lector de inventario diario.");
assert.match(
  dashboard,
  /calculateDailyForecast\(\{[\s\S]*?dailyBranchStock: effectiveDailyBranchStock,\s*dailyColdRoom: effectiveDailyColdRoom,\s*branchStockTargets,/,
  "La tabla diaria recibe el inventario cargado y el stock fijo por sucursal."
);
assert.match(dashboard, /buildBranchStockTargetMap\(stockRows\)/, "El stock fijo por sucursal sale de la carga «Stock fijo».");
assert.match(dashboard, /setStockRows\(attachBranchStockTargets\(parseStock\(workbook\), parseBranchStockTargets\(workbook\)\)\)/, "Al subir Stock fijo se leen las hojas por sucursal.");
assert.match(page, /formatNumber\(dailySummary\.aProducirMensual, 0\)/, "El KPI mensual suma «Mandar a producir».");

assert.doesNotMatch(app, /<FreezeReadinessStrip[\s\S]*<\/header>/, "Las tiras de freeze/salud no hinchan el header.");

assert.match(css, /\.main > \.loaded-files-section \{ order: 2; \}/, "Los datos van primero en el orden visual.");
assert.match(css, /\.main > \.decision-section \{ order: 6; \}/, "La salud del pronóstico sigue a la carga.");
assert.match(css, /\.main > \.daily-section \{ order: 8; \}/, "La tabla de planta queda en el flujo primario.");
assert.match(css, /\.daily-table \{\s*min-width: 0;/, "La tabla diaria ya no fuerza scroll horizontal.");

async function loadApp() {
  const built = await esbuild.build({
    entryPoints: [path.join(ROOT, "App.jsx")], bundle: true, platform: "node", format: "cjs", write: false,
    loader: { ".css": "text" },
    define: { "import.meta.env.VITE_API_URL": '""', "import.meta.env.DEV": "false", "import.meta.env.PROD": "true" },
    logLevel: "silent",
  });
  const m = new Module("ui-hierarchy-test");
  m.filename = path.join(__dirname, "ui-hierarchy-test.bundle.cjs");
  m.paths = module.paths;
  m._compile(built.outputFiles[0].text, m.filename);
  return m.exports;
}

function monthClose(month, producto, cantidad) {
  const [year, monthNumber] = month.split("-").map(Number);
  return { fecha: new Date(year, monthNumber - 1, 1), producto, cantidad, monthlyTotal: true, monthDays: new Date(year, monthNumber, 0).getDate() };
}

(async () => {
  const { calculateForecast, calculateDailyForecast } = await loadApp();
  const products = ["GELATINA IND FRESA", "FRUTAS GDE"];
  const monthlyRows = calculateForecast({
    stockRows: products.map((producto, i) => ({ producto, stock: 30, orden: i + 1 })),
    historicalVentas: products.flatMap((p, i) => ["2026-04", "2026-05", "2026-06"].map((m, j) => monthClose(m, p, 200 + 40 * i + 10 * j))),
    bajas: [], existencias: [], realProduction: [], selectedMonth: "2026-07", dailyBufferPct: 10,
  });
  const daily = calculateDailyForecast({
    monthlyRows, ventasReales: [], realProduction: [], selectedMonth: "2026-07", dailyBufferPct: 10,
    activePromos: [], dailyBranchStock: [], dailyColdRoom: [],
  });
  assert(daily.length > 0, "Hay filas diarias.");
  for (const row of daily) {
    assert.strictEqual(row.aProducirDia, row.produccionBrutaDia, `${row.producto} ${row.fecha}: sin inventario, producir = bruto`);
    assert.strictEqual(row.produccionSugeridaDia, row.produccionBrutaDia, `${row.producto} ${row.fecha}: sin inventario, pedido = bruto`);
    assert.strictEqual(row.promedioUsado, row.pronosticoVentaDia, "Promedio aplicado es el mismo número que el pronóstico de venta");
  }
  // Con inventario: Mandar a producir = bruto − (sucursales + cuarto frío), nunca menos de 0.
  const fecha = "2026-07-02";
  const conInv = calculateDailyForecast({
    monthlyRows, ventasReales: [], realProduction: [], selectedMonth: "2026-07", dailyBufferPct: 10, activePromos: [],
    dailyBranchStock: [{ fecha, sucursal: "Suc. Plaza", producto: "FRUTAS GDE", cantidad: 3 }, { fecha, sucursal: "Suc. Vistas", producto: "FRUTAS GDE", cantidad: 2 }, { fecha, sucursal: "Suc. Plaza", producto: "GELATINA IND FRESA", cantidad: 999 }],
    dailyColdRoom: [{ fecha, producto: "FRUTAS GDE", cantidad: 1 }],
  });
  const sin = new Map(daily.map((r) => [`${r.fecha}|${r.producto}`, r]));
  for (const row of conInv) {
    const base = sin.get(`${row.fecha}|${row.producto}`);
    const esperado = row.fecha !== fecha ? base.aProducirDia
      : row.producto === "FRUTAS GDE" ? Math.max(0, base.produccionBrutaDia - 6) : 0;
    assert.strictEqual(row.aProducirDia, esperado, `${row.producto} ${row.fecha}: producir con inventario`);
  }
  // Regla elegida (sep 2026): de sucursales solo se resta lo que sobra de su stock fijo.
  const {
    parseBranchStockTargets, attachBranchStockTargets, buildBranchStockTargetMap, branchStockKey, parseStock,
  } = await loadApp();
  const XLSX = require(path.join(ROOT, "node_modules", "xlsx"));
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet([
    ["PRODUCTO", "STOCK"], ["FRUTAS GDE", 40], ["GELATINA IND FRESA", 60],
  ]), "STOCK DE SUCURSALES");
  XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet([
    ["2026-09-01", "", "", "SUCURSAL", "", "Suc. Plaza"],
    ["PRODUCTO ", "CODIGO", "EXISTENCIA EN SUCURSAL", "STOCK DETERMINADO"],
    ["FRUTAS GDE", "1", 0, 2], ["GELATINA IND FRESA", "2", 0, 0],
  ]), "Suc. Plaza");
  XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet([
    ["2026-09-01", "", "", "SUCURSAL", "", "Vistas"],
    ["PRODUCTO", "CODIGO", "EXISTENCIA EN SUCURSAL", "STOCK DETERMINADO FIN SEMANA"],
    ["FRUTAS GDE", "1", 0, 5], ["GELATINA IND FRESA", "2", 0, ""],
  ]), "Suc. Vistas");
  const targets = parseBranchStockTargets(wb);
  assert.strictEqual(targets.length, 3, "Celda en blanco = sin stock fijo; 0 sí cuenta como stock fijo.");
  assert.strictEqual(branchStockKey("Suc. Vistas"), branchStockKey("VISTAS"), "Sucursal se cruza sin «Suc.» ni mayúsculas.");
  const stockConSucursal = attachBranchStockTargets(parseStock(wb), targets);
  const targetMap = buildBranchStockTargetMap(JSON.parse(JSON.stringify(stockConSucursal)));
  assert.strictEqual(targetMap.get(`FRUTAS GDE|${branchStockKey("Suc. Plaza")}`), 2, "El stock fijo por sucursal sobrevive el respaldo (JSON).");
  const excedente = calculateDailyForecast({
    monthlyRows, ventasReales: [], realProduction: [], selectedMonth: "2026-07", dailyBufferPct: 10, activePromos: [],
    dailyBranchStock: [
      { fecha, sucursal: "Suc. Plaza", producto: "FRUTAS GDE", cantidad: 3 },
      { fecha, sucursal: "Suc. Vistas", producto: "FRUTAS GDE", cantidad: 2 },
      { fecha, sucursal: "Suc. Allende", producto: "FRUTAS GDE", cantidad: 9 },
      { fecha, sucursal: "Suc. Plaza", producto: "GELATINA IND FRESA", cantidad: 4 },
      { fecha, sucursal: "Suc. Vistas", producto: "GELATINA IND FRESA", cantidad: 999 },
    ],
    dailyColdRoom: [{ fecha, producto: "FRUTAS GDE", cantidad: 1 }],
    branchStockTargets: targetMap,
  });
  for (const row of excedente) {
    const base = sin.get(`${row.fecha}|${row.producto}`);
    if (row.fecha !== fecha) {
      assert.strictEqual(row.aProducirDia, base.aProducirDia, `${row.producto} ${row.fecha}: otros días sin cambio`);
      continue;
    }
    // FRUTAS: Plaza 3−2 = 1; Vistas 2−5 → 0 (no compensa); Allende sin stock fijo → 0. CF 1 completo.
    // GELATINA: Plaza 4−0 = 4; Vistas sin stock fijo (en blanco) → no se resta nada de sus 999.
    const descuento = row.producto === "FRUTAS GDE" ? 1 + 1 : 4;
    assert.strictEqual(row.aProducirDia, Math.max(0, base.produccionBrutaDia - descuento), `${row.producto}: solo se resta el excedente`);
    assert.strictEqual(row.sucursalesSinStockFijo, 1, `${row.producto}: una sucursal sin stock fijo`);
    assert.match(row.reglaOperativa, /que sobra del stock fijo en sucursales/, "La leyenda dice qué se descontó.");
    assert.match(row.reglaOperativa, /sin stock fijo: no se restan/, "La leyenda avisa la sucursal sin stock fijo.");
  }
  const frutas = excedente.find((row) => row.fecha === fecha && row.producto === "FRUTAS GDE");
  assert.strictEqual(frutas.excedenteSucursalesDia, 1);
  assert.match(frutas.reglaOperativa, /menos 1 que sobra del stock fijo en sucursales \(piso 5 vs stock fijo 7\)/);
  assert.match(frutas.reglaOperativa, /menos 1 en cuarto frío/);
  // Respaldo viejo sin stock fijo por sucursal: no se resta nada de sucursales, solo el cuarto frío.
  const sinMapa = calculateDailyForecast({
    monthlyRows, ventasReales: [], realProduction: [], selectedMonth: "2026-07", dailyBufferPct: 10, activePromos: [],
    dailyBranchStock: [{ fecha, sucursal: "Suc. Plaza", producto: "FRUTAS GDE", cantidad: 30 }],
    dailyColdRoom: [{ fecha, producto: "FRUTAS GDE", cantidad: 1 }],
    branchStockTargets: new Map(),
  }).find((row) => row.fecha === fecha && row.producto === "FRUTAS GDE");
  assert.strictEqual(sinMapa.aProducirDia, Math.max(0, sin.get(`${fecha}|FRUTAS GDE`).produccionBrutaDia - 1));

  console.log("ui-hierarchy-test: página simplificada (carga, salud mínima, promo y «Mandar a producir») intacta");
})().catch((error) => {
  console.error(error);
  process.exit(1);
});
