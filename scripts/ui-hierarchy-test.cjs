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

// Sin captura de inventario en la página, el cálculo diario no resta inventario.
assert.match(
  dashboard,
  /calculateDailyForecast\(\{[\s\S]*?dailyBranchStock: \[\],\s*dailyColdRoom: \[\],/,
  "La tabla diaria se calcula con inventario vacío."
);

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
  console.log("ui-hierarchy-test: página simplificada (carga, salud mínima, promo y «Mandar a producir») intacta");
})().catch((error) => {
  console.error(error);
  process.exit(1);
});
