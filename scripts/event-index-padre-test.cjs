/**
 * Prueba del índice de evento del año anterior en Madres (mayo, blend completo) y
 * Padre (junio, media distancia) con DATOS SINTÉTICOS (cifras redondas inventadas,
 * no son ventas de Pastelería Pepes). Protege:
 *  - junio (Día del Padre) sube hacia mes previo × (junio/mayo del año anterior),
 *    la mitad del camino;
 *  - mayo (Día de las Madres) llega a mes previo × (mayo/abril del año anterior);
 *  - solo sube: si el año anterior el evento fue menor, no toca el pronóstico;
 *  - sin año anterior no hay índice;
 *  - no usa el mes pronosticado (ventas del mes no cambian nada);
 *  - meses sin evento (julio) no llevan índice.
 */
const path = require("path");
const Module = require("module");
const esbuild = require("esbuild");

const ROOT = path.join(__dirname, "..");

async function loadApp() {
  const built = await esbuild.build({
    entryPoints: [path.join(ROOT, "App.jsx")], bundle: true, platform: "node", format: "cjs", write: false,
    loader: { ".css": "text" },
    define: { "import.meta.env.VITE_API_URL": JSON.stringify(""), "import.meta.env.DEV": "false", "import.meta.env.PROD": "true" },
    logLevel: "silent",
  });
  const m = new Module("event-index-padre-test");
  m.filename = path.join(__dirname, "event-index-padre-test.bundle.cjs");
  m.paths = module.paths;
  m._compile(built.outputFiles[0].text, m.filename);
  return m.exports;
}

function assert(condition, message) {
  if (!condition) throw new Error(`[índice Madres/Padre sintético] ${message}`);
}
const near = (a, b, tol = 1e-6) => Math.abs(a - b) <= tol;

function close(monthKey, producto, cantidad) {
  const [y, mo] = monthKey.split("-").map(Number);
  return { fecha: new Date(y, mo - 1, 15, 12), producto, cantidad, monthlyTotal: true, monthDays: new Date(y, mo, 0).getDate() };
}
function records(series, producto, beforeMonth) {
  return Object.entries(series).filter(([k]) => k < beforeMonth).map(([k, v]) => close(k, producto, v));
}
// Modelo plano de prueba: `total` piezas repartidas en el mes.
function flatModel(app, month, total) {
  const days = new Date(Number(month.slice(0, 4)), Number(month.slice(5, 7)), 0).getDate();
  return { averages: app.uniformWeekdayAverages(total / days), trend: 1, method: "Plano", recentMonths: [] };
}

(async () => {
  const app = await loadApp();
  for (const fn of ["applyEventPriorYearIndex", "uniformWeekdayAverages", "forecastTotalFromAverages", "calendarEventForMonth"]) {
    assert(typeof app[fn] === "function", `falta ${fn}`);
  }
  assert(app.calendarEventForMonth("2025-06")?.id === "padre", "junio es Día del Padre");

  // Año anterior: abr 400, may 600 (Madres ×1.5), jun 500 (Padre ×500/600). Año en curso: abr 500, may 700.
  const S = { "2024-04": 400, "2024-05": 600, "2024-06": 500, "2024-07": 300, "2025-04": 500, "2025-05": 700 };

  // Junio: meta = 700 × (500/600) = 583.33; modelo 400 → 400 + (583.33 − 400) × 0.5 = 491.67.
  const jun = app.applyEventPriorYearIndex(flatModel(app, "2025-06", 400), records(S, "X", "2025-06"), "2025-06");
  const junTotal = app.forecastTotalFromAverages(jun.averages, "2025-06");
  assert(near(junTotal, 400 + (700 * 500 / 600 - 400) * 0.5, 1e-6), `junio con índice Padre = 491.67, salió ${junTotal}`);
  assert(/índice Día del Padre año anterior/.test(jun.method), `junio debe decir índice Día del Padre (${jun.method})`);

  // Solo sube: modelo 700 ya está arriba de la meta 583.33 → no cambia.
  const junHigh = app.applyEventPriorYearIndex(flatModel(app, "2025-06", 700), records(S, "X", "2025-06"), "2025-06");
  assert(near(app.forecastTotalFromAverages(junHigh.averages, "2025-06"), 700), "junio: el índice solo sube");

  // Sin año anterior: no hay índice.
  const noPrior = Object.fromEntries(Object.entries(S).filter(([k]) => k >= "2025-01"));
  const junNo = app.applyEventPriorYearIndex(flatModel(app, "2025-06", 400), records(noPrior, "X", "2025-06"), "2025-06");
  assert(near(app.forecastTotalFromAverages(junNo.averages, "2025-06"), 400) && !/índice/.test(junNo.method || ""), "sin año anterior no hay índice");

  // Sin fuga: una venta de junio 2025 en los registros no se usa (se filtra antes del mes).
  const leak = records({ ...S, "2025-06": 5000 }, "X", "2025-06");
  const junLeak = app.applyEventPriorYearIndex(flatModel(app, "2025-06", 400), leak, "2025-06");
  assert(near(app.forecastTotalFromAverages(junLeak.averages, "2025-06"), junTotal), "junio no usa ventas del mes pronosticado");

  // Mayo (Madres, blend completo): meta = 500 × (600/400) = 750; modelo 600 → 750.
  const may = app.applyEventPriorYearIndex(flatModel(app, "2025-05", 600), records(S, "X", "2025-05"), "2025-05");
  assert(near(app.forecastTotalFromAverages(may.averages, "2025-05"), 750, 1e-6), `mayo con índice Madres completo = 750, salió ${app.forecastTotalFromAverages(may.averages, "2025-05")}`);

  // Julio no es mes de evento: no cambia.
  const jul = app.applyEventPriorYearIndex(flatModel(app, "2025-07", 300), records({ ...S, "2025-06": 600 }, "X", "2025-07"), "2025-07");
  assert(near(app.forecastTotalFromAverages(jul.averages, "2025-07"), 300) && !/índice/.test(jul.method || ""), "julio sin índice de evento");

  console.log("event-index-padre-test ok");
})().catch((error) => {
  console.error(error);
  process.exit(1);
});
