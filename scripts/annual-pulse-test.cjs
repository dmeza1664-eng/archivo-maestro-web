/**
 * Prueba de la regla "pulso anual de fecha fija" con DATOS SINTÉTICOS
 * (cifras redondas inventadas, no son ventas de Pastelería Pepes). Protege:
 *  - pulso repetido en el mismo mes dos años seguidos -> sube al nivel del año anterior;
 *  - sin historia de hace dos años -> basta el año anterior;
 *  - con historia de hace dos años y sin pulso entonces -> no se toca;
 *  - pulso que cambió de mes (Cuaresma) sin antecedente de dos años -> no se toca;
 *  - mes con vecinos altos (no es pulso) y producto estable -> no se toca;
 *  - solo sube: nunca baja el pronóstico del modelo.
 * El pulso arranca con pocas piezas el mes previo (como en la vida real), así el
 * modelo base no lo toma como reactivación estacional y la prueba ve la regla.
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
  const m = new Module("annual-pulse-test");
  m.filename = path.join(__dirname, "annual-pulse-test.bundle.cjs");
  m.paths = module.paths;
  m._compile(built.outputFiles[0].text, m.filename);
  return m.exports;
}

function assert(condition, message) {
  if (!condition) throw new Error(`[pulso anual sintético] ${message}`);
}

function close(monthKey, producto, cantidad) {
  const [y, mo] = monthKey.split("-").map(Number);
  return { fecha: new Date(y, mo - 1, 15, 12), producto, cantidad, monthlyTotal: true, monthDays: new Date(y, mo, 0).getDate() };
}

const MONTHS = [];
for (const y of [2024, 2025, 2026]) for (let mo = 1; mo <= 12; mo += 1) MONTHS.push(`${y}-${String(mo).padStart(2, "0")}`);

// cierres mensuales sintéticos (lo que no aparece es 0)
const SERIES = {
  "PULSO REPETIDO": { "2024-02": 200, "2025-01": 15, "2025-02": 300, "2026-01": 25 },
  "PULSO UN ANIO": { "2025-02": 300, "2026-01": 25 },
  "PULSO CAMBIO DE MES": { "2025-03": 600, "2026-02": 250 },
  "NO ES PULSO": { "2025-01": 100, "2025-02": 300, "2025-03": 100, "2026-01": 100 },
  "ESTABLE": Object.fromEntries(MONTHS.map((k) => [k, 100])),
};

function forecastFor(app, month, producto) {
  const ventas = [];
  for (const [p, months] of Object.entries(SERIES)) {
    for (const k of MONTHS) {
      if (k >= month) continue;
      ventas.push(close(k, p, months[k] || 0));
    }
  }
  const stockRows = Object.keys(SERIES).map((p, i) => ({ producto: p, stock: 10, orden: i + 1 }));
  const rows = app.calculateForecast({
    stockRows, historicalVentas: app.filterVentasBeforeMonth(ventas, month), bajas: [], existencias: [], realProduction: [],
    selectedMonth: month, dailyBufferPct: 10,
  });
  const row = rows.find((r) => r.producto === producto);
  assert(row, `falta ${producto}`);
  return row;
}

(async () => {
  const app = await loadApp();
  assert(typeof app.applyAnnualFixedDatePulse === "function", "falta applyAnnualFixedDatePulse");

  // feb-2026: pulso en feb-2024 y feb-2025 -> sube al nivel de feb-2025
  const repeated = forecastFor(app, "2026-02", "PULSO REPETIDO");
  assert(Math.abs(repeated.pronosticoVenta - 300) < 1, `pulso repetido debe quedar en 300, quedó ${repeated.pronosticoVenta}`);
  assert(/pulso anual/.test(repeated.metodoPronostico), "el método debe explicar la regla de pulso anual");

  // feb-2025: no hay historia de feb-2023 -> basta con feb-2024
  const oneYearOnly = forecastFor(app, "2025-02", "PULSO REPETIDO");
  assert(Math.abs(oneYearOnly.pronosticoVenta - 200) < 1, `sin historia de dos años debe quedar en 200, quedó ${oneYearOnly.pronosticoVenta}`);
  assert(/pulso anual/.test(oneYearOnly.metodoPronostico), "sin historia de dos años la regla sí actúa");

  // feb-2026: feb-2024 sí está en la historia y no fue pulso -> no se toca
  const single = forecastFor(app, "2026-02", "PULSO UN ANIO");
  assert(!/pulso anual/.test(single.metodoPronostico) && single.pronosticoVenta < 150, `pulso de un solo año con historia de dos no debe subir (${single.pronosticoVenta})`);

  // mar-2026: pulso en mar-2025 pero no en mar-2024 (Cuaresma cambió de mes) -> no se toca
  const moved = forecastFor(app, "2026-03", "PULSO CAMBIO DE MES");
  assert(!/pulso anual/.test(moved.metodoPronostico), "pulso que cambió de mes no debe usar la regla");

  // feb-2026: feb-2025 con vecinos altos no es pulso
  const notPulse = forecastFor(app, "2026-02", "NO ES PULSO");
  assert(!/pulso anual/.test(notPulse.metodoPronostico), "mes con vecinos altos no es pulso");

  const stable = forecastFor(app, "2026-02", "ESTABLE");
  assert(!/pulso anual/.test(stable.metodoPronostico), "producto estable no se toca");

  // solo sube: si el modelo ya está arriba del año anterior, no se toca
  const records = [];
  for (const k of MONTHS) if (k < "2026-02") records.push(close(k, "PULSO REPETIDO", SERIES["PULSO REPETIDO"][k] || 0));
  const high = { averages: new Map([0, 1, 2, 3, 4, 5, 6].map((d) => [d, 20])), trend: 1, method: "Modelo" };
  const complete = MONTHS.filter((k) => k < "2026-02");
  const same = app.applyAnnualFixedDatePulse(high, records, "2026-02", complete);
  assert(same === high, "si el modelo ya supera el año anterior, no se toca");

  console.log("annual-pulse-test OK (sintético): repetido 300, un año 200, sin subir cuando no aplica");
})().catch((error) => {
  console.error(error);
  process.exit(1);
});
