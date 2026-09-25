/**
 * Prueba de la regla "producto de pulso con antecedente" (#25) con DATOS SINTÉTICOS
 * (cifras redondas inventadas, no son ventas de Pastelería Pepes). Protege:
 *  - pulso que el año anterior se apagó al mes siguiente -> pronóstico a la mitad;
 *  - sin año anterior, o si el año anterior siguió vendiendo -> no se toca;
 *  - si la Cuaresma cae distinto en el mes objetivo (Semana Santa movible) -> no se toca;
 *  - producto estable -> no se toca;
 *  - días de Cuaresma por mes (2024: Pascua 31-mar; 2025: Pascua 20-abr).
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
  const m = new Module("pulse-fade-test");
  m.filename = path.join(__dirname, "pulse-fade-test.bundle.cjs");
  m.paths = module.paths;
  m._compile(built.outputFiles[0].text, m.filename);
  return m.exports;
}

function assert(condition, message) {
  if (!condition) throw new Error(`[pulso sintético] ${message}`);
}

function close(monthKey, producto, cantidad) {
  const [y, mo] = monthKey.split("-").map(Number);
  return { fecha: new Date(y, mo - 1, 15, 12), producto, cantidad, monthlyTotal: true, monthDays: new Date(y, mo, 0).getDate() };
}

const MONTHS = [];
for (const y of [2024, 2025]) for (let mo = 1; mo <= 12; mo += 1) MONTHS.push(`${y}-${String(mo).padStart(2, "0")}`);

// cierres mensuales sintéticos (lo que no aparece es 0)
const SERIES = {
  "PULSO CON ANTECEDENTE": { "2024-02": 200, "2025-02": 300 },
  "PULSO SIN ANTECEDENTE": { "2025-02": 300 },
  "PULSO QUE SIGUIO": { "2024-02": 200, "2024-03": 150, "2025-02": 300 },
  "PULSO CUARESMA": { "2024-03": 900, "2025-03": 600 },
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
  assert(app.lentDaysInMonth("2024-02") === 16 && app.lentDaysInMonth("2024-03") === 30 && app.lentDaysInMonth("2024-04") === 0, "Cuaresma 2024");
  assert(app.lentDaysInMonth("2025-02") === 0 && app.lentDaysInMonth("2025-03") === 27 && app.lentDaysInMonth("2025-04") === 19, "Cuaresma 2025");

  const faded = forecastFor(app, "2025-03", "PULSO CON ANTECEDENTE");
  assert(/pulso/.test(faded.metodoPronostico), "el método debe explicar la regla de pulso");

  // la regla deja el modelo a la mitad (no en 0)
  const march = "2025-03";
  const records = [];
  for (const k of MONTHS) if (k < march) records.push(close(k, "PULSO CON ANTECEDENTE", SERIES["PULSO CON ANTECEDENTE"][k] || 0));
  const averages = new Map([0, 1, 2, 3, 4, 5, 6].map((d) => [d, 10]));
  const halved = app.applyPriorYearPulseFade({ averages, trend: 1, method: "Modelo" }, records, march);
  assert([...halved.averages.values()].every((v) => Math.abs(v - 5) < 1e-9), "el pulso con antecedente debe bajar el modelo a la mitad");
  assert(halved.trend === 0.5 && /pulso/.test(halved.method), "tendencia a la mitad y método explicado");

  const noPrior = forecastFor(app, "2025-03", "PULSO SIN ANTECEDENTE");
  assert(noPrior.pronosticoVenta > 0, "sin año anterior no se toca");
  assert(!/pulso/.test(noPrior.metodoPronostico), "sin año anterior no debe marcar pulso");

  const kept = forecastFor(app, "2025-03", "PULSO QUE SIGUIO");
  assert(kept.pronosticoVenta > 0 && !/pulso/.test(kept.metodoPronostico), "si el año anterior siguió vendiendo no se toca");

  const lent = forecastFor(app, "2025-04", "PULSO CUARESMA");
  assert(lent.pronosticoVenta > 0 && !/pulso/.test(lent.metodoPronostico), "Semana Santa movida: abril 2025 no se apaga");

  const steady = forecastFor(app, "2025-03", "ESTABLE");
  assert(steady.pronosticoVenta > 0 && !/pulso/.test(steady.metodoPronostico), "producto estable no se toca");

  console.log("pulse-fade-test ok");
})().catch((e) => { console.error(e); process.exit(1); });
