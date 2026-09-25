/**
 * Prueba del tope de "estreno" en la limpieza de catálogo con DATOS SINTÉTICOS
 * (cifras redondas inventadas, no son ventas de Pastelería Pepes). Protege:
 *  - febrero con el año abierto: un producto "Otros" que vendió parejo todo el año
 *    anterior no se toma como estreno (antes quedaba en 35% de enero);
 *  - un estreno de verdad (sin venta el mes anterior) sigue con el tope.
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
  const m = new Module("false-debut-test");
  m.filename = path.join(__dirname, "false-debut-test.bundle.cjs");
  m.paths = module.paths;
  m._compile(built.outputFiles[0].text, m.filename);
  return m.exports;
}

function assert(condition, message) {
  if (!condition) throw new Error(`[estreno sintético] ${message}`);
}

function close(monthKey, producto, cantidad) {
  const [y, mo] = monthKey.split("-").map(Number);
  return { fecha: new Date(y, mo - 1, 15, 12), producto, cantidad, monthlyTotal: true, monthDays: new Date(y, mo, 0).getDate() };
}

const MONTHS = [];
for (const y of [2025, 2026]) for (let mo = 1; mo <= 12; mo += 1) MONTHS.push(`${y}-${String(mo).padStart(2, "0")}`);

const SERIES = {
  "POSTRE DE TODO EL ANIO": Object.fromEntries(MONTHS.map((k) => [k, 600])),
  "POSTRE ESTRENO": { "2026-04": 600 },
  "MOKA GDE": Object.fromEntries(MONTHS.map((k) => [k, 300])),
};

function forecastFor(app, month, producto) {
  const ventas = [];
  for (const [p, months] of Object.entries(SERIES)) {
    for (const k of MONTHS) {
      if (k >= month) continue;
      if (months[k]) ventas.push(close(k, p, months[k]));
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
  const feb = forecastFor(app, "2026-02", "POSTRE DE TODO EL ANIO");
  assert(!/pico sin soporte/.test(feb.metodoPronostico || ""), `febrero no debe tratar enero como estreno (${feb.metodoPronostico})`);
  assert(feb.pronosticoVenta > 450, `febrero debe quedar cerca de 600, quedó ${feb.pronosticoVenta}`);

  const may = forecastFor(app, "2026-05", "POSTRE ESTRENO");
  assert(/pico sin soporte/.test(may.metodoPronostico || ""), `un estreno de verdad sigue con tope (${may.metodoPronostico})`);
  assert(may.pronosticoVenta < 300, `estreno debe quedar topado, quedó ${may.pronosticoVenta}`);

  console.log(`false-debut-test OK (sintético): febrero ${feb.pronosticoVenta.toFixed(0)}, estreno ${may.pronosticoVenta.toFixed(0)}`);
})().catch((error) => {
  console.error(error);
  process.exit(1);
});
