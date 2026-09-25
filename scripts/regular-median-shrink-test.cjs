/**
 * Prueba de la regla "25% hacia el nivel de meses regulares del año" con DATOS
 * SINTÉTICOS (cifras redondas inventadas, no son ventas de Pastelería Pepes).
 * Protege:
 *  - producto que alterna alto/bajo: en un mes sin evento el pronóstico se
 *    acerca 25% a la mediana (últimos 2 meses regulares del año + mes anterior);
 *  - meses con evento (enero, febrero, mayo, junio, diciembre, Semana Santa) no se tocan;
 *  - producto que dejó de vender en la ventana: no se toca (no revive apagados);
 *  - solo mira el año en curso: con y sin el año anterior da lo mismo.
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
  const m = new Module("regular-median-shrink-test");
  m.filename = path.join(__dirname, "regular-median-shrink-test.bundle.cjs");
  m.paths = module.paths;
  m._compile(built.outputFiles[0].text, m.filename);
  return m.exports;
}

function assert(condition, message) {
  if (!condition) throw new Error(`[mediana meses regulares sintético] ${message}`);
}

function close(monthKey, producto, cantidad) {
  const [y, mo] = monthKey.split("-").map(Number);
  return { fecha: new Date(y, mo - 1, 15, 12), producto, cantidad, monthlyTotal: true, monthDays: new Date(y, mo, 0).getDate() };
}

const Y = (months) => Object.fromEntries(months.map(([k, v]) => [k, v]));
const SERIES = {
  // 2025: Semana Santa en abril. Enero, marzo, julio... son meses regulares.
  "ALTERNA": Y([["2025-01", 400], ["2025-02", 600], ["2025-03", 800], ["2025-04", 400], ["2025-05", 800],
    ["2025-06", 400], ["2025-07", 800], ["2025-08", 400], ["2025-09", 800], ["2025-10", 400], ["2025-11", 800]]),
  "APAGADO": Y([["2025-01", 300], ["2025-02", 300], ["2025-03", 300], ["2025-04", 300], ["2025-05", 300],
    ["2025-06", 300], ["2025-07", 300], ["2025-08", 0], ["2025-09", 0]]),
  "PREVIO": Y([["2024-07", 900], ["2024-08", 100], ["2024-09", 900], ["2024-10", 100], ["2024-11", 900], ["2024-12", 100],
    ["2025-01", 400], ["2025-02", 600], ["2025-03", 800], ["2025-04", 400], ["2025-05", 800],
    ["2025-06", 400], ["2025-07", 800], ["2025-08", 400], ["2025-09", 800]]),
};
const MONTHS = [];
for (const y of [2024, 2025]) for (let mo = 1; mo <= 12; mo += 1) MONTHS.push(`${y}-${String(mo).padStart(2, "0")}`);

function forecast(app, month, producto, { dropPriorYear = false } = {}) {
  const ventas = [];
  for (const [p, months] of Object.entries(SERIES)) {
    for (const k of MONTHS) {
      if (k >= month) continue;
      if (dropPriorYear && k < "2025-01") continue;
      if (!(k in months) && k < "2025-01") continue;
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
  assert(typeof app.applyRegularMonthMedianShrink === "function", "falta applyRegularMonthMedianShrink");

  // Unidad: modelo de 800 en octubre; ventana = sep (800, mes anterior y regular) y ago (400) -> mediana 600.
  const records = Object.entries(SERIES.ALTERNA).map(([k, v]) => close(k, "ALTERNA", v)).filter((r) => r.fecha < new Date(2025, 9, 1));
  assert(typeof app.uniformWeekdayAverages === "function" && typeof app.forecastTotalFromAverages === "function", "faltan helpers exportados");
  const model = { averages: app.uniformWeekdayAverages(800 / 31), trend: 1, method: "prueba" };
  {
    const out = app.applyRegularMonthMedianShrink(model, records, "2025-10");
    const total = app.forecastTotalFromAverages(out.averages, "2025-10");
    assert(Math.abs(total - 750) < 0.5, `octubre: 800 + 25% x (600 - 800) = 750, salió ${total.toFixed(2)}`);
    assert(/meses regulares/.test(out.method), "el método debe explicar la regla");
    for (const m of ["2025-05", "2025-06", "2025-12", "2025-04"]) {
      const rec = records.filter((r) => r.fecha < new Date(Number(m.slice(0, 4)), Number(m.slice(5)) - 1, 1));
      const same = app.applyRegularMonthMedianShrink(model, rec, m);
      assert(same === model, `${m} es mes de evento (o Semana Santa 2025): no se toca`);
    }
    const jan = app.applyRegularMonthMedianShrink(model, records, "2026-01");
    assert(jan === model, "enero (arranque del año) no se toca");
  }

  // Integración: el método final del pronóstico lleva la regla en septiembre y no en mayo.
  const sep = forecast(app, "2025-09", "ALTERNA");
  assert(/meses regulares/.test(sep.metodoPronostico), `septiembre ALTERNA debe llevar la regla (${sep.metodoPronostico})`);
  const may = forecast(app, "2025-05", "ALTERNA");
  assert(!/meses regulares/.test(may.metodoPronostico), "mayo no lleva la regla");

  // Apagado: agosto y septiembre en 0 -> no revive.
  const off = forecast(app, "2025-10", "APAGADO");
  assert(!/meses regulares/.test(off.metodoPronostico) && off.pronosticoVenta <= 1, `apagado no debe revivir (fc ${off.pronosticoVenta.toFixed(1)})`);

  // Solo año en curso: con y sin 2024 igual en octubre.
  const withPrior = forecast(app, "2025-10", "PREVIO");
  const withoutPrior = forecast(app, "2025-10", "PREVIO", { dropPriorYear: true });
  assert(Math.abs(withPrior.pronosticoVenta - withoutPrior.pronosticoVenta) < 1e-6,
    `octubre: con 2024 (${withPrior.pronosticoVenta.toFixed(2)}) debe ser igual que sin 2024 (${withoutPrior.pronosticoVenta.toFixed(2)})`);

  console.log(`regular-median-shrink-test OK (sintético): septiembre ${sep.pronosticoVenta.toFixed(0)}, mayo sin regla, apagado ${off.pronosticoVenta.toFixed(0)}`);
})().catch((error) => {
  console.error(error);
  process.exit(1);
});
