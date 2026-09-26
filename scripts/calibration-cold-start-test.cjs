/**
 * Calibración parcial (15%) y arranque en frío acotado a productos con UN año de
 * antecedente del mes objetivo. DATOS SINTÉTICOS (cifras redondas inventadas, no son
 * ventas de Pastelería Pepes). Protege:
 *  - el ajuste por error del mes anterior aplica CALIBRATION_SHRINK = 0.15 cuando el
 *    mes de validación no es evento, y completo cuando sí lo es (mayo = Madres);
 *  - el impulso frío de Madres (y el "sin recorte pre-Madres") sigue para productos
 *    sin mayo del año anterior o con un solo año de antecedente;
 *  - con mayo de hace dos años (> 40) el impulso frío ya no aplica;
 *  - con menos de 2 meses seguidos del año en curso y año anterior, tampoco (#22);
 *  - nada usa ventas del mes pronosticado.
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
  const m = new Module("calibration-cold-start-test");
  m.filename = path.join(__dirname, "calibration-cold-start-test.bundle.cjs");
  m.paths = module.paths;
  m._compile(built.outputFiles[0].text, m.filename);
  return m.exports;
}

function assert(condition, message) {
  if (!condition) throw new Error(`[calibración / arranque en frío sintético] ${message}`);
}
const near = (a, b, tol = 1e-9) => Math.abs(a - b) <= tol;
const clamp = (v, lo, hi) => Math.min(hi, Math.max(lo, v));

function close(monthKey, producto, cantidad) {
  const [y, mo] = monthKey.split("-").map(Number);
  return { fecha: new Date(y, mo - 1, 15, 12), producto, cantidad, monthlyTotal: true, monthDays: new Date(y, mo, 0).getDate() };
}
function records(series, producto, beforeMonth) {
  return Object.entries(series).filter(([k]) => k < beforeMonth).map(([k, v]) => close(k, producto, v));
}
function monthsBetween(from, to) {
  const out = [];
  let [y, m] = from.split("-").map(Number);
  const [ty, tm] = to.split("-").map(Number);
  while (y < ty || (y === ty && m <= tm)) {
    out.push(`${y}-${String(m).padStart(2, "0")}`);
    m += 1; if (m > 12) { m = 1; y += 1; }
  }
  return out;
}
function flatModel(app, month, total, trend = 1) {
  const days = new Date(Number(month.slice(0, 4)), Number(month.slice(5, 7)), 0).getDate();
  return { averages: app.uniformWeekdayAverages(total / days), trend, method: "Plano", recentMonths: [] };
}

(async () => {
  const app = await loadApp();
  for (const fn of ["calculateForecastModelLegacy", "coldStartHasTwoPriorYears", "coldStartEventUplift", "liftColdStartMadresCalibration", "forecastTotalFromAverages", "uniformWeekdayAverages"]) {
    assert(typeof app[fn] === "function", `falta ${fn}`);
  }
  assert(app.CALIBRATION_SHRINK === 0.15, `CALIBRATION_SHRINK debe ser 0.15 (es ${app.CALIBRATION_SHRINK})`);

  // --- Calibración: serie que crece 5%/mes; el mes de validación (agosto) queda por
  // encima de lo que predecía el modelo, así que la calibración cruda es > 1.
  const grow = {};
  monthsBetween("2024-01", "2025-08").forEach((k, i) => { grow[k] = Math.round(300 * Math.pow(1.05, i)); });
  const sep = app.calculateForecastModelLegacy(records(grow, "X", "2025-09"), "2025-09");
  assert(sep.backtestMonth === "2025-08", `validación de septiembre = agosto (${sep.backtestMonth})`);
  const rawSep = clamp(sep.backtestActual / sep.backtestForecast, 0.85, 1.15);
  assert(rawSep > 1.01, `la serie sintética debe dar calibración cruda > 1 (${rawSep})`);
  assert(near(sep.trend, 1 + (rawSep - 1) * 0.15), `septiembre aplica el 15% del ajuste: esperado ${1 + (rawSep - 1) * 0.15}, salió ${sep.trend}`);

  // Sin fuga: ventas del mes pronosticado no cambian la calibración.
  const sepLeak = app.calculateForecastModelLegacy(records({ ...grow, "2025-09": 9000 }, "X", "2025-10"), "2025-09");
  assert(near(sepLeak.trend, sep.trend), "la calibración no usa el mes pronosticado");

  // Validación en evento (mayo = Madres): el ajuste va completo.
  const jun = app.calculateForecastModelLegacy(records(grow, "X", "2025-06"), "2025-06");
  assert(jun.backtestMonth === "2025-05", `validación de junio = mayo (${jun.backtestMonth})`);
  const rawJun = jun.backtestActual > 0 && jun.backtestForecast > 0 ? clamp(jun.backtestActual / jun.backtestForecast, 0.85, 1.15) : 1;
  assert(near(jun.trend, rawJun), `con validación en Madres la calibración es completa (${jun.trend} vs ${rawJun})`);

  // --- Arranque en frío (Madres, categoría Mini medianos por nombre genérico).
  const P = "MINI PRUEBA";
  const base2025 = { "2025-01": 200, "2025-02": 200, "2025-03": 200, "2025-04": 200 };
  const oneYear = { "2024-04": 200, "2024-05": 260, ...base2025 };
  const twoYears = { "2023-04": 200, "2023-05": 260, ...oneYear };
  const newProduct = { ...base2025 };
  const gap = { "2024-04": 200, "2024-05": 260, "2025-01": 200, "2025-02": 200, "2025-04": 200 };

  assert(!app.coldStartHasTwoPriorYears(records(oneYear, P, "2025-05"), "2025-05"), "un año de antecedente");
  assert(app.coldStartHasTwoPriorYears(records(twoYears, P, "2025-05"), "2025-05"), "dos años de antecedente");

  const upNew = app.coldStartEventUplift(records(newProduct, P, "2025-05"), "2025-05", P);
  const upOne = app.coldStartEventUplift(records(oneYear, P, "2025-05"), "2025-05", P);
  const upTwo = app.coldStartEventUplift(records(twoYears, P, "2025-05"), "2025-05", P);
  const upGap = app.coldStartEventUplift(records(gap, P, "2025-05"), "2025-05", P);
  assert(upNew?.factor === 1.18, "producto sin mayo anterior: impulso frío Madres 1.18");
  assert(upOne?.factor === 1.18, "producto con un año de antecedente y 4 meses seguidos: impulso frío sigue");
  assert(upTwo === null, "producto con mayo de hace dos años: sin impulso frío");
  assert(upGap === null, "año anterior y solo 1 mes seguido del año en curso: sin impulso frío (#22)");

  // Sin fuga: mayo 2025 en los registros no cambia la decisión.
  const upTwoLeak = app.coldStartEventUplift(records({ ...twoYears, "2025-05": 5000 }, P, "2025-05"), "2025-05", P);
  assert(upTwoLeak === null, "el arranque en frío no usa el mes pronosticado");

  // "Sin recorte pre-Madres": deshace una calibración < 1 solo con un año de antecedente.
  const liftOne = app.liftColdStartMadresCalibration(flatModel(app, "2025-05", 190, 0.95), records(oneYear, P, "2025-05"), "2025-05", P);
  const liftTwo = app.liftColdStartMadresCalibration(flatModel(app, "2025-05", 190, 0.95), records(twoYears, P, "2025-05"), "2025-05", P);
  assert(near(app.forecastTotalFromAverages(liftOne.averages, "2025-05"), 200, 1e-6) && liftOne.trend === 1, "un año de antecedente: sin recorte pre-Madres");
  assert(near(app.forecastTotalFromAverages(liftTwo.averages, "2025-05"), 190, 1e-6) && liftTwo.trend === 0.95, "dos años de antecedente: se respeta la calibración");

  console.log("calibration-cold-start-test ok");
})().catch((error) => {
  console.error(error);
  process.exit(1);
});
