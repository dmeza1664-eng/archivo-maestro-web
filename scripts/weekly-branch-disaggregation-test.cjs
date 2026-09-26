/**
 * Prueba del reparto semanal / por sucursal del pronóstico mensual con DATOS SINTÉTICOS
 * (cifras redondas inventadas, no son datos de Pastelería Pepes). Protege:
 *  - la suma de semanas y de sucursales = pronóstico mensual (no cambia el mensual);
 *  - semanas ISO recortadas al mes (primera/última parciales);
 *  - perfil de día de la semana y participación por sucursal salen del histórico;
 *  - ventas con fecha dentro del mes pronosticado se ignoran (sin fuga);
 *  - sucursal sin venta en los últimos 14 días no recibe reparto;
 *  - factor de fecha especial (10 de mayo) tomado del año anterior;
 *  - víspera solo para el SKU que históricamente sube la víspera (el otro no cambia).
 */
const path = require("path");
const { disaggregateMonthlyForecast, monthWeeks, isoWeekKey } = require(path.join(__dirname, "lib", "weekly-branch-disaggregation.cjs"));
function assert(c, msg) { if (!c) throw new Error(`[semanal-sucursal sintético] ${msg}`); }
const near = (a, b, tol = 1e-6) => Math.abs(a - b) <= tol;
const DAY = 86400000;
const iso = (t) => new Date(t).toISOString().slice(0, 10);
function days(from, to) { const out = []; for (let t = Date.parse(from + "T00:00:00Z"); t <= Date.parse(to + "T00:00:00Z"); t += DAY) out.push(t); return out; }

// Semanas ISO: septiembre 2025 empieza lunes 1 → W36..W40, la última (29–30) parcial.
const w = monthWeeks("2025-09");
assert(w.length === 5 && w[0].semana === "2025-W36" && w[0].dias === 7 && w[4].dias === 2, `semanas sep-2025 ${JSON.stringify(w)}`);
assert(isoWeekKey(Date.UTC(2024, 11, 30)) === "2025-W01", "30-dic-2024 es semana 2025-W01");
assert(monthWeeks("2025-02").reduce((a, x) => a + x.dias, 0) === 28, "febrero 2025 tiene 28 días");

// Histórico: sucursal A vende 30/día, B 10/día; sábados el triple. C dejó de vender el 10-ago (más de 14 días antes del mes).
const daily = [];
for (const t of days("2025-06-09", "2025-08-31")) { // 84 días = 12 semanas exactas
  const sat = new Date(t).getUTCDay() === 6 ? 3 : 1;
  daily.push({ fecha: iso(t), sucursal: "A", producto: "PASTEL X", cantidad: 30 * sat });
  daily.push({ fecha: iso(t), sucursal: "B", producto: "PASTEL X", cantidad: 10 * sat });
  if (t < Date.UTC(2025, 7, 11)) daily.push({ fecha: iso(t), sucursal: "C", producto: "PASTEL X", cantidad: 50 * sat });
}
const base = disaggregateMonthlyForecast({ month: "2025-09", monthlyForecast: { "PASTEL X": 1000 }, dailySales: daily });
const sumW = base.porSkuSemana.reduce((a, r) => a + r.cantidad, 0);
const sumB = base.porSucursalSkuSemana.reduce((a, r) => a + r.cantidad, 0);
assert(near(sumW, 1000) && near(sumB, 1000), `suma semanas ${sumW} / sucursales ${sumB} debe ser 1000`);
assert(base.sucursales.join(",") === "A,B", `solo sucursales activas A,B (C sin venta en 14 días): ${base.sucursales}`);
const shareA = base.porSucursalSkuSemana.filter((r) => r.sucursal === "A").reduce((a, r) => a + r.cantidad, 0) / 1000;
assert(near(shareA, 0.75, 1e-9), `participación de A = 0.75 (30 de 40), salió ${shareA}`);
// Semana completa (7 días, 1 sábado) vs semana parcial de 2 días (lun–mar): proporción por perfil 9:2
const wFull = base.porSkuSemana.find((r) => r.semana === "2025-W36").cantidad;
const wPart = base.porSkuSemana.find((r) => r.semana === "2025-W40").cantidad;
assert(near(wFull / wPart, 9 / 2, 1e-9), `perfil día de semana: semana completa/parcial = 4.5, salió ${wFull / wPart}`);

// Sin fuga: ventas dentro del mes pronosticado no cambian nada.
const leak = daily.concat(days("2025-09-01", "2025-09-30").map((t) => ({ fecha: iso(t), sucursal: "B", producto: "PASTEL X", cantidad: 999 })));
const withLeak = disaggregateMonthlyForecast({ month: "2025-09", monthlyForecast: { "PASTEL X": 1000 }, dailySales: leak });
assert(JSON.stringify(withLeak.porSucursalSkuSemana) === JSON.stringify(base.porSucursalSkuSemana), "ventas del mes pronosticado no deben usarse");

// Producto sin historia: se reparte con el perfil general y la participación general de sucursal.
const nuevo = disaggregateMonthlyForecast({ month: "2025-09", monthlyForecast: { "NUEVO": 100 }, dailySales: daily });
assert(near(nuevo.porSkuSemana.reduce((a, r) => a + r.cantidad, 0), 100), "producto nuevo conserva el mensual");

// 10 de mayo: el año anterior ese día vendió 5× el promedio de su día de la semana.
const may = [];
for (const t of days("2024-05-01", "2024-05-31")) may.push({ fecha: iso(t), sucursal: "A", producto: "PASTEL X", cantidad: iso(t) === "2024-05-10" ? 50 : 10 });
for (const t of days("2025-03-01", "2025-04-30")) may.push({ fecha: iso(t), sucursal: "A", producto: "PASTEL X", cantidad: 10 });
const ev = disaggregateMonthlyForecast({ month: "2025-05", monthlyForecast: { "PASTEL X": 1000 }, dailySales: may });
const noEv = disaggregateMonthlyForecast({ month: "2025-05", monthlyForecast: { "PASTEL X": 1000 }, dailySales: may, options: { events: false } });
const wk = isoWeekKey(Date.UTC(2025, 4, 10));
const fEv = ev.porSkuSemana.find((r) => r.semana === wk).cantidad, fNo = noEv.porSkuSemana.find((r) => r.semana === wk).cantidad;
assert(fEv > fNo * 1.3, `semana del 10 de mayo debe subir con el factor del año anterior (${fEv} vs ${fNo})`);
assert(near(ev.porSkuSemana.reduce((a, r) => a + r.cantidad, 0), 1000), "con evento, el mensual se conserva");

// Víspera por SKU: 2 años de historia (dic-2022 a ene-2025) en la sucursal A. "SUBE" vende 10 en días
// normales y 30 en cada víspera (1–2 días antes de los eventos); "PLANO" vende 10 siempre.
const { visperaDaysOfYear } = require(path.join(__dirname, "lib", "weekly-branch-disaggregation.cjs"));
const vsp = new Set([2022, 2023, 2024, 2025].flatMap((y) => [...visperaDaysOfYear(y)]));
const hist = [];
for (const t of days("2023-02-01", "2025-01-31")) {
  hist.push({ fecha: iso(t), sucursal: "A", producto: "SUBE", cantidad: vsp.has(t) ? 30 : 10 });
  hist.push({ fecha: iso(t), sucursal: "A", producto: "PLANO", cantidad: 10 });
}
const FC = { SUBE: 280, PLANO: 280 };
const conV = disaggregateMonthlyForecast({ month: "2025-02", monthlyForecast: FC, dailySales: hist });
const sinV = disaggregateMonthlyForecast({ month: "2025-02", monthlyForecast: FC, dailySales: hist, options: { visperas: false } });
assert(conV.skuVispera.join(",") === "SUBE", `solo SUBE lleva víspera, salió ${conV.skuVispera}`);
const wkV = isoWeekKey(Date.UTC(2025, 1, 12)); // semana del 12–13 de febrero (víspera de San Valentín)
const q = (r, p) => r.porSkuSemana.find((x) => x.producto === p && x.semana === wkV).cantidad;
assert(q(conV, "SUBE") > q(sinV, "SUBE") * 1.1, `la semana de la víspera sube para SUBE (${q(conV, "SUBE")} vs ${q(sinV, "SUBE")})`);
assert(near(q(conV, "PLANO"), q(sinV, "PLANO"), 1e-9), "PLANO no sube la víspera: no cambia");
for (const p of ["SUBE", "PLANO"]) assert(near(conV.porSkuSemana.filter((x) => x.producto === p).reduce((a, x) => a + x.cantidad, 0), 280), `${p}: con víspera el mensual se conserva`);
const leakV = hist.concat(days("2025-02-01", "2025-02-28").map((t) => ({ fecha: iso(t), sucursal: "A", producto: "PLANO", cantidad: vsp.has(t) ? 500 : 10 })));
const conLeak = disaggregateMonthlyForecast({ month: "2025-02", monthlyForecast: FC, dailySales: leakV });
assert(JSON.stringify(conLeak.porSkuSemana) === JSON.stringify(conV.porSkuSemana), "la víspera no usa ventas del mes pronosticado");

let threw = false; try { disaggregateMonthlyForecast({ month: "2025-9", monthlyForecast: {}, dailySales: [] }); } catch { threw = true; }
assert(threw, "mes con formato inválido debe fallar");
console.log("weekly-branch-disaggregation-test ok");
