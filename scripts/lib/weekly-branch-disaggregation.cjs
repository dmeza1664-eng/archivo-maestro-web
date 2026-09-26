/**
 * Desagregación del pronóstico MENSUAL (por SKU) a SEMANA y a SUCURSAL.
 *
 * No cambia el pronóstico mensual: reparte cada cantidad mensual con perfiles
 * históricos calculados SOLO con ventas diarias ANTERIORES al primer día del mes
 * (las filas con fecha >= inicio del mes se ignoran; no hay fuga de datos).
 *
 *  - Perfil de día de la semana por SKU (ventana `windowDays`, 12 semanas), encogido hacia el
 *    perfil de todos los SKU con peso `beta` (piezas equivalentes).
 *  - Factor de fechas especiales de fecha fija (14-feb, 30-abr, 10-may, 24/31-dic,
 *    6-ene, 2-nov, 15-sep y Día del Padre = 3er domingo de junio): venta de ese día
 *    el año anterior contra el promedio del mismo día de la semana de ese mes,
 *    por SKU, encogido hacia el factor de todos los SKU con peso `eventK`.
 *  - Víspera de evento (1–2 días antes de esas fechas) SOLO para los SKU que
 *    históricamente suben la víspera: con la venta de los `vispWindowDays` días
 *    previos (2 años) se compara, por SKU, la venta de cada víspera contra el
 *    promedio del mismo día de la semana en días normales de ese mes (sin
 *    evento, víspera ni 15–17). Razón encogida hacia 1 con `vispK` piezas; el
 *    SKU entra si la razón es >= `vispMin` con al menos `vispMinObs` vísperas
 *    observadas, y su víspera se multiplica por esa razón (tope `vispMax`). Los
 *    demás SKU no se tocan. Sin nombres de SKU; sin datos del mes.
 *  - Participación por sucursal de las últimas `shareWindowDays` (por SKU),
 *    encogida hacia la participación general de la sucursal con peso `alpha`.
 *    Solo sucursales que vendieron en los últimos `activeDays` días.
 *  - Semanas ISO (lunes a domingo) recortadas al mes: la primera y la última
 *    semana pueden ser parciales. La suma de semanas = pronóstico mensual.
 */
"use strict";

const DEFAULTS = Object.freeze({
  windowDays: 84,
  shareWindowDays: 84,
  alpha: 30,
  beta: 50,
  activeDays: 14,
  events: true,
  eventK: 20,
  eventMin: 0.2,
  eventMax: 8,
  visperas: true,
  vispWindowDays: 730,
  vispK: 30,
  vispMin: 1.5,
  vispMax: 3,
  vispMinObs: 4,
});
const FIXED_EVENTS = [[1, 6], [2, 14], [4, 30], [5, 10], [9, 15], [11, 2], [12, 24], [12, 31]];
const DAY_MS = 86400000;

function toUTC(iso) { const [y, m, d] = iso.slice(0, 10).split("-").map(Number); return Date.UTC(y, m - 1, d); }
function isoOf(t) { return new Date(t).toISOString().slice(0, 10); }
function weekdayMon0(t) { return (new Date(t).getUTCDay() + 6) % 7; } // lunes=0 … domingo=6
function isoWeekKey(t) {
  const d = new Date(t); const day = weekdayMon0(t);
  const thu = new Date(t + (3 - day) * DAY_MS); const y = thu.getUTCFullYear();
  const jan4 = Date.UTC(y, 0, 4); const w1 = jan4 - weekdayMon0(jan4) * DAY_MS;
  const wk = 1 + Math.floor((t - w1) / (7 * DAY_MS));
  void d; return `${y}-W${String(wk).padStart(2, "0")}`;
}
function monthDays(monthKey) {
  const [y, m] = monthKey.split("-").map(Number); const out = [];
  for (let t = Date.UTC(y, m - 1, 1); new Date(t).getUTCMonth() === m - 1; t += DAY_MS) out.push(t);
  return out;
}
function thirdSundayJune(y) { let n = 0; for (let t = Date.UTC(y, 5, 1); ; t += DAY_MS) { if (weekdayMon0(t) === 6 && ++n === 3) return t; } }

/** Semanas ISO del mes, recortadas al mes: [{semana, inicio, fin, dias}] */
function monthWeeks(monthKey) {
  const seg = new Map();
  for (const t of monthDays(monthKey)) { const k = isoWeekKey(t); if (!seg.has(k)) seg.set(k, []); seg.get(k).push(t); }
  return [...seg.entries()].map(([semana, ds]) => ({ semana, inicio: isoOf(ds[0]), fin: isoOf(ds[ds.length - 1]), dias: ds.length, _days: ds }));
}

function indexDaily(dailySales, beforeT) {
  const byDay = new Map(); // t -> Map(`${suc}\u0000${prod}` -> q)
  for (const r of dailySales || []) {
    const q = Number(r.cantidad); if (!r.fecha || !Number.isFinite(q) || q === 0) continue;
    const t = toUTC(String(r.fecha)); if (!(t < beforeT)) continue; // sin fuga: solo antes del mes
    const k = `${r.sucursal}\u0000${r.producto}`;
    if (!byDay.has(t)) byDay.set(t, new Map());
    const m = byDay.get(t); m.set(k, (m.get(k) || 0) + q);
  }
  return byDay;
}

function eventFactors(monthKey, byDay, o) {
  const f = new Map(); if (!o.events) return f;
  const [y, m] = monthKey.split("-").map(Number);
  const evs = FIXED_EVENTS.filter(([mm]) => mm === m).map(([mm, dd]) => [Date.UTC(y, mm - 1, dd), Date.UTC(y - 1, mm - 1, dd)]);
  if (m === 6) evs.push([thirdSundayJune(y), thirdSundayJune(y - 1)]);
  const clamp = (x) => Math.max(o.eventMin, Math.min(o.eventMax, x));
  for (const [ev, ly] of evs) {
    const lyMonth = isoOf(ly).slice(0, 7);
    const same = monthDays(lyMonth).filter((t) => weekdayMon0(t) === weekdayMon0(ly) && t !== ly);
    const tot = (t, p) => { let s = 0; for (const [k, q] of byDay.get(t) || []) if (p == null || k.split("\u0000")[1] === p) s += q; return s; };
    const baseAll = same.reduce((a, t) => a + tot(t), 0) / Math.max(1, same.length);
    const vAll = tot(ly); const facAll = baseAll > 0 ? vAll / baseAll : 1;
    if (baseAll <= 0 && vAll <= 0) continue; // sin historia del año anterior: no hay factor
    f.set(`${ev}\u0000*`, clamp(facAll));
    const prods = new Set();
    for (const t of [...same, ly]) for (const k of (byDay.get(t) || new Map()).keys()) prods.add(k.split("\u0000")[1]);
    for (const p of prods) {
      const b = same.reduce((a, t) => a + tot(t, p), 0) / Math.max(1, same.length); const v = tot(ly, p);
      f.set(`${ev}\u0000${p}`, clamp((v + facAll * o.eventK) / (b + o.eventK)));
    }
  }
  return f;
}

function eventDaysOfYear(y) {
  const s = new Set(FIXED_EVENTS.map(([mm, dd]) => Date.UTC(y, mm - 1, dd)));
  s.add(thirdSundayJune(y));
  return s;
}
// Vísperas (1 y 2 días antes) de los eventos del año y del siguiente, sin días de evento del año.
function visperaDaysOfYear(y) {
  const ev = new Set([...eventDaysOfYear(y), ...eventDaysOfYear(y + 1)]);
  const out = new Set();
  for (const e of ev) for (const k of [1, 2]) { const t = e - k * DAY_MS; if (!ev.has(t)) out.add(t); }
  return out;
}

/**
 * Factor de víspera por SKU (solo SKU que históricamente suben la víspera).
 * Devuelve Map(`${t}\u0000${producto}` -> factor) para las vísperas del mes.
 */
function visperaFactors(monthKey, byDay, products, o) {
  const f = new Map(); if (!o.visperas) return f;
  const days = monthDays(monthKey); const start = days[0]; const y = Number(monthKey.slice(0, 4));
  const evAll = new Set(); const vsAll = new Set();
  for (let yy = y - 3; yy <= y; yy++) { for (const t of eventDaysOfYear(yy)) evAll.add(t); for (const t of visperaDaysOfYear(yy)) vsAll.add(t); }
  const winDays = [...byDay.keys()].filter((t) => t >= start - o.vispWindowDays * DAY_MS && t < start);
  const normal = new Map();
  for (const t of winDays) {
    const d = new Date(t); const dom = d.getUTCDate();
    if (evAll.has(t) || vsAll.has(t) || (dom >= 15 && dom <= 17)) continue;
    const key = `${d.getUTCFullYear()}-${d.getUTCMonth()}-${weekdayMon0(t)}`;
    if (!normal.has(key)) normal.set(key, []);
    normal.get(key).push(t);
  }
  const byProd = (t) => { const m = new Map(); for (const [k, q] of byDay.get(t) || []) { const p = k.split("\u0000")[1]; m.set(p, (m.get(p) || 0) + q); } return m; };
  const sv = new Map(); const sb = new Map(); const nv = new Map();
  for (const t of winDays) {
    if (!vsAll.has(t)) continue;
    const d = new Date(t);
    const nd = normal.get(`${d.getUTCFullYear()}-${d.getUTCMonth()}-${weekdayMon0(t)}`);
    if (!nd || nd.length < 2) continue;
    const base = new Map();
    for (const x of nd) for (const [p, q] of byProd(x)) base.set(p, (base.get(p) || 0) + q / nd.length);
    const cur = byProd(t);
    for (const p of new Set([...base.keys(), ...cur.keys()])) {
      const b = base.get(p) || 0; if (!(b > 0)) continue;
      sv.set(p, (sv.get(p) || 0) + (cur.get(p) || 0)); sb.set(p, (sb.get(p) || 0) + b); nv.set(p, (nv.get(p) || 0) + 1);
    }
  }
  const evMonth = eventDaysOfYear(y); const vsMonth = visperaDaysOfYear(y);
  const vd = days.filter((t) => vsMonth.has(t) && !evMonth.has(t));
  if (!vd.length) return f;
  for (const p of products) {
    if ((nv.get(p) || 0) < o.vispMinObs || !((sb.get(p) || 0) > 0)) continue;
    const r = ((sv.get(p) || 0) + o.vispK) / (sb.get(p) + o.vispK);
    if (r < o.vispMin) continue;
    const fac = Math.max(1, Math.min(o.vispMax, r));
    for (const t of vd) f.set(`${t}\u0000${p}`, fac);
  }
  return f;
}

/**
 * @param {object} args
 * @param {string} args.month 'YYYY-MM'
 * @param {Object<string,number>|Array<{producto:string,cantidad:number}>} args.monthlyForecast
 * @param {Array<{fecha:string,sucursal:string,producto:string,cantidad:number}>} args.dailySales ventas de piso de venta de sucursales (no traspasos de planta)
 * @param {object} [args.options]
 * @returns {{semanas:Array, porSkuSemana:Array, porSucursalSkuSemana:Array, sucursales:Array}}
 */
function disaggregateMonthlyForecast({ month, monthlyForecast, dailySales, options = {} }) {
  if (!/^\d{4}-\d{2}$/.test(String(month || ""))) throw new Error("month debe ser YYYY-MM");
  const o = { ...DEFAULTS, ...options };
  const F = Array.isArray(monthlyForecast)
    ? Object.fromEntries(monthlyForecast.map((r) => [r.producto, Number(r.cantidad) || 0]))
    : { ...(monthlyForecast || {}) };
  const days = monthDays(month); const start = days[0];
  const byDay = indexDaily(dailySales, start);

  const dowP = new Map(); const dowAll = new Array(7).fill(0);
  const bp = new Map(); const bAll = new Map(); const pTot = new Map(); const lastSale = new Map();
  for (let i = o.windowDays; i >= 1; i--) {
    const t = start - i * DAY_MS; const inShare = i <= o.shareWindowDays;
    for (const [k, q] of byDay.get(t) || []) {
      const [s, p] = k.split("\u0000"); const wd = weekdayMon0(t);
      if (!dowP.has(p)) dowP.set(p, new Array(7).fill(0));
      dowP.get(p)[wd] += q; dowAll[wd] += q;
      if (inShare) { bp.set(k, (bp.get(k) || 0) + q); bAll.set(s, (bAll.get(s) || 0) + q); pTot.set(p, (pTot.get(p) || 0) + q); }
      if (q > 0) lastSale.set(s, Math.max(lastSale.get(s) || 0, t));
    }
  }
  const recent = [...lastSale.entries()].filter(([, t]) => t >= start - o.activeDays * DAY_MS).map(([s]) => s).sort();
  const totRecent = recent.reduce((a, s) => a + Math.max(0, bAll.get(s) || 0), 0);
  const sb = new Map(recent.map((s) => [s, totRecent > 0 ? Math.max(0, bAll.get(s) || 0) / totRecent : 1 / recent.length]));
  const sa = dowAll.reduce((a, b) => a + b, 0);
  const wall = sa > 0 ? dowAll.map((x) => x / sa) : new Array(7).fill(1 / 7);
  const ev = eventFactors(month, byDay, o);
  const visp = visperaFactors(month, byDay, Object.keys(F), o);
  const weeks = monthWeeks(month);

  const porSkuSemana = []; const porSucursalSkuSemana = [];
  for (const [p, fmRaw] of Object.entries(F)) {
    const Fm = Number(fmRaw) || 0; if (Fm <= 0) continue;
    const dp = dowP.get(p); const n = dp ? dp.reduce((a, b) => a + b, 0) : 0;
    const wp = n > 0 ? dp.map((x, k) => Math.max(0, (x + o.beta * wall[k]) / (n + o.beta))) : wall;
    const dw = new Map(days.map((t) => [t, wp[weekdayMon0(t)] * (ev.get(`${t}\u0000${p}`) ?? ev.get(`${t}\u0000*`) ?? 1) * (visp.get(`${t}\u0000${p}`) ?? 1)]));
    let sdw = 0; for (const v of dw.values()) sdw += v;
    const npk = Math.max(0, pTot.get(p) || 0);
    const shares = recent.map((s) => [s, (Math.max(0, bp.get(`${s}\u0000${p}`) || 0) + o.alpha * sb.get(s)) / (npk + o.alpha)]);
    const ss = shares.reduce((a, [, v]) => a + v, 0);
    for (const w of weeks) {
      const fw = sdw > 0 ? Fm * w._days.reduce((a, t) => a + dw.get(t), 0) / sdw : Fm * w.dias / days.length;
      porSkuSemana.push({ producto: p, semana: w.semana, inicio: w.inicio, fin: w.fin, cantidad: fw });
      for (const [s, sh] of shares) porSucursalSkuSemana.push({ sucursal: s, producto: p, semana: w.semana, inicio: w.inicio, fin: w.fin, cantidad: ss > 0 ? fw * sh / ss : 0 });
    }
  }
  return {
    semanas: weeks.map(({ _days, ...w }) => w),
    porSkuSemana, porSucursalSkuSemana, sucursales: recent, opciones: o,
    skuVispera: [...new Set([...visp.keys()].map((k) => k.split("\u0000")[1]))].sort(),
  };
}

module.exports = { disaggregateMonthlyForecast, monthWeeks, isoWeekKey, visperaDaysOfYear, DEFAULTS };
