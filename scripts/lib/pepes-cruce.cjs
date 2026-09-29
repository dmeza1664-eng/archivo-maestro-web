/**
 * Pronóstico walk-forward y cruce producción contra pronóstico (Pastelería Pepes).
 *
 * Cada mes se pronostica solo con meses anteriores (filterVentasBeforeMonth) y
 * con el mismo llamado a calculateForecast del backtest oficial (colchón 10%,
 * sin existencias ni bajas), así que las cifras coinciden con el arnés oficial.
 * Plan = producción sugerida del modelo (pronóstico de sucursales + colchón,
 * con la regla operativa de pasteles).
 */
const r2 = (x) => +Number(x || 0).toFixed(2);
const pct = (num, den) => (den ? +((num / den) * 100).toFixed(2) : null);

function pronosticarMes(app, stockRows, serie, mes, colchonPct) {
  const historia = app.filterVentasBeforeMonth(serie, mes);
  return app.calculateForecast({
    stockRows, historicalVentas: historia, bajas: [], existencias: [], realProduction: [],
    selectedMonth: mes, dailyBufferPct: colchonPct,
  });
}

/** Evalúa un mes cerrado. Devuelve null si el mes no tiene venta. */
function evaluarMes(app, stockRows, serie, mes, colchonPct) {
  const fr = pronosticarMes(app, stockRows, serie, mes, colchonPct);
  const am = new Map();
  for (const r of serie) if (r.mes === mes) am.set(r.producto, (am.get(r.producto) || 0) + r.cantidad);
  if (![...am.values()].some((v) => v > 0)) return null;
  const an = app.analyzeForecastProductErrors(fr, am, { topN: 15 });
  const plan = new Map(fr.map((r) => [r.producto, r.produccionSugerida]));
  const metodo = new Map(fr.map((r) => [r.producto, r.metodoPronostico]));
  const abs = an.rows.reduce((s, r) => s + r.absoluteError, 0);
  return {
    mes,
    wape: r2(an.wape), actual: r2(an.actual), forecast: r2(an.forecast), abs: r2(abs),
    productos: an.rows.map((r) => ({ producto: r.producto, pronostico: r2(r.forecast), venta: r2(r.actual), plan: plan.get(r.producto) ?? null, metodo: metodo.get(r.producto) || "" })),
  };
}

function metricasDeMeses(evaluados, filtro = () => true) {
  let abs = 0; let act = 0; let absTotal = 0; let neto = 0; let n = 0;
  for (const e of evaluados) {
    if (!filtro(e.mes)) continue;
    abs += e.abs; act += e.actual; n += 1;
    const F = e.productos.reduce((s, p) => s + p.pronostico, 0);
    const A = e.productos.reduce((s, p) => s + p.venta, 0);
    absTotal += Math.abs(F - A); neto += F - A;
  }
  return n ? { meses: n, wape: pct(abs, act), errorTotalMensual: pct(absTotal, act), netoPct: pct(neto, act) } : null;
}

function metricasPorAnio(evaluados) {
  const anios = [...new Set(evaluados.map((e) => e.mes.slice(0, 4)))].sort();
  const out = {};
  for (const y of anios) {
    const delAnio = evaluados.filter((e) => e.mes.startsWith(y));
    const base = metricasDeMeses(delAnio);
    const sinEne = metricasDeMeses(delAnio, (m) => !m.endsWith("-01"));
    out[y] = {
      primerMes: delAnio[0].mes, ultimoMes: delAnio[delAnio.length - 1].mes, meses: base.meses,
      wape: base.wape, wapeSinEnero: sinEne ? sinEne.wape : null, errorTotalMensual: base.errorTotalMensual, netoPct: base.netoPct,
      porMes: Object.fromEntries(delAnio.map((e) => [e.mes.slice(5), e.wape])),
    };
  }
  return out;
}

/** Filas producto × mes del cruce (venta de sucursales, pronóstico, plan, producción real). */
function cruceProductoMes(evaluados, produccion) {
  const filas = [];
  for (const e of evaluados) {
    for (const p of e.productos) {
      const prod = produccion.get(`${e.mes}\u0000${p.producto}`);
      const producida = prod ? prod.producida : 0;
      const plan = p.plan || 0;
      if (!p.venta && !p.pronostico && !producida && !plan) continue;
      filas.push({
        mes: e.mes, producto: p.producto,
        venta_sucursales: p.venta, pronostico: p.pronostico, plan_modelo: plan, produccion_real: producida,
        error_pronostico: r2(p.pronostico - p.venta), error_abs: r2(Math.abs(p.pronostico - p.venta)),
        error_pct: p.venta > 0 ? pct(Math.abs(p.pronostico - p.venta), p.venta) : null,
        produccion_menos_venta: r2(producida - p.venta),
        sobrante_real: r2(Math.max(0, producida - p.venta)), faltante_real: r2(Math.max(0, p.venta - producida)),
        sobrante_plan: r2(Math.max(0, plan - p.venta)), faltante_plan: r2(Math.max(0, p.venta - plan)),
        plan_menos_produccion: r2(plan - producida),
        lineas_produccion_abiertas: prod && prod.lineasAbiertas != null ? prod.lineasAbiertas : "",
        metodo: p.metodo,
      });
    }
  }
  return filas;
}

const CAMPOS_SUMA = ["venta_sucursales", "pronostico", "plan_modelo", "produccion_real", "error_abs", "sobrante_real", "faltante_real", "sobrante_plan", "faltante_plan"];

function sumar(filas) {
  const s = Object.fromEntries(CAMPOS_SUMA.map((k) => [k, 0]));
  for (const f of filas) for (const k of CAMPOS_SUMA) s[k] += Number(f[k]) || 0;
  return s;
}

function redondearResumen(s) {
  const out = {};
  for (const [k, v] of Object.entries(s)) out[k] = typeof v === "number" ? Math.round(v) : v;
  return out;
}

function cruceMensual(evaluados, filas) {
  return evaluados.map((e) => {
    const s = sumar(filas.filter((f) => f.mes === e.mes));
    const F = e.productos.reduce((a, p) => a + p.pronostico, 0);
    const A = e.productos.reduce((a, p) => a + p.venta, 0);
    return {
      mes: e.mes, ...redondearResumen(s),
      wape_producto_mes: e.wape, error_total_mes_pct: pct(Math.abs(F - A), A), sesgo_pct: pct(F - A, A),
      produccion_vs_venta_pct: pct(s.produccion_real - s.venta_sucursales, s.venta_sucursales),
      plan_vs_produccion_pct: pct(s.plan_modelo - s.produccion_real, s.produccion_real),
    };
  });
}

function cruceProductoAnual(filas) {
  const grupos = new Map();
  for (const f of filas) {
    const k = `${f.mes.slice(0, 4)}\u0000${f.producto}`;
    if (!grupos.has(k)) grupos.set(k, []);
    grupos.get(k).push(f);
  }
  return [...grupos.entries()].map(([k, fs]) => {
    const [anio, producto] = k.split("\u0000");
    const s = sumar(fs);
    return {
      anio, producto, meses: fs.length, ...redondearResumen(s),
      wape_pct: pct(s.error_abs, s.venta_sucursales),
      produccion_vs_venta_pct: pct(s.produccion_real - s.venta_sucursales, s.venta_sucursales),
    };
  }).sort((a, b) => a.anio.localeCompare(b.anio) || b.venta_sucursales - a.venta_sucursales || a.producto.localeCompare(b.producto, "es"));
}

module.exports = { pronosticarMes, evaluarMes, metricasDeMeses, metricasPorAnio, cruceProductoMes, cruceMensual, cruceProductoAnual, sumar, pct, r2 };
