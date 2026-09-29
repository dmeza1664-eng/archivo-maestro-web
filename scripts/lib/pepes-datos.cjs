/**
 * Datos del Pronóstico (Pastelería Pepes): lectura de extractos CSV, mapeo de
 * productos y construcción de la serie oficial de demanda (venta de sucursales
 * al público, mes × producto del catálogo stock_ideal).
 *
 * Reproduce exacto los constructores que se usaron para las cifras oficiales
 * (build_series.py 2023–2024 y build2026.py 2024–2026):
 *  - el producto sale de mapeo-productos.csv por nProductoPK (solo filas con
 *    inCatalog = true y producto);
 *  - un nProductoPK que no está en el mapeo solo se asigna, desde 2026-01, por
 *    código exacto (stock-codigo-map.json), la misma regla del mapeo oficial;
 *  - la venta de sucursales excluye el surtido de Planta León y a Suc. Amado
 *    Nervo antes de ago-2024. Esa decisión la toma la app (filterDemandSales),
 *    así la serie y el modelo usan una sola definición.
 * Sin reglas por nombre de SKU; no mira el mes pronosticado.
 */
const fs = require("fs");

const VENTAS_COLUMNAS = ["mes", "nProductoPK", "cCodigo", "cDescripcion", "nBodegaPK", "bodegaNombre", "nSucursalPK", "sucursalNombre", "qty", "subtotal", "total", "lineas"];
const PRODUCCION_COLUMNAS = ["mes", "nProductoPK", "cCodigo", "cDescripcion", "qtyPedida", "qtyProducida", "qtyCancelada", "qtySalida", "lineas"];
const CODIGO_FALLBACK_DESDE = "2026-01";

// CSV (RFC 4180): comas, comillas dobles y saltos de línea dentro de comillas.
function parseCsv(text) {
  const src = String(text || "").replace(/^\uFEFF/, "");
  const rows = [];
  let field = "";
  let row = [];
  let quoted = false;
  for (let i = 0; i < src.length; i += 1) {
    const ch = src[i];
    if (quoted) {
      if (ch === '"') {
        if (src[i + 1] === '"') { field += '"'; i += 1; } else quoted = false;
      } else field += ch;
    } else if (ch === '"') quoted = true;
    else if (ch === ",") { row.push(field); field = ""; }
    else if (ch === "\n" || ch === "\r") {
      if (ch === "\r" && src[i + 1] === "\n") i += 1;
      row.push(field); field = "";
      if (row.length > 1 || row[0] !== "") rows.push(row);
      row = [];
    } else field += ch;
  }
  if (field !== "" || row.length) { row.push(field); if (row.length > 1 || row[0] !== "") rows.push(row); }
  if (!rows.length) return [];
  const header = rows[0];
  return rows.slice(1).map((values) => Object.fromEntries(header.map((h, idx) => [h, values[idx] ?? ""])));
}

function readCsv(file) {
  return parseCsv(fs.readFileSync(file, "utf8"));
}

function csvCell(value) {
  if (value == null) return "";
  const s = value instanceof Date ? value.toISOString() : String(value);
  return /[",\n\r]/.test(s) ? `"${s.replace(/"/g, '""')}"` : s;
}

function toCsv(rows, columns) {
  return [columns.join(","), ...rows.map((r) => columns.map((c) => csvCell(r[c])).join(","))].join("\n") + "\n";
}

function loadMapeo(mapeoCsv, codigosJson) {
  const porPK = new Map(readCsv(mapeoCsv).map((r) => [String(r.nProductoPK), r]));
  const porCodigo = codigosJson && fs.existsSync(codigosJson) ? JSON.parse(fs.readFileSync(codigosJson, "utf8")) : {};
  return { porPK, porCodigo };
}

function resolverProducto(row, mapeo) {
  const m = mapeo.porPK.get(String(row.nProductoPK));
  if (m) return m.inCatalog === "true" && m.producto ? m.producto : null;
  if (String(row.mes) >= CODIGO_FALLBACK_DESDE) {
    const c = mapeo.porCodigo[String(row.cCodigo ?? "")];
    if (c && c.inCatalog && c.producto) return c.producto;
  }
  return null;
}

function diasDelMes(mes) {
  const [y, m] = String(mes).split("-").map(Number);
  return new Date(y, m, 0).getDate();
}

function siguienteMes(mes, delta = 1) {
  const [y, m] = String(mes).split("-").map(Number);
  const d = new Date(y, m - 1 + delta, 1);
  return `${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, "0")}`;
}

// Clasificador de canal con la definición de la app (filterDemandSales), con
// caché por sucursal y mes.
function crearClasificadorDemanda(app) {
  const cache = new Map();
  return (sucursalNombre, mes) => {
    const key = `${sucursalNombre}\u0000${mes}`;
    if (!cache.has(key)) {
      const [y, m] = String(mes).split("-").map(Number);
      const probe = { sucursal: sucursalNombre, canal: sucursalNombre, fecha: new Date(y, m - 1, 15, 12), producto: "_", cantidad: 1 };
      cache.set(key, app.filterDemandSales([probe]).length === 1);
    }
    return cache.get(key);
  };
}

/**
 * Serie oficial (mes × producto) de venta de sucursales.
 * ventas: filas mes × producto × bodega (extracto o base), en el orden de lectura.
 * desde/hasta: meses "YYYY-MM" inclusivos (null = sin límite).
 */
function construirSerieSucursal(ventas, mapeo, esDemanda, { desde = null, hasta = null } = {}) {
  const agg = new Map();
  const sinMapeo = new Map();
  for (const r of ventas) {
    const mes = String(r.mes);
    if ((desde && mes < desde) || (hasta && mes > hasta)) continue;
    const producto = resolverProducto(r, mapeo);
    if (!producto) {
      if (!mapeo.porPK.has(String(r.nProductoPK))) {
        const k = String(r.nProductoPK);
        const cur = sinMapeo.get(k) || { nProductoPK: k, cCodigo: r.cCodigo, cDescripcion: r.cDescripcion, piezas: 0 };
        cur.piezas += Number(r.qty) || 0;
        sinMapeo.set(k, cur);
      }
      continue;
    }
    if (!esDemanda(r.sucursalNombre, mes)) continue;
    const key = `${mes}\u0000${producto}`;
    agg.set(key, (agg.get(key) || 0) + (Number(r.qty) || 0));
  }
  const keys = [...agg.keys()].sort((a, b) => (a < b ? -1 : a > b ? 1 : 0));
  const rows = keys.map((key) => {
    const [mes, producto] = key.split("\u0000");
    return {
      producto, productoOriginal: producto, fecha: `${mes}-15T19:00:00.000Z`, cantidad: agg.get(key),
      monthlyTotal: true, monthDays: diasDelMes(mes), mes, source: "pepes_devBI",
    };
  });
  return { rows, sinMapeo: [...sinMapeo.values()].sort((a, b) => b.piezas - a.piezas) };
}

// Filas de la serie listas para la app (fecha como Date), igual que el arnés oficial.
function hidratarSerie(rows) {
  return rows.map((r) => ({ ...r, fecha: new Date(r.fecha), cantidad: Number(r.cantidad) || 0, monthlyTotal: true, monthDays: r.monthDays || diasDelMes(r.mes) }));
}

/** Producción real mes × producto del catálogo (qtyProducida = SUM(nProducidos)). */
function agregarProduccion(produccion, mapeo) {
  const out = new Map();
  for (const r of produccion) {
    const producto = resolverProducto(r, mapeo);
    if (!producto) continue;
    const key = `${r.mes}\u0000${producto}`;
    const cur = out.get(key) || { producida: 0, pedida: 0, cancelada: 0, lineasAbiertas: null };
    cur.producida += Number(r.qtyProducida) || 0;
    cur.pedida += Number(r.qtyPedida) || 0;
    cur.cancelada += Number(r.qtyCancelada) || 0;
    if (r.lineasAbiertas != null && r.lineasAbiertas !== "") cur.lineasAbiertas = (cur.lineasAbiertas || 0) + (Number(r.lineasAbiertas) || 0);
    out.set(key, cur);
  }
  return out;
}

/** Venta de sucursales mes × producto (incluye meses incompletos; sirve para el avance del mes). */
function agregarVentaSucursal(ventas, mapeo, esDemanda) {
  const out = new Map();
  for (const r of ventas) {
    const producto = resolverProducto(r, mapeo);
    if (!producto || !esDemanda(r.sucursalNombre, String(r.mes))) continue;
    const key = `${r.mes}\u0000${producto}`;
    out.set(key, (out.get(key) || 0) + (Number(r.qty) || 0));
  }
  return out;
}

/** Último mes completo según la fecha de corte (última venta registrada, YYYY-MM-DD). */
function ultimoMesCompleto(fechaCorte) {
  const [y, m, d] = String(fechaCorte).slice(0, 10).split("-").map(Number);
  const mes = `${y}-${String(m).padStart(2, "0")}`;
  return d >= new Date(y, m, 0).getDate() ? mes : siguienteMes(mes, -1);
}

module.exports = {
  VENTAS_COLUMNAS, PRODUCCION_COLUMNAS, CODIGO_FALLBACK_DESDE,
  parseCsv, readCsv, toCsv, csvCell, loadMapeo, resolverProducto, diasDelMes, siguienteMes,
  crearClasificadorDemanda, construirSerieSucursal, hidratarSerie, agregarProduccion, agregarVentaSucursal, ultimoMesCompleto,
};
