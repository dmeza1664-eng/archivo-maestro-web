/**
 * Fuente "archivos": lee los extractos CSV de pepes_devBI que ya existen
 * (mismas consultas de solo lectura que el modo base):
 *   ventas-YYYY-mensual-producto-bodega.csv   o  part-ventas-bodega-YYYY-MM.csv
 *   produccion-YYYY-mensual-producto.csv      o  part-produccion-YYYY-MM.csv
 *   extract-meta*.json (maxVentaDCreada = última venta registrada)
 * Si para un mes hay archivo anual y archivo parcial, se usa el anual (nunca los dos).
 */
const fs = require("fs");
const path = require("path");
const { readCsv } = require("./pepes-datos.cjs");

const PATRONES = [
  { tipo: "ventas", re: /^ventas-(\d{4})-mensual-producto-bodega\.csv$/, anual: true },
  { tipo: "ventas", re: /^part-ventas-bodega-(\d{4})-(\d{2})\.csv$/, anual: false },
  { tipo: "produccion", re: /^produccion-(\d{4})-mensual-producto\.csv$/, anual: true },
  { tipo: "produccion", re: /^part-produccion-(\d{4})-(\d{2})\.csv$/, anual: false },
];

function separarDirectorios(valor) {
  const sep = process.platform === "win32" ? /[;,]/ : /[:;,]/;
  return String(valor || "").split(sep).map((s) => s.trim()).filter(Boolean);
}

function descubrir(dirs) {
  const archivos = [];
  const metas = [];
  for (const dir of dirs) {
    if (!fs.existsSync(dir) || !fs.statSync(dir).isDirectory()) throw new Error(`No existe la carpeta de datos: ${dir}`);
    for (const name of fs.readdirSync(dir).sort()) {
      const file = path.join(dir, name);
      if (/^extract-meta.*\.json$/.test(name)) { metas.push(file); continue; }
      for (const p of PATRONES) {
        const m = name.match(p.re);
        if (m) archivos.push({ tipo: p.tipo, anual: p.anual, anio: m[1], mes: m[2] ? `${m[1]}-${m[2]}` : null, file });
      }
    }
  }
  return { archivos, metas };
}

function leerTipo(archivos, tipo) {
  const anuales = archivos.filter((a) => a.tipo === tipo && a.anual).sort((a, b) => a.anio.localeCompare(b.anio) || a.file.localeCompare(b.file));
  const parciales = archivos.filter((a) => a.tipo === tipo && !a.anual).sort((a, b) => a.mes.localeCompare(b.mes) || a.file.localeCompare(b.file));
  const filas = [];
  const usados = [];
  const mesesCubiertos = new Set();
  const aniosAnuales = new Set();
  for (const a of anuales) {
    if (aniosAnuales.has(a.anio)) throw new Error(`Hay dos archivos anuales de ${tipo} ${a.anio}; deja solo uno (${a.file}).`);
    aniosAnuales.add(a.anio);
    const rows = readCsv(a.file);
    for (const r of rows) mesesCubiertos.add(String(r.mes));
    filas.push(...rows);
    usados.push({ archivo: a.file, filas: rows.length });
  }
  const mesesParciales = new Set();
  for (const a of parciales) {
    if (mesesCubiertos.has(a.mes) || aniosAnuales.has(a.anio)) continue;
    if (mesesParciales.has(a.mes)) throw new Error(`Hay dos archivos parciales de ${tipo} ${a.mes}; deja solo uno (${a.file}).`);
    mesesParciales.add(a.mes);
    const rows = readCsv(a.file);
    filas.push(...rows);
    usados.push({ archivo: a.file, filas: rows.length });
  }
  return { filas, usados };
}

function fechaCorteDeMetas(metas) {
  let max = null;
  for (const file of metas) {
    try {
      const meta = JSON.parse(fs.readFileSync(file, "utf8"));
      const v = meta.maxVentaDCreada ? String(meta.maxVentaDCreada).slice(0, 10) : null;
      if (v && (!max || v > max)) max = v;
    } catch { /* meta ilegible: se ignora */ }
  }
  return max;
}

/** Lee todos los extractos de las carpetas indicadas. */
function leerArchivos({ dirs }) {
  if (!dirs.length) throw new Error("Modo archivos: falta PEPES_DATOS_DIRS (carpetas con los extractos CSV).");
  const { archivos, metas } = descubrir(dirs);
  const ventas = leerTipo(archivos, "ventas");
  const produccion = leerTipo(archivos, "produccion");
  if (!ventas.filas.length) throw new Error(`Modo archivos: no se encontraron extractos de ventas en ${dirs.join(", ")}.`);
  const advertencias = [];
  let fechaCorte = fechaCorteDeMetas(metas);
  let origenCorte = "maxVentaDCreada de extract-meta";
  if (!fechaCorte) {
    const ultimo = ventas.filas.reduce((m, r) => (String(r.mes) > m ? String(r.mes) : m), "");
    const [y, mo] = ultimo.split("-").map(Number);
    fechaCorte = `${ultimo}-${String(new Date(y, mo, 0).getDate()).padStart(2, "0")}`;
    origenCorte = "último mes con venta (sin extract-meta; se supone completo)";
    advertencias.push(`No hay maxVentaDCreada en extract-meta; se supone completo ${ultimo}. Usa PEPES_FECHA_CORTE si no lo está.`);
  }
  return {
    modo: "archivos",
    fuente: `extractos CSV en ${dirs.join(", ")}`,
    ventas: ventas.filas,
    produccion: produccion.filas,
    fechaCorte,
    origenCorte,
    archivosLeidos: [...ventas.usados, ...produccion.usados],
    advertencias,
  };
}

module.exports = { leerArchivos, separarDirectorios, descubrir };
