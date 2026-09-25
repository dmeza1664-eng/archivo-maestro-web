#!/usr/bin/env node
/**
 * CLI: reparte el pronóstico mensual por SKU a semana ISO (recortada al mes) y a sucursal.
 * No modifica el pronóstico mensual oficial; solo lo desagrega.
 *
 * Uso:
 *   node scripts/weekly-branch-forecast.cjs --month 2025-10 \
 *        --forecast pronostico-mensual.json   (objeto {producto: cantidad} o arreglo [{producto, cantidad}])
 *        --daily ventas-diarias.csv           (columnas: fecha,sucursal,producto,cantidad; ventas de piso de sucursales)
 *        [--out salida.csv] [--level sucursal|sku] [--alpha 30] [--beta 50] [--window 56] [--no-events]
 * Solo se usan las filas diarias con fecha anterior al mes pedido.
 */
"use strict";
const fs = require("fs");
const path = require("path");
const { disaggregateMonthlyForecast } = require(path.join(__dirname, "lib", "weekly-branch-disaggregation.cjs"));

function parseArgs(argv) {
  const a = {}; for (let i = 0; i < argv.length; i++) {
    const k = argv[i]; if (!k.startsWith("--")) continue;
    const key = k.slice(2); if (key === "no-events") { a.events = false; continue; }
    a[key] = argv[++i];
  } return a;
}
function parseCsvLine(line) {
  const out = []; let cur = ""; let q = false;
  for (let i = 0; i < line.length; i++) {
    const c = line[i];
    if (q) { if (c === '"' && line[i + 1] === '"') { cur += '"'; i++; } else if (c === '"') q = false; else cur += c; }
    else if (c === '"') q = true; else if (c === ",") { out.push(cur); cur = ""; } else cur += c;
  }
  out.push(cur); return out;
}
function readDaily(file) {
  const lines = fs.readFileSync(file, "utf8").replace(/^\uFEFF/, "").split(/\r?\n/).filter(Boolean);
  const head = parseCsvLine(lines[0]).map((h) => h.trim().toLowerCase());
  const ix = (n) => { const i = head.indexOf(n); if (i < 0) throw new Error(`falta la columna ${n} en ${file}`); return i; };
  const [f, s, p, c] = ["fecha", "sucursal", "producto", "cantidad"].map(ix);
  return lines.slice(1).map((l) => { const r = parseCsvLine(l); return { fecha: r[f], sucursal: r[s], producto: r[p], cantidad: Number(r[c]) }; });
}
const csvCell = (v) => (/[",\n]/.test(String(v)) ? `"${String(v).replace(/"/g, '""')}"` : String(v));

function main() {
  const a = parseArgs(process.argv.slice(2));
  if (!a.month || !a.forecast || !a.daily) { console.error("Uso: --month YYYY-MM --forecast archivo.json --daily ventas.csv [--out salida.csv]"); process.exit(2); }
  const options = {};
  for (const [k, o] of [["alpha", "alpha"], ["beta", "beta"], ["window", "windowDays"], ["share-window", "shareWindowDays"]]) if (a[k] != null) options[o] = Number(a[k]);
  if (a.events === false) options.events = false;
  const res = disaggregateMonthlyForecast({ month: a.month, monthlyForecast: JSON.parse(fs.readFileSync(a.forecast, "utf8")), dailySales: readDaily(a.daily), options });
  const bySku = a.level === "sku";
  const rows = bySku ? res.porSkuSemana : res.porSucursalSkuSemana;
  const cols = bySku ? ["producto", "semana", "inicio", "fin", "cantidad"] : ["sucursal", "producto", "semana", "inicio", "fin", "cantidad"];
  const text = [cols.join(","), ...rows.map((r) => cols.map((c) => csvCell(c === "cantidad" ? Math.round(r[c] * 100) / 100 : r[c])).join(","))].join("\n") + "\n";
  if (a.out) { fs.writeFileSync(a.out, text); console.log(`${rows.length} filas → ${a.out} (${res.semanas.length} semanas, ${res.sucursales.length} sucursales)`); }
  else process.stdout.write(text);
}
if (require.main === module) main();
module.exports = { readDaily, parseCsvLine };
