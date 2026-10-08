/**
 * Prueba del prototipo de CSV para CRM. Solo usa archivos locales del repo.
 */
const fs = require("fs");
const os = require("os");
const path = require("path");
const { readCsv } = require("../lib/pepes-datos.cjs");
const crm = require("./generar-csv-importacion-crm.cjs");

function assert(condition, message) {
  if (!condition) throw new Error(`[crm-csv] ${message}`);
}

const tmp = fs.mkdtempSync(path.join(os.tmpdir(), "crm-csv-"));
const r = crm.generar({ outDir: tmp });

assert(r.empresas.length === crm.SUCURSALES_FIXTURE.length, "todas las sucursales fixture salen como empresas");
assert(r.empresas.every((e) => e.phone === "" && e.email === ""), "las sucursales no inventan teléfono ni correo");
assert(r.empresas.some((e) => e.record_type === "planta" && e.name === "Planta León"), "Planta León queda marcada como planta");
assert(r.productos.length > 50, "el mapeo del repo trae decenas de productos de catálogo");
assert(r.productos.every((p) => p.external_id.startsWith("pepes-prod-") && p.sku && p.name), "productos con id externo, sku y nombre");
assert(r.contactos.length === 3, "tres tiendas sintéticas de galletas");
assert(r.contactos.every((c) => /^\+52550000000\d$/.test(c.phone)), "teléfonos de ejemplo, no reales");
assert(r.contactos.every((c) => /ejemplo/i.test(c.lastname) || /ejemplo/i.test(c.company)), "contactos marcados como ejemplo");

for (const e of r.escritos) {
  assert(fs.existsSync(e.file), `se escribió ${e.file}`);
  const rows = readCsv(e.file);
  assert(rows.length === e.filas, `${path.basename(e.file)} tiene ${e.filas} filas de datos`);
}

const manifiesto = JSON.parse(fs.readFileSync(path.join(tmp, "manifiesto.json"), "utf8"));
assert(/no hay clientes/i.test(manifiesto.aviso), "el manifiesto avisa que no hay clientes de mostrador");

const headerEmpresas = fs.readFileSync(path.join(tmp, "empresas-sucursales.csv"), "utf8").split("\n")[0];
assert(headerEmpresas === crm.COLUMNAS_EMPRESAS.join(","), "encabezado de empresas estable para mapear en el CRM");

console.log(`ok: ${r.empresas.length} empresas, ${r.productos.length} productos, ${r.contactos.length} contactos sintéticos`);
