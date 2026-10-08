/**
 * Prototipo seguro: arma CSV de ejemplo listos para importar en un CRM
 * (HubSpot, Zoho, Kommo, Bitrix24, Odoo, Pipedrive) a partir de catálogos
 * que YA usa este repo. No pide credenciales, no se conecta a SQL Server
 * ni a Firebase, no lee producción.
 *
 * Fuentes (todas locales):
 *   - datos/catalogo/mapeo-productos.csv  → productos inCatalog
 *   - sucursales de los fixtures de prueba del repo (nombres que ya aparecen
 *     en scripts/pronostico-automatico-test.cjs y scripts/demand-series-sucursal-test.cjs)
 *   - scripts/crm/fixtures/tiendas-galletas-sinteticas.json (tiendas B2B inventadas)
 *
 * Uso:
 *   node scripts/crm/generar-csv-importacion-crm.cjs
 *   node scripts/crm/generar-csv-importacion-crm.cjs --out /tmp/crm-ejemplo
 *
 * Salida (tres CSV + un manifiesto):
 *   empresas-sucursales.csv   → Companies / Accounts / Organizations
 *   productos.csv             → Products
 *   contactos-tiendas-galletas.csv → Contacts / Persons (encargados de tienda)
 */
const fs = require("fs");
const path = require("path");
const { readCsv, toCsv } = require("../lib/pepes-datos.cjs");

const ROOT = path.join(__dirname, "..", "..");

/** Sucursales que el repo ya usa en fixtures (no es el catálogo real de TI). */
const SUCURSALES_FIXTURE = [
  { nSucursalPK: "1", cNombre: "Suc. Centro", tipo: "sucursal", ciudad: "León" },
  { nSucursalPK: "2", cNombre: "Suc. Norte", tipo: "sucursal", ciudad: "León" },
  { nSucursalPK: "3", cNombre: "Suc. Amado Nervo", tipo: "sucursal", ciudad: "León" },
  { nSucursalPK: "9", cNombre: "Planta León", tipo: "planta", ciudad: "León" },
  { nSucursalPK: "ejemplo-leon", cNombre: "Suc. León", tipo: "sucursal", ciudad: "León" },
  { nSucursalPK: "ejemplo-cayetano", cNombre: "San Cayetano", tipo: "sucursal", ciudad: "León" },
  { nSucursalPK: "ejemplo-hidalgo", cNombre: "Villa Hidalgo", tipo: "sucursal", ciudad: "León" },
  { nSucursalPK: "ejemplo-plaza", cNombre: "Suc. Plaza", tipo: "sucursal", ciudad: "León" },
  { nSucursalPK: "ejemplo-vistas", cNombre: "Suc. Vistas", tipo: "sucursal", ciudad: "León" },
  { nSucursalPK: "ejemplo-allende", cNombre: "Suc. Allende", tipo: "sucursal", ciudad: "León" },
];

const COLUMNAS_EMPRESAS = [
  "external_id",
  "name",
  "record_type",
  "city",
  "country",
  "phone",
  "email",
  "industry",
  "description",
  "lifecyclestage",
];

const COLUMNAS_PRODUCTOS = [
  "external_id",
  "sku",
  "name",
  "description",
  "promotional",
  "in_wape_universe",
  "qty_total_2025",
];

const COLUMNAS_CONTACTOS = [
  "external_id",
  "firstname",
  "lastname",
  "phone",
  "email",
  "company",
  "company_external_id",
  "city",
  "address",
  "zone",
  "jobtitle",
  "lifecyclestage",
];

function parseArgs(argv) {
  const outIdx = argv.indexOf("--out");
  return {
    outDir: outIdx >= 0 ? path.resolve(argv[outIdx + 1] || "") : path.join(__dirname, "ejemplos"),
  };
}

function productosDesdeMapeo(mapeoCsv) {
  return readCsv(mapeoCsv)
    .filter((r) => r.inCatalog === "true" && r.producto)
    .map((r) => ({
      external_id: `pepes-prod-${r.nProductoPK}`,
      sku: r.cCodigo,
      name: r.producto,
      description: r.cDescripcion,
      promotional: r.promotional,
      in_wape_universe: r.inWapeUniverse,
      qty_total_2025: r.qtyTotal2025,
    }));
}

function empresasDesdeSucursales(sucursales) {
  return sucursales.map((s) => ({
    external_id: `pepes-suc-${s.nSucursalPK}`,
    name: s.cNombre,
    record_type: s.tipo,
    city: s.ciudad,
    country: "Mexico",
    phone: "",
    email: "",
    industry: "Food Production",
    description: s.tipo === "planta"
      ? "Planta / surtido interno. No es sucursal de venta al público."
      : "Sucursal de Pastelería Pepe (fixture del repo; no es un dump de pepes_devBI).",
    lifecyclestage: "customer",
  }));
}

function contactosDesdeTiendas(tiendas) {
  return tiendas.map((t) => {
    const partes = String(t.encargado || "Encargado Ejemplo").trim().split(/\s+/);
    const firstname = partes[0] || "Encargado";
    const lastname = partes.slice(1).join(" ") || "Ejemplo";
    return {
      external_id: `pepes-galleta-${t.storeId}`,
      firstname,
      lastname,
      phone: t.telefono || "",
      email: "",
      company: t.nombre,
      company_external_id: `pepes-galleta-tienda-${t.storeId}`,
      city: t.municipio || "",
      address: [t.direccion, t.colonia].filter(Boolean).join(", "),
      zone: t.zona || "",
      jobtitle: "Encargado de tienda (consignación galletas)",
      lifecyclestage: "lead",
    };
  });
}

function escribir(dir, nombre, rows, columns) {
  const file = path.join(dir, nombre);
  fs.writeFileSync(file, toCsv(rows, columns), "utf8");
  return { file, filas: rows.length };
}

function generar({ outDir, mapeoCsv, tiendasJson } = {}) {
  const mapeo = mapeoCsv || path.join(ROOT, "datos", "catalogo", "mapeo-productos.csv");
  const tiendasPath = tiendasJson || path.join(__dirname, "fixtures", "tiendas-galletas-sinteticas.json");
  if (!fs.existsSync(mapeo)) throw new Error(`No existe el mapeo de productos: ${mapeo}`);
  if (!fs.existsSync(tiendasPath)) throw new Error(`No existe el fixture de tiendas: ${tiendasPath}`);

  const fixture = JSON.parse(fs.readFileSync(tiendasPath, "utf8"));
  if (!/SINT[ÉE]TICO/i.test(fixture._SINTETICO || "")) {
    throw new Error("El fixture de tiendas debe estar marcado como sintético (_SINTETICO).");
  }

  const empresas = empresasDesdeSucursales(SUCURSALES_FIXTURE);
  const productos = productosDesdeMapeo(mapeo);
  const contactos = contactosDesdeTiendas(fixture.tiendas || []);
  if (!productos.length) throw new Error("El mapeo no trajo productos inCatalog.");
  if (!contactos.length) throw new Error("El fixture de tiendas está vacío.");

  fs.mkdirSync(outDir, { recursive: true });
  const escritos = [
    escribir(outDir, "empresas-sucursales.csv", empresas, COLUMNAS_EMPRESAS),
    escribir(outDir, "productos.csv", productos, COLUMNAS_PRODUCTOS),
    escribir(outDir, "contactos-tiendas-galletas.csv", contactos, COLUMNAS_CONTACTOS),
  ];

  const manifiesto = {
    generado: new Date().toISOString(),
    aviso: "Datos de ejemplo / catálogo. No hay clientes de mostrador identificados. No conectar a producción.",
    fuentes: {
      sucursales: "fixtures de prueba del repo (no es dbo.Sucursales de pepes_devBI)",
      productos: path.relative(ROOT, mapeo),
      tiendas_galletas: path.relative(ROOT, tiendasPath),
    },
    filas: Object.fromEntries(escritos.map((e) => [path.basename(e.file), e.filas])),
    mapeo_crm: {
      HubSpot: "empresas → Companies; contactos → Contacts; productos → Products. Import API /crm/imports.",
      "Zoho CRM": "empresas → Accounts; contactos → Contacts (Last Name obligatorio); productos → Products.",
      Kommo: "empresas → Companies; contactos → Contacts (phone como identificador).",
      Bitrix24: "empresas → Companies; contactos → Contacts; CSV nativo o crm.item.batchImport.",
      Odoo: "empresas/contactos → res.partner (is_company); productos → product.template. Usar External ID.",
      Pipedrive: "empresas → Organizations; contactos → Persons; productos → Products.",
    },
  };
  const manifiestoFile = path.join(outDir, "manifiesto.json");
  fs.writeFileSync(manifiestoFile, JSON.stringify(manifiesto, null, 2) + "\n", "utf8");
  return { outDir, escritos, manifiesto, empresas, productos, contactos };
}

function main() {
  const { outDir } = parseArgs(process.argv.slice(2));
  const r = generar({ outDir });
  for (const e of r.escritos) {
    console.log(`${e.filas} filas → ${e.file}`);
  }
  console.log(`manifiesto → ${path.join(r.outDir, "manifiesto.json")}`);
}

if (require.main === module) main();

module.exports = {
  SUCURSALES_FIXTURE,
  COLUMNAS_EMPRESAS,
  COLUMNAS_PRODUCTOS,
  COLUMNAS_CONTACTOS,
  productosDesdeMapeo,
  empresasDesdeSucursales,
  contactosDesdeTiendas,
  generar,
};
