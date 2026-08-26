const path = require("path");
const XLSX = require("xlsx");

const DOWNLOADS = path.resolve(__dirname, "..");
const FILE = path.join(DOWNLOADS, "ESTATUS_PRODUCTOS_PARA_LLENAR.xlsx");
const SHEET = "Estatus productos";

const CHANGES = new Map([
  ["PAN DE MUERTO IND CHOCOLATE", { estatus: "ESTACIONAL", nota: "Temporada de octubre. Aparecio en la lista de octubre 2025 con venta 0" }],
  ["PAN DE MUERTO CHOCOLATE GDE", { estatus: "ESTACIONAL", nota: "Temporada de octubre. Aparecio en la lista de octubre 2025 con venta 0" }],
]);

function normalized(value) {
  return String(value || "")
    .toUpperCase().normalize("NFD").replace(/[\u0300-\u036f]/g, "")
    .replace(/[.,/\\_-]+/g, " ").replace(/\s+/g, " ").trim();
}

const workbook = XLSX.readFile(FILE, { cellStyles: true });
const sheet = workbook.Sheets[SHEET];
const range = XLSX.utils.decode_range(sheet["!ref"]);

const headers = {};
for (let col = range.s.c; col <= range.e.c; col += 1) {
  const cell = sheet[XLSX.utils.encode_cell({ r: 0, c: col })];
  if (cell) headers[String(cell.v).trim()] = col;
}

const applied = [];
for (let row = 1; row <= range.e.r; row += 1) {
  const productCell = sheet[XLSX.utils.encode_cell({ r: row, c: headers.Producto })];
  if (!productCell) continue;
  const change = CHANGES.get(normalized(productCell.v));
  if (!change) continue;

  const statusAddress = XLSX.utils.encode_cell({ r: row, c: headers.Estatus });
  const previous = sheet[statusAddress]?.v;
  sheet[statusAddress] = { ...(sheet[statusAddress] || {}), t: "s", v: change.estatus };

  const desdeAddress = XLSX.utils.encode_cell({ r: row, c: headers.Desde });
  if (sheet[desdeAddress]) sheet[desdeAddress] = { ...sheet[desdeAddress], t: "s", v: "" };

  if (headers.Nota !== undefined) {
    const notaAddress = XLSX.utils.encode_cell({ r: row, c: headers.Nota });
    sheet[notaAddress] = { ...(sheet[notaAddress] || {}), t: "s", v: change.nota };
  }

  applied.push({ producto: productCell.v, de: previous, a: change.estatus });
}

XLSX.writeFile(workbook, FILE);

const rows = XLSX.utils.sheet_to_json(workbook.Sheets[SHEET], { defval: "" });
const counts = rows.reduce((map, row) => {
  const key = String(row.Estatus || "").trim().toUpperCase() || "(vacio)";
  map[key] = (map[key] || 0) + 1;
  return map;
}, {});

console.log(JSON.stringify({ archivo: FILE, cambios: applied, estatusFinal: counts }, null, 2));
