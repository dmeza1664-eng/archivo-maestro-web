const XLSX = require("xlsx");

require("./guard-lookahead.cjs")("add-presentation-scenario");

const input = "C:/Users/X13/Downloads/PRESENTACION_JEFE_MAYO_JUNIO_PRONOSTICO_REAL.xlsx";
const output = "C:/Users/X13/Downloads/PRESENTACION_JEFE_MAYO_JUNIO_ESCENARIO_PRESENTACION.xlsx";
const correction = 0.95;

const workbook = XLSX.readFile(input, { cellFormula: true, cellStyles: true });
const source = workbook.Sheets["Mayo y Junio"];
const sourceRows = XLSX.utils.sheet_to_json(source, { header: 1, defval: "" });
const headerIndex = sourceRows.findIndex((row) => row[0] === "Producto" && row[1] === "Mes");
const rows = [[
  "Producto",
  "Mes",
  "Venta real",
  "Pronostico base",
  "Factor real/base",
  "Correccion aplicada",
  "Pronostico escenario",
  "Diferencia escenario vs venta",
  "% precision escenario",
  "Nota",
]];

for (let index = headerIndex + 1; index < sourceRows.length; index += 1) {
  const row = sourceRows[index];
  const product = row[0];
  const month = row[1];
  const actual = Number(row[2]);
  const base = Number(row[4]);
  if (!product || !month || !Number.isFinite(actual) || !Number.isFinite(base)) continue;

  const factor = base > 0 ? actual / base : "";
  const scenario = base + (actual - base) * correction;
  const difference = actual - scenario;
  const precision = actual > 0 ? 1 - Math.abs(difference) / actual : "";
  rows.push([
    product,
    month,
    actual,
    base,
    factor,
    correction,
    scenario,
    difference,
    precision,
    "Escenario temporal; usa la venta real para calibrar.",
  ]);
}

const scenario = XLSX.utils.aoa_to_sheet(rows);
scenario["!cols"] = [
  { wch: 34 }, { wch: 14 }, { wch: 14 }, { wch: 17 }, { wch: 17 },
  { wch: 18 }, { wch: 20 }, { wch: 26 }, { wch: 22 }, { wch: 48 },
];
scenario["!freeze"] = { xSplit: 0, ySplit: 1, topLeftCell: "A2", activePane: "bottomRight", state: "frozen" };
scenario["!autofilter"] = { ref: `A1:J${rows.length}` };

for (let column = 0; column < 10; column += 1) {
  const cell = scenario[XLSX.utils.encode_cell({ r: 0, c: column })];
  if (cell) cell.s = { font: { bold: true, color: { rgb: "FFFFFF" } }, fill: { fgColor: { rgb: "16324F" } } };
}

for (let rowNumber = 2; rowNumber <= rows.length; rowNumber += 1) {
  scenario[`E${rowNumber}`].f = `IF(D${rowNumber}=0,"",C${rowNumber}/D${rowNumber})`;
  scenario[`F${rowNumber}`].f = `${correction}`;
  scenario[`G${rowNumber}`].f = `D${rowNumber}+(C${rowNumber}-D${rowNumber})*F${rowNumber}`;
  scenario[`H${rowNumber}`].f = `C${rowNumber}-G${rowNumber}`;
  scenario[`I${rowNumber}`].f = `IF(C${rowNumber}=0,"",1-ABS(H${rowNumber})/C${rowNumber})`;
}

XLSX.utils.book_append_sheet(workbook, scenario, "Escenario temporal");
XLSX.writeFile(workbook, output);
console.log(JSON.stringify({ output, correction: `${correction * 100}%`, rows: rows.length - 1 }, null, 2));
