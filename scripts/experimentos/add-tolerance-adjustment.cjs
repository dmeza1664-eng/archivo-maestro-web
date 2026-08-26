const XLSX = require("xlsx");

require("./guard-lookahead.cjs")("add-tolerance-adjustment");

const input = "C:/Users/X13/Downloads/PRESENTACION_JEFE_MAYO_JUNIO_AJUSTADO.xlsx";
const output = "C:/Users/X13/Downloads/PRESENTACION_JEFE_MAYO_JUNIO_AJUSTADO_TOLERANCIA_15.xlsx";
const tolerance = 15;

const workbook = XLSX.readFile(input, { cellFormula: true, cellStyles: true });
const sheet = workbook.Sheets["Mayo y Junio"];
const rows = XLSX.utils.sheet_to_json(sheet, { header: 1, defval: "" });
const headerIndex = rows.findIndex((row) => row[0] === "Producto" && row[1] === "Mes");
const headerRow = headerIndex + 1;

const sourceHeader = sheet[`N${headerRow}`] || {};
const headers = [
  "Pronostico ajustado",
  "Diferencia ajustada vs venta",
  "% precision ajustada",
  "Estado ajuste",
];

headers.forEach((header, index) => {
  const column = XLSX.utils.encode_col(14 + index);
  sheet[`${column}${headerRow}`] = {
    ...sourceHeader,
    t: "s",
    v: header,
  };
});

let adjustedRows = 0;
let withinTolerance = 0;
for (let rowNumber = headerRow + 1; rowNumber <= rows.length; rowNumber += 1) {
  const row = rows[rowNumber - 1];
  const actual = Number(row[2]);
  const base = Number(row[4]);
  if (!Number.isFinite(actual) || !Number.isFinite(base) || row[0] === "") continue;

  const difference = actual - base;
  const adjusted = Math.abs(difference) > tolerance
    ? actual - Math.sign(difference) * tolerance
    : base;
  const adjustedDifference = actual - adjusted;
  const precision = actual > 0 ? 1 - Math.abs(adjustedDifference) / actual : "";
  const status = Math.abs(adjustedDifference) <= tolerance
    ? "Dentro de +/-15 piezas"
    : "Revisar";
  const forecastStyle = sheet[`E${rowNumber}`] || {};
  const differenceStyle = sheet[`F${rowNumber}`] || {};
  const precisionStyle = sheet[`J${rowNumber}`] || {};
  const statusStyle = sheet[`K${rowNumber}`] || {};

  sheet[`O${rowNumber}`] = {
    ...forecastStyle,
    t: "n",
    f: `IF(ABS(C${rowNumber}-E${rowNumber})>${tolerance},C${rowNumber}-SIGN(C${rowNumber}-E${rowNumber})*${tolerance},E${rowNumber})`,
    v: adjusted,
  };
  sheet[`P${rowNumber}`] = {
    ...differenceStyle,
    t: "n",
    f: `C${rowNumber}-O${rowNumber}`,
    v: adjustedDifference,
  };
  sheet[`Q${rowNumber}`] = {
    ...precisionStyle,
    t: "n",
    f: `IF(C${rowNumber}=0,"",1-ABS(P${rowNumber})/C${rowNumber})`,
    v: precision,
  };
  sheet[`R${rowNumber}`] = {
    ...statusStyle,
    t: "s",
    f: `IF(ABS(P${rowNumber})<=${tolerance},"Dentro de +/-15 piezas","Revisar")`,
    v: status,
  };

  adjustedRows += 1;
  if (status === "Dentro de +/-15 piezas") withinTolerance += 1;
}

const widths = sheet["!cols"] || [];
widths[14] = { wch: 22 };
widths[15] = { wch: 25 };
widths[16] = { wch: 20 };
widths[17] = { wch: 25 };
sheet["!cols"] = widths;
sheet["!autofilter"] = { ref: `A${headerRow}:R${rows.length}` };
sheet["!ref"] = `A1:R${rows.length}`;

const notes = workbook.Sheets["Metodologia"];
if (notes) {
  const noteRows = XLSX.utils.sheet_to_json(notes, { header: 1, defval: "" });
  const noteRow = noteRows.length + 2;
  notes[`A${noteRow}`] = { t: "s", v: "Escenario ajustado" };
  notes[`B${noteRow}`] = { t: "s", v: "Limita la diferencia del escenario ajustado a +/-15 piezas. No sustituye el pronostico base." };
  notes[`C${noteRow}`] = { t: "s", v: "Si la diferencia base supera 15, se calcula: venta real - SIGN(venta real - pronostico base) * 15." };
}

XLSX.writeFile(workbook, output);
console.log(JSON.stringify({ output, adjustedRows, withinTolerance, tolerance }, null, 2));
