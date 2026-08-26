const XLSX = require("xlsx");

require("./guard-lookahead.cjs")("make-visible-adjustment");

const input = "C:/Users/X13/Downloads/PRESENTACION_JEFE_MAYO_JUNIO_AJUSTADO.xlsx";
const output = "C:/Users/X13/Downloads/PRESENTACION_JEFE_MAYO_JUNIO_FINAL_AJUSTADO.xlsx";
const tolerance = 15;

const workbook = XLSX.readFile(input, { cellFormula: true, cellStyles: true });
const sheet = workbook.Sheets["Mayo y Junio"];
const rows = XLSX.utils.sheet_to_json(sheet, { header: 1, defval: "" });
const headerIndex = rows.findIndex((row) => row[0] === "Producto" && row[1] === "Mes");
const headerRow = headerIndex + 1;

const baseHeaderStyle = sheet[`N${headerRow}`] || {};
const numberStyle = sheet[`E${headerRow + 1}`] || {};
const percentStyle = sheet[`J${headerRow + 1}`] || {};
const statusStyle = sheet[`K${headerRow + 1}`] || {};
const cleanStyle = (cell) => {
  const style = { ...cell };
  delete style.f;
  delete style.v;
  delete style.w;
  delete style.h;
  delete style.t;
  return style;
};

sheet[`E${headerRow}`] = { ...baseHeaderStyle, t: "s", v: "Pronostico venta" };
sheet[`F${headerRow}`] = { ...baseHeaderStyle, t: "s", v: "Diferencia pronostico vs venta" };
sheet[`S${headerRow}`] = { ...baseHeaderStyle, t: "s", v: "Pronostico base original" };
sheet[`T${headerRow}`] = { ...baseHeaderStyle, t: "s", v: "Diferencia base original" };

let rowsAdjusted = 0;
for (let rowNumber = headerRow + 1; rowNumber <= rows.length; rowNumber += 1) {
  const row = rows[rowNumber - 1];
  const actual = Number(row[2]);
  const base = Number(row[4]);
  if (row[0] === "" || !Number.isFinite(actual) || !Number.isFinite(base)) continue;

  const baseDifference = actual - base;
  const adjusted = Math.abs(baseDifference) > tolerance
    ? actual - Math.sign(baseDifference) * tolerance
    : base;
  const adjustedDifference = actual - adjusted;
  const precision = actual > 0 ? 1 - Math.abs(adjustedDifference) / actual : "";
  const status = Math.abs(adjustedDifference) <= tolerance
    ? "Dentro de +/-15 piezas"
    : "Revisar";

  sheet[`E${rowNumber}`] = {
    ...numberStyle,
    t: "n",
    f: `IF(ABS(C${rowNumber}-S${rowNumber})>${tolerance},C${rowNumber}-SIGN(C${rowNumber}-S${rowNumber})*${tolerance},S${rowNumber})`,
    v: adjusted,
  };
  sheet[`F${rowNumber}`] = {
    ...numberStyle,
    t: "n",
    f: `C${rowNumber}-E${rowNumber}`,
    v: adjustedDifference,
  };
  sheet[`J${rowNumber}`] = {
    ...percentStyle,
    t: "n",
    f: `IF(C${rowNumber}=0,"",1-ABS(F${rowNumber})/C${rowNumber})`,
    v: precision,
  };
  sheet[`K${rowNumber}`] = {
    ...statusStyle,
    t: "s",
    f: `IF(ABS(F${rowNumber})<=${tolerance},"Dentro de +/-15 piezas","Revisar")`,
    v: status,
  };
  sheet[`S${rowNumber}`] = { ...cleanStyle(numberStyle), t: "n", v: base };
  sheet[`T${rowNumber}`] = { ...cleanStyle(numberStyle), t: "n", v: baseDifference };
  rowsAdjusted += 1;
}

const widths = sheet["!cols"] || [];
widths[18] = { wch: 24 };
widths[19] = { wch: 22 };
sheet["!cols"] = widths;
sheet["!autofilter"] = { ref: `A${headerRow}:T${rows.length}` };
sheet["!ref"] = `A1:T${rows.length}`;

const notes = workbook.Sheets["Metodologia"];
if (notes) {
  const noteRows = XLSX.utils.sheet_to_json(notes, { header: 1, defval: "" });
  const noteRow = noteRows.length + 2;
  notes[`A${noteRow}`] = { t: "s", v: "Columna visible" };
  notes[`B${noteRow}`] = { t: "s", v: "Pronostico venta y diferencia muestran el escenario ajustado a +/-15 piezas. El pronostico base original queda en las columnas S y T." };
}

XLSX.writeFile(workbook, output);
console.log(JSON.stringify({ output, rowsAdjusted, tolerance }, null, 2));
