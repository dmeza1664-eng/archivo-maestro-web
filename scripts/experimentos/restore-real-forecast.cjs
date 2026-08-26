const XLSX = require("xlsx");

const input = "C:/Users/X13/Downloads/PRESENTACION_JEFE_MAYO_JUNIO_AJUSTADO.xlsx";
const output = "C:/Users/X13/Downloads/PRESENTACION_JEFE_MAYO_JUNIO_PRONOSTICO_REAL.xlsx";

const workbook = XLSX.readFile(input, { cellFormula: true, cellStyles: true });
const sheet = workbook.Sheets["Mayo y Junio"];
const rows = XLSX.utils.sheet_to_json(sheet, { header: 1, defval: "" });
const headerIndex = rows.findIndex((row) => row[0] === "Producto" && row[1] === "Mes");
const headerRow = headerIndex + 1;

sheet[`E${headerRow}`] = { ...(sheet[`N${headerRow}`] || {}), t: "s", v: "Pronostico venta" };
sheet[`F${headerRow}`] = { ...(sheet[`N${headerRow}`] || {}), t: "s", v: "Diferencia pronostico vs venta" };

let restoredRows = 0;
for (let rowNumber = headerRow + 1; rowNumber <= rows.length; rowNumber += 1) {
  const row = rows[rowNumber - 1];
  const actual = Number(row[2]);
  const base = Number(row[18]);
  if (row[0] === "" || !Number.isFinite(actual) || !Number.isFinite(base)) continue;

  const difference = actual - base;
  const precision = actual > 0 ? 1 - Math.abs(difference) / actual : "";
  const status = actual === 0 && base === 0
    ? "Sin movimiento"
    : Math.abs(difference) <= 15
      ? "Dentro de +/-15 piezas"
      : "Revisar pronostico";
  const numberStyle = sheet[`E${rowNumber}`] || {};
  const percentStyle = sheet[`J${rowNumber}`] || {};
  const statusStyle = sheet[`K${rowNumber}`] || {};

  sheet[`E${rowNumber}`] = {
    ...numberStyle,
    t: "n",
    f: `S${rowNumber}`,
    v: base,
  };
  sheet[`F${rowNumber}`] = {
    ...numberStyle,
    t: "n",
    f: `C${rowNumber}-E${rowNumber}`,
    v: difference,
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
    f: `IF(AND(C${rowNumber}=0,E${rowNumber}=0),"Sin movimiento",IF(ABS(F${rowNumber})<=15,"Dentro de +/-15 piezas","Revisar pronostico"))`,
    v: status,
  };
  restoredRows += 1;
}

const notes = workbook.Sheets["Metodologia"];
if (notes) {
  const noteRows = XLSX.utils.sheet_to_json(notes, { header: 1, defval: "" });
  const existing = noteRows.findIndex((row) => row[0] === "Columna visible");
  const noteRow = existing >= 0 ? existing + 1 : noteRows.length + 2;
  notes[`A${noteRow}`] = { t: "s", v: "Pronostico real" };
  notes[`B${noteRow}`] = { t: "s", v: "El pronostico visible usa unicamente historico anterior al mes. La venta real del mismo mes solo se utiliza para validar el resultado." };
  notes[`C${noteRow}`] = { t: "s", v: "No se limita ni se corrige la diferencia usando la venta real del mismo mes." };
}

XLSX.writeFile(workbook, output);
XLSX.writeFile(workbook, input);
console.log(JSON.stringify({ output, restoredRows }, null, 2));
