const XLSX = require("xlsx");

const DOWNLOADS = "C:/Users/X13/Downloads";
const OUTPUT = `${DOWNLOADS}/PRESENTACION_JEFE_MAYO_JUNIO_PRONOSTICO_MEJORADO.xlsx`;

const files = {
  enero: `${DOWNLOADS}/VENTA ENERO 2026.xlsx`,
  febrero: `${DOWNLOADS}/VENTAS FEEEBRERO 2026.xlsx`,
  marzo: `${DOWNLOADS}/MARZO VENTAS.xlsx`,
  abril: `${DOWNLOADS}/venta abril 2026.xlsx`,
  mayoJunio: `${DOWNLOADS}/VENTAS DE MAYO Y JUNIO - ANGEL.xlsx`,
  producidoMayo: "C:/Users/X13/Documents/PRODUCIDO/PRODUCIDO MAYO.xlsx",
  producidoJunio: "C:/Users/X13/Documents/PRODUCIDO/PRODUCIDO JUNIO.xlsx",
  bajasJunio: "C:/Users/X13/Documents/DEVOLUCION Y BAJAS/JUNIO BAJAS.xlsx",
  bajasJulio: "C:/Users/X13/Documents/DEVOLUCION Y BAJAS/JULIO BAJAS.xlsx",
  stock: `${DOWNLOADS}/STOCK IDEAL SUCURSALES.xlsx`,
};

const MONTHS = ["Enero", "Febrero", "Marzo", "Abril", "Mayo", "Junio", "Julio"];
const WEEKDAYS = [1, 2, 3, 4, 5, 6, 0];

function norm(value) {
  return String(value ?? "")
    .trim()
    .toUpperCase()
    .normalize("NFD")
    .replace(/[\u0300-\u036f]/g, "");
}

function normalizeProduct(value) {
  const normalized = norm(value).replace(/[.,/\\_-]+/g, " ").replace(/\s+/g, " ").trim();
  const compact = normalized.replace(/[^A-Z0-9]/g, "");
  if (compact === "PINAGDE" || compact === "PINAGRANDE") return "PINA GDE";
  return normalized;
}

function isValidProduct(value) {
  const product = norm(value);
  if (!product || /^\d+$/.test(product)) return false;
  if (["TOTAL", "SUBTOTAL", "SUMA", "SUMAS", "GRANDE", "MEDIANO", "CHICO"].includes(product)) return false;
  if (product.startsWith("TOTAL") || product.includes("PRODUCTO") || product.includes("ESPECIALIDAD")) return false;
  if (product.includes("VENTA DE HOY") || product.includes("MODIFICACION DE PRECIO")) return false;
  if (/REBANADA|REBANADAS/.test(product)) return false;
  return true;
}

function number(value) {
  const parsed = Number(String(value ?? "").replace(/[$,\s]/g, ""));
  return Number.isFinite(parsed) ? parsed : 0;
}

function dateRecord(year, month, day, cantidad, producto, extra = {}) {
  return { year, month, day, cantidad, producto, ...extra };
}

function monthKey(year, month) {
  return `${year}-${String(month).padStart(2, "0")}`;
}

function dateFromRecord(record) {
  return new Date(record.year, record.month - 1, record.day);
}

function addToMap(map, product, quantity) {
  map.set(product, (map.get(product) || 0) + quantity);
}

function parseMonthlySummary(path, year, month) {
  const workbook = XLSX.readFile(path);
  const sheet = workbook.Sheets[workbook.SheetNames[0]];
  const rows = XLSX.utils.sheet_to_json(sheet, { header: 1, defval: "" });
  const header = rows[0].map(norm);
  const productCol = header.findIndex((cell) => cell.includes("PRODUCTO"));
  const quantityCol = header.findIndex((cell) => cell.includes("CANT"));
  const records = [];
  for (let row = 1; row < rows.length; row += 1) {
    const original = rows[row][productCol];
    if (!isValidProduct(original)) continue;
    const product = normalizeProduct(original);
    records.push(dateRecord(year, month, 1, number(rows[row][quantityCol]), product, {
      monthlyTotal: true,
      monthDays: new Date(year, month, 0).getDate(),
    }));
  }
  return records;
}

function parseMarch(path) {
  const workbook = XLSX.readFile(path);
  const records = [];
  for (const sheetName of workbook.SheetNames) {
    const day = Number(sheetName);
    if (!Number.isInteger(day) || day < 1 || day > 31) continue;
    const rows = XLSX.utils.sheet_to_json(workbook.Sheets[sheetName], { header: 1, defval: "" });
    for (let row = 1; row < rows.length; row += 1) {
      if (!isValidProduct(rows[row][0])) continue;
      records.push(dateRecord(2026, 3, day, number(rows[row][11]), normalizeProduct(rows[row][0])));
    }
  }
  return records;
}

function weekdayIndex(value) {
  const day = norm(value);
  if (day.startsWith("LUN")) return 1;
  if (day.startsWith("MAR")) return 2;
  if (day.startsWith("MIE")) return 3;
  if (day.startsWith("JUE")) return 4;
  if (day.startsWith("VIE")) return 5;
  if (day.startsWith("SAB")) return 6;
  if (day.startsWith("DOM")) return 0;
  return null;
}

function parseWide(path, sheetName, year, month) {
  const workbook = XLSX.readFile(path);
  const rows = XLSX.utils.sheet_to_json(workbook.Sheets[sheetName], { header: 1, defval: "" });
  let weekdayRow = -1;
  let dateRow = -1;
  let bestWeekdayScore = 0;
  let bestDateScore = 0;
  for (let row = 0; row < Math.min(rows.length, 5); row += 1) {
    const weekdayScore = rows[row].slice(1).filter((cell) => weekdayIndex(cell) !== null).length;
    const dateScore = rows[row].slice(1).filter((cell) => /^\d{1,2}$/.test(String(cell).trim())).length;
    if (weekdayScore > bestWeekdayScore) {
      bestWeekdayScore = weekdayScore;
      weekdayRow = row;
    }
    if (dateScore > bestDateScore) {
      bestDateScore = dateScore;
      dateRow = row;
    }
  }
  if (weekdayRow < 0 || dateRow < 0) return [];

  const records = [];
  for (let row = Math.max(weekdayRow, dateRow) + 1; row < rows.length; row += 1) {
    if (!isValidProduct(rows[row][0])) continue;
    const product = normalizeProduct(rows[row][0]);
    for (let col = 1; col < rows[row].length; col += 1) {
      const day = Number(rows[dateRow][col]);
      if (!Number.isInteger(day) || day < 1 || day > new Date(year, month, 0).getDate()) continue;
      if (weekdayIndex(rows[weekdayRow][col]) === null) continue;
      records.push(dateRecord(year, month, day, number(rows[row][col]), product));
    }
  }
  return records;
}

function parseProduction(path) {
  const workbook = XLSX.readFile(path);
  const rows = XLSX.utils.sheet_to_json(workbook.Sheets[workbook.SheetNames[0]], { header: 1, defval: "" });
  const result = new Map();
  for (let row = 1; row < rows.length; row += 1) {
    if (!isValidProduct(rows[row][1])) continue;
    addToMap(result, normalizeProduct(rows[row][1]), number(rows[row][0]));
  }
  return result;
}

function parseBajas(path, sheetName) {
  const workbook = XLSX.readFile(path);
  const rows = XLSX.utils.sheet_to_json(workbook.Sheets[sheetName], { header: 1, defval: "" });
  const result = new Map();
  for (let row = 2; row < rows.length; row += 1) {
    if (!isValidProduct(rows[row][0])) continue;
    addToMap(result, normalizeProduct(rows[row][0]), number(rows[row][1]));
  }
  return result;
}

function parseStock(path) {
  const workbook = XLSX.readFile(path);
  const sheet = workbook.Sheets["TOTAL A TENER SUC.(EXIST.+DIST)"];
  const rows = XLSX.utils.sheet_to_json(sheet, { header: 1, defval: "" });
  const header = rows[1].map(norm);
  const productCol = header.findIndex((cell) => cell.includes("PRODUCTO"));
  const stockCol = header.findIndex((cell) => cell === "STOCK");
  const result = new Map();
  for (let row = 2; row < rows.length; row += 1) {
    if (!isValidProduct(rows[row][productCol])) continue;
    result.set(normalizeProduct(rows[row][productCol]), number(rows[row][stockCol]));
  }
  return result;
}

function collectSales() {
  return [
    ...parseMonthlySummary(files.enero, 2026, 1),
    ...parseMonthlySummary(files.febrero, 2026, 2),
    ...parseMarch(files.marzo),
    ...parseWide(files.abril, "POR DÍA - SEMANA", 2026, 4),
    ...parseWide(files.mayoJunio, "MAYO 2026", 2026, 5),
    ...parseWide(files.mayoJunio, "JULIO 2026", 2026, 6),
  ];
}

function addAverage(bucketMap, weekday, value) {
  const bucket = bucketMap.get(weekday) || { total: 0, count: 0 };
  bucket.total += value;
  bucket.count += 1;
  bucketMap.set(weekday, bucket);
}

function buildAverages(records) {
  const buckets = new Map();
  for (const record of records) {
    if (record.monthlyTotal) {
      const dailyValue = record.cantidad / record.monthDays;
      for (let day = 1; day <= record.monthDays; day += 1) {
        addAverage(buckets, new Date(record.year, record.month - 1, day).getDay(), dailyValue);
      }
    } else {
      addAverage(buckets, dateFromRecord(record).getDay(), record.cantidad);
    }
  }
  return new Map([...buckets.entries()].map(([weekday, value]) => [weekday, value.count ? value.total / value.count : 0]));
}

function overallDailyAverage(records) {
  let total = 0;
  let days = 0;
  for (const record of records) {
    if (record.monthlyTotal) {
      total += record.cantidad;
      days += record.monthDays;
    } else {
      total += record.cantidad;
      days += 1;
    }
  }
  return days ? total / days : 0;
}

function clamp(value, min, max) {
  return Math.min(max, Math.max(min, value));
}

function improvedForecastModel(records) {
  const monthKeys = [...new Set(records.map((record) => monthKey(record.year, record.month)))].sort().slice(-3);
  const data = new Map();
  for (const record of records) {
    const key = monthKey(record.year, record.month);
    if (!monthKeys.includes(key)) continue;
    const monthData = data.get(key) || { total: 0, days: new Set(), weekdays: new Map() };
    if (record.monthlyTotal) {
      const dailyValue = record.cantidad / record.monthDays;
      monthData.total += record.cantidad;
      for (let day = 1; day <= record.monthDays; day += 1) {
        const date = new Date(record.year, record.month - 1, day);
        const weekday = date.getDay();
        const bucket = monthData.weekdays.get(weekday) || { total: 0, count: 0 };
        bucket.total += dailyValue;
        bucket.count += 1;
        monthData.weekdays.set(weekday, bucket);
        monthData.days.add(`${record.year}-${record.month}-${day}`);
      }
    } else {
      const date = new Date(record.year, record.month - 1, record.day);
      const weekday = date.getDay();
      const bucket = monthData.weekdays.get(weekday) || { total: 0, count: 0 };
      bucket.total += record.cantidad;
      bucket.count += 1;
      monthData.weekdays.set(weekday, bucket);
      monthData.total += record.cantidad;
      monthData.days.add(`${record.year}-${record.month}-${record.day}`);
    }
    data.set(key, monthData);
  }

  const weights = monthKeys.length === 3 ? [0.2, 0.3, 0.5] : monthKeys.length === 2 ? [0.4, 0.6] : [1];
  const dailyAverages = monthKeys
    .map((key) => {
      const monthData = data.get(key);
      return monthData && monthData.days.size ? monthData.total / monthData.days.size : 0;
    })
    .filter((value) => value > 0);
  const recentAverage = dailyAverages.at(-1) || 0;
  const previousValues = dailyAverages.slice(0, -1);
  const previousAverage = previousValues.length
    ? previousValues.reduce((sum, value) => sum + value, 0) / previousValues.length
    : 0;
  const trend = recentAverage > 0 && previousAverage > 0
    ? clamp(recentAverage / previousAverage, 0.85, 1.15)
    : 1;
  const averages = new Map();

  for (const weekday of [1, 2, 3, 4, 5, 6, 0]) {
    let total = 0;
    let weightTotal = 0;
    monthKeys.forEach((key, index) => {
      const bucket = data.get(key)?.weekdays.get(weekday);
      if (!bucket || !bucket.count) return;
      total += (bucket.total / bucket.count) * weights[index];
      weightTotal += weights[index];
    });
    if (weightTotal) averages.set(weekday, (total / weightTotal) * trend);
  }

  return { averages, trend, monthKeys };
}

function forecastForMonth(records, year, month, product) {
  const productRecords = records.filter((record) => record.producto === product);
  const model = improvedForecastModel(productRecords);
  const averages = model.averages;
  const fallback = overallDailyAverage(productRecords);
  let forecast = 0;
  const days = new Date(year, month, 0).getDate();
  for (let day = 1; day <= days; day += 1) {
    const weekday = new Date(year, month - 1, day).getDay();
    forecast += averages.has(weekday) ? averages.get(weekday) : fallback;
  }
  return forecast;
}

function actualForMonth(records, year, month, product) {
  return records
    .filter((record) => record.producto === product && record.year === year && record.month === month)
    .reduce((sum, record) => sum + record.cantidad, 0);
}

function monthlyTotals(records, product) {
  return MONTHS.slice(0, 6).map((_, index) => actualForMonth(records, 2026, index + 1, product));
}

function round(value) {
  return Number(value.toFixed(2));
}

function statusForDifference(value, actual, forecast) {
  if (actual === 0 && forecast === 0) return "Sin movimiento";
  return Math.abs(value) <= 15 ? "Dentro de +/-15 piezas" : "Fuera de +/-15 piezas";
}

function rowsToSheet(rows) {
  return XLSX.utils.json_to_sheet(rows);
}

function styleSheet(sheet, widths, freeze = "A2") {
  sheet["!cols"] = widths.map((wch) => ({ wch }));
  sheet["!freeze"] = { xSplit: 0, ySplit: 1, topLeftCell: freeze, activePane: "bottomRight", state: "frozen" };
  const range = XLSX.utils.decode_range(sheet["!ref"] || "A1");
  for (let col = range.s.c; col <= range.e.c; col += 1) {
    const cell = sheet[XLSX.utils.encode_cell({ r: 0, c: col })];
    if (cell) {
      cell.s = { font: { bold: true, color: { rgb: "FFFFFF" } }, fill: { fgColor: { rgb: "16324F" } } };
    }
  }
  sheet["!autofilter"] = { ref: sheet["!ref"] };
}

function setFormula(sheet, address, formula, value = undefined) {
  sheet[address] = { t: typeof value === "string" ? "s" : "n", f: formula };
  if (value !== undefined) sheet[address].v = value;
}

const sales = collectSales();
const products = [...new Set(sales.map((record) => record.producto))]
  .filter((product) => sales.some((record) => record.producto === product && record.cantidad > 0))
  .sort((a, b) => a.localeCompare(b, "es"));
const productionMay = parseProduction(files.producidoMayo);
const productionJune = parseProduction(files.producidoJunio);
const bajasJune = parseBajas(files.bajasJunio, "BAJAS ERICK");
const bajasJuly = parseBajas(files.bajasJulio, "ERIC BAJAS");
const stock = parseStock(files.stock);

function buildValidationRows(month, monthName) {
  return products.map((product) => {
    const history = sales.filter((record) => record.year < 2026 || record.month < month);
    const forecast = forecastForMonth(history, 2026, month, product);
    const actual = actualForMonth(sales, 2026, month, product);
    const difference = actual - forecast;
    return {
      Producto: product,
      [`Pronostico ${monthName}`]: round(forecast),
      [`Venta real ${monthName}`]: round(actual),
      "Diferencia piezas": round(difference),
      "Diferencia absoluta piezas": round(Math.abs(difference)),
      "Error %": actual ? round((Math.abs(difference) / actual) * 100) : "",
      Estado: statusForDifference(difference, actual, forecast),
    };
  });
}

const mayRows = buildValidationRows(5, "mayo");
const juneRows = buildValidationRows(6, "junio");

const comparisonRows = products.map((product, index) => ({
  Producto: product,
  "Pronostico mayo": mayRows[index]["Pronostico mayo"],
  "Venta real mayo": mayRows[index]["Venta real mayo"],
  "Diferencia mayo": mayRows[index]["Diferencia piezas"],
  "Diferencia absoluta mayo": mayRows[index]["Diferencia absoluta piezas"],
  "Estado mayo": mayRows[index].Estado,
  "Pronostico junio": juneRows[index]["Pronostico junio"],
  "Venta real junio": juneRows[index]["Venta real junio"],
  "Diferencia junio": juneRows[index]["Diferencia piezas"],
  "Diferencia absoluta junio": juneRows[index]["Diferencia absoluta piezas"],
  "Estado junio": juneRows[index].Estado,
}));

const julyRows = products.map((product) => {
  const forecast = forecastForMonth(sales, 2026, 7, product);
  const bajaRateBase = actualForMonth(sales, 2026, 5, product) || actualForMonth(sales, 2026, 6, product);
  const bajaRate = bajaRateBase ? bajasJune.get(product) / bajaRateBase : 0;
  const expectedBajas = forecast * (Number.isFinite(bajaRate) ? bajaRate : 0);
  return {
    Producto: product,
    "Pronostico venta julio": round(forecast),
    "Pronostico redondeado": Math.ceil(forecast),
    "Bajas junio": bajasJune.get(product) || 0,
    "Bajas julio registradas": bajasJuly.get(product) || 0,
    "Bajas esperadas julio": round(expectedBajas),
    "Produccion base sugerida": Math.ceil(forecast),
    "Produccion con bajas esperadas": Math.ceil(forecast + expectedBajas),
    "Stock ideal": stock.get(product) || 0,
  };
});

const historyRows = products.map((product) => {
  const totals = monthlyTotals(sales, product);
  return {
    Producto: product,
    Enero: totals[0],
    Febrero: totals[1],
    Marzo: totals[2],
    Abril: totals[3],
    Mayo: totals[4],
    Junio: totals[5],
    "Total enero-junio": totals.reduce((sum, value) => sum + value, 0),
  };
});

const operationsRows = products.map((product) => ({
  Producto: product,
  "Producido mayo": productionMay.get(product) || 0,
  "Producido junio": productionJune.get(product) || 0,
  "Bajas junio": bajasJune.get(product) || 0,
  "Bajas julio": bajasJuly.get(product) || 0,
  "Stock ideal": stock.get(product) || 0,
}));

function validationStats(rows, monthName) {
  const actualKey = `Venta real ${monthName}`;
  const forecastKey = `Pronostico ${monthName}`;
  const comparableRows = rows.filter((row) => row[actualKey] > 0 || row[forecastKey] > 0);
  const activeRows = rows.filter((row) => row[actualKey] > 0);
  const actual = rows.reduce((sum, row) => sum + row[actualKey], 0);
  const forecast = rows.reduce((sum, row) => sum + row[forecastKey], 0);
  const absoluteError = comparableRows.reduce((sum, row) => sum + row["Diferencia absoluta piezas"], 0);
  const activeAbsoluteError = activeRows.reduce((sum, row) => sum + row["Diferencia absoluta piezas"], 0);
  return {
    actual,
    forecast,
    activeRows,
    comparableRows,
    absoluteError,
    activeAbsoluteError,
    inside: activeRows.filter((row) => row["Diferencia absoluta piezas"] <= 15).length,
  };
}

const mayStats = validationStats(mayRows, "mayo");
const juneStats = validationStats(juneRows, "junio");
const julyForecast = julyRows.reduce((sum, row) => sum + row["Pronostico venta julio"], 0);

const summaryRows = [
  { Indicador: "Archivos de ventas utilizados", Valor: "Enero, febrero, marzo, abril, mayo y junio 2026" },
  { Indicador: "Metodologia", Valor: "Promedio ponderado de los 3 meses anteriores por dia de semana + tendencia limitada 85%-115%" },
  { Indicador: "Pronostico mayo (enero-abril)", Valor: round(mayStats.forecast) },
  { Indicador: "Venta real mayo", Valor: round(mayStats.actual) },
  { Indicador: "Diferencia total mayo", Valor: round(mayStats.actual - mayStats.forecast) },
  { Indicador: "Pronostico junio (enero-mayo)", Valor: round(juneStats.forecast) },
  { Indicador: "Venta real junio", Valor: round(juneStats.actual) },
  { Indicador: "Diferencia total junio", Valor: round(juneStats.actual - juneStats.forecast) },
  { Indicador: "Productos con venta real mayo", Valor: mayStats.activeRows.length },
  { Indicador: "Dentro de +/-15 piezas mayo", Valor: `${mayStats.inside} de ${mayStats.activeRows.length}` },
  { Indicador: "Productos con venta real junio", Valor: juneStats.activeRows.length },
  { Indicador: "Dentro de +/-15 piezas junio", Valor: `${juneStats.inside} de ${juneStats.activeRows.length}` },
  { Indicador: "Pronostico julio", Valor: round(julyForecast) },
  { Indicador: "Bajas junio registradas", Valor: [...bajasJune.values()].reduce((sum, value) => sum + value, 0) },
  { Indicador: "Bajas julio registradas", Valor: [...bajasJuly.values()].reduce((sum, value) => sum + value, 0) },
];

const notesRows = [
  { Tema: "Tolerancia", Detalle: "Se marca dentro de rango cuando la diferencia absoluta entre pronostico y venta real es de maximo 15 piezas." },
  { Tema: "Pronostico real", Detalle: "Usa los tres meses historicos anteriores disponibles, pondera mas el mes reciente y aplica una tendencia limitada entre 85% y 115%." },
  { Tema: "Enero y febrero", Detalle: "Los archivos solo tienen totales mensuales; se distribuyen proporcionalmente entre los dias del mes para calcular promedios por dia de semana." },
  { Tema: "Marzo", Detalle: "El archivo contiene ventas del dia 1 al 15; se usa como historico parcial." },
  { Tema: "Produccion", Detalle: "La produccion base sugerida es una referencia de demanda; no se fuerza para que coincida con la venta real." },
  { Tema: "Bajas", Detalle: "Las bajas provienen de las hojas BAJAS ERICK y ERIC BAJAS y pueden ser parciales respecto a toda la empresa." },
];

const workbook = XLSX.utils.book_new();
const summarySheet = rowsToSheet(summaryRows);
const maySheet = rowsToSheet(mayRows);
const juneSheet = rowsToSheet(juneRows);
const comparisonSheet = rowsToSheet(comparisonRows);
const julySheet = rowsToSheet(julyRows);
const historySheet = rowsToSheet(historyRows);
const operationsSheet = rowsToSheet(operationsRows);
const notesSheet = rowsToSheet(notesRows);

function addValidationFormulas(sheet, rows) {
  for (let row = 2; row <= rows.length + 1; row += 1) {
    const source = rows[row - 2];
    setFormula(sheet, `D${row}`, `C${row}-B${row}`, source["Diferencia piezas"]);
    setFormula(sheet, `E${row}`, `ABS(D${row})`, source["Diferencia absoluta piezas"]);
    setFormula(sheet, `F${row}`, `IF(C${row}>0,E${row}/C${row}*100,\"\")`, source["Error %"] === "" ? undefined : source["Error %"]);
    setFormula(sheet, `G${row}`, `IF(AND(B${row}=0,C${row}=0),\"Sin movimiento\",IF(E${row}<=15,\"Dentro de +/-15 piezas\",\"Fuera de +/-15 piezas\"))`, source.Estado);
  }
}

addValidationFormulas(maySheet, mayRows);
addValidationFormulas(juneSheet, juneRows);

for (let row = 2; row <= comparisonRows.length + 1; row += 1) {
  const source = comparisonRows[row - 2];
  setFormula(comparisonSheet, `D${row}`, `C${row}-B${row}`, source["Diferencia mayo"]);
  setFormula(comparisonSheet, `E${row}`, `ABS(D${row})`, source["Diferencia absoluta mayo"]);
  setFormula(comparisonSheet, `F${row}`, `IF(AND(B${row}=0,C${row}=0),\"Sin movimiento\",IF(E${row}<=15,\"Dentro de +/-15 piezas\",\"Fuera de +/-15 piezas\"))`, source["Estado mayo"]);
  setFormula(comparisonSheet, `I${row}`, `H${row}-G${row}`, source["Diferencia junio"]);
  setFormula(comparisonSheet, `J${row}`, `ABS(I${row})`, source["Diferencia absoluta junio"]);
  setFormula(comparisonSheet, `K${row}`, `IF(AND(G${row}=0,H${row}=0),\"Sin movimiento\",IF(J${row}<=15,\"Dentro de +/-15 piezas\",\"Fuera de +/-15 piezas\"))`, source["Estado junio"]);
}

for (let row = 2; row <= julyRows.length + 1; row += 1) {
  const source = julyRows[row - 2];
  setFormula(julySheet, `C${row}`, `ROUNDUP(B${row},0)`, source["Pronostico redondeado"]);
  setFormula(julySheet, `G${row}`, `ROUNDUP(B${row},0)`, source["Produccion base sugerida"]);
  setFormula(julySheet, `H${row}`, `ROUNDUP(B${row}+F${row},0)`, source["Produccion con bajas esperadas"]);
}

for (let row = 2; row <= historyRows.length + 1; row += 1) {
  const source = historyRows[row - 2];
  setFormula(historySheet, `H${row}`, `SUM(B${row}:G${row})`, source["Total enero-junio"]);
}

styleSheet(summarySheet, [46, 70], "A2");
styleSheet(maySheet, [38, 18, 18, 18, 25, 12, 25], "A2");
styleSheet(juneSheet, [38, 18, 18, 18, 25, 12, 25], "A2");
styleSheet(comparisonSheet, [38, 18, 18, 18, 25, 25, 18, 18, 18, 25, 25], "A2");
styleSheet(julySheet, [38, 22, 20, 14, 24, 22, 25, 30, 14], "A2");
styleSheet(historySheet, [38, 12, 12, 12, 12, 12, 12, 20], "A2");
styleSheet(operationsSheet, [38, 18, 18, 14, 14, 14], "A2");
styleSheet(notesSheet, [22, 110], "A2");

XLSX.utils.book_append_sheet(workbook, summarySheet, "Resumen ejecutivo");
XLSX.utils.book_append_sheet(workbook, maySheet, "Validacion mayo");
XLSX.utils.book_append_sheet(workbook, juneSheet, "Validacion junio");
XLSX.utils.book_append_sheet(workbook, comparisonSheet, "Real vs pronostico M-J");
XLSX.utils.book_append_sheet(workbook, julySheet, "Pronostico julio");
XLSX.utils.book_append_sheet(workbook, historySheet, "Historico mensual");
XLSX.utils.book_append_sheet(workbook, operationsSheet, "Producido y bajas");
XLSX.utils.book_append_sheet(workbook, notesSheet, "Notas metodo");

XLSX.writeFile(workbook, OUTPUT);
console.log(JSON.stringify({ output: OUTPUT, salesRecords: sales.length, products: products.length, mayForecast: round(mayStats.forecast), mayActual: round(mayStats.actual), juneForecast: round(juneStats.forecast), juneActual: round(juneStats.actual), julyForecast: round(julyForecast) }, null, 2));
