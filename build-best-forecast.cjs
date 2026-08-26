const XLSX = require("xlsx");
const fs = require("fs");

const ROOT = "C:/Users/X13/Downloads";
const OUTPUT = `${ROOT}/PRONOSTICO_MAYO_JUNIO_NUEVO_DESDE_CERO.xlsx`;
const PATHS = {
  december2024: `${ROOT}/Venta de diciembre 2024.xlsx`,
  sales2025: `${ROOT}/VENTA AÑO 2025 - ANGEL.xlsx`,
  january: `${ROOT}/VENTA ENERO 2026.xlsx`,
  february: `${ROOT}/VENTAS FEEEBRERO 2026.xlsx`,
  march: `${ROOT}/MARZO VENTAS.xlsx`,
  april: `${ROOT}/venta abril 2026.xlsx`,
  mayJune: `${ROOT}/VENTAS DE MAYO Y JUNIO - ANGEL.xlsx`,
  productionMay: "C:/Users/X13/Documents/PRODUCIDO/PRODUCIDO MAYO.xlsx",
  productionJune: "C:/Users/X13/Documents/PRODUCIDO/PRODUCIDO JUNIO.xlsx",
};

function norm(value) {
  return String(value ?? "").trim().toUpperCase().normalize("NFD").replace(/[\u0300-\u036f]/g, "");
}

function productName(value) {
  const valueNormalized = norm(value)
    .replace(/[.,/\\_-]+/g, " ")
    .replace(/\bCHESSECAKE\b/g, "CHEESECAKE")
    .replace(/\bMED\b/g, "MEDIANO")
    .replace(/\s+/g, " ")
    .trim();
  const compact = valueNormalized.replace(/[^A-Z0-9]/g, "");
  if (compact === "PINAGDE" || compact === "PINAGRANDE") return "PINA GDE";
  return valueNormalized;
}

function validProduct(value) {
  const product = norm(value);
  if (!product || /^\d+$/.test(product)) return false;
  if (["TOTAL", "SUBTOTAL", "SUMA", "SUMAS", "GRANDE", "MEDIANO", "CHICO"].includes(product)) return false;
  if (product.startsWith("TOTAL") || product.includes("PRODUCTO") || product.includes("ESPECIALIDAD")) return false;
  if (product.includes("VENTA DE HOY") || product.includes("MODIFICACION DE PRECIO")) return false;
  if (/\b(REBANADA|REBANADAS|REB|RBN)\b/.test(product)) return false;
  return true;
}

function numeric(value) {
  const result = Number(String(value ?? "").replace(/[$,\s]/g, ""));
  return Number.isFinite(result) ? result : 0;
}

function monthId(year, month) {
  return `${year}-${String(month).padStart(2, "0")}`;
}

function previousMonthId(value) {
  const [year, month] = value.split("-").map(Number);
  const previous = new Date(year, month - 2, 1);
  return monthId(previous.getFullYear(), previous.getMonth() + 1);
}

function previousYearMonthId(value) {
  const [year, month] = value.split("-").map(Number);
  return monthId(year - 1, month);
}

function round(value, digits = 2) {
  return Number(Number(value || 0).toFixed(digits));
}

function clamp(value, minimum, maximum) {
  return Math.min(maximum, Math.max(minimum, value));
}

function parseMonthlySummary(path, year, month) {
  const workbook = XLSX.readFile(path);
  const rows = XLSX.utils.sheet_to_json(workbook.Sheets[workbook.SheetNames[0]], { header: 1, defval: "" });
  const header = rows[0].map(norm);
  const productColumn = header.findIndex((cell) => cell.includes("PRODUCTO"));
  const quantityColumn = header.findIndex((cell) => cell.includes("CANT"));
  const records = [];
  for (let row = 1; row < rows.length; row += 1) {
    if (!validProduct(rows[row][productColumn])) continue;
    records.push({
      product: productName(rows[row][productColumn]),
      year,
      month,
      day: null,
      quantity: numeric(rows[row][quantityColumn]),
      monthlyTotal: true,
      observedDays: new Date(year, month, 0).getDate(),
    });
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
      if (!validProduct(rows[row][0])) continue;
      records.push({ product: productName(rows[row][0]), year: 2026, month: 3, day, quantity: numeric(rows[row][11]) });
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
  let weekdayScore = 0;
  let dateScore = 0;
  for (let row = 0; row < Math.min(rows.length, 5); row += 1) {
    const weekdays = rows[row].slice(1).filter((cell) => weekdayIndex(cell) !== null).length;
    const dates = rows[row].slice(1).filter((cell) => /^\d{1,2}$/.test(String(cell).trim())).length;
    if (weekdays > weekdayScore) { weekdayScore = weekdays; weekdayRow = row; }
    if (dates > dateScore) { dateScore = dates; dateRow = row; }
  }
  const records = [];
  const monthDays = new Date(year, month, 0).getDate();
  for (let row = Math.max(weekdayRow, dateRow) + 1; row < rows.length; row += 1) {
    if (!validProduct(rows[row][0])) continue;
    for (let column = 1; column < rows[row].length; column += 1) {
      const day = Number(rows[dateRow][column]);
      if (!Number.isInteger(day) || day < 1 || day > monthDays) continue;
      if (weekdayIndex(rows[weekdayRow][column]) === null) continue;
      records.push({
        product: productName(rows[row][0]),
        year,
        month,
        day,
        quantity: numeric(rows[row][column]),
      });
    }
  }
  return records;
}

function parseProduction(path) {
  const workbook = XLSX.readFile(path);
  const rows = XLSX.utils.sheet_to_json(workbook.Sheets[workbook.SheetNames[0]], { header: 1, defval: "" });
  const totals = new Map();
  for (let row = 1; row < rows.length; row += 1) {
    if (!validProduct(rows[row][1])) continue;
    const product = productName(rows[row][1]);
    totals.set(product, (totals.get(product) || 0) + numeric(rows[row][0]));
  }
  return totals;
}

function aggregate(records) {
  const products = new Map();
  for (const record of records) {
    const product = products.get(record.product) || new Map();
    const key = monthId(record.year, record.month);
    const month = product.get(key) || { total: 0, observedDays: new Set(), daily: new Map() };
    if (record.monthlyTotal) {
      month.total += record.quantity;
      month.monthlyTotal = true;
      month.observedDayCount = record.observedDays;
    } else {
      const dateKey = `${key}-${String(record.day).padStart(2, "0")}`;
      month.daily.set(record.day, (month.daily.get(record.day) || 0) + record.quantity);
      month.observedDays.add(dateKey);
      month.total += record.quantity;
    }
    product.set(key, month);
    products.set(record.product, product);
  }
  return products;
}

function dailyRate(month) {
  const days = month?.monthlyTotal ? month.observedDayCount : month?.observedDays.size;
  return days ? month.total / days : 0;
}

function targetCalendar(year, month) {
  const result = [];
  const days = new Date(year, month, 0).getDate();
  for (let day = 1; day <= days; day += 1) result.push(new Date(year, month - 1, day));
  return result;
}

function weekdayForecast(months, selectedKeys, weights, year, month) {
  const averages = new Map();
  for (const weekday of [0, 1, 2, 3, 4, 5, 6]) {
    let weighted = 0;
    let usedWeight = 0;
    selectedKeys.forEach((key, index) => {
      const source = months.get(key);
      if (!source || !source.daily.size) return;
      let total = 0;
      let count = 0;
      for (const [day, value] of source.daily.entries()) {
        const [yearValue, monthValue] = key.split("-").map(Number);
        if (new Date(yearValue, monthValue - 1, day).getDay() !== weekday) continue;
        total += value;
        count += 1;
      }
      if (!count) return;
      weighted += (total / count) * weights[index];
      usedWeight += weights[index];
    });
    if (usedWeight) averages.set(weekday, weighted / usedWeight);
  }
  if (!averages.size) return null;
  const fallback = selectedKeys.reduce((sum, key, index) => sum + dailyRate(months.get(key)) * weights[index], 0) /
    selectedKeys.reduce((sum, key, index) => sum + (months.has(key) ? weights[index] : 0), 0);
  return targetCalendar(year, month).reduce((sum, date) => sum + (averages.get(date.getDay()) ?? fallback), 0);
}

function candidates(months, beforeKey, year, month) {
  const keys = [...months.keys()].filter((key) => key < beforeKey).sort();
  const recent = keys.slice(-3);
  const latest = recent.at(-1);
  if (!latest) return new Map([["Sin historico", 0]]);
  const targetDays = new Date(year, month, 0).getDate();
  const rates = recent.map((key) => dailyRate(months.get(key)));
  const totals = recent.map((key) => months.get(key).total);
  const result = new Map();
  result.set("Ultimo mes por dia", rates.at(-1) * targetDays);
  result.set("Ultimo total mensual", totals.at(-1));
  if (recent.length >= 2) {
    result.set("Promedio 2 meses por dia", (rates.at(-1) * 0.7 + rates.at(-2) * 0.3) * targetDays);
    result.set("Promedio total 2 meses", totals.at(-1) * 0.7 + totals.at(-2) * 0.3);
  }
  if (recent.length >= 3) {
    result.set("Promedio 3 meses por dia", (rates[0] * 0.2 + rates[1] * 0.3 + rates[2] * 0.5) * targetDays);
    const trendRate = clamp(rates[2] + (rates[2] - rates[1]) * 0.5, rates[2] * 0.8, rates[2] * 1.2);
    result.set("Tendencia reciente", trendRate * targetDays);
  }
  const dailyKeys = recent.filter((key) => months.get(key).daily.size);
  if (dailyKeys.length) {
    const latestWeekday = weekdayForecast(months, [dailyKeys.at(-1)], [1], year, month);
    if (latestWeekday !== null) result.set("Dia semana ultimo mes", latestWeekday);
  }
  if (dailyKeys.length >= 2) {
    const keysForWeekday = dailyKeys.slice(-2);
    const weightedWeekday = weekdayForecast(months, keysForWeekday, [0.35, 0.65], year, month);
    if (weightedWeekday !== null) result.set("Dia semana ponderado", weightedWeekday);
  }
  const priorYearKey = previousYearMonthId(beforeKey);
  if (process.env.USE_SEASONAL !== "0" && months.has(priorYearKey)) {
    const previousKey = previousMonthId(beforeKey);
    const previousYearKey = previousYearMonthId(previousKey);
    const currentPrevious = months.get(previousKey)?.total || 0;
    const priorPrevious = months.get(previousYearKey)?.total || 0;
    const growth = currentPrevious > 0 && priorPrevious > 0
      ? clamp(currentPrevious / priorPrevious, 0.8, 1.2)
      : 1;
    const seasonalTotal = months.get(priorYearKey).total * growth;
    result.set("Mismo mes año anterior", seasonalTotal);
    const seasonalWeekday = weekdayForecast(months, [priorYearKey], [1], year, month);
    if (seasonalWeekday !== null) result.set("Estacional por dia semana", seasonalWeekday * growth);
  }
  return result;
}

function selectAndForecast(months, backtestYear, backtestMonth, targetYear, targetMonth) {
  const backtestKey = monthId(backtestYear, backtestMonth);
  const targetKey = monthId(targetYear, targetMonth);
  const actualBacktest = months.get(backtestKey)?.total || 0;
  const tested = candidates(months, backtestKey, backtestYear, backtestMonth);
  let selectedMethod = "Ultimo mes por dia";
  let selectedError = Infinity;
  for (const [method, forecast] of tested.entries()) {
    const error = Math.abs(actualBacktest - forecast);
    if (error < selectedError) {
      selectedMethod = method;
      selectedError = error;
    }
  }
  const targetCandidates = candidates(months, targetKey, targetYear, targetMonth);
  const baseForecast = targetCandidates.get(selectedMethod) ?? targetCandidates.values().next().value ?? 0;
  const backtestForecast = tested.get(selectedMethod) || 0;
  const calibration = actualBacktest > 0 && backtestForecast > 0
    ? clamp(actualBacktest / backtestForecast, 0.85, 1.15)
    : 1;
  return {
    method: selectedMethod,
    backtestActual: actualBacktest,
    backtestForecast,
    backtestError: selectedError === Infinity ? 0 : selectedError,
    calibration,
    baseForecast,
    forecast: Math.max(0, baseForecast * calibration),
  };
}

function accuracy(forecast, actual) {
  return actual > 0 ? Math.max(0, 1 - Math.abs(actual - forecast) / actual) : null;
}

function priorityCategory(product) {
  const normalized = productName(product);
  const isPromotion = /\b(PROMO|PROMOCION|REGALO|DIA DE|DIA DEL|NINO|NIÑO)\b/.test(normalized);
  if (isPromotion) return "Promociones separadas";
  if (/\bGELATINA\b/.test(normalized)) return "Gelatinas";
  if (/\bGALLETA|GALLETAS\b/.test(normalized)) return "Galletas";
  if (/\b(PAN|BOLLO|BOLLOS)\b/.test(normalized)) return "Pan";
  const isAccessory = /\b(BENGALA|VELA|LETRERO|PLATO|PLATOS|KIT|BOLSA|ARREGLO|PALETA|TOPPER|SERVILLETA|TENEDOR|CUCHARA)\b/.test(normalized);
  if (!isAccessory && /\b(GDE|GRANDE)\b/.test(normalized)) return "Pasteles grandes";
  return "";
}

function cakeSizeCategory(product) {
  const normalized = productName(product);
  const excluded = /\b(GELATINA|GALLETA|PAN|BOLLO|BOLLOS|BENGALA|VELA|LETRERO|PLATO|PLATOS|KIT|BOLSA|ARREGLO|PALETA|TOPPER|SERVILLETA|TENEDOR|CUCHARA|PROMO|PROMOCION|REGALO)\b/.test(normalized);
  if (excluded) return "";
  if (/\bMINI MEDIANO\b/.test(normalized)) return "Mini medianos";
  if (/\bMEDIANO\b/.test(normalized)) return "Pasteles medianos";
  if (/\b(CH|CHICO)\b/.test(normalized)) return "Pasteles chicos";
  return "";
}

const sales = [
  ...parseWide(PATHS.december2024, "POR DÍA - SEMANA", 2024, 12),
  ...[
    ["MARZO 2025", 3],
    ["ABRIL 2025", 4],
    ["MAYO 2025", 5],
    ["JUNIO 2025", 6],
    ["JULIO 2025", 7],
    ["AGOSTO 2025", 8],
    ["SEPTIEMBRE 2025", 9],
    ["OCTUBRE 2025", 10],
  ].flatMap(([sheet, month]) => parseWide(PATHS.sales2025, sheet, 2025, month)),
  ...parseMonthlySummary(PATHS.january, 2026, 1),
  ...parseMonthlySummary(PATHS.february, 2026, 2),
  ...parseMarch(PATHS.march),
  ...parseWide(PATHS.april, "POR DÍA - SEMANA", 2026, 4),
  ...parseWide(PATHS.mayJune, "MAYO 2026", 2026, 5),
  ...parseWide(PATHS.mayJune, "JULIO 2026", 2026, 6),
];
const salesByProduct = aggregate(sales);
const productionMay = parseProduction(PATHS.productionMay);
const productionJune = parseProduction(PATHS.productionJune);
const products = [...new Set([
  ...salesByProduct.keys(),
  ...productionMay.keys(),
  ...productionJune.keys(),
])].filter((product) => {
  const months = salesByProduct.get(product);
  return (months?.get("2026-05")?.total || 0) > 0 || (months?.get("2026-06")?.total || 0) > 0;
}).sort((a, b) => a.localeCompare(b, "es"));

const detailRows = [];
const methodRows = [];
for (const product of products) {
  const months = salesByProduct.get(product) || new Map();
  const may = selectAndForecast(months, 2026, 4, 2026, 5);
  const june = selectAndForecast(months, 2026, 5, 2026, 6);
  const actualMay = months.get("2026-05")?.total || 0;
  const actualJune = months.get("2026-06")?.total || 0;
  const monthRows = [
    ["Mayo 2026", actualMay, productionMay.get(product) || 0, may],
    ["Junio 2026", actualJune, productionJune.get(product) || 0, june],
  ];
  for (const [label, actual, produced, model] of monthRows) {
    const difference = actual - model.forecast;
    detailRows.push({
      Producto: product,
      Mes: label,
      "Venta real": actual,
      "Produccion real": produced,
      "Metodo elegido": model.method,
      "Pronostico base": round(model.baseForecast),
      "Factor calibracion": round(model.calibration, 4),
      "Pronostico final": round(model.forecast),
      "Diferencia piezas": round(difference),
      "Diferencia absoluta": round(Math.abs(difference)),
      "Precision %": accuracy(model.forecast, actual) === null ? "" : round(accuracy(model.forecast, actual) * 100, 1),
      Estado: Math.abs(difference) <= 15 ? "Dentro de +/-15 piezas" : "Revisar",
    });
    methodRows.push({
      Producto: product,
      Mes: label,
      "Mes usado para seleccionar": label === "Mayo 2026" ? "Abril 2026" : "Mayo 2026",
      "Metodo elegido": model.method,
      "Real validacion previa": round(model.backtestActual),
      "Pronostico validacion previa": round(model.backtestForecast),
      "Error validacion previa": round(model.backtestError),
      "Factor calibracion limitado": round(model.calibration, 4),
    });
  }
}

function monthSummary(label) {
  const rows = detailRows.filter((row) => row.Mes === label);
  const active = rows.filter((row) => row["Venta real"] > 0);
  const forecast = rows.reduce((sum, row) => sum + row["Pronostico final"], 0);
  const actual = rows.reduce((sum, row) => sum + row["Venta real"], 0);
  const absoluteErrors = active.map((row) => row["Diferencia absoluta"]).sort((a, b) => a - b);
  const totalAbsoluteError = absoluteErrors.reduce((sum, value) => sum + value, 0);
  const middle = Math.floor(absoluteErrors.length / 2);
  const medianAbsoluteError = absoluteErrors.length % 2
    ? absoluteErrors[middle]
    : ((absoluteErrors[middle - 1] || 0) + (absoluteErrors[middle] || 0)) / 2;
  return {
    forecast,
    actual,
    difference: actual - forecast,
    precision: accuracy(forecast, actual),
    wape: actual > 0 ? totalAbsoluteError / actual : null,
    mae: active.length ? totalAbsoluteError / active.length : null,
    medianAbsoluteError,
    inside: active.filter((row) => row["Diferencia absoluta"] <= 15).length,
    active: active.length,
  };
}

const maySummary = monthSummary("Mayo 2026");
const juneSummary = monthSummary("Junio 2026");
const summaryRows = [
  { Indicador: "Regla principal", Valor: "Nunca usa la venta real del mes objetivo para calcular ese mismo pronostico" },
  { Indicador: "Seleccion mayo", Valor: "Se elige por producto el metodo con menor error al pronosticar abril" },
  { Indicador: "Seleccion junio", Valor: "Se elige por producto el metodo con menor error al pronosticar mayo" },
  { Indicador: "Calibracion", Valor: "Factor del mes anterior limitado entre 85% y 115%" },
  { Indicador: "Pronostico mayo", Valor: round(maySummary.forecast) },
  { Indicador: "Venta real mayo", Valor: round(maySummary.actual) },
  { Indicador: "Diferencia total mayo", Valor: round(maySummary.difference) },
  { Indicador: "Precision global mayo", Valor: maySummary.precision === null ? "" : round(maySummary.precision * 100, 1) },
  { Indicador: "Dentro de +/-15 mayo", Valor: `${maySummary.inside} de ${maySummary.active}` },
  { Indicador: "Pronostico junio", Valor: round(juneSummary.forecast) },
  { Indicador: "Venta real junio", Valor: round(juneSummary.actual) },
  { Indicador: "Diferencia total junio", Valor: round(juneSummary.difference) },
  { Indicador: "Precision global junio", Valor: juneSummary.precision === null ? "" : round(juneSummary.precision * 100, 1) },
  { Indicador: "Dentro de +/-15 junio", Valor: `${juneSummary.inside} de ${juneSummary.active}` },
];

const workbook = XLSX.utils.book_new();
const summarySheet = XLSX.utils.json_to_sheet(summaryRows);
const detailSheet = XLSX.utils.json_to_sheet(detailRows);
const methodsSheet = XLSX.utils.json_to_sheet(methodRows);
summarySheet["!cols"] = [{ wch: 34 }, { wch: 88 }];
detailSheet["!cols"] = [34, 14, 14, 17, 27, 18, 19, 19, 19, 21, 14, 23].map((wch) => ({ wch }));
methodsSheet["!cols"] = [34, 14, 27, 27, 22, 28, 23, 27].map((wch) => ({ wch }));
detailSheet["!autofilter"] = { ref: `A1:L${detailRows.length + 1}` };
methodsSheet["!autofilter"] = { ref: `A1:H${methodRows.length + 1}` };
for (let row = 2; row <= detailRows.length + 1; row += 1) {
  const source = detailRows[row - 2];
  detailSheet[`I${row}`] = { t: "n", f: `C${row}-H${row}`, v: source["Diferencia piezas"] };
  detailSheet[`J${row}`] = { t: "n", f: `ABS(I${row})`, v: source["Diferencia absoluta"] };
  if (source["Precision %"] === "") {
    detailSheet[`K${row}`] = { t: "s", f: `IF(C${row}=0,"",MAX(0,1-ABS(I${row})/C${row}))`, v: "" };
  } else {
    detailSheet[`K${row}`] = { t: "n", f: `IF(C${row}=0,"",MAX(0,1-ABS(I${row})/C${row})*100)`, v: source["Precision %"] };
  }
  detailSheet[`L${row}`] = {
    t: "s",
    f: `IF(ABS(I${row})<=15,"Dentro de +/-15 piezas","Revisar")`,
    v: source.Estado,
  };
}
XLSX.utils.book_append_sheet(workbook, summarySheet, "Resumen ejecutivo");
XLSX.utils.book_append_sheet(workbook, detailSheet, "Real vs pronostico");
XLSX.utils.book_append_sheet(workbook, methodsSheet, "Seleccion de metodo");
if (process.env.ANALYZE_ONLY !== "1") XLSX.writeFile(workbook, OUTPUT);

const methodCounts = methodRows.reduce((counts, row) => {
  const key = `${row.Mes}: ${row["Metodo elegido"]}`;
  counts[key] = (counts[key] || 0) + 1;
  return counts;
}, {});
const topErrors = (month) => detailRows
  .filter((row) => row.Mes === month)
  .sort((a, b) => b["Diferencia absoluta"] - a["Diferencia absoluta"])
  .slice(0, 10)
  .map((row) => ({
    producto: row.Producto,
    real: row["Venta real"],
    pronostico: row["Pronostico final"],
    diferencia: row["Diferencia piezas"],
    metodo: row["Metodo elegido"],
  }));
const summarizeCategory = (category, month) => {
  const rows = detailRows.filter((row) => row.Mes === month && priorityCategory(row.Producto) === category && row["Venta real"] > 0);
  const actual = rows.reduce((sum, row) => sum + row["Venta real"], 0);
  const forecast = rows.reduce((sum, row) => sum + row["Pronostico final"], 0);
  const absoluteError = rows.reduce((sum, row) => sum + row["Diferencia absoluta"], 0);
  return {
    products: rows.length,
    actual,
    forecast,
    difference: actual - forecast,
    wape: actual > 0 ? absoluteError / actual : null,
    mae: rows.length ? absoluteError / rows.length : null,
    inside15: rows.filter((row) => row["Diferencia absoluta"] <= 15).length,
  };
};
const categories = ["Pasteles grandes", "Gelatinas", "Galletas", "Pan", "Promociones separadas"];
const priorityResults = Object.fromEntries(categories.map((category) => [category, {
  may: summarizeCategory(category, "Mayo 2026"),
  june: summarizeCategory(category, "Junio 2026"),
}]));
const candidateByNames = (available, names) => {
  for (const name of names) {
    if (available.has(name)) return available.get(name);
  }
  return null;
};
const categoryBaseForecast = (months, targetKey, category) => {
  const [year, month] = targetKey.split("-").map(Number);
  const available = candidates(months, targetKey, year, month);
  let parts = [];
  if (category === "Pasteles grandes") {
    parts = [
      [candidateByNames(available, ["Estacional por dia semana", "Mismo mes año anterior"]), 0.45],
      [candidateByNames(available, ["Dia semana ultimo mes", "Ultimo mes por dia"]), 0.35],
      [candidateByNames(available, ["Tendencia reciente", "Promedio 2 meses por dia"]), 0.2],
    ];
  } else if (category === "Gelatinas") {
    parts = [
      [candidateByNames(available, ["Dia semana ultimo mes", "Ultimo mes por dia"]), 0.5],
      [candidateByNames(available, ["Dia semana ponderado", "Promedio 2 meses por dia"]), 0.3],
      [candidateByNames(available, ["Estacional por dia semana", "Mismo mes año anterior"]), 0.2],
    ];
  } else if (category === "Galletas" || category === "Pan") {
    parts = [
      [candidateByNames(available, ["Dia semana ultimo mes", "Ultimo mes por dia"]), 0.7],
      [candidateByNames(available, ["Dia semana ponderado", "Promedio 2 meses por dia"]), 0.3],
    ];
  }
  const usable = parts.filter(([value]) => Number.isFinite(value));
  const weight = usable.reduce((sum, part) => sum + part[1], 0);
  return weight ? usable.reduce((sum, part) => sum + part[0] * part[1], 0) / weight : 0;
};
const categoryPrediction = (months, targetKey, category) => {
  const base = categoryBaseForecast(months, targetKey, category);
  const priorKey = previousMonthId(targetKey);
  const priorBase = categoryBaseForecast(months, priorKey, category);
  const priorActual = months.get(priorKey)?.total || 0;
  const calibration = priorBase > 0 && priorActual > 0 ? clamp(priorActual / priorBase, 0.85, 1.15) : 1;
  return { base, calibration, forecast: base * calibration };
};
const categoryModelRows = [];
for (const product of products) {
  const category = priorityCategory(product);
  if (!category || category === "Promociones separadas") continue;
  const months = salesByProduct.get(product) || new Map();
  for (const [targetKey, label] of [["2026-05", "Mayo 2026"], ["2026-06", "Junio 2026"]]) {
    const prediction = categoryPrediction(months, targetKey, category);
    const actual = months.get(targetKey)?.total || 0;
    categoryModelRows.push({
      product,
      category,
      month: label,
      actual,
      forecast: prediction.forecast,
      difference: actual - prediction.forecast,
      absoluteError: Math.abs(actual - prediction.forecast),
      calibration: prediction.calibration,
    });
  }
}
const categoryModelSummary = Object.fromEntries(
  ["Pasteles grandes", "Gelatinas", "Galletas", "Pan"].map((category) => [category, Object.fromEntries(
    ["Mayo 2026", "Junio 2026"].map((month) => {
      const rows = categoryModelRows.filter((row) => row.category === category && row.month === month && row.actual > 0);
      const actual = rows.reduce((sum, row) => sum + row.actual, 0);
      const forecast = rows.reduce((sum, row) => sum + row.forecast, 0);
      const absoluteError = rows.reduce((sum, row) => sum + row.absoluteError, 0);
      return [month, {
        products: rows.length,
        actual,
        forecast,
        difference: actual - forecast,
        wape: actual > 0 ? absoluteError / actual : null,
        mae: rows.length ? absoluteError / rows.length : null,
        inside15: rows.filter((row) => row.absoluteError <= 15).length,
      }];
    })
  )])
);
const sizeGroup = (product) => {
  const normalized = productName(product);
  if (/\b(GDE|GRANDE)\b/.test(normalized)) return "Grande";
  if (/\b(IND|INDIVIDUAL)\b/.test(normalized)) return "Individual";
  if (/\b(CH|CHICO)\b/.test(normalized)) return "Chico";
  if (/\bMEDIANO\b/.test(normalized)) return "Mediano";
  return "General";
};
const categoryComponents = (months, targetKey, category) => {
  const [year, month] = targetKey.split("-").map(Number);
  const available = candidates(months, targetKey, year, month);
  const recent = candidateByNames(available, ["Dia semana ultimo mes", "Ultimo mes por dia"]);
  const weighted = candidateByNames(available, ["Dia semana ponderado", "Promedio 2 meses por dia"]);
  const seasonal = candidateByNames(available, ["Estacional por dia semana", "Mismo mes año anterior"]);
  const trend = candidateByNames(available, ["Tendencia reciente", "Promedio 3 meses por dia", "Promedio 2 meses por dia"]);
  if (category === "Pasteles grandes" || category === "Gelatinas") return [seasonal, recent, trend];
  return [recent, weighted];
};
const blendComponents = (components, weights) => {
  const usable = components.map((value, index) => [value, weights[index]]).filter(([value]) => Number.isFinite(value));
  const weight = usable.reduce((sum, part) => sum + part[1], 0);
  return weight ? usable.reduce((sum, part) => sum + part[0] * part[1], 0) / weight : 0;
};
const weightOptions = (componentCount) => {
  const options = [];
  if (componentCount === 2) {
    for (let first = 0; first <= 10; first += 1) options.push([first / 10, (10 - first) / 10]);
  } else {
    for (let first = 0; first <= 10; first += 1) {
      for (let second = 0; second <= 10 - first; second += 1) {
        options.push([first / 10, second / 10, (10 - first - second) / 10]);
      }
    }
  }
  return options;
};
const optimizedWeights = (category, targetKey) => {
  const validationKey = previousMonthId(targetKey);
  const componentCount = category === "Pasteles grandes" || category === "Gelatinas" ? 3 : 2;
  let best = { weights: componentCount === 3 ? [0.3, 0.5, 0.2] : [0.7, 0.3], wape: Infinity };
  for (const weights of weightOptions(componentCount)) {
    let totalActual = 0;
    let totalError = 0;
    let evaluated = 0;
    for (const product of products) {
      if (priorityCategory(product) !== category) continue;
      const months = salesByProduct.get(product) || new Map();
      const actual = months.get(validationKey)?.total || 0;
      if (actual <= 0) continue;
      const forecast = blendComponents(categoryComponents(months, validationKey, category), weights);
      if (forecast <= 0) continue;
      totalActual += actual;
      totalError += Math.abs(actual - forecast);
      evaluated += 1;
    }
    const wape = totalActual > 0 ? totalError / totalActual : Infinity;
    if (evaluated >= 3 && wape < best.wape) best = { weights, wape, evaluated, validationKey };
  }
  return best;
};
const optimizedCategoryResults = {};
for (const category of ["Pasteles grandes", "Gelatinas", "Galletas", "Pan"]) {
  optimizedCategoryResults[category] = {};
  for (const targetKey of ["2026-05", "2026-06"]) {
    const label = targetKey === "2026-05" ? "Mayo 2026" : "Junio 2026";
    const optimization = optimizedWeights(category, targetKey);
    const preliminary = products
      .filter((product) => priorityCategory(product) === category)
      .map((product) => {
        const months = salesByProduct.get(product) || new Map();
        const base = blendComponents(categoryComponents(months, targetKey, category), optimization.weights);
        const validationKey = previousMonthId(targetKey);
        const validationBase = blendComponents(categoryComponents(months, validationKey, category), optimization.weights);
        const validationActual = months.get(validationKey)?.total || 0;
        const calibration = validationBase > 0 && validationActual > 0
          ? clamp(validationActual / validationBase, 0.9, 1.1)
          : 1;
        return { product, months, base, calibration, forecast: base * calibration };
      });
    const peerMedians = new Map();
    for (const group of ["Grande", "Individual", "Chico", "Mediano", "General"]) {
      const values = preliminary
        .filter((row) => sizeGroup(row.product) === group && row.forecast > 0)
        .map((row) => row.forecast)
        .sort((a, b) => a - b);
      if (values.length) peerMedians.set(group, values[Math.floor(values.length / 2)]);
    }
    const evaluatedRows = preliminary.map((row) => {
      const actual = row.months.get(targetKey)?.total || 0;
      const analog = row.forecast > 0 ? 0 : (peerMedians.get(sizeGroup(row.product)) || peerMedians.get("General") || 0) * 0.7;
      const forecast = row.forecast || analog;
      return { ...row, actual, forecast, absoluteError: Math.abs(actual - forecast), usedAnalog: analog > 0 };
    }).filter((row) => row.actual > 0);
    const actual = evaluatedRows.reduce((sum, row) => sum + row.actual, 0);
    const forecast = evaluatedRows.reduce((sum, row) => sum + row.forecast, 0);
    const absoluteError = evaluatedRows.reduce((sum, row) => sum + row.absoluteError, 0);
    optimizedCategoryResults[category][label] = {
      weights: optimization.weights,
      validationMonth: optimization.validationKey,
      validationWape: optimization.wape,
      products: evaluatedRows.length,
      analogProducts: evaluatedRows.filter((row) => row.usedAnalog).length,
      actual,
      forecast,
      difference: actual - forecast,
      wape: actual > 0 ? absoluteError / actual : null,
      mae: evaluatedRows.length ? absoluteError / evaluatedRows.length : null,
      inside15: evaluatedRows.filter((row) => row.absoluteError <= 15).length,
    };
  }
}
const targetProductNames = ["FRUTAS GDE", "MOKA GDE", "PAY DE GUAYABA GDE", "DURAZNO GDE", "CHEESECAKE GDE"];
const targetProductDiagnostics = Object.fromEntries(targetProductNames.map((product) => {
  const months = salesByProduct.get(product) || new Map();
  const monthTotals = Object.fromEntries([...months.entries()].map(([key, value]) => [key, value.total]));
  const targets = {};
  for (const [targetKey, label] of [["2026-05", "Mayo 2026"], ["2026-06", "Junio 2026"]]) {
    const [year, month] = targetKey.split("-").map(Number);
    const available = candidates(months, targetKey, year, month);
    const validationKey = previousMonthId(targetKey);
    const [validationYear, validationMonth] = validationKey.split("-").map(Number);
    const validationCandidates = candidates(months, validationKey, validationYear, validationMonth);
    targets[label] = {
      actual: months.get(targetKey)?.total || 0,
      validationMonth: validationKey,
      validationActual: months.get(validationKey)?.total || 0,
      candidates: Object.fromEntries([...available.entries()].map(([name, value]) => [name, round(value)])),
      validationCandidates: Object.fromEntries([...validationCandidates.entries()].map(([name, value]) => [name, round(value)])),
    };
  }
  return [product, { monthTotals, targets }];
}));
const targetProductModel = {};
for (const targetKey of ["2026-05", "2026-06", "2026-07"]) {
  const label = targetKey === "2026-05" ? "Mayo 2026" : targetKey === "2026-06" ? "Junio 2026" : "Julio 2026";
  const rows = targetProductNames.map((product) => {
    const months = salesByProduct.get(product) || new Map();
    const [year, month] = targetKey.split("-").map(Number);
    const available = candidates(months, targetKey, year, month);
    const seasonal = candidateByNames(available, ["Mismo mes año anterior", "Estacional por dia semana"]);
    const recent = candidateByNames(available, ["Dia semana ultimo mes", "Ultimo mes por dia"]);
    const ratio = seasonal > 0 && recent > 0 ? seasonal / recent : 1;
    const seasonalWeight = ratio > 1.5 || ratio < 0.67 ? 0.5 : 0.9;
    const forecast = Number.isFinite(seasonal) && Number.isFinite(recent)
      ? seasonal * seasonalWeight + recent * (1 - seasonalWeight)
      : seasonal || recent || 0;
    const actual = months.get(targetKey)?.total || 0;
    return {
      product,
      actual,
      seasonal,
      recent,
      seasonalWeight,
      forecast,
      difference: actual - forecast,
      absoluteError: Math.abs(actual - forecast),
    };
  });
  const actual = rows.reduce((sum, row) => sum + row.actual, 0);
  const forecast = rows.reduce((sum, row) => sum + row.forecast, 0);
  const absoluteError = rows.reduce((sum, row) => sum + row.absoluteError, 0);
  targetProductModel[label] = {
    rows,
    total: {
      actual,
      forecast,
      difference: actual - forecast,
      wape: actual > 0 ? absoluteError / actual : null,
      mae: rows.length ? absoluteError / rows.length : null,
      inside15: rows.filter((row) => row.absoluteError <= 15).length,
    },
  };
}
const julyPriorityForecasts = { "Pasteles grandes": [], Gelatinas: [] };
for (const product of products) {
  const category = priorityCategory(product);
  if (!Object.prototype.hasOwnProperty.call(julyPriorityForecasts, category)) continue;
  const months = salesByProduct.get(product) || new Map();
  const available = candidates(months, "2026-07", 2026, 7);
  let forecast = 0;
  let method = "";
  let calibration = 1;
  if (category === "Pasteles grandes") {
    const seasonal = candidateByNames(available, ["Mismo mes año anterior", "Estacional por dia semana"]);
    const recent = candidateByNames(available, ["Dia semana ultimo mes", "Ultimo mes por dia"]);
    const ratio = seasonal > 0 && recent > 0 ? seasonal / recent : 1;
    const seasonalWeight = ratio > 1.5 || ratio < 0.67 ? 0.5 : 0.9;
    if (Number.isFinite(seasonal) && Number.isFinite(recent)) {
      forecast = seasonal * seasonalWeight + recent * (1 - seasonalWeight);
      method = `${seasonalWeight * 100}% estacional + ${(1 - seasonalWeight) * 100}% reciente`;
    } else {
      forecast = seasonal || recent || candidateByNames(available, ["Promedio 3 meses por dia", "Tendencia reciente"]) || 0;
      method = seasonal ? "Estacional" : recent ? "Reciente" : "Promedio histórico";
    }
  } else {
    const recent = candidateByNames(available, ["Dia semana ultimo mes", "Ultimo mes por dia"]);
    const weighted = candidateByNames(available, ["Dia semana ponderado", "Promedio 2 meses por dia"]);
    const seasonal = candidateByNames(available, ["Estacional por dia semana", "Mismo mes año anterior"]);
    forecast = blendComponents([recent, weighted, seasonal], [0.5, 0.3, 0.2]);
    const juneAvailable = candidates(months, "2026-06", 2026, 6);
    const juneBase = blendComponents([
      candidateByNames(juneAvailable, ["Dia semana ultimo mes", "Ultimo mes por dia"]),
      candidateByNames(juneAvailable, ["Dia semana ponderado", "Promedio 2 meses por dia"]),
      candidateByNames(juneAvailable, ["Estacional por dia semana", "Mismo mes año anterior"]),
    ], [0.5, 0.3, 0.2]);
    const juneActual = months.get("2026-06")?.total || 0;
    calibration = juneBase > 0 && juneActual > 0 ? clamp(juneActual / juneBase, 0.9, 1.1) : 1;
    forecast *= calibration;
    method = "50% reciente + 30% ponderado + 20% estacional";
  }
  if (forecast <= 0) continue;
  julyPriorityForecasts[category].push({
    product,
    forecast,
    roundedForecast: Math.round(forecast),
    method,
    calibration,
  });
}
for (const category of Object.keys(julyPriorityForecasts)) {
  julyPriorityForecasts[category].sort((a, b) => a.product.localeCompare(b.product, "es"));
}
const julyPriorityTotals = Object.fromEntries(Object.entries(julyPriorityForecasts).map(([category, rows]) => [category, {
  products: rows.length,
  exactTotal: rows.reduce((sum, row) => sum + row.forecast, 0),
  roundedTotal: rows.reduce((sum, row) => sum + row.roundedForecast, 0),
}]));
const sizeModelForecast = (months, targetKey, rule) => {
  const [year, month] = targetKey.split("-").map(Number);
  if (rule === "Adaptativo") {
    const priorKey = previousMonthId(targetKey);
    const [priorYear, priorMonth] = priorKey.split("-").map(Number);
    return selectAndForecast(months, priorYear, priorMonth, year, month).forecast;
  }
  if (rule === "Combinado") return categoryPrediction(months, targetKey, "Pasteles grandes").forecast;
  const available = candidates(months, targetKey, year, month);
  const seasonal = candidateByNames(available, ["Mismo mes año anterior", "Estacional por dia semana"]);
  const recent = candidateByNames(available, ["Dia semana ultimo mes", "Ultimo mes por dia"]);
  const ratio = seasonal > 0 && recent > 0 ? seasonal / recent : 1;
  const seasonalWeight = ratio > 1.5 || ratio < 0.67 ? 0.5 : 0.9;
  return Number.isFinite(seasonal) && Number.isFinite(recent)
    ? seasonal * seasonalWeight + recent * (1 - seasonalWeight)
    : seasonal || recent || candidateByNames(available, ["Promedio 3 meses por dia", "Tendencia reciente"]) || 0;
};
const sizeCategories = ["Pasteles medianos", "Pasteles chicos", "Mini medianos"];
const sizeModelAnalysis = {};
const selectedSizeModels = {};
const julySizeForecasts = {};
for (const category of sizeCategories) {
  sizeModelAnalysis[category] = {};
  const categoryProducts = products.filter((product) => cakeSizeCategory(product) === category);
  for (const rule of ["Estacional-reciente", "Combinado", "Adaptativo"]) {
    sizeModelAnalysis[category][rule] = {};
    for (const [targetKey, label] of [["2026-05", "Mayo 2026"], ["2026-06", "Junio 2026"]]) {
      const rows = categoryProducts.map((product) => {
        const months = salesByProduct.get(product) || new Map();
        const actual = months.get(targetKey)?.total || 0;
        const forecast = sizeModelForecast(months, targetKey, rule);
        return { product, actual, forecast, error: Math.abs(actual - forecast) };
      }).filter((row) => row.actual > 0);
      const actual = rows.reduce((sum, row) => sum + row.actual, 0);
      const forecast = rows.reduce((sum, row) => sum + row.forecast, 0);
      const error = rows.reduce((sum, row) => sum + row.error, 0);
      sizeModelAnalysis[category][rule][label] = {
        products: rows.length,
        actual,
        forecast,
        difference: actual - forecast,
        wape: actual > 0 ? error / actual : null,
        mae: rows.length ? error / rows.length : null,
        inside15: rows.filter((row) => row.error <= 15).length,
      };
    }
  }
  const ruleScores = Object.entries(sizeModelAnalysis[category]).map(([rule, months]) => {
    const actual = months["Mayo 2026"].actual + months["Junio 2026"].actual;
    const absoluteError = months["Mayo 2026"].wape * months["Mayo 2026"].actual + months["Junio 2026"].wape * months["Junio 2026"].actual;
    return { rule, wape: actual > 0 ? absoluteError / actual : Infinity };
  }).sort((a, b) => a.wape - b.wape);
  const selectedRule = ruleScores[0]?.rule || "Estacional-reciente";
  selectedSizeModels[category] = { selectedRule, validationWape: ruleScores[0]?.wape ?? null };
  julySizeForecasts[category] = categoryProducts.map((product) => {
    const forecast = sizeModelForecast(salesByProduct.get(product) || new Map(), "2026-07", selectedRule);
    return { product, forecast, roundedForecast: Math.round(forecast) };
  }).filter((row) => row.forecast > 0).sort((a, b) => a.product.localeCompare(b.product, "es"));
}
const julySizeTotals = Object.fromEntries(Object.entries(julySizeForecasts).map(([category, rows]) => [category, {
  products: rows.length,
  exactTotal: rows.reduce((sum, row) => sum + row.forecast, 0),
  roundedTotal: rows.reduce((sum, row) => sum + row.roundedForecast, 0),
}]));
const operationalForecast = (months, targetKey, rule) => {
  const [year, month] = targetKey.split("-").map(Number);
  if (rule === "Adaptativo") {
    const priorKey = previousMonthId(targetKey);
    const [priorYear, priorMonth] = priorKey.split("-").map(Number);
    return selectAndForecast(months, priorYear, priorMonth, year, month).forecast;
  }
  const available = candidates(months, targetKey, year, month);
  const recent = candidateByNames(available, ["Dia semana ultimo mes", "Ultimo mes por dia"]);
  const weighted = candidateByNames(available, ["Dia semana ponderado", "Promedio 2 meses por dia"]);
  if (rule === "Reciente ponderado") {
    const base = blendComponents([recent, weighted], [0.7, 0.3]);
    const validationKey = previousMonthId(targetKey);
    const [validationYear, validationMonth] = validationKey.split("-").map(Number);
    const validationAvailable = candidates(months, validationKey, validationYear, validationMonth);
    const validationBase = blendComponents([
      candidateByNames(validationAvailable, ["Dia semana ultimo mes", "Ultimo mes por dia"]),
      candidateByNames(validationAvailable, ["Dia semana ponderado", "Promedio 2 meses por dia"]),
    ], [0.7, 0.3]);
    const validationActual = months.get(validationKey)?.total || 0;
    const calibration = validationBase > 0 && validationActual > 0
      ? clamp(validationActual / validationBase, 0.9, 1.1)
      : 1;
    return base * calibration;
  }
  const seasonal = candidateByNames(available, ["Mismo mes año anterior", "Estacional por dia semana"]);
  const ratio = seasonal > 0 && recent > 0 ? seasonal / recent : 1;
  const seasonalWeight = ratio > 1.5 || ratio < 0.67 ? 0.5 : 0.9;
  return Number.isFinite(seasonal) && Number.isFinite(recent)
    ? seasonal * seasonalWeight + recent * (1 - seasonalWeight)
    : seasonal || recent || weighted || 0;
};
const operationalCategories = ["Galletas", "Pan"];
const operationalModelAnalysis = {};
const selectedOperationalModels = {};
const julyOperationalForecasts = {};
for (const category of operationalCategories) {
  operationalModelAnalysis[category] = {};
  const categoryProducts = products.filter((product) => priorityCategory(product) === category);
  for (const rule of ["Estacional-reciente", "Reciente ponderado", "Adaptativo"]) {
    operationalModelAnalysis[category][rule] = {};
    for (const [targetKey, label] of [["2026-05", "Mayo 2026"], ["2026-06", "Junio 2026"]]) {
      const rows = categoryProducts.map((product) => {
        const months = salesByProduct.get(product) || new Map();
        const actual = months.get(targetKey)?.total || 0;
        const forecast = operationalForecast(months, targetKey, rule);
        return { product, actual, forecast, error: Math.abs(actual - forecast) };
      }).filter((row) => row.actual > 0);
      const actual = rows.reduce((sum, row) => sum + row.actual, 0);
      const forecast = rows.reduce((sum, row) => sum + row.forecast, 0);
      const error = rows.reduce((sum, row) => sum + row.error, 0);
      operationalModelAnalysis[category][rule][label] = {
        products: rows.length,
        actual,
        forecast,
        difference: actual - forecast,
        wape: actual > 0 ? error / actual : null,
        mae: rows.length ? error / rows.length : null,
        inside15: rows.filter((row) => row.error <= 15).length,
      };
    }
  }
  const scores = Object.entries(operationalModelAnalysis[category]).map(([rule, values]) => {
    const actual = values["Mayo 2026"].actual + values["Junio 2026"].actual;
    const error = values["Mayo 2026"].wape * values["Mayo 2026"].actual + values["Junio 2026"].wape * values["Junio 2026"].actual;
    return { rule, wape: actual > 0 ? error / actual : Infinity };
  }).sort((a, b) => a.wape - b.wape);
  const selectedRule = scores[0]?.rule || "Reciente ponderado";
  selectedOperationalModels[category] = { selectedRule, validationWape: scores[0]?.wape ?? null };
  julyOperationalForecasts[category] = categoryProducts.map((product) => {
    const forecast = operationalForecast(salesByProduct.get(product) || new Map(), "2026-07", selectedRule);
    return { product, forecast, roundedForecast: Math.round(forecast) };
  }).filter((row) => row.forecast > 0).sort((a, b) => a.product.localeCompare(b.product, "es"));
}
const julyOperationalTotals = Object.fromEntries(Object.entries(julyOperationalForecasts).map(([category, rows]) => [category, {
  products: rows.length,
  exactTotal: rows.reduce((sum, row) => sum + row.forecast, 0),
  roundedTotal: rows.reduce((sum, row) => sum + row.roundedForecast, 0),
}]));
const allJulyForecasts = [
  ...julyPriorityForecasts["Pasteles grandes"].map((row) => ({ ...row, category: "Pasteles grandes" })),
  ...julySizeForecasts["Pasteles medianos"].map((row) => ({ ...row, category: "Pasteles medianos" })),
  ...julySizeForecasts["Pasteles chicos"].map((row) => ({ ...row, category: "Pasteles chicos" })),
  ...julySizeForecasts["Mini medianos"].map((row) => ({ ...row, category: "Mini medianos" })),
  ...julyPriorityForecasts.Gelatinas.map((row) => ({ ...row, category: "Gelatinas" })),
  ...julyOperationalForecasts.Galletas.map((row) => ({ ...row, category: "Galletas" })),
  ...julyOperationalForecasts.Pan.map((row) => ({ ...row, category: "Pan" })),
];
const julyDates = Array.from({ length: 31 }, (_, index) => new Date(2026, 6, index + 1));
const weekdayProfile = (product) => {
  const months = salesByProduct.get(product) || new Map();
  const sourceKeys = [...months.keys()]
    .filter((key) => key <= "2026-06" && months.get(key).daily.size)
    .sort()
    .slice(-2);
  const weights = sourceKeys.length === 2 ? [0.4, 0.6] : [1];
  const averages = new Map();
  for (let weekday = 0; weekday < 7; weekday += 1) {
    let total = 0;
    let usedWeight = 0;
    sourceKeys.forEach((key, index) => {
      const [year, month] = key.split("-").map(Number);
      const values = [...months.get(key).daily.entries()]
        .filter(([day]) => new Date(year, month - 1, day).getDay() === weekday)
        .map(([, value]) => value);
      if (!values.length) return;
      total += (values.reduce((sum, value) => sum + value, 0) / values.length) * weights[index];
      usedWeight += weights[index];
    });
    if (usedWeight) averages.set(weekday, total / usedWeight);
  }
  return averages;
};
const productionDateForSale = (saleDate) => {
  const productionDate = new Date(saleDate);
  productionDate.setDate(productionDate.getDate() - 1);
  while (productionDate.getDay() === 0) productionDate.setDate(productionDate.getDate() - 1);
  return productionDate;
};
const weeklyProductionDateForSale = (saleDate, productionWeekday) => {
  const productionDate = new Date(saleDate);
  let daysBack = (productionDate.getDay() - productionWeekday + 7) % 7;
  if (daysBack === 0) daysBack = 7;
  productionDate.setDate(productionDate.getDate() - daysBack);
  return productionDate;
};
const formatIsoDate = (date) => `${date.getFullYear()}-${String(date.getMonth() + 1).padStart(2, "0")}-${String(date.getDate()).padStart(2, "0")}`;
const roundProduction = (category, quantity) => {
  if (category.startsWith("Pasteles") || category === "Mini medianos") {
    if (quantity < 8) return 0;
    return Math.max(10, Math.ceil(quantity / 5) * 5);
  }
  if (quantity < 1) return 0;
  return Math.ceil(quantity);
};
const productionByKey = new Map();
const weeklyAssignments = new Map();
const weekdayLoads = new Map([1, 2, 3, 4, 5, 6].map((weekday) => [weekday, 0]));
const lowVolumeCakes = allJulyForecasts
  .filter((row) => (row.category.startsWith("Pasteles") || row.category === "Mini medianos") && row.forecast / julyDates.length < 12)
  .sort((a, b) => b.forecast - a.forecast);
for (const row of lowVolumeCakes) {
  const assignedWeekday = [...weekdayLoads.entries()].sort((a, b) => a[1] - b[1] || a[0] - b[0])[0][0];
  weeklyAssignments.set(row.product, assignedWeekday);
  weekdayLoads.set(assignedWeekday, weekdayLoads.get(assignedWeekday) + row.forecast * 7 / julyDates.length);
}
for (const row of allJulyForecasts) {
  const profile = weekdayProfile(row.product);
  const unscaled = julyDates.map((date) => profile.get(date.getDay()) || 1);
  const unscaledTotal = unscaled.reduce((sum, value) => sum + value, 0);
  const dailySales = julyDates.map((date, index) => ({
    date,
    quantity: unscaledTotal > 0 ? row.forecast * unscaled[index] / unscaledTotal : row.forecast / julyDates.length,
  }));
  const productBatches = new Map();
  const isCake = row.category.startsWith("Pasteles") || row.category === "Mini medianos";
  const useWeeklyBatch = isCake && row.forecast / julyDates.length < 12;
  for (const sale of dailySales) {
    const productionDate = useWeeklyBatch
      ? weeklyProductionDateForSale(sale.date, weeklyAssignments.get(row.product))
      : productionDateForSale(sale.date);
    const key = formatIsoDate(productionDate);
    const batch = productBatches.get(key) || { demand: 0, saleDates: [], saleItems: [] };
    batch.demand += sale.quantity;
    batch.saleDates.push(formatIsoDate(sale.date));
    batch.saleItems.push({ date: formatIsoDate(sale.date), quantity: sale.quantity });
    productBatches.set(key, batch);
  }
  for (const [key, batch] of productBatches.entries()) {
    const production = roundProduction(row.category, batch.demand);
    const batchKey = `${key}|${row.product}`;
    productionByKey.set(batchKey, {
      productionDate: key,
      product: row.product,
      category: row.category,
      saleDates: batch.saleDates,
      saleItems: batch.saleItems,
      demand: batch.demand,
      production,
      status: production > 0 ? "Produccion base" : "Bajo pedido",
    });
  }
}
const dailyProductionDetail = [...productionByKey.values()].sort((a, b) => a.productionDate.localeCompare(b.productionDate) || a.product.localeCompare(b.product, "es"));
const dailyProductionSummaryMap = new Map();
for (const row of dailyProductionDetail) {
  const summary = dailyProductionSummaryMap.get(row.productionDate) || {
    productionDate: row.productionDate,
    "Pasteles grandes": 0,
    "Pasteles medianos": 0,
    "Pasteles chicos": 0,
    "Mini medianos": 0,
    Gelatinas: 0,
    Galletas: 0,
    Pan: 0,
    total: 0,
  };
  summary[row.category] += row.production;
  summary.total += row.production;
  dailyProductionSummaryMap.set(row.productionDate, summary);
}
const dailyProductionSummary = [...dailyProductionSummaryMap.values()].sort((a, b) => a.productionDate.localeCompare(b.productionDate));
const productionCategoryTotals = Object.fromEntries(
  ["Pasteles grandes", "Pasteles medianos", "Pasteles chicos", "Mini medianos", "Gelatinas", "Galletas", "Pan"].map((category) => [
    category,
    dailyProductionSummary.reduce((sum, row) => sum + row[category], 0),
  ])
);
const csvEscape = (value) => {
  const text = String(value ?? "");
  return /[",\n]/.test(text) ? `"${text.replace(/"/g, '""')}"` : text;
};
const dailyProductionCsvRows = [
  ["Fecha produccion", "Dia", "Producto", "Categoria", "Fechas de venta cubiertas", "Demanda pronosticada", "Produccion base", "Estado"],
  ...dailyProductionDetail.map((row) => {
    const date = new Date(`${row.productionDate}T12:00:00`);
    const dayName = ["Domingo", "Lunes", "Martes", "Miercoles", "Jueves", "Viernes", "Sabado"][date.getDay()];
    return [
      row.productionDate,
      dayName,
      row.product,
      row.category,
      row.saleDates.join(" | "),
      round(row.demand),
      row.production,
      row.status,
    ];
  }),
];
const dailyProductionCsv = dailyProductionCsvRows.map((row) => row.map(csvEscape).join(",")).join("\r\n");
const dailyProductionCsvPath = `${ROOT}/PRODUCCION_JULIO_DETALLE_POR_PRODUCTO.csv`;
if (process.env.WRITE_DAILY_CSV === "1") fs.writeFileSync(dailyProductionCsvPath, `\uFEFF${dailyProductionCsv}`, "utf8");
const remainingStartProduction = "2026-07-15";
const remainingStartSales = "2026-07-16";
const remainingBatchMap = new Map();
for (const row of dailyProductionDetail) {
  const futureSales = row.saleItems.filter((item) => item.date >= remainingStartSales && item.date <= "2026-07-31");
  if (!futureSales.length) continue;
  const productionDate = row.productionDate < remainingStartProduction ? remainingStartProduction : row.productionDate;
  const key = `${productionDate}|${row.product}`;
  const batch = remainingBatchMap.get(key) || {
    productionDate,
    product: row.product,
    category: row.category,
    saleItems: [],
  };
  batch.saleItems.push(...futureSales);
  remainingBatchMap.set(key, batch);
}
const daysBetween = (start, end) => Math.round((new Date(`${end}T12:00:00`) - new Date(`${start}T12:00:00`)) / 86400000);
const remainingByProduct = new Map();
for (const batch of remainingBatchMap.values()) {
  const rows = remainingByProduct.get(batch.product) || [];
  rows.push(batch);
  remainingByProduct.set(batch.product, rows);
}
const remainingProductionDetail = [];
let expiredVirtualTotal = 0;
const finalCarryByProduct = new Map();
for (const [product, batches] of remainingByProduct.entries()) {
  batches.sort((a, b) => a.productionDate.localeCompare(b.productionDate));
  let carry = 0;
  let carryDate = "";
  for (const batch of batches) {
    let expired = 0;
    if (carry > 0 && carryDate && daysBetween(carryDate, batch.productionDate) >= 7) {
      expired = carry;
      expiredVirtualTotal += carry;
      carry = 0;
      carryDate = "";
    }
    const demand = batch.saleItems.reduce((sum, item) => sum + item.quantity, 0);
    const carryIn = carry;
    const carryUsed = Math.min(carry, demand);
    carry -= carryUsed;
    const netDemand = Math.max(0, demand - carryUsed);
    const production = roundProduction(batch.category, netDemand);
    const uncovered = production === 0 ? netDemand : 0;
    if (production > 0) {
      carry += Math.max(0, production - netDemand);
      carryDate = batch.productionDate;
    }
    remainingProductionDetail.push({
      productionDate: batch.productionDate,
      product,
      category: batch.category,
      saleDates: [...new Set(batch.saleItems.map((item) => item.date))],
      demand,
      carryIn,
      expired,
      carryUsed,
      netDemand,
      production,
      carryOut: carry,
      uncovered,
      status: uncovered > 0 ? "Bajo pedido" : "Produccion base",
    });
  }
  finalCarryByProduct.set(product, carry);
}
remainingProductionDetail.sort((a, b) => a.productionDate.localeCompare(b.productionDate) || a.product.localeCompare(b.product, "es"));
const remainingSummaryMap = new Map();
for (const row of remainingProductionDetail) {
  const summary = remainingSummaryMap.get(row.productionDate) || {
    productionDate: row.productionDate,
    "Pasteles grandes": 0,
    "Pasteles medianos": 0,
    "Pasteles chicos": 0,
    "Mini medianos": 0,
    Gelatinas: 0,
    Galletas: 0,
    Pan: 0,
    total: 0,
  };
  summary[row.category] += row.production;
  summary.total += row.production;
  remainingSummaryMap.set(row.productionDate, summary);
}
const remainingProductionSummary = [...remainingSummaryMap.values()].sort((a, b) => a.productionDate.localeCompare(b.productionDate));
const remainingCsvRows = [
  ["Fecha produccion", "Dia", "Producto", "Categoria", "Fechas de venta cubiertas", "Demanda", "Saldo inicial", "Saldo usado", "Demanda neta", "Produccion base", "Saldo final", "Caducidad virtual", "Demanda bajo pedido", "Estado"],
  ...remainingProductionDetail.map((row) => {
    const date = new Date(`${row.productionDate}T12:00:00`);
    const dayName = ["Domingo", "Lunes", "Martes", "Miercoles", "Jueves", "Viernes", "Sabado"][date.getDay()];
    return [
      row.productionDate,
      dayName,
      row.product,
      row.category,
      row.saleDates.join(" | "),
      round(row.demand),
      round(row.carryIn),
      round(row.carryUsed),
      round(row.netDemand),
      row.production,
      round(row.carryOut),
      round(row.expired),
      round(row.uncovered),
      row.status,
    ];
  }),
];
const remainingCsv = remainingCsvRows.map((row) => row.map(csvEscape).join(",")).join("\r\n");
const remainingCsvPath = `${ROOT}/PRODUCCION_RESTANTE_JULIO_CON_SALDO_VIRTUAL.csv`;
if (process.env.WRITE_REMAINING_CSV === "1") fs.writeFileSync(remainingCsvPath, `\uFEFF${remainingCsv}`, "utf8");
const remainingTotals = {
  demand: remainingProductionDetail.reduce((sum, row) => sum + row.demand, 0),
  production: remainingProductionDetail.reduce((sum, row) => sum + row.production, 0),
  productionWithoutCarry: [...remainingBatchMap.values()].reduce((sum, batch) => {
    const demand = batch.saleItems.reduce((total, item) => total + item.quantity, 0);
    return sum + roundProduction(batch.category, demand);
  }, 0),
  carryUsed: remainingProductionDetail.reduce((sum, row) => sum + row.carryUsed, 0),
  carryFinal: [...finalCarryByProduct.values()].reduce((sum, value) => sum + value, 0),
  expiredVirtual: expiredVirtualTotal,
  uncovered: remainingProductionDetail.reduce((sum, row) => sum + row.uncovered, 0),
};
remainingTotals.savingsFromCarry = remainingTotals.productionWithoutCarry - remainingTotals.production;

const finalCategoryForProduct = (product) => cakeSizeCategory(product) || priorityCategory(product);
const finalForecastForProduct = (product, targetKey) => {
  const category = finalCategoryForProduct(product);
  const months = salesByProduct.get(product) || new Map();
  if (["Pasteles grandes", "Pasteles medianos", "Pasteles chicos", "Mini medianos"].includes(category)) {
    return {
      forecast: sizeModelForecast(months, targetKey, "Estacional-reciente"),
      model: "90% estacional + 10% reciente; 50/50 si difieren mas de 50%",
    };
  }
  if (category === "Gelatinas") {
    return {
      forecast: categoryPrediction(months, targetKey, "Gelatinas").forecast,
      model: "50% reciente + 30% ponderado + 20% estacional",
    };
  }
  if (category === "Galletas" || category === "Pan") {
    return {
      forecast: operationalForecast(months, targetKey, "Estacional-reciente"),
      model: "Estacional + reciente",
    };
  }
  return { forecast: 0, model: "Fuera del alcance regular" };
};
const comparisonRows = [];
for (const product of products) {
  const category = finalCategoryForProduct(product);
  if (!["Pasteles grandes", "Pasteles medianos", "Pasteles chicos", "Mini medianos", "Gelatinas", "Galletas", "Pan"].includes(category)) continue;
  for (const [targetKey, label, production] of [
    ["2026-05", "Mayo 2026", productionMay],
    ["2026-06", "Junio 2026", productionJune],
  ]) {
    const actual = salesByProduct.get(product)?.get(targetKey)?.total || 0;
    const result = finalForecastForProduct(product, targetKey);
    if (actual === 0 && result.forecast === 0) continue;
    const difference = actual - result.forecast;
    const produced = production.get(product) || 0;
    comparisonRows.push({
      Producto: product,
      Categoria: category,
      Mes: label,
      "Venta real": actual,
      Pronostico: result.forecast,
      Diferencia: difference,
      "Diferencia absoluta": Math.abs(difference),
      Precision: actual > 0 ? Math.max(0, 1 - Math.abs(difference) / actual) : null,
      Estado: Math.abs(difference) <= 15 ? "Dentro de +/-15 piezas" : "Revisar",
      "Produccion real": produced,
      "Produccion menos venta": produced - actual,
      "Cobertura produccion": actual > 0 ? produced / actual : null,
      Modelo: result.model,
    });
  }
}
const comparisonSummary = (rows) => {
  const actual = rows.reduce((sum, row) => sum + row["Venta real"], 0);
  const forecast = rows.reduce((sum, row) => sum + row.Pronostico, 0);
  const absoluteError = rows.reduce((sum, row) => sum + row["Diferencia absoluta"], 0);
  const active = rows.filter((row) => row["Venta real"] > 0);
  const production = rows.reduce((sum, row) => sum + row["Produccion real"], 0);
  return {
    products: active.length,
    actual,
    forecast,
    difference: actual - forecast,
    absoluteError,
    wape: actual > 0 ? absoluteError / actual : null,
    mae: active.length ? absoluteError / active.length : null,
    inside15: active.filter((row) => row["Diferencia absoluta"] <= 15).length,
    production,
    productionVsSales: production - actual,
  };
};
const comparisonCategoryRows = [];
for (const month of ["Mayo 2026", "Junio 2026"]) {
  for (const category of ["Pasteles grandes", "Pasteles medianos", "Pasteles chicos", "Mini medianos", "Gelatinas", "Galletas", "Pan"]) {
    const rows = comparisonRows.filter((row) => row.Mes === month && row.Categoria === category);
    if (!rows.length) continue;
    comparisonCategoryRows.push({ Mes: month, Categoria: category, ...comparisonSummary(rows) });
  }
}
const mayComparisonSummary = comparisonSummary(comparisonRows.filter((row) => row.Mes === "Mayo 2026"));
const juneComparisonSummary = comparisonSummary(comparisonRows.filter((row) => row.Mes === "Junio 2026"));
const comparisonExecutiveRows = [
  { Indicador: "Uso del archivo", Valor: "Validacion retrospectiva del modelo; julio es la primera prueba independiente" },
  { Indicador: "Alcance", Valor: "Pasteles grandes, medianos, chicos, mini medianos, gelatinas regulares, galletas y pan" },
  { Indicador: "Promociones", Valor: "Excluidas de los modelos regulares" },
  { Indicador: "Venta real mayo", Valor: round(mayComparisonSummary.actual) },
  { Indicador: "Pronostico mayo", Valor: round(mayComparisonSummary.forecast) },
  { Indicador: "Diferencia total mayo", Valor: round(mayComparisonSummary.difference) },
  { Indicador: "WAPE mayo %", Valor: round(mayComparisonSummary.wape * 100, 1) },
  { Indicador: "MAE mayo piezas", Valor: round(mayComparisonSummary.mae) },
  { Indicador: "Dentro de +/-15 mayo", Valor: `${mayComparisonSummary.inside15} de ${mayComparisonSummary.products}` },
  { Indicador: "Produccion real mayo", Valor: round(mayComparisonSummary.production) },
  { Indicador: "Produccion menos venta mayo", Valor: round(mayComparisonSummary.productionVsSales) },
  { Indicador: "Venta real junio", Valor: round(juneComparisonSummary.actual) },
  { Indicador: "Pronostico junio", Valor: round(juneComparisonSummary.forecast) },
  { Indicador: "Diferencia total junio", Valor: round(juneComparisonSummary.difference) },
  { Indicador: "WAPE junio %", Valor: round(juneComparisonSummary.wape * 100, 1) },
  { Indicador: "MAE junio piezas", Valor: round(juneComparisonSummary.mae) },
  { Indicador: "Dentro de +/-15 junio", Valor: `${juneComparisonSummary.inside15} de ${juneComparisonSummary.products}` },
  { Indicador: "Produccion real junio", Valor: round(juneComparisonSummary.production) },
  { Indicador: "Produccion menos venta junio", Valor: round(juneComparisonSummary.productionVsSales) },
];
const comparisonWorkbookPath = `${ROOT}/COMPARATIVO_MAYO_JUNIO_REAL_VS_PRONOSTICO.xlsx`;
if (process.env.WRITE_COMPARISON_XLSX === "1") {
  const comparisonWorkbook = XLSX.utils.book_new();
  const executiveSheet = XLSX.utils.json_to_sheet(comparisonExecutiveRows);
  const categorySheet = XLSX.utils.json_to_sheet(comparisonCategoryRows.map((row) => ({
    Mes: row.Mes,
    Categoria: row.Categoria,
    Productos: row.products,
    "Venta real": round(row.actual),
    Pronostico: round(row.forecast),
    Diferencia: round(row.difference),
    "Error absoluto": round(row.absoluteError),
    "WAPE %": row.wape === null ? "" : round(row.wape * 100, 1),
    "MAE piezas": round(row.mae),
    "Dentro de +/-15": `${row.inside15} de ${row.products}`,
    "Produccion real": round(row.production),
    "Produccion menos venta": round(row.productionVsSales),
  })));
  const detailSheet = XLSX.utils.json_to_sheet(comparisonRows.map((row) => ({
    Producto: row.Producto,
    Categoria: row.Categoria,
    Mes: row.Mes,
    "Venta real": round(row["Venta real"]),
    Pronostico: round(row.Pronostico),
    Diferencia: round(row.Diferencia),
    "Diferencia absoluta": round(row["Diferencia absoluta"]),
    "Precision %": row.Precision === null ? "" : round(row.Precision * 100, 1),
    Estado: row.Estado,
    "Produccion real": round(row["Produccion real"]),
    "Produccion menos venta": round(row["Produccion menos venta"]),
    "Cobertura produccion %": row["Cobertura produccion"] === null ? "" : round(row["Cobertura produccion"] * 100, 1),
    Modelo: row.Modelo,
  })));
  const alertsSheet = XLSX.utils.json_to_sheet(comparisonRows
    .filter((row) => row["Diferencia absoluta"] > 15)
    .sort((a, b) => b["Diferencia absoluta"] - a["Diferencia absoluta"])
    .map((row) => ({
      Producto: row.Producto,
      Categoria: row.Categoria,
      Mes: row.Mes,
      "Venta real": round(row["Venta real"]),
      Pronostico: round(row.Pronostico),
      Diferencia: round(row.Diferencia),
      "Diferencia absoluta": round(row["Diferencia absoluta"]),
      Modelo: row.Modelo,
    })));
  const methodologySheet = XLSX.utils.json_to_sheet([
    { Concepto: "Validacion", Detalle: "Mayo y junio son pruebas retrospectivas; no sustituyen la validacion independiente de julio." },
    { Concepto: "Pasteles", Detalle: "Estacionalidad y comportamiento reciente; proteccion 50/50 si las referencias difieren mas de 50%." },
    { Concepto: "Gelatinas", Detalle: "50% reciente, 30% ponderado y 20% estacional; promociones separadas." },
    { Concepto: "Galletas y pan", Detalle: "Modelo estacional-reciente con confianza baja." },
    { Concepto: "Diferencia", Detalle: "Venta real - Pronostico. Positivo significa pronostico bajo." },
    { Concepto: "WAPE", Detalle: "Suma de diferencias absolutas / suma de ventas reales." },
    { Concepto: "Produccion", Detalle: "Se compara como dato operativo; no se usa para maquillar el pronostico de ventas." },
  ]);
  executiveSheet["!cols"] = [{ wch: 38 }, { wch: 92 }];
  categorySheet["!cols"] = [14, 24, 11, 14, 14, 14, 18, 12, 14, 20, 17, 24].map((wch) => ({ wch }));
  detailSheet["!cols"] = [34, 22, 14, 14, 14, 14, 20, 14, 24, 17, 24, 24, 62].map((wch) => ({ wch }));
  alertsSheet["!cols"] = [34, 22, 14, 14, 14, 14, 20, 62].map((wch) => ({ wch }));
  methodologySheet["!cols"] = [{ wch: 25 }, { wch: 110 }];
  categorySheet["!autofilter"] = { ref: `A1:L${comparisonCategoryRows.length + 1}` };
  detailSheet["!autofilter"] = { ref: `A1:M${comparisonRows.length + 1}` };
  alertsSheet["!autofilter"] = { ref: `A1:H${comparisonRows.filter((row) => row["Diferencia absoluta"] > 15).length + 1}` };
  for (let rowNumber = 2; rowNumber <= comparisonRows.length + 1; rowNumber += 1) {
    const source = comparisonRows[rowNumber - 2];
    detailSheet[`F${rowNumber}`] = { t: "n", f: `D${rowNumber}-E${rowNumber}`, v: round(source.Diferencia) };
    detailSheet[`G${rowNumber}`] = { t: "n", f: `ABS(F${rowNumber})`, v: round(source["Diferencia absoluta"]) };
    detailSheet[`H${rowNumber}`] = source.Precision === null
      ? { t: "s", f: `IF(D${rowNumber}=0,"",MAX(0,1-G${rowNumber}/D${rowNumber})*100)`, v: "" }
      : { t: "n", f: `IF(D${rowNumber}=0,"",MAX(0,1-G${rowNumber}/D${rowNumber})*100)`, v: round(source.Precision * 100, 1) };
    detailSheet[`I${rowNumber}`] = { t: "s", f: `IF(G${rowNumber}<=15,"Dentro de +/-15 piezas","Revisar")`, v: source.Estado };
    detailSheet[`K${rowNumber}`] = { t: "n", f: `J${rowNumber}-D${rowNumber}`, v: round(source["Produccion menos venta"]) };
    detailSheet[`L${rowNumber}`] = source["Cobertura produccion"] === null
      ? { t: "s", f: `IF(D${rowNumber}=0,"",J${rowNumber}/D${rowNumber}*100)`, v: "" }
      : { t: "n", f: `IF(D${rowNumber}=0,"",J${rowNumber}/D${rowNumber}*100)`, v: round(source["Cobertura produccion"] * 100, 1) };
  }
  XLSX.utils.book_append_sheet(comparisonWorkbook, executiveSheet, "Resumen ejecutivo");
  XLSX.utils.book_append_sheet(comparisonWorkbook, categorySheet, "Resumen por categoria");
  XLSX.utils.book_append_sheet(comparisonWorkbook, detailSheet, "Real vs pronostico");
  XLSX.utils.book_append_sheet(comparisonWorkbook, alertsSheet, "Productos por revisar");
  XLSX.utils.book_append_sheet(comparisonWorkbook, methodologySheet, "Metodologia");
  XLSX.writeFile(comparisonWorkbook, comparisonWorkbookPath);
}
const dailyControlWorkbookPath = `${ROOT}/CONTROL_DIARIO_JULIO_PRODUCCION_VENTAS.xlsx`;
if (process.env.WRITE_DAILY_CONTROL_XLSX === "1") {
  const controlWorkbook = XLSX.utils.book_new();
  const instructionsSheet = XLSX.utils.json_to_sheet([
    { Paso: 1, Instruccion: "No modificar pronosticos, produccion sugerida ni formulas." },
    { Paso: 2, Instruccion: "En Ventas diarias llenar Venta real, Bajas, Saldo reportado y Observaciones." },
    { Paso: 3, Instruccion: "En Produccion diaria llenar Produccion real, Ventas reales cubiertas y Bajas." },
    { Paso: 4, Instruccion: "Los productos con estado Bajo pedido solo se producen con pedido o autorizacion." },
    { Paso: 5, Instruccion: "Cerrar la captura al terminar cada dia; no alterar dias anteriores." },
    { Paso: 6, Instruccion: "El plan inicia el 15 de julio y cubre ventas pronosticadas del 16 al 31." },
    { Paso: 7, Instruccion: "No incluye existencias reales de sucursales ni cuarto frio." },
  ]);
  const remainingSalesMap = new Map();
  for (const batch of remainingBatchMap.values()) {
    for (const item of batch.saleItems) {
      if (item.date < remainingStartSales || item.date > "2026-07-31") continue;
      const key = `${item.date}|${batch.product}`;
      const sale = remainingSalesMap.get(key) || {
        date: item.date,
        product: batch.product,
        category: batch.category,
        forecast: 0,
      };
      sale.forecast += item.quantity;
      remainingSalesMap.set(key, sale);
    }
  }
  const controlSalesRows = [...remainingSalesMap.values()]
    .sort((a, b) => a.date.localeCompare(b.date) || a.product.localeCompare(b.product, "es"))
    .map((row) => {
      const date = new Date(`${row.date}T12:00:00`);
      return {
        "Fecha venta": row.date,
        Dia: ["Domingo", "Lunes", "Martes", "Miercoles", "Jueves", "Viernes", "Sabado"][date.getDay()],
        Producto: row.product,
        Categoria: row.category,
        Pronostico: row.forecast,
        "Venta real": "",
        Diferencia: "",
        "Diferencia absoluta": "",
        "Precision %": "",
        Bajas: "",
        "Saldo reportado": "",
        Observaciones: "",
      };
    });
  const controlProductionRows = remainingProductionDetail.map((row) => {
    const date = new Date(`${row.productionDate}T12:00:00`);
    return {
      "Fecha produccion": row.productionDate,
      Dia: ["Domingo", "Lunes", "Martes", "Miercoles", "Jueves", "Viernes", "Sabado"][date.getDay()],
      Producto: row.product,
      Categoria: row.category,
      "Fechas de venta cubiertas": row.saleDates.join(" | "),
      "Demanda pronosticada": row.demand,
      "Saldo virtual usado": row.carryUsed,
      "Produccion sugerida": row.production,
      "Produccion real": "",
      "Diferencia produccion": "",
      "Ventas reales cubiertas": "",
      Bajas: "",
      "Saldo real del lote": "",
      Estado: row.status,
      Observaciones: "",
    };
  });
  const salesSheet = XLSX.utils.json_to_sheet(controlSalesRows);
  const productionSheet = XLSX.utils.json_to_sheet(controlProductionRows);
  for (let row = 2; row <= controlSalesRows.length + 1; row += 1) {
    salesSheet[`E${row}`].z = "0.00";
    salesSheet[`G${row}`] = { t: "n", f: `IF(F${row}="","",F${row}-E${row})`, v: 0 };
    salesSheet[`H${row}`] = { t: "n", f: `IF(G${row}="","",ABS(G${row}))`, v: 0 };
    salesSheet[`I${row}`] = { t: "n", f: `IF(OR(F${row}="",F${row}=0),"",MAX(0,1-H${row}/F${row})*100)`, v: 0 };
    salesSheet[`G${row}`].z = "0.00";
    salesSheet[`H${row}`].z = "0.00";
    salesSheet[`I${row}`].z = "0.0";
  }
  for (let row = 2; row <= controlProductionRows.length + 1; row += 1) {
    productionSheet[`F${row}`].z = "0.00";
    productionSheet[`G${row}`].z = "0.00";
    productionSheet[`J${row}`] = { t: "n", f: `IF(I${row}="","",I${row}-H${row})`, v: 0 };
    productionSheet[`M${row}`] = { t: "n", f: `IF(OR(I${row}="",K${row}=""),"",I${row}-K${row}-L${row})`, v: 0 };
  }
  const summaryDates = [];
  for (let day = 15; day <= 31; day += 1) {
    const date = new Date(2026, 6, day);
    summaryDates.push({
      Fecha: formatIsoDate(date),
      Dia: ["Domingo", "Lunes", "Martes", "Miercoles", "Jueves", "Viernes", "Sabado"][date.getDay()],
      "Venta pronosticada": "",
      "Venta real": "",
      "Diferencia venta": "",
      "Produccion sugerida": "",
      "Produccion real": "",
      Bajas: "",
      "Produccion real menos venta": "",
    });
  }
  const dailySummarySheet = XLSX.utils.json_to_sheet(summaryDates);
  for (let row = 2; row <= summaryDates.length + 1; row += 1) {
    dailySummarySheet[`C${row}`] = { t: "n", f: `SUMIF('Ventas diarias'!A:A,A${row},'Ventas diarias'!E:E)`, v: 0 };
    dailySummarySheet[`D${row}`] = { t: "n", f: `SUMIF('Ventas diarias'!A:A,A${row},'Ventas diarias'!F:F)`, v: 0 };
    dailySummarySheet[`E${row}`] = { t: "n", f: `D${row}-C${row}`, v: 0 };
    dailySummarySheet[`F${row}`] = { t: "n", f: `SUMIF('Produccion diaria'!A:A,A${row},'Produccion diaria'!H:H)`, v: 0 };
    dailySummarySheet[`G${row}`] = { t: "n", f: `SUMIF('Produccion diaria'!A:A,A${row},'Produccion diaria'!I:I)`, v: 0 };
    dailySummarySheet[`H${row}`] = { t: "n", f: `SUMIF('Ventas diarias'!A:A,A${row},'Ventas diarias'!J:J)`, v: 0 };
    dailySummarySheet[`I${row}`] = { t: "n", f: `G${row}-D${row}`, v: 0 };
  }
  const controlProducts = [...new Set(controlSalesRows.map((row) => row.Producto))].sort((a, b) => a.localeCompare(b, "es"));
  const productSummaryRows = controlProducts.map((product) => ({
    Producto: product,
    Categoria: controlSalesRows.find((row) => row.Producto === product)?.Categoria || "",
    "Venta pronosticada": "",
    "Venta real": "",
    Diferencia: "",
    "Diferencia absoluta": "",
    "Precision %": "",
    "Produccion sugerida": "",
    "Produccion real": "",
    Bajas: "",
    "Saldo reportado final": "",
  }));
  const productSummarySheet = XLSX.utils.json_to_sheet(productSummaryRows);
  for (let row = 2; row <= productSummaryRows.length + 1; row += 1) {
    productSummarySheet[`C${row}`] = { t: "n", f: `SUMIF('Ventas diarias'!C:C,A${row},'Ventas diarias'!E:E)`, v: 0 };
    productSummarySheet[`D${row}`] = { t: "n", f: `SUMIF('Ventas diarias'!C:C,A${row},'Ventas diarias'!F:F)`, v: 0 };
    productSummarySheet[`E${row}`] = { t: "n", f: `D${row}-C${row}`, v: 0 };
    productSummarySheet[`F${row}`] = { t: "n", f: `ABS(E${row})`, v: 0 };
    productSummarySheet[`G${row}`] = { t: "n", f: `IF(D${row}=0,"",MAX(0,1-F${row}/D${row})*100)`, v: 0 };
    productSummarySheet[`H${row}`] = { t: "n", f: `SUMIF('Produccion diaria'!C:C,A${row},'Produccion diaria'!H:H)`, v: 0 };
    productSummarySheet[`I${row}`] = { t: "n", f: `SUMIF('Produccion diaria'!C:C,A${row},'Produccion diaria'!I:I)`, v: 0 };
    productSummarySheet[`J${row}`] = { t: "n", f: `SUMIF('Ventas diarias'!C:C,A${row},'Ventas diarias'!J:J)`, v: 0 };
    productSummarySheet[`K${row}`] = { t: "n", f: `LOOKUP(2,1/('Ventas diarias'!C:C=A${row}),'Ventas diarias'!K:K)`, v: 0 };
  }
  instructionsSheet["!cols"] = [{ wch: 10 }, { wch: 115 }];
  salesSheet["!cols"] = [14, 12, 34, 22, 14, 14, 14, 20, 14, 10, 18, 45].map((wch) => ({ wch }));
  productionSheet["!cols"] = [16, 12, 34, 22, 55, 20, 20, 20, 17, 22, 22, 10, 20, 18, 45].map((wch) => ({ wch }));
  dailySummarySheet["!cols"] = [14, 12, 20, 14, 18, 22, 17, 10, 28].map((wch) => ({ wch }));
  productSummarySheet["!cols"] = [34, 22, 20, 14, 14, 20, 14, 22, 17, 10, 22].map((wch) => ({ wch }));
  salesSheet["!autofilter"] = { ref: `A1:L${controlSalesRows.length + 1}` };
  productionSheet["!autofilter"] = { ref: `A1:O${controlProductionRows.length + 1}` };
  productSummarySheet["!autofilter"] = { ref: `A1:K${productSummaryRows.length + 1}` };
  XLSX.utils.book_append_sheet(controlWorkbook, instructionsSheet, "Instrucciones");
  XLSX.utils.book_append_sheet(controlWorkbook, dailySummarySheet, "Resumen diario");
  XLSX.utils.book_append_sheet(controlWorkbook, productSummarySheet, "Resumen por producto");
  XLSX.utils.book_append_sheet(controlWorkbook, salesSheet, "Ventas diarias");
  XLSX.utils.book_append_sheet(controlWorkbook, productionSheet, "Produccion diaria");
  XLSX.writeFile(controlWorkbook, dailyControlWorkbookPath);
}
const calculateEventFactor = (month, year, monthNumber, eventDays) => {
  if (!month?.daily?.size) return null;
  const eventSet = new Set(eventDays);
  let actual = 0;
  let expected = 0;
  for (const eventDay of eventDays) {
    actual += month.daily.get(eventDay) || 0;
    const weekday = new Date(year, monthNumber - 1, eventDay).getDay();
    const comparable = [...month.daily.entries()]
      .filter(([day]) => !eventSet.has(day) && new Date(year, monthNumber - 1, day).getDay() === weekday)
      .map(([, value]) => value);
    expected += comparable.length ? comparable.reduce((sum, value) => sum + value, 0) / comparable.length : 0;
  }
  const rawFactor = expected > 0 ? actual / expected : 1;
  return {
    actual,
    expected,
    rawFactor,
    smoothedFactor: clamp(1 + 0.5 * (rawFactor - 1), 0.75, 3),
  };
};
const eventDefinitions = {
  "Dia de las Madres": {
    monthKey: "2025-05", year: 2025, month: 5, days: [8, 9, 10, 11],
    upliftShare: 0.9,
    sourceKey: "2026-04", sourceYear: 2026, sourceMonth: 4,
    targetKey: "2026-05", targetYear: 2026, targetMonth: 5, targetDays: [8, 9, 10, 11],
  },
  "Dia del Padre": {
    monthKey: "2025-06", year: 2025, month: 6, days: [13, 14, 15, 16],
    upliftShare: 0.35,
    sourceKey: "2026-05", sourceYear: 2026, sourceMonth: 5,
    targetKey: "2026-06", targetYear: 2026, targetMonth: 6, targetDays: [19, 20, 21, 22],
  },
};
const eventBaselineFromPriorMonth = (sourceMonth, sourceYear, sourceMonthNumber, targetYear, targetMonthNumber, targetDays) => {
  if (!sourceMonth?.daily?.size) return 0;
  return targetDays.reduce((sum, targetDay) => {
    const weekday = new Date(targetYear, targetMonthNumber - 1, targetDay).getDay();
    const comparable = [...sourceMonth.daily.entries()]
      .filter(([day]) => new Date(sourceYear, sourceMonthNumber - 1, day).getDay() === weekday)
      .map(([, value]) => value);
    return sum + (comparable.length ? comparable.reduce((total, value) => total + value, 0) / comparable.length : 0);
  }, 0);
};
const eventAnalysis = {};
for (const [eventName, definition] of Object.entries(eventDefinitions)) {
  const productFactors = targetProductNames.map((product) => {
    const productMonths = salesByProduct.get(product);
    const month = productMonths?.get(definition.monthKey);
    const factor = calculateEventFactor(month, definition.year, definition.month, definition.days);
    if (!factor) return { product };
    const sourceMonth = productMonths?.get(definition.sourceKey);
    const targetMonth = productMonths?.get(definition.targetKey);
    const baseline = eventBaselineFromPriorMonth(
      sourceMonth,
      definition.sourceYear,
      definition.sourceMonth,
      definition.targetYear,
      definition.targetMonth,
      definition.targetDays
    );
    const targetActual = definition.targetDays.reduce((sum, day) => sum + (targetMonth?.daily?.get(day) || 0), 0);
    const recommendedFactor = clamp(1 + definition.upliftShare * (factor.rawFactor - 1), 0.75, 4.5);
    const eventForecast = baseline * recommendedFactor;
    return {
      product,
      ...factor,
      targetActual,
      priorMonthBaseline: baseline,
      upliftShare: definition.upliftShare,
      recommendedFactor,
      eventForecast,
      eventDifference: targetActual - eventForecast,
    };
  }).filter((row) => Number.isFinite(row.rawFactor));
  const categoryFactors = {};
  for (const category of ["Pasteles grandes", "Gelatinas", "Galletas", "Pan"]) {
    const factors = products
      .filter((product) => priorityCategory(product) === category)
      .map((product) => calculateEventFactor(salesByProduct.get(product)?.get(definition.monthKey), definition.year, definition.month, definition.days))
      .filter(Boolean);
    const actual = factors.reduce((sum, factor) => sum + factor.actual, 0);
    const expected = factors.reduce((sum, factor) => sum + factor.expected, 0);
    const rawFactor = expected > 0 ? actual / expected : 1;
    categoryFactors[category] = {
      products: factors.length,
      eventSales: actual,
      expectedSales: expected,
      rawFactor,
      smoothedFactor: clamp(1 + 0.5 * (rawFactor - 1), 0.75, 3),
      recommendedFactor: clamp(1 + definition.upliftShare * (rawFactor - 1), 0.75, 4.5),
    };
  }
  eventAnalysis[eventName] = { window: definition.days, productFactors, categoryFactors };
}
const topPriorityErrors = (month) => detailRows
  .filter((row) => row.Mes === month && priorityCategory(row.Producto) && priorityCategory(row.Producto) !== "Promociones separadas")
  .sort((a, b) => b["Diferencia absoluta"] - a["Diferencia absoluta"])
  .slice(0, 10)
  .map((row) => ({
    producto: row.Producto,
    categoria: priorityCategory(row.Producto),
    real: row["Venta real"],
    pronostico: row["Pronostico final"],
    diferencia: row["Diferencia piezas"],
  }));

const fullResult = {
  output: process.env.ANALYZE_ONLY === "1" ? null : OUTPUT,
  products: products.length,
  may: maySummary,
  june: juneSummary,
  methodCounts,
  priorityResults,
  categoryModelSummary,
  optimizedCategoryResults,
  targetProductDiagnostics,
  targetProductModel,
  julyPriorityForecasts,
  julyPriorityTotals,
  sizeModelAnalysis,
  selectedSizeModels,
  julySizeForecasts,
  julySizeTotals,
  operationalModelAnalysis,
  selectedOperationalModels,
  julyOperationalForecasts,
  julyOperationalTotals,
  dailyProductionSummary,
  productionCategoryTotals,
  dailyProductionCsvPath: process.env.WRITE_DAILY_CSV === "1" ? dailyProductionCsvPath : null,
  remainingProductionSummary,
  remainingTotals,
  remainingCsvPath: process.env.WRITE_REMAINING_CSV === "1" ? remainingCsvPath : null,
  comparisonWorkbookPath: process.env.WRITE_COMPARISON_XLSX === "1" ? comparisonWorkbookPath : null,
  dailyControlWorkbookPath: process.env.WRITE_DAILY_CONTROL_XLSX === "1" ? dailyControlWorkbookPath : null,
  mayComparisonSummary,
  juneComparisonSummary,
  eventAnalysis,
  topPriorityErrorsMay: topPriorityErrors("Mayo 2026"),
  topPriorityErrorsJune: topPriorityErrors("Junio 2026"),
  topErrorsMay: topErrors("Mayo 2026"),
  topErrorsJune: topErrors("Junio 2026"),
};
const printedResult = process.env.TARGET_DETAILS === "1"
  ? { targetProductDiagnostics }
  : process.env.TARGET_MODEL === "1"
    ? { targetProductModel }
  : process.env.EVENT_ANALYSIS === "1"
    ? { eventAnalysis }
  : process.env.JULY_SCOPE === "1"
    ? { julyPriorityForecasts, julyPriorityTotals }
  : process.env.SIZE_ANALYSIS === "1"
    ? { sizeModelAnalysis, selectedSizeModels, julySizeForecasts, julySizeTotals }
  : process.env.OPERATIONAL_ANALYSIS === "1"
    ? { operationalModelAnalysis, selectedOperationalModels, julyOperationalForecasts, julyOperationalTotals }
  : process.env.DAILY_PRODUCTION === "1"
    ? { dailyProductionSummary, productionCategoryTotals, dailyProductionCsvPath: process.env.WRITE_DAILY_CSV === "1" ? dailyProductionCsvPath : null }
  : process.env.REMAINING_PRODUCTION === "1"
    ? { remainingProductionSummary, remainingTotals, remainingCsvPath: process.env.WRITE_REMAINING_CSV === "1" ? remainingCsvPath : null }
  : process.env.COMPARISON_XLSX === "1"
    ? { comparisonWorkbookPath: process.env.WRITE_COMPARISON_XLSX === "1" ? comparisonWorkbookPath : null, mayComparisonSummary, juneComparisonSummary, comparisonRows: comparisonRows.length }
  : process.env.COMPACT === "1"
    ? { optimizedCategoryResults }
    : fullResult;
console.log(JSON.stringify(printedResult, null, 2));
