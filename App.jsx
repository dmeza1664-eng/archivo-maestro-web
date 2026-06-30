import React, { useEffect, useMemo, useState } from "react";
import { createRoot } from "react-dom/client";
import JSZip from "jszip";
import * as XLSX from "xlsx";
import {
  BarChart3,
  CheckCircle2,
  Database,
  Download,
  FileSpreadsheet,
  PackageCheck,
  Search,
  ShieldCheck,
  Target,
  Upload,
} from "lucide-react";
import "./style.css";

const INVALID_PRODUCTS = new Set([
  "",
  "NAN",
  "TOTAL",
  "SUBTOTAL",
  "SUMA",
  "SUMAS",
  "PRODUCTO",
  "TOTAL AREA",
  "TOTAL ÁREA",
  "ESPECIALIDAD",
  "GRANDE",
  "MEDIANO",
  "CHICO",
  "FECHA",
  "DIA",
  "DÍA",
  "CALENDARIO",
  "SEMANA",
]);

const CALENDAR_WORDS = new Set([
  "LUNES",
  "MARTES",
  "MIERCOLES",
  "MIÉRCOLES",
  "JUEVES",
  "VIERNES",
  "SABADO",
  "SÁBADO",
  "DOMINGO",
  "MONDAY",
  "TUESDAY",
  "WEDNESDAY",
  "THURSDAY",
  "FRIDAY",
  "SATURDAY",
  "SUNDAY",
  "MON",
  "TUE",
  "WED",
  "THU",
  "FRI",
  "SAT",
  "SUN",
  "ENE",
  "FEB",
  "MAR",
  "ABR",
  "MAY",
  "JUN",
  "JUL",
  "AGO",
  "SEP",
  "OCT",
  "NOV",
  "DIC",
  "JAN",
  "APR",
  "AUG",
  "DEC",
]);

const STATUS_META = {
  "Sin dato real": { className: "muted", label: "Sin dato real" },
  "No producir": { className: "muted", label: "No producir" },
  "Dentro de rango": { className: "ok", label: "Dentro de rango" },
  "Riesgo faltante": { className: "danger", label: "Riesgo faltante" },
  Sobreproduccion: { className: "warn", label: "Sobreproducción" },
  Revisar: { className: "warn", label: "Revisar" },
};

const WEEKDAYS = [
  { index: 1, label: "Lunes" },
  { index: 2, label: "Martes" },
  { index: 3, label: "Miércoles" },
  { index: 4, label: "Jueves" },
  { index: 5, label: "Viernes" },
  { index: 6, label: "Sábado" },
  { index: 0, label: "Domingo" },
];

const WEEKDAY_ALIASES = [
  { names: ["LUNES", "LUN"], index: 1 },
  { names: ["MARTES", "MAR"], index: 2 },
  { names: ["MIERCOLES", "MIE"], index: 3 },
  { names: ["JUEVES", "JUE"], index: 4 },
  { names: ["VIERNES", "VIE"], index: 5 },
  { names: ["SABADO", "SAB"], index: 6 },
  { names: ["DOMINGO", "DOM"], index: 0 },
];

function norm(value) {
  return String(value ?? "")
    .trim()
    .toUpperCase()
    .normalize("NFD")
    .replace(/[\u0300-\u036f]/g, "");
}

function normalizeProduct(value) {
  const normalized = norm(value)
    .replace(/[.,/\\_-]+/g, " ")
    .replace(/\s+/g, " ")
    .trim();
  const compact = normalized.replace(/[^A-Z0-9]/g, "");
  if (compact === "PINAGDE" || compact === "PINAGRANDE") return "PINA GDE";
  return normalized;
}

function isSliceProduct(value) {
  const normalized = normalizeProduct(value);
  return /\b(REBANADA|REBANADAS|REB|RBN)\b/.test(normalized);
}

function isOperationalCakeProduct(value) {
  const normalized = normalizeProduct(value);
  return /\b(GDE|GRANDE|MED|MEDIANO|CH|CHICO)\b/.test(normalized);
}

function getProduccionSugeridaPastel(value) {
  const numericValue = Number(value) || 0;
  if (numericValue < 8) return 0;
  return 10 + Math.floor((numericValue - 8) / 5) * 5;
}

function getProduccionSugerida(producto, value) {
  if (isOperationalCakeProduct(producto)) {
    return getProduccionSugeridaPastel(value);
  }
  return Math.max(0, Math.ceil(Number(value) || 0));
}

function getReglaOperativaLabel(producto, value) {
  if (!isOperationalCakeProduct(producto)) return "Redondeo normal";
  const produccionSugerida = getProduccionSugeridaPastel(value);
  if (produccionSugerida === 0) return "Menor a 8: no producir";
  return `Mínimo 10 y múltiplos de 5: ${produccionSugerida}`;
}

const WEEKDAY_BY_NORM = new Map(
  WEEKDAYS.flatMap((day) => [
    [norm(day.label), day.index],
    [norm(day.label).slice(0, 3), day.index],
  ])
);

function weekdayIndexFromText(value) {
  const normalized = norm(value);
  if (!normalized) return null;
  if (WEEKDAY_BY_NORM.has(normalized)) return WEEKDAY_BY_NORM.get(normalized);
  for (const alias of WEEKDAY_ALIASES) {
    if (
      alias.names.some(
        (name) =>
          normalized === name ||
          normalized.startsWith(`${name} `) ||
          normalized.startsWith(`${name}-`) ||
          normalized.startsWith(`${name}/`) ||
          (name.length > 3 && normalized.startsWith(name)) ||
          normalized.includes(` ${name} `)
      )
    ) {
      return alias.index;
    }
  }
  return null;
}

function getWeekdayAverage(row, weekday) {
  switch (norm(weekday)) {
    case "LUNES":
      return Number(row.promedioLunes || 0);
    case "MARTES":
      return Number(row.promedioMartes || 0);
    case "MIERCOLES":
      return Number(row.promedioMiercoles || 0);
    case "JUEVES":
      return Number(row.promedioJueves || 0);
    case "VIERNES":
      return Number(row.promedioViernes || 0);
    case "SABADO":
      return Number(row.promedioSabado || 0);
    case "DOMINGO":
      return Number(row.promedioDomingo || 0);
    default:
      return 0;
  }
}

function isDateLikeValue(value) {
  if (value instanceof Date) return true;
  if (typeof value === "number") return value > 20000 && value < 80000;
  const raw = String(value ?? "").trim();
  const p = norm(raw);
  if (!p) return false;
  if (/^\d{1,2}[/-]\d{1,2}[/-]\d{2,4}$/.test(raw)) return true;
  if (/^\d{4}[/-]\d{1,2}[/-]\d{1,2}$/.test(raw)) return true;
  if (/^(MON|TUE|WED|THU|FRI|SAT|SUN)\s+[A-Z]{3}\s+\d{1,2}\s+\d{4}/.test(p)) return true;
  if (/\b(ENE|FEB|MAR|ABR|MAY|JUN|JUL|AGO|SEP|OCT|NOV|DIC|JAN|APR|AUG|DEC)\b/.test(p) && /\b(19|20)\d{2}\b/.test(p)) {
    return true;
  }
  return false;
}

function looksLikeCalendarHeader(value) {
  const p = norm(value);
  if (!p) return false;
  if (CALENDAR_WORDS.has(p)) return true;
  if (p.includes("CALENDARIO") || p.includes("SEMANA") || p.includes("FECHA")) return true;
  return false;
}

function isCalendarRow(row = []) {
  const nonEmpty = row.filter((cell) => String(cell ?? "").trim() !== "");
  if (nonEmpty.length < 3) return false;
  const calendarCells = nonEmpty.filter((cell) => isDateLikeValue(cell) || looksLikeCalendarHeader(cell)).length;
  return calendarCells >= 3 && calendarCells / nonEmpty.length >= 0.5;
}

function isValidProduct(value, row = []) {
  if (isDateLikeValue(value) || isCalendarRow(row)) return false;
  const p = norm(value);
  if (!p) return false;
  if (INVALID_PRODUCTS.has(p)) return false;
  if (p.startsWith("TOTAL")) return false;
  if (p.includes("PRODUCTO") || p.includes("ESPECIALIDAD")) return false;
  if (looksLikeCalendarHeader(p)) return false;
  if (/^\d+$/.test(p)) return false;
  return true;
}

function toNumber(value) {
  if (typeof value === "number") return Number.isFinite(value) ? value : 0;
  const cleaned = String(value ?? "")
    .replace(/\s/g, "")
    .replace(/\$/g, "")
    .replace(/,/g, "");
  const n = Number(cleaned);
  return Number.isFinite(n) ? n : 0;
}

function formatNumber(value, digits = 0) {
  return new Intl.NumberFormat("es-MX", {
    maximumFractionDigits: digits,
    minimumFractionDigits: digits,
  }).format(Number.isFinite(value) ? value : 0);
}

function formatPercent(value, digits = 0) {
  if (!Number.isFinite(value)) return "0%";
  return `${formatNumber(value, digits)}%`;
}

function parseDateCell(value) {
  if (value instanceof Date && !Number.isNaN(value.getTime())) return value;
  if (typeof value === "number" && value > 20000 && value < 80000) {
    const utcDays = Math.floor(value - 25569);
    return new Date(utcDays * 86400 * 1000);
  }
  const raw = String(value ?? "").trim();
  if (!raw) return null;
  if (/^\d{1,2}$/.test(raw)) return null;
  const parsed = new Date(raw);
  return Number.isNaN(parsed.getTime()) ? null : parsed;
}

function parseDayNumber(value) {
  if (value instanceof Date && !Number.isNaN(value.getTime())) return value.getDate();
  const raw = String(value ?? "").trim();
  if (!/^\d{1,2}$/.test(raw)) return null;
  const day = Number(raw);
  return day >= 1 && day <= 31 ? day : null;
}

function inferMonthYearFromWideHeaders(weekdayHeaders = [], dateHeaders = []) {
  const anchors = [];
  for (let c = 1; c < Math.max(weekdayHeaders.length, dateHeaders.length); c++) {
    const weekday = weekdayIndexFromText(weekdayHeaders[c]);
    const day = parseDayNumber(dateHeaders[c]);
    if (weekday !== null && day !== null) anchors.push({ weekday, day });
  }
  if (!anchors.length) return null;

  const currentYear = new Date().getFullYear();
  const candidateYears = [...new Set([currentYear - 1, currentYear, currentYear + 1, 2025, 2026, 2027])];
  let best = null;
  for (const year of candidateYears) {
    for (let monthIndex = 0; monthIndex < 12; monthIndex++) {
      const daysInMonth = new Date(year, monthIndex + 1, 0).getDate();
      let score = 0;
      for (const anchor of anchors) {
        if (anchor.day > daysInMonth) continue;
        if (new Date(year, monthIndex, anchor.day).getDay() === anchor.weekday) score += 1;
      }
      if (!best || score > best.score) best = { year, monthIndex, score };
    }
  }
  return best && best.score > 0 ? best : null;
}

function dateKey(date) {
  const d = parseDateCell(date);
  if (!d) return "";
  const year = d.getFullYear();
  const month = String(d.getMonth() + 1).padStart(2, "0");
  const day = String(d.getDate()).padStart(2, "0");
  return `${year}-${month}-${day}`;
}

function displayDate(date) {
  const key = dateKey(date);
  if (!key) return "";
  const [year, month, day] = key.split("-");
  return `${day}/${month}/${year}`;
}

function weekdayLabel(index) {
  return WEEKDAYS.find((day) => day.index === index)?.label || "";
}

function defaultMonthValue() {
  const today = new Date();
  return `${today.getFullYear()}-${String(today.getMonth() + 1).padStart(2, "0")}`;
}

function addDays(date, days) {
  const next = new Date(date);
  next.setDate(next.getDate() + days);
  return next;
}

function datesForMonth(monthValue) {
  const [year, month] = String(monthValue || defaultMonthValue())
    .split("-")
    .map(Number);
  if (!year || !month) return [];
  const dates = [];
  const cursor = new Date(year, month - 1, 1);
  while (cursor.getMonth() === month - 1) {
    dates.push(new Date(cursor));
    cursor.setDate(cursor.getDate() + 1);
  }
  return dates;
}

function detectDominantMonth(records) {
  const counts = new Map();
  for (const record of records) {
    const key = dateKey(record.fecha);
    if (!key) continue;
    const monthKey = key.slice(0, 7);
    counts.set(monthKey, (counts.get(monthKey) || 0) + 1);
  }
  let bestMonth = "";
  let bestCount = 0;
  for (const [monthKey, count] of counts.entries()) {
    if (count > bestCount) {
      bestMonth = monthKey;
      bestCount = count;
    }
  }
  return bestMonth;
}

function recordWeekday(record) {
  const weekdayFromHeader = weekdayIndexFromText(record.weekday);
  if (weekdayFromHeader !== null) return weekdayFromHeader;
  const parsedHeaderDate = parseDateCell(record.weekday);
  if (parsedHeaderDate) return parsedHeaderDate.getDay();
  const parsedDate = parseDateCell(record.fecha);
  if (parsedDate) return parsedDate.getDay();
  return null;
}

function horizonWeekendFactor(days, weekendBoost) {
  const today = new Date();
  let factor = 0;
  for (let i = 0; i < Math.max(1, days); i++) {
    const day = addDays(today, i).getDay();
    factor += [0, 6].includes(day) ? weekendBoost : 1;
  }
  return factor / Math.max(1, days);
}

async function readWorkbook(file) {
  const data = await file.arrayBuffer();
  return XLSX.read(data, { type: "array", cellDates: true });
}

function rowsFromFirstSheet(workbook) {
  const sheet = workbook.Sheets[workbook.SheetNames[0]];
  return XLSX.utils.sheet_to_json(sheet, { header: 1, defval: "" });
}

function parseStock(workbook) {
  const rows = rowsFromFirstSheet(workbook);
  const parsed = [];
  for (let i = 0; i < rows.length; i++) {
    if (!isValidProduct(rows[i][0], rows[i])) continue;
    const productoOriginal = String(rows[i][0] ?? "").trim();
    const product = normalizeProduct(productoOriginal);
    const stock = toNumber(rows[i][1]);
    parsed.push({ producto: product, productoOriginal, stock, orden: parsed.length + 1 });
  }
  return parsed;
}

function parseExistencias(workbook) {
  const rows = rowsFromFirstSheet(workbook);

  let headerIndex = -1;
  let productCol = -1;
  let totalSucCol = -1;
  let cfCol = -1;
  let sumaCol = -1;

  for (let i = 0; i < Math.min(rows.length, 15); i++) {
    const row = rows[i].map(norm);
    const p = row.findIndex((x) => x.includes("PRODUCTO"));
    const total = row.findIndex((x) => x.includes("TOTAL") && (x.includes("GRAL") || x.includes("SUC")));
    const cf = row.findIndex((x) => x.includes("CUARTO") || x === "C.F." || x === "CF");
    const suma = row.findIndex((x) => x.includes("SUMA") && x.includes("SUC"));
    if (p >= 0 && total >= 0 && cf >= 0 && suma >= 0) {
      headerIndex = i;
      productCol = p;
      totalSucCol = total;
      cfCol = cf;
      sumaCol = suma;
      break;
    }
  }

  if (headerIndex < 0) return [];

  const parsed = [];
  for (let i = headerIndex + 1; i < rows.length; i++) {
    if (!isValidProduct(rows[i][productCol], rows[i])) continue;
    const productoOriginal = String(rows[i][productCol] ?? "").trim();
    const product = normalizeProduct(productoOriginal);
    parsed.push({
      producto: product,
      productoOriginal,
      totalSuc: toNumber(rows[i][totalSucCol]),
      cf: toNumber(rows[i][cfCol]),
      sumaSucCf: toNumber(rows[i][sumaCol]),
    });
  }
  return parsed;
}

function parseMonthlyDailySheets(workbook, type = "ventas") {
  const out = [];
  const skip = new Set(["RESUMEN", "REPORTE", "TOTAL", "TOTALES", "CONCENTRADO", "HOJA1"]);
  for (const sheetName of workbook.SheetNames) {
    if (skip.has(norm(sheetName))) continue;

    const sheet = workbook.Sheets[sheetName];
    const rows = XLSX.utils.sheet_to_json(sheet, { header: 1, defval: "" });
    if (!rows.length) continue;

    let fecha = new Date(sheetName);
    const dayOnly = String(sheetName).match(/\d{1,2}/);
    if (Number.isNaN(fecha.getTime()) && dayOnly) fecha = new Date(2026, 0, Number(dayOnly[0]));
    if (Number.isNaN(fecha.getTime())) fecha = new Date();

    let headerIndex = -1;
    let productCol = -1;
    let qtyCol = -1;
    let amountCol = -1;

    for (let i = 0; i < Math.min(rows.length, 12); i++) {
      const row = rows[i].map(norm);
      const p = row.findIndex((x) => x.includes("PRODUCTO") || x.includes("DESCRIPCION") || x.includes("ARTICULO"));
      const q = row.findIndex((x) => x.includes("CANT") || x.includes("VENTA") || x.includes("PIEZAS") || x.includes("UNIDADES"));
      const a = row.findIndex((x) => x.includes("IMPORTE") || x.includes("TOTAL"));
      if (p >= 0 && q >= 0) {
        headerIndex = i;
        productCol = p;
        qtyCol = q;
        amountCol = a;
        break;
      }
    }

    if (headerIndex < 0 && rows[0]?.length >= 2) {
      headerIndex = 0;
      productCol = 0;
      qtyCol = 1;
      amountCol = 2;
    }

    for (let i = headerIndex + 1; i < rows.length; i++) {
      if (!isValidProduct(rows[i][productCol], rows[i])) continue;
      const productoOriginal = String(rows[i][productCol] ?? "").trim();
      const product = normalizeProduct(productoOriginal);
      const rawCantidad = rows[i][qtyCol];
      if (type === "ventas" && String(rawCantidad ?? "").trim() === "") continue;
      const cantidad = toNumber(rawCantidad);
      const importe = amountCol >= 0 ? toNumber(rows[i][amountCol]) : 0;
      if (type !== "ventas" && cantidad === 0 && importe === 0) continue;
      out.push({ fecha, producto: product, productoOriginal, cantidad, importe, tipo: type });
    }
  }
  return out;
}

function parseWideSales(workbook, type = "ventas") {
  const rows = rowsFromFirstSheet(workbook);
  if (rows.length < 4) return [];

  let weekdayHeaderIndex = 0;
  let bestWeekdayScore = -1;
  let dateHeaderIndex = -1;
  let bestDateScore = 0;
  const topRows = Math.min(rows.length, 4);
  for (let i = 0; i < topRows; i++) {
    let weekdayScore = 0;
    let dateScore = 0;
    for (let c = 1; c < rows[i].length; c++) {
      if (weekdayIndexFromText(rows[i][c]) !== null) weekdayScore += 1;
      if (parseDateCell(rows[i][c]) || parseDayNumber(rows[i][c]) !== null) dateScore += 1;
    }
    if (weekdayScore > bestWeekdayScore) {
      bestWeekdayScore = weekdayScore;
      weekdayHeaderIndex = i;
    }
    if (dateScore > bestDateScore) {
      bestDateScore = dateScore;
      dateHeaderIndex = i;
    }
  }

  const weekdayHeaders = rows[weekdayHeaderIndex] || [];
  const dateHeaders = dateHeaderIndex >= 0 ? rows[dateHeaderIndex] || [] : [];
  const inferredMonth = inferMonthYearFromWideHeaders(weekdayHeaders, dateHeaders);
  const parsed = [];
  const startRow = Math.max(weekdayHeaderIndex, dateHeaderIndex) + 1;
  for (let r = startRow; r < rows.length; r++) {
    if (!isValidProduct(rows[r][0], rows[r])) continue;
    const productoOriginal = String(rows[r][0] ?? "").trim();
    const product = normalizeProduct(productoOriginal);
    const hasDailyValue = rows[r].some(
      (cell, c) => c > 0 && weekdayIndexFromText(weekdayHeaders[c]) !== null && String(cell ?? "").trim() !== ""
    );
    if (!hasDailyValue) continue;
    for (let c = 1; c < rows[r].length; c++) {
      const weekdayHeader = weekdayHeaders[c];
      const weekday = weekdayIndexFromText(weekdayHeader);
      if (weekday === null) continue;
      const parsedHeaderDate = parseDateCell(dateHeaders[c]);
      const dayNumber = parseDayNumber(dateHeaders[c]);
      const fecha =
        parsedHeaderDate ||
        (inferredMonth && dayNumber ? new Date(inferredMonth.year, inferredMonth.monthIndex, dayNumber) : new Date(2026, 0, c));
      const cantidad = toNumber(rows[r][c]);
      if (type !== "ventas" && cantidad === 0) continue;
      parsed.push({
        fecha,
        producto: product,
        productoOriginal,
        cantidad,
        importe: 0,
        weekday: weekdayHeader,
        tipo: type,
      });
    }
  }
  return parsed;
}

function parseSalesOrReturns(workbook, type) {
  const wide = parseWideSales(workbook, type);
  if (wide.length > 0) return wide;
  const bySheets = parseMonthlyDailySheets(workbook, type);
  if (bySheets.length > 0) return bySheets;
  return [];
}

function parseProductionReal(workbook) {
  const rows = rowsFromFirstSheet(workbook);
  let headerIndex = -1;
  let productCol = 0;
  let qtyCol = 1;
  let dateCol = -1;

  for (let i = 0; i < Math.min(rows.length, 15); i++) {
    const row = rows[i].map(norm);
    const p = row.findIndex((x) => x.includes("PRODUCTO") || x.includes("DESCRIPCION") || x.includes("ARTICULO"));
    const d = row.findIndex((x) => x === "FECHA" || x.includes("FECHA") || x === "DIA" || x === "DÍA");
    const q = row.findIndex(
      (x) =>
        x.includes("PRODUCCION") ||
        x.includes("REAL") ||
        x.includes("CANT") ||
        x.includes("PIEZAS") ||
        x.includes("UNIDADES")
    );
    if (p >= 0 && q >= 0 && p !== q) {
      headerIndex = i;
      productCol = p;
      qtyCol = q;
      dateCol = d;
      break;
    }
  }

  const start = headerIndex >= 0 ? headerIndex + 1 : 0;
  const records = [];
  for (let i = start; i < rows.length; i++) {
    if (!isValidProduct(rows[i][productCol], rows[i])) continue;
    const productoOriginal = String(rows[i][productCol] ?? "").trim();
    const product = normalizeProduct(productoOriginal);
    const cantidad = toNumber(rows[i][qtyCol]);
    if (cantidad === 0) continue;
    const fecha = dateCol >= 0 ? parseDateCell(rows[i][dateCol]) : null;
    records.push({
      producto: product,
      productoOriginal,
      cantidad,
      fecha,
      fechaKey: fecha ? dateKey(fecha) : "",
    });
  }
  return records;
}

function aggregateProductionRows(records) {
  const map = new Map();
  for (const item of records) {
    if (!isValidProduct(item.producto)) continue;
    const product = normalizeProduct(item.producto);
    if (isSliceProduct(product)) continue;
    map.set(product, (map.get(product) || 0) + toNumber(item.cantidad));
  }
  return [...map.entries()].map(([producto, cantidad]) => ({ producto, cantidad }));
}

function aggregateDailyProductionRows(records) {
  const map = new Map();
  for (const item of records) {
    if (!item.fechaKey || !isValidProduct(item.producto)) continue;
    const product = normalizeProduct(item.producto);
    if (isSliceProduct(product)) continue;
    const key = `${product}|${item.fechaKey}`;
    map.set(key, (map.get(key) || 0) + toNumber(item.cantidad));
  }
  return map;
}

async function parseProductionRealFile(file) {
  if (!file) return [];
  const lowerName = file.name.toLowerCase();
  if (lowerName.endsWith(".zip")) {
    const zip = await JSZip.loadAsync(await file.arrayBuffer());
    const records = [];
    for (const entry of Object.values(zip.files)) {
      if (entry.dir || !/\.(xlsx|xls)$/i.test(entry.name)) continue;
      const data = await entry.async("arraybuffer");
      const workbook = XLSX.read(data, { type: "array", cellDates: true });
      records.push(...parseProductionReal(workbook));
    }
    return records;
  }

  const workbook = await readWorkbook(file);
  return parseProductionReal(workbook);
}

function groupByProduct(records) {
  const map = new Map();
  for (const r of records) {
    const product = normalizeProduct(r.producto);
    if (isSliceProduct(product)) continue;
    const current = map.get(product) || [];
    current.push(r);
    map.set(product, current);
  }
  return map;
}

function calculateForecast({ stockRows, ventas, bajas, existencias, realProduction, days, weekendBoost }) {
  const ventasByProduct = groupByProduct(ventas);
  const bajasByProduct = groupByProduct(bajas);
  const existMap = new Map(existencias.map((e) => [e.producto, e]));
  const realMap = new Map(aggregateProductionRows(realProduction).map((e) => [e.producto, e.cantidad]));
  const horizonFactor = horizonWeekendFactor(days, weekendBoost);

  return stockRows.filter((s) => !isSliceProduct(s.producto)).map((s) => {
    const v = ventasByProduct.get(s.producto) || [];
    const b = bajasByProduct.get(s.producto) || [];

    const values = v.map((x) => x.cantidad);
    const productWeekdayAverages = calculateProductWeekdayAverages(v);
    const recentValues = values.slice(-28);
    const promedioReciente = recentValues.length ? recentValues.reduce((a, n) => a + n, 0) / recentValues.length : 0;
    const promedioHistorico = values.length ? values.reduce((a, n) => a + n, 0) / values.length : 0;
    const promedioDiario = promedioReciente * 0.65 + promedioHistorico * 0.35;
    const pronosticoVenta = promedioDiario * days * horizonFactor;

    const bajasTotal = b.reduce((a, n) => a + n.cantidad, 0);
    const ventasTotal = v.reduce((a, n) => a + n.cantidad, 0);
    const tasaBajas = ventasTotal > 0 ? bajasTotal / ventasTotal : 0;
    const bajasEsperadas = pronosticoVenta * tasaBajas;
    const colchonOperativo = pronosticoVenta > 0 ? (pronosticoVenta >= 300 ? 15 : 10) : 0;
    const baseConColchon = pronosticoVenta + colchonOperativo;
    const produccionSugerida = getProduccionSugerida(s.producto, baseConColchon);

    const ex = existMap.get(s.producto) || { totalSuc: 0, cf: 0, sumaSucCf: 0 };
    const sumaSucCf = ex.sumaSucCf || ex.totalSuc + ex.cf;
    const inventarioObjetivo = s.stock;
    const baseProduccionRecomendada = Math.max(0, inventarioObjetivo + produccionSugerida - sumaSucCf);
    const produccionRecomendada = getProduccionSugerida(s.producto, baseProduccionRecomendada);
    const hasRealData = realMap.has(s.producto);
    const produccionReal = hasRealData ? realMap.get(s.producto) : 0;
    const diferenciaReal = produccionReal - produccionSugerida;
    const precision =
      hasRealData && produccionReal > 0
        ? (1 - Math.abs(produccionSugerida - produccionReal) / produccionReal) * 100
        : null;

    const confianza =
      promedioHistorico > 0
        ? Math.max(0, Math.min(100, 100 - (Math.abs(promedioReciente - promedioHistorico) / promedioHistorico) * 100))
        : values.length > 0
          ? 50
          : 0;

    let estatus = "Sin dato real";
    if (!hasRealData) estatus = "Sin dato real";
    else if (produccionSugerida === 0 && produccionReal === 0) estatus = "No producir";
    else if (produccionReal < produccionSugerida) estatus = "Riesgo faltante";
    else if (produccionReal > produccionSugerida) estatus = "Sobreproduccion";
    else if (precision !== null && precision < 80) estatus = "Revisar";
    else estatus = "Dentro de rango";

    return {
      producto: s.producto,
      orden: s.orden,
      promedioReciente,
      promedioHistorico,
      promedioDiario,
      promedioLunes: Number(productWeekdayAverages.get(1) || 0),
      promedioMartes: Number(productWeekdayAverages.get(2) || 0),
      promedioMiercoles: Number(productWeekdayAverages.get(3) || 0),
      promedioJueves: Number(productWeekdayAverages.get(4) || 0),
      promedioViernes: Number(productWeekdayAverages.get(5) || 0),
      promedioSabado: Number(productWeekdayAverages.get(6) || 0),
      promedioDomingo: Number(productWeekdayAverages.get(0) || 0),
      demandaPronosticada: pronosticoVenta,
      pronosticoVenta,
      tasaBajas,
      bajasEsperadas,
      colchonOperativo,
      baseConColchon,
      reglaOperativa: getReglaOperativaLabel(s.producto, baseConColchon),
      produccionSugerida,
      inventarioObjetivo,
      baseProduccionRecomendada,
      totalSuc: ex.totalSuc || 0,
      cf: ex.cf || 0,
      sumaSucCf,
      produccionRecomendada,
      produccionReal,
      hasRealData,
      diferenciaReal,
      precision,
      confianza,
      estatus,
    };
  });
}

function calculateProductWeekdayAverages(records) {
  const buckets = new Map();
  for (const record of records) {
    const weekday = recordWeekday(record);
    if (weekday === null) continue;
    const bucket = buckets.get(weekday) || { total: 0, count: 0 };
    bucket.total += toNumber(record.cantidad);
    bucket.count += 1;
    buckets.set(weekday, bucket);
  }

  const averages = new Map();
  for (const [weekday, bucket] of buckets.entries()) {
    averages.set(weekday, bucket.count ? bucket.total / bucket.count : 0);
  }
  return averages;
}

function calculateDailyForecast({ monthlyRows, ventas, realProduction, selectedMonth, dailyBufferPct }) {
  const realDailyMap = aggregateDailyProductionRows(realProduction);
  const monthDates = datesForMonth(selectedMonth);
  const productRows = monthlyRows.filter((row) => isValidProduct(row.producto) && !isSliceProduct(row.producto));

  return productRows.flatMap((productRow) => {
    const product = productRow.producto;
    return monthDates.map((date) => {
      const key = dateKey(date);
      const weekday = date.getDay();
      const dayName = weekdayLabel(weekday);
      const pronosticoVentaDia = getWeekdayAverage(productRow, dayName);
      const colchonDiario = pronosticoVentaDia * (dailyBufferPct / 100);
      const baseConColchonDia = pronosticoVentaDia + colchonDiario;
      const produccionSugeridaDia = getProduccionSugerida(product, baseConColchonDia);
      const realKey = `${product}|${key}`;
      const hasRealData = realDailyMap.has(realKey);
      const produccionRealDia = hasRealData ? realDailyMap.get(realKey) : null;
      const diferenciaPiezas = hasRealData ? produccionRealDia - produccionSugeridaDia : null;

      let estatus = "Sin dato real";
      if (hasRealData && diferenciaPiezas < 0) estatus = "Riesgo faltante";
      else if (hasRealData && diferenciaPiezas > 0) estatus = "Sobreproduccion";
      else if (hasRealData) estatus = "Dentro de rango";

      return {
        fecha: key,
        fechaDisplay: displayDate(date),
        weekday,
        dia: weekdayLabel(weekday),
        producto: product,
        promedioUsado: pronosticoVentaDia,
        pronosticoVentaDia,
        colchonDiario,
        baseConColchonDia,
        reglaOperativa: getReglaOperativaLabel(product, baseConColchonDia),
        produccionSugeridaDia,
        produccionRealDia,
        hasRealData,
        diferenciaPiezas,
        estatus,
      };
    });
  });
}

function summarizeDailyMonth(rows) {
  const pronosticoVentaMensual = rows.reduce((sum, row) => sum + row.pronosticoVentaDia, 0);
  const colchonDiarioMensual = rows.reduce((sum, row) => sum + row.colchonDiario, 0);
  const baseConColchonMensual = rows.reduce((sum, row) => sum + row.baseConColchonDia, 0);
  const produccionSugeridaMensual = rows.reduce((sum, row) => sum + row.produccionSugeridaDia, 0);
  const produccionRealMensual = rows.reduce((sum, row) => sum + (row.produccionRealDia || 0), 0);
  const diferenciaMensual = produccionRealMensual - produccionSugeridaMensual;
  const precision =
    produccionRealMensual > 0
      ? (1 - Math.abs(produccionSugeridaMensual - produccionRealMensual) / produccionRealMensual) * 100
      : 0;

  return {
    pronosticoVentaMensual,
    colchonDiarioMensual,
    baseConColchonMensual,
    produccionSugeridaMensual,
    produccionRealMensual,
    diferenciaMensual,
    precision,
  };
}

function exportToExcel(rows, summary) {
  const detalle = rows.map((r) => ({
    Producto: r.producto,
    "Promedio reciente": Number(r.promedioReciente.toFixed(2)),
    "Promedio historico": Number(r.promedioHistorico.toFixed(2)),
    "Pronostico venta": Number(r.pronosticoVenta.toFixed(2)),
    "Colchon operativo": r.colchonOperativo,
    "Base con colchon": Number((r.baseConColchon || 0).toFixed(2)),
    "Regla operativa": r.reglaOperativa,
    "Produccion sugerida": r.produccionSugerida,
    "Precision %": r.precision === null ? "" : Number(r.precision.toFixed(1)),
    Confianza: Number(r.confianza.toFixed(1)),
    Estatus: STATUS_META[r.estatus]?.label || r.estatus,
  }));

  const resumen = [
    { Indicador: "Produccion sugerida con regla operativa", Valor: summary.totalPronosticada },
    { Indicador: "Produccion recomendada", Valor: summary.totalRecomendada },
    { Indicador: "Produccion real", Valor: summary.totalReal },
    { Indicador: "Brecha real vs sugerida", Valor: summary.brechaTotal },
    { Indicador: "Precision ejecutiva", Valor: `${summary.precisionEjecutiva.toFixed(1)}%` },
    { Indicador: "Productos en riesgo", Valor: summary.riesgoFaltante },
    { Indicador: "Productos con sobreproduccion", Valor: summary.sobreproduccion },
  ];

  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(resumen), "Dashboard");
  XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(detalle), "Detalle");
  XLSX.writeFile(wb, "dashboard_produccion_archivo_maestro.xlsx");
}

function exportDailyToExcel(rows, summary) {
  const resumen = [
    { Indicador: "Pronostico venta mensual", Valor: Number(summary.pronosticoVentaMensual.toFixed(2)) },
    { Indicador: "Colchon aplicado mensual", Valor: Number(summary.colchonDiarioMensual.toFixed(2)) },
    { Indicador: "Base con colchon mensual", Valor: Number(summary.baseConColchonMensual.toFixed(2)) },
    { Indicador: "Produccion sugerida mensual", Valor: summary.produccionSugeridaMensual },
    { Indicador: "Produccion real mensual", Valor: summary.produccionRealMensual },
    { Indicador: "Diferencia mensual", Valor: summary.diferenciaMensual },
    { Indicador: "Precision %", Valor: Number(summary.precision.toFixed(1)) },
  ];

  const detalle = rows.map((row) => ({
    Fecha: row.fechaDisplay,
    Dia: row.dia,
    Producto: row.producto,
    "Promedio aplicado": Number(row.promedioUsado.toFixed(2)),
    "Pronostico venta dia": Number(row.pronosticoVentaDia.toFixed(2)),
    "Colchon diario": Number(row.colchonDiario.toFixed(2)),
    "Base con colchon": Number((row.baseConColchonDia || 0).toFixed(2)),
    "Regla operativa": row.reglaOperativa,
    "Produccion sugerida dia": row.produccionSugeridaDia,
    "Produccion real dia": row.produccionRealDia ?? "",
    "Diferencia piezas": row.diferenciaPiezas ?? "",
    Estatus: STATUS_META[row.estatus]?.label || row.estatus,
  }));

  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(resumen), "Resumen mensual");
  XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(detalle), "Pronostico diario");
  XLSX.writeFile(wb, "produccion_diaria_sugerida.xlsx");
}

function UploadBox({ title, description, onFile, fileName, required, accept = ".xlsx,.xls" }) {
  return (
    <div className="upload-card">
      <div className="upload-heading">
        <div className="upload-icon">
          <Upload size={20} />
        </div>
        {required && <span className="tag">Base</span>}
      </div>
      <h3>{title}</h3>
      <p>{description}</p>
      <label className="upload-button">
        <FileSpreadsheet size={17} />
        Seleccionar Excel
        <input type="file" accept={accept} onChange={(e) => onFile(e.target.files?.[0])} />
      </label>
      {fileName && <span className="file-name">{fileName}</span>}
    </div>
  );
}

function KpiCard({ icon: Icon, label, value, tone, caption }) {
  return (
    <div className={`kpi-card ${tone || ""}`}>
      <div className="kpi-icon">
        <Icon size={21} />
      </div>
      <span>{label}</span>
      <strong>{value}</strong>
      {caption && <small>{caption}</small>}
    </div>
  );
}

function App() {
  const [stockRows, setStockRows] = useState([]);
  const [ventas, setVentas] = useState([]);
  const [bajas, setBajas] = useState([]);
  const [existencias, setExistencias] = useState([]);
  const [realProduction, setRealProduction] = useState([]);
  const [files, setFiles] = useState({});
  const [query, setQuery] = useState("");
  const [weekendBoost, setWeekendBoost] = useState(1.15);
  const [days, setDays] = useState(30);
  const [showMissingReal, setShowMissingReal] = useState(false);
  const [selectedMonth, setSelectedMonth] = useState(defaultMonthValue());
  const [selectedMonthTouched, setSelectedMonthTouched] = useState(false);
  const [dailyBufferPct, setDailyBufferPct] = useState(0);
  const [dailyDateFilter, setDailyDateFilter] = useState("");
  const [dailyProductQuery, setDailyProductQuery] = useState("");
  const [dailyWeekdayFilter, setDailyWeekdayFilter] = useState("");
  const [onlyDailyShortage, setOnlyDailyShortage] = useState(false);
  const [onlyDailyOverproduction, setOnlyDailyOverproduction] = useState(false);
  const [validationProduct, setValidationProduct] = useState("");

  async function handleFile(file, parser, key, setter) {
    if (!file) return;
    const wb = await readWorkbook(file);
    setter(parser(wb));
    setFiles((f) => ({ ...f, [key]: file.name }));
  }

  async function handleProductionReal(file) {
    if (!file) return;
    const parsed = await parseProductionRealFile(file);
    setRealProduction(parsed);
    setFiles((f) => ({ ...f, real: file.name }));
  }

  useEffect(() => {
    if (selectedMonthTouched || !ventas.length) return;
    const detectedMonth = detectDominantMonth(ventas);
    if (detectedMonth) setSelectedMonth(detectedMonth);
  }, [ventas, selectedMonthTouched]);

  const forecast = useMemo(
    () =>
      calculateForecast({
        stockRows,
        ventas,
        bajas,
        existencias,
        realProduction,
        days,
        weekendBoost,
      }),
    [stockRows, ventas, bajas, existencias, realProduction, days, weekendBoost]
  );

  const comparableForecast = showMissingReal ? forecast : forecast.filter((r) => r.hasRealData);
  const filtered = comparableForecast.filter((r) => r.producto.includes(norm(query)));

  const dailyRows = useMemo(
    () =>
      calculateDailyForecast({
        monthlyRows: forecast,
        ventas,
        realProduction,
        selectedMonth,
        dailyBufferPct,
      }),
    [forecast, ventas, realProduction, selectedMonth, dailyBufferPct]
  );

  const filteredDailyRows = dailyRows.filter((row) => {
    if (dailyDateFilter && row.fecha !== dailyDateFilter) return false;
    if (dailyProductQuery && !row.producto.includes(norm(dailyProductQuery))) return false;
    if (dailyWeekdayFilter !== "" && row.weekday !== Number(dailyWeekdayFilter)) return false;
    if (onlyDailyShortage && row.estatus !== "Riesgo faltante") return false;
    if (onlyDailyOverproduction && row.estatus !== "Sobreproduccion") return false;
    return true;
  });

  const dailySummary = useMemo(() => summarizeDailyMonth(dailyRows), [dailyRows]);

  const validationProducts = useMemo(
    () => [...forecast].sort((a, b) => a.producto.localeCompare(b.producto, "es")),
    [forecast]
  );

  useEffect(() => {
    if (!validationProducts.length) {
      setValidationProduct("");
      return;
    }
    if (validationProducts.some((row) => row.producto === validationProduct)) return;
    const example =
      validationProducts.find((row) => normalizeProduct(row.producto) === "PINA GDE") ||
      validationProducts[0];
    setValidationProduct(example.producto);
  }, [validationProducts, validationProduct]);

  const validationForecast = forecast.find((row) => row.producto === validationProduct) || null;
  const validationSales = ventas
    .filter((row) => normalizeProduct(row.producto) === normalizeProduct(validationProduct))
    .map((row) => {
      const weekday = recordWeekday(row);
      return {
        ...row,
        fechaDisplay: displayDate(row.fecha),
        dia: weekday === null ? norm(row.weekday) || "Sin día" : weekdayLabel(weekday),
      };
    });
  const validationThursdaySales = validationSales.filter((sale) => recordWeekday(sale) === 4);
  const validationThursdayAverage = validationThursdaySales.length
    ? validationThursdaySales.reduce((sum, sale) => sum + toNumber(sale.cantidad), 0) / validationThursdaySales.length
    : 0;
  const validationSourceNames = [...new Set(validationSales.map((sale) => sale.productoOriginal || sale.producto))];
  const validationDailyRows = dailyRows.filter(
    (row) => normalizeProduct(row.producto) === normalizeProduct(validationProduct)
  );
  const validationSummary = summarizeDailyMonth(validationDailyRows);
  const validationWeekdayAverages = WEEKDAYS.map((day) => ({
    ...day,
    value: validationForecast ? getWeekdayAverage(validationForecast, day.label) : 0,
    registros: validationSales.filter((sale) => recordWeekday(sale) === day.index).length,
  }));

  const summary = useMemo(() => {
    const totalPronosticada = comparableForecast.reduce((a, r) => a + r.produccionSugerida, 0);
    const totalRecomendada = comparableForecast.reduce((a, r) => a + r.produccionRecomendada, 0);
    const totalReal = comparableForecast.reduce((a, r) => a + r.produccionReal, 0);
    const totalColchon = comparableForecast.reduce((a, r) => a + r.colchonOperativo, 0);
    const brechaTotal = totalReal - totalPronosticada;
    const precisionEjecutiva = totalReal > 0 ? (1 - Math.abs(totalPronosticada - totalReal) / totalReal) * 100 : 0;
    const confianza = comparableForecast.length ? comparableForecast.reduce((a, r) => a + r.confianza, 0) / comparableForecast.length : 0;
    const riesgoFaltante = comparableForecast.filter((r) => r.estatus === "Riesgo faltante").length;
    const sobreproduccion = comparableForecast.filter((r) => r.estatus === "Sobreproduccion").length;
    const sinDatoReal = forecast.filter((r) => r.estatus === "Sin dato real").length;
    return {
      totalPronosticada,
      totalRecomendada,
      totalReal,
      totalColchon,
      brechaTotal,
      precisionEjecutiva,
      confianza,
      riesgoFaltante,
      sobreproduccion,
      sinDatoReal,
    };
  }, [comparableForecast, forecast]);

  const loadedFileItems = [
    { label: "Ventas", loaded: Boolean(files.ventas || ventas.length) },
    { label: "Stock fijo", loaded: Boolean(files.stock || stockRows.length) },
    { label: "Producción real", loaded: Boolean(files.real || realProduction.length) },
    { label: "Bajas/devoluciones", loaded: Boolean(files.bajas || bajas.length) },
    { label: "Existencias", loaded: Boolean(files.existencias || existencias.length) },
  ];

  return (
    <div className="app">
      <aside className="sidebar">
        <div className="brand">
          <div className="brand-icon">
            <PackageCheck />
          </div>
          <div>
            <h1>Archivo Maestro</h1>
            <p>Planeación ejecutiva de producción</p>
          </div>
        </div>

        <div className="formula">
          <strong>Modelo operativo</strong>
          <span>Producción sugerida = pronóstico de venta + colchón operativo, ajustado a la regla operativa de mínimo 10 piezas y múltiplos de 5.</span>
          <span>Colchón: 15 piezas si el pronóstico es 300 o más; 10 piezas si es menor.</span>
          <span>Pasteles: producción sugerida mínima de 10 piezas y múltiplos de 5; si es menor a 8, no producir.</span>
          <span>Recomendación = inventario objetivo + pronóstico con colchón - existencias.</span>
        </div>

        <div className="sidebar-status">
          <span>{files.stock ? "Stock cargado" : "Falta stock fijo"}</span>
          <span>{files.ventas ? "Ventas cargadas" : "Faltan ventas"}</span>
          <span>{files.real ? "Producción real cargada" : "Real opcional"}</span>
        </div>
      </aside>

      <main className="main">
        <header className="top">
          <div>
            <span className="eyebrow">Dashboard ejecutivo</span>
            <h2>Producción diaria sugerida</h2>
            <p>Planea la producción por día usando el promedio histórico del mismo día de semana.</p>
          </div>
          <button className="primary" onClick={() => exportToExcel(filtered, summary)} disabled={!forecast.length}>
            <Download size={18} /> Exportar datos para MySQL
          </button>
        </header>

        <section className="executive-summary-section">
          <div className="section-heading">
            <div>
              <span className="eyebrow">Vista general</span>
              <h3>Resumen Ejecutivo</h3>
              <p>Resumen operativo del mes seleccionado, sin comparativos contra producción real incompleta.</p>
            </div>
            <BarChart3 size={24} />
          </div>

          <section className="executive-summary-kpis">
            <KpiCard
              icon={PackageCheck}
              label="Productos analizados"
              value={formatNumber(forecast.length)}
              caption="Catálogo filtrado sin rebanadas"
            />
            <KpiCard
              icon={BarChart3}
              label="Pronóstico de venta mensual"
              value={formatNumber(dailySummary.pronosticoVentaMensual, 0)}
              caption="Suma de pronósticos diarios"
            />
            <KpiCard
              icon={ShieldCheck}
              label="Producción sugerida mensual"
              value={formatNumber(dailySummary.produccionSugeridaMensual, 0)}
              caption="Con regla operativa"
            />
            <KpiCard
              icon={Target}
              label="Mes analizado"
              value={selectedMonth || "Sin mes"}
              caption={`Colchón operativo: ${dailyBufferPct}%`}
            />
          </section>

          <div className="forecast-concepts" role="note">
            <p><strong>Pronóstico de venta:</strong> cantidad estimada que se espera vender.</p>
            <p><strong>Producción sugerida:</strong> cantidad recomendada a producir después de aplicar colchón operativo y regla de múltiplos de 5.</p>
          </div>
        </section>

        <section className="loaded-files-section">
          <div className="section-heading compact-heading">
            <div>
              <span className="eyebrow">Estado de carga</span>
              <h3>Archivos cargados</h3>
            </div>
            <div className="loaded-context">
              <span>Mes: {selectedMonth || "Sin mes"}</span>
              <span>Colchón: {dailyBufferPct}%</span>
            </div>
          </div>
          <div className="file-status-grid">
            {loadedFileItems.map((item) => (
              <div className="file-status-item" key={item.label}>
                <span className={`status-dot ${item.loaded ? "loaded" : ""}`} />
                <strong>{item.label}</strong>
                <small>{item.loaded ? "Cargado" : "Pendiente"}</small>
              </div>
            ))}
          </div>
        </section>

        <section className="uploads">
          <UploadBox
            title="Stock fijo"
            description="Productos oficiales, stock objetivo y orden."
            required
            onFile={(file) => handleFile(file, parseStock, "stock", setStockRows)}
            fileName={files.stock}
          />
          <UploadBox
            title="Ventas"
            description="Histórico de venta por día y producto."
            required
            onFile={(file) => handleFile(file, (wb) => parseSalesOrReturns(wb, "ventas"), "ventas", setVentas)}
            fileName={files.ventas}
          />
          <UploadBox
            title="Bajas"
            description="Merma, devoluciones o bajas por producto."
            onFile={(file) => handleFile(file, (wb) => parseSalesOrReturns(wb, "bajas"), "bajas", setBajas)}
            fileName={files.bajas}
          />
          <UploadBox
            title="Existencias"
            description="TOTAL SUC., C.F. y SUMA SUC. Y C.F."
            onFile={(file) => handleFile(file, parseExistencias, "existencias", setExistencias)}
            fileName={files.existencias}
          />
          <UploadBox
            title="Producción real"
            description="Excel consolidado o ZIP con producción real."
            accept=".xlsx,.xls,.zip"
            onFile={handleProductionReal}
            fileName={files.real}
          />
        </section>

        <section className="controls">
          <div className="search">
            <Search size={18} />
            <input placeholder="Buscar producto..." value={query} onChange={(e) => setQuery(e.target.value)} />
          </div>
          <label>
            Horizonte mensual
            <input min="1" type="number" value={days} onChange={(e) => setDays(Math.max(1, Number(e.target.value)))} />
          </label>
          <label>
            Factor fin semana
            <input
              min="1"
              step="0.01"
              type="number"
              value={weekendBoost}
              onChange={(e) => setWeekendBoost(Math.max(1, Number(e.target.value)))}
            />
          </label>
          <label className="check-control">
            <input
              type="checkbox"
              checked={showMissingReal}
              onChange={(e) => setShowMissingReal(e.target.checked)}
            />
            Mostrar productos sin dato real
          </label>
        </section>

        <section className="daily-section">
          <div className="section-heading">
            <div>
              <span className="eyebrow">Planeación por día</span>
              <h3>Producción diaria sugerida</h3>
              <p>Promedia la venta por día de semana, incluyendo días sin venta, y asigna el pronóstico a cada fecha del mes seleccionado.</p>
              <strong className="row-counter">{formatNumber(dailyRows.length)} filas diarias generadas</strong>
            </div>
            <button className="primary" onClick={() => exportDailyToExcel(filteredDailyRows, dailySummary)} disabled={!filteredDailyRows.length}>
              <Download size={18} /> Exportar diario
            </button>
          </div>

          <section className="controls daily-controls">
            <label>
              Mes
              <input
                type="month"
                value={selectedMonth}
                onChange={(e) => {
                  setSelectedMonthTouched(true);
                  setSelectedMonth(e.target.value);
                }}
              />
            </label>
            <label>
              Fecha
              <input type="date" value={dailyDateFilter} onChange={(e) => setDailyDateFilter(e.target.value)} />
            </label>
            <div className="search">
              <Search size={18} />
              <input
                placeholder="Producto diario..."
                value={dailyProductQuery}
                onChange={(e) => setDailyProductQuery(e.target.value)}
              />
            </div>
            <label>
              Día de semana
              <select value={dailyWeekdayFilter} onChange={(e) => setDailyWeekdayFilter(e.target.value)}>
                <option value="">Todos</option>
                {WEEKDAYS.map((day) => (
                  <option value={day.index} key={day.index}>
                    {day.label}
                  </option>
                ))}
              </select>
            </label>
            <label>
              Colchón %
              <input
                min="0"
                type="number"
                value={dailyBufferPct}
                onChange={(e) => setDailyBufferPct(Math.max(0, Number(e.target.value)))}
              />
            </label>
            <label className="check-control">
              <input
                type="checkbox"
                checked={onlyDailyShortage}
                onChange={(e) => setOnlyDailyShortage(e.target.checked)}
              />
              Ver solo productos con faltante
            </label>
            <label className="check-control">
              <input
                type="checkbox"
                checked={onlyDailyOverproduction}
                onChange={(e) => setOnlyDailyOverproduction(e.target.checked)}
              />
              Ver solo productos con sobreproducción
            </label>
          </section>

          <section className="table-card daily-table-card">
            <table className="daily-table">
              <thead>
                <tr>
                  <th>Fecha</th>
                  <th>Día</th>
                  <th>Producto</th>
                  <th>Promedio aplicado</th>
                  <th>Pronóstico de venta</th>
                  <th>Colchón diario</th>
                  <th>Base con colchón</th>
                  <th>Producción sugerida</th>
                </tr>
              </thead>
              <tbody>
                {filteredDailyRows.map((row) => (
                  <tr key={`${row.fecha}-${row.producto}`}>
                    <td>{row.fecha}</td>
                    <td>{row.dia}</td>
                    <td>{row.producto}</td>
                    <td>{row.promedioUsado.toFixed(2)}</td>
                    <td>{row.pronosticoVentaDia.toFixed(2)}</td>
                    <td>{row.colchonDiario.toFixed(2)}</td>
                    <td>{row.baseConColchonDia.toFixed(2)}</td>
                    <td className="strong">{row.produccionSugeridaDia}</td>
                  </tr>
                ))}
              </tbody>
            </table>
            {!dailyRows.length && (
              <div className="empty">
                Carga <strong>stock fijo</strong> y <strong>ventas</strong> para calcular producción diaria.
              </div>
            )}
            {dailyRows.length > 0 && !filteredDailyRows.length && (
              <div className="empty">
                No hay datos diarios para los filtros seleccionados. Revisa el mes, la fecha o el producto.
              </div>
            )}
          </section>
        </section>

        <section className="notes-section">
          <div className="section-heading compact-heading">
            <div>
              <span className="eyebrow">Lectura rápida</span>
              <h3>Notas de interpretación</h3>
            </div>
          </div>
          <div className="notes-list">
            <p>El pronóstico usa el promedio histórico por día de semana.</p>
            <p>La producción sugerida aplica el colchón operativo.</p>
            <p>Para pasteles GDE, MED y CH, la producción se ajusta a mínimo 10 y múltiplos de 5.</p>
            <p>Si el cálculo es menor a 8, se sugiere no producir.</p>
            <p>La vista Validación de cálculos permite auditar cada producto.</p>
          </div>
        </section>

        <section className="validation-section">
          <div className="section-heading">
            <div>
              <span className="eyebrow">Auditoría paso a paso</span>
              <h3>Validación de cálculos</h3>
              <p>Revisa las ventas leídas, los promedios aplicados y el cálculo diario completo de un producto.</p>
            </div>
            <label className="validation-product-select">
              Producto
              <select value={validationProduct} onChange={(e) => setValidationProduct(e.target.value)}>
                {!validationProducts.length && <option value="">Carga productos</option>}
                {validationProducts.map((row) => (
                  <option value={row.producto} key={row.producto}>
                    {row.producto}
                  </option>
                ))}
              </select>
            </label>
          </div>

          {validationForecast ? (
            <>
              <section className="validation-summary">
                <KpiCard
                  icon={FileSpreadsheet}
                  label="Ventas diarias leídas"
                  value={formatNumber(validationSales.length)}
                  caption={
                    validationSourceNames.length
                      ? `Excel: ${validationSourceNames.join(", ")}`
                      : `Producto homologado: ${validationProduct}`
                  }
                />
                <KpiCard
                  icon={BarChart3}
                  label="Pronóstico venta mensual"
                  value={formatNumber(validationSummary.pronosticoVentaMensual, 2)}
                  caption={`Suma diaria de ${selectedMonth}`}
                />
                <KpiCard
                  icon={ShieldCheck}
                  label="Producción sugerida mensual"
                  value={formatNumber(validationSummary.produccionSugeridaMensual)}
                  caption={`Regla operativa con ${dailyBufferPct}% de colchón`}
                />
                <KpiCard
                  icon={Database}
                  label="Producción real"
                  value={formatNumber(validationForecast.produccionReal)}
                  caption={`Diferencia: ${formatNumber(
                    validationForecast.produccionReal - validationSummary.produccionSugeridaMensual
                  )}`}
                  tone={
                    validationForecast.produccionReal < validationSummary.produccionSugeridaMensual
                      ? "danger"
                      : validationForecast.produccionReal > validationSummary.produccionSugeridaMensual
                        ? "warn"
                        : "ok"
                  }
                />
              </section>

              <section className="validation-grid">
                <div className="panel">
                  <div className="panel-title">
                    <div>
                      <h3>Promedio por día de semana</h3>
                      <p>Promedio = suma de cantidades del día / registros encontrados.</p>
                    </div>
                    <CheckCircle2 size={22} />
                  </div>
                  <div className="weekday-average-list">
                    {validationWeekdayAverages.map((day) => (
                      <div className="weekday-average-row" key={day.index}>
                        <span>{day.label}</span>
                        <small>{day.registros} registros</small>
                        <strong>{formatNumber(day.value, 2)}</strong>
                      </div>
                    ))}
                  </div>
                </div>

                <div className="panel">
                  <div className="panel-title">
                    <div>
                      <h3>Comprobación mensual</h3>
                      <p>La suma utiliza cada fecha del mes seleccionado.</p>
                    </div>
                    <Target size={22} />
                  </div>
                  <div className="calculation-checks">
                    <div>
                      <span>Pronóstico venta</span>
                      <strong>{formatNumber(validationSummary.pronosticoVentaMensual, 2)}</strong>
                    </div>
                    <div>
                    </div>
                    <div>
                      <span>Producción real</span>
                      <strong>{formatNumber(validationForecast.produccionReal)}</strong>
                    </div>
                    <div>
                      <span>Precisión contra real</span>
                      <strong>
                        {validationForecast.produccionReal > 0
                          ? formatPercent(
                              (1 -
                                Math.abs(
                                  validationSummary.produccionSugeridaMensual - validationForecast.produccionReal
                                ) /
                                  validationForecast.produccionReal) *
                                100,
                              1
                            )
                          : "Sin dato real"}
                      </strong>
                    </div>
                  </div>
                </div>
              </section>

              <div className="validation-block">
                <div className="validation-block-heading">
                  <div>
                    <h4>1. Ventas diarias leídas del Excel</h4>
                    <p>Estos son los registros usados para calcular los promedios de {validationProduct}.</p>
                  </div>
                  <strong>{validationSales.length} registros</strong>
                </div>
                <section className="table-card validation-sales-table-card">
                  <table className="validation-sales-table">
                    <thead>
                      <tr>
                        <th>Fecha leída</th>
                        <th>Día leído</th>
                        <th>Nombre en Excel</th>
                        <th>Producto homologado</th>
                        <th>Cantidad</th>
                      </tr>
                    </thead>
                    <tbody>
                      {validationSales.map((row, index) => (
                        <tr key={`${row.producto}-${row.fechaDisplay}-${index}`}>
                          <td>{row.fechaDisplay || "-"}</td>
                          <td>{row.dia}</td>
                          <td>{row.productoOriginal || row.producto}</td>
                          <td>{row.producto}</td>
                          <td className="strong">{formatNumber(row.cantidad, 2)}</td>
                        </tr>
                      ))}
                    </tbody>
                  </table>
                  {!validationSales.length && <div className="empty">No se encontraron ventas leídas para este producto.</div>}
                </section>
              </div>

              <div className="validation-block">
                <div className="validation-block-heading">
                  <div>
                    <h4>2. Registros usados para el promedio de jueves</h4>
                    <p>Promedio jueves = suma de cantidades de jueves / registros de jueves.</p>
                  </div>
                  <strong>
                    {validationThursdaySales.length} registros · Promedio {formatNumber(validationThursdayAverage, 2)}
                  </strong>
                </div>
                <section className="table-card validation-sales-table-card">
                  <table className="validation-sales-table">
                    <thead>
                      <tr>
                        <th>Fecha leída</th>
                        <th>Día leído</th>
                        <th>Nombre en Excel</th>
                        <th>Producto homologado</th>
                        <th>Cantidad usada</th>
                      </tr>
                    </thead>
                    <tbody>
                      {validationThursdaySales.map((row, index) => (
                        <tr key={`thursday-${row.producto}-${row.fechaDisplay}-${index}`}>
                          <td>{row.fechaDisplay || "-"}</td>
                          <td>{row.dia}</td>
                          <td>{row.productoOriginal || row.producto}</td>
                          <td>{row.producto}</td>
                          <td className="strong">{formatNumber(row.cantidad, 2)}</td>
                        </tr>
                      ))}
                    </tbody>
                  </table>
                  {!validationThursdaySales.length && (
                    <div className="empty">No se encontraron registros de jueves para este producto.</div>
                  )}
                </section>
              </div>

              <div className="validation-block">
                <div className="validation-block-heading">
                  <div>
                    <h4>3. Pronóstico diario y producción sugerida</h4>
                    <p>Cada fila muestra colchón, base con colchón y la regla operativa aplicada.</p>
                  </div>
                  <strong>{validationDailyRows.length} días</strong>
                </div>
                <section className="table-card validation-daily-table-card">
                  <table className="validation-daily-table">
                    <thead>
                      <tr>
                        <th>Fecha</th>
                        <th>Día</th>
                        <th>Promedio aplicado</th>
                        <th>Pronóstico de venta</th>
                        <th>Colchón aplicado</th>
                        <th>Base con colchón</th>
                        <th>Producción sugerida final</th>
                        <th>Producción real diaria</th>
                        <th>Diferencia</th>
                      </tr>
                    </thead>
                    <tbody>
                      {validationDailyRows.map((row) => (
                        <tr key={`validation-${row.fecha}-${row.producto}`}>
                          <td>{row.fechaDisplay}</td>
                          <td>{row.dia}</td>
                          <td>{formatNumber(row.promedioUsado, 2)}</td>
                          <td>{formatNumber(row.pronosticoVentaDia, 2)}</td>
                          <td>{formatNumber(row.colchonDiario, 2)}</td>
                          <td>{formatNumber(row.baseConColchonDia, 2)}</td>
                          <td className="strong">{formatNumber(row.produccionSugeridaDia)}</td>
                          <td>{row.produccionRealDia === null ? "-" : formatNumber(row.produccionRealDia)}</td>
                          <td>{row.diferenciaPiezas === null ? "-" : formatNumber(row.diferenciaPiezas)}</td>
                        </tr>
                      ))}
                    </tbody>
                  </table>
                </section>
              </div>
            </>
          ) : (
            <div className="empty validation-empty">
              Carga stock fijo y ventas para validar paso a paso un producto como <strong>PIÑA GDE</strong>.
            </div>
          )}
        </section>

        <section className="table-card">
          <table>
            <thead>
              <tr>
                <th>Producto</th>
                <th>Prom. reciente</th>
                <th>Pronóstico de venta</th>
                <th>Colchón</th>
                <th>Producción sugerida</th>
                <th>Stock objetivo</th>
                <th>Existencias</th>
                <th>Producción recomendada</th>
                <th>Producción real</th>
                <th>Brecha</th>
                <th>Precisión</th>
                <th>Estatus</th>
              </tr>
            </thead>
            <tbody>
              {filtered.map((r) => {
                const meta = STATUS_META[r.estatus] || STATUS_META.Revisar;
                return (
                  <tr key={r.producto}>
                    <td>{r.producto}</td>
                    <td>{r.promedioReciente.toFixed(2)}</td>
                    <td>{r.pronosticoVenta.toFixed(1)}</td>
                    <td>{r.colchonOperativo}</td>
                    <td>{r.produccionSugerida}</td>
                    <td>{r.inventarioObjetivo}</td>
                    <td>{r.sumaSucCf}</td>
                    <td className="strong">{r.produccionRecomendada}</td>
                    <td>{r.produccionReal}</td>
                    <td className={r.diferenciaReal < 0 ? "negative" : r.diferenciaReal > 0 ? "positive" : ""}>
                      {r.diferenciaReal > 0 ? "+" : ""}
                      {formatNumber(r.diferenciaReal)}
                    </td>
                    <td>{r.precision === null ? "-" : formatPercent(r.precision, 0)}</td>
                    <td>
                      <span className={`pill ${meta.className}`}>{meta.label}</span>
                    </td>
                  </tr>
                );
              })}
            </tbody>
          </table>
          {!forecast.length && (
            <div className="empty">
              Carga <strong>stock fijo</strong> y <strong>ventas</strong> para generar el dashboard de producción.
            </div>
          )}
          {forecast.length > 0 && !filtered.length && (
            <div className="empty">
              No hay productos comparables con producción real. Activa <strong>mostrar productos sin dato real</strong> para revisar todo el catálogo.
            </div>
          )}
        </section>
      </main>
    </div>
  );
}

createRoot(document.getElementById("root")).render(<App />);
