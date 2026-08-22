import React, { useEffect, useMemo, useState } from "react";
import { createRoot } from "react-dom/client";
import JSZip from "jszip";
import * as XLSX from "xlsx";
import {
  BarChart3,
  CalendarRange,
  CheckCircle2,
  Database,
  Download,
  FileSpreadsheet,
  LogOut,
  PackageCheck,
  Save,
  Search,
  ShieldCheck,
  Target,
  TrendingUp,
  Upload,
  UserRound,
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

const API_URL = String(
  import.meta.env.VITE_API_URL ?? (import.meta.env.DEV ? "http://localhost:4000" : "")
).replace(/\/$/, "");
const SESSION_STORAGE_KEY = "archivoMaestroSession";
const FORECAST_MODEL_VERSION = "categorySeasonal";
const OPERATIONAL_MARGIN_PCT = 12;
const API_PAGE_SIZE = 4000;
const API_UPLOAD_BATCH_SIZE = 1500;
const MAX_SNAPSHOT_BYTES = 3.5 * 1024 * 1024;

function loadStoredSession() {
  try {
    const value = JSON.parse(localStorage.getItem(SESSION_STORAGE_KEY) || "null");
    return value?.token && value?.user ? value : null;
  } catch {
    return null;
  }
}

async function apiRequest(path, { token, method = "GET", body } = {}) {
  const response = await fetch(`${API_URL}${path}`, {
    method,
    headers: {
      ...(body ? { "Content-Type": "application/json" } : {}),
      ...(token ? { Authorization: `Bearer ${token}` } : {}),
    },
    ...(body ? { body: JSON.stringify(body) } : {}),
  });
  const payload = await response.json().catch(() => ({}));
  if (!response.ok) {
    const error = new Error(payload.error || `Error del servidor (${response.status})`);
    error.status = response.status;
    throw error;
  }
  return payload;
}

async function apiRequestAllRows(path, { token } = {}) {
  const rows = [];
  let cursor = 0;
  for (let page = 0; page < 10000; page += 1) {
    const separator = path.includes("?") ? "&" : "?";
    const response = await apiRequest(
      `${path}${separator}cursor=${encodeURIComponent(cursor)}&limit=${API_PAGE_SIZE}`,
      { token }
    );
    rows.push(...(response.rows || []));
    if (!response.hasMore) return { ...response, rows };
    const nextCursor = Number(response.nextCursor);
    if (!Number.isFinite(nextCursor) || nextCursor <= cursor) {
      throw new Error("La sincronización devolvió un cursor inválido");
    }
    cursor = nextCursor;
  }
  throw new Error("La sincronización excedió el máximo de páginas permitido");
}

function consolidateRowsForUpload(rows, keyForRow, sumImporte = false) {
  const consolidated = new Map();
  for (const row of rows.filter((value) => !value.monthlyTotal)) {
    const key = keyForRow(row);
    const current = consolidated.get(key);
    if (!current) {
      consolidated.set(key, { ...row });
      continue;
    }
    current.cantidad += row.cantidad;
    if (sumImporte && row.importe !== null) current.importe = Number(current.importe || 0) + row.importe;
  }
  return [...consolidated.values()];
}

function consolidateSalesRowsForUpload(rows) {
  return consolidateRowsForUpload(
    rows,
    (row) => [row.fecha, norm(row.producto_codigo), norm(row.sucursal), norm(row.cliente)].join("|"),
    true
  );
}

function consolidateOperationalRowsForUpload(rows, isProduction) {
  return consolidateRowsForUpload(
    rows,
    (row) => [
      row.fecha,
      norm(row.producto_codigo),
      isProduction ? norm(row.turno) : norm(row.sucursal),
      isProduction ? "" : norm(row.motivo),
    ].join("|")
  );
}

async function uploadRowsInBatches({ endpoint, bodyKey, rows, archivo, token, onProgress }) {
  const batches = [];
  for (let offset = 0; offset < rows.length; offset += API_UPLOAD_BATCH_SIZE) {
    batches.push(rows.slice(offset, offset + API_UPLOAD_BATCH_SIZE));
  }
  const totals = {
    received: 0,
    valid: 0,
    rejected: 0,
    consolidated: 0,
    duplicatesInFile: 0,
    inserted: 0,
    updated: 0,
    issues: [],
  };
  for (const [index, batch] of batches.entries()) {
    onProgress?.(index + 1, batches.length);
    const response = await apiRequest(endpoint, {
      token,
      method: "POST",
      body: {
        archivo: batches.length > 1 ? `${archivo} [lote ${index + 1}/${batches.length}]` : archivo,
        [bodyKey]: batch,
      },
    });
    for (const field of ["received", "valid", "rejected", "consolidated", "duplicatesInFile", "inserted", "updated"]) {
      totals[field] += Number(response[field] || 0);
    }
    totals.issues.push(...(response.issues || []).slice(0, Math.max(0, 25 - totals.issues.length)));
  }
  return totals;
}

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
    .trim()
    // El catalogo escribe CHESSECAKE; las ventas usan ambas grafias.
    .replace(/CHE{1,2}S{1,2}ECAKE/g, "CHESSECAKE");
  const compact = normalized.replace(/[^A-Z0-9]/g, "");
  if (compact === "PINAGDE" || compact === "PINAGRANDE") return "PINA GDE";
  return normalized;
}

function productMatchKey(value) {
  return normalizeProduct(value)
    .replace(/\b(PASTEL|TARTA|PANQUE|PAY|DE|DEL|LA|EL)\b/g, " ")
    .replace(/\bGRANDE\b/g, "GDE")
    .replace(/\bMEDIANO\b/g, "MED")
    .replace(/\bCHICO\b/g, "CH")
    .replace(/\s+/g, " ")
    .trim();
}

function loadStoredProductAliases() {
  try {
    const parsed = JSON.parse(localStorage.getItem(PRODUCT_ALIAS_STORAGE_KEY) || "{}");
    if (!parsed || typeof parsed !== "object" || Array.isArray(parsed)) return {};
    return Object.fromEntries(
      Object.entries(parsed)
        .map(([alias, official]) => [normalizeProduct(alias), normalizeProduct(official)])
        .filter(([alias, official]) => alias && official)
    );
  } catch {
    return {};
  }
}

function getOfficialProducts(stockRows) {
  const seen = new Set();
  return stockRows
    .map((row) => normalizeProduct(row.producto))
    .filter((product) => {
      if (!product || seen.has(product)) return false;
      seen.add(product);
      return true;
    });
}

function findOfficialProduct(product, officialProducts) {
  const normalized = normalizeProduct(product);
  if (!normalized) return "";
  if (officialProducts.includes(normalized)) return normalized;

  const matchKey = productMatchKey(normalized);
  const officialByKey = officialProducts.find((official) => productMatchKey(official) === matchKey);
  return officialByKey || "";
}

function resolveOfficialProduct(product, productAliases, officialProducts) {
  const normalized = normalizeProduct(product);
  if (!normalized) return "";
  const manual = productAliases[normalized];
  if (manual && officialProducts.includes(manual)) return manual;
  return findOfficialProduct(normalized, officialProducts) || normalized;
}

function applyProductAliases(records, productAliases, officialProducts) {
  if (!officialProducts.length) return records;
  return records.map((record) => {
    const official = resolveOfficialProduct(record.producto, productAliases, officialProducts);
    return official === record.producto ? record : { ...record, producto: official };
  });
}

function buildHomologationRows({ ventas, bajas, existencias, realProduction, productAliases, officialProducts }) {
  if (!officialProducts.length) return [];
  const byProduct = new Map();
  const sources = [
    ["Ventas", ventas],
    ["Bajas", bajas],
    ["Existencias", existencias],
    ["Producción real", realProduction],
  ];

  for (const [sourceName, records] of sources) {
    for (const record of records) {
      const product = normalizeProduct(record.producto);
      if (!product || isSliceProduct(product)) continue;
      const row = byProduct.get(product) || {
        product,
        originalNames: new Set(),
        sources: new Set(),
        count: 0,
      };
      row.originalNames.add(record.productoOriginal || record.producto);
      row.sources.add(sourceName);
      row.count += 1;
      byProduct.set(product, row);
    }
  }

  return [...byProduct.values()]
    .map((row) => {
      const manual = productAliases[row.product] || "";
      const automatic = findOfficialProduct(row.product, officialProducts);
      const official = manual || automatic;
      return {
        ...row,
        originalNames: [...row.originalNames].slice(0, 4),
        sources: [...row.sources],
        official,
        status: official ? (manual ? "Manual" : "Automática") : "Pendiente",
      };
    })
    .filter((row) => row.status !== "Automática" || row.product !== row.official)
    .sort((a, b) => {
      if (a.status === "Pendiente" && b.status !== "Pendiente") return -1;
      if (a.status !== "Pendiente" && b.status === "Pendiente") return 1;
      return a.product.localeCompare(b.product, "es");
    });
}

function isSliceProduct(value) {
  const normalized = normalizeProduct(value);
  return /\b(REBANADA|REBANADAS|REB|RBN)\b/.test(normalized);
}

function isPromotionalProduct(value) {
  return /\b(PROMO|PROMOCION|PROMOCIONAL)\b/.test(normalizeProduct(value));
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

const PRODUCT_ALIAS_STORAGE_KEY = "archivoMaestroProductAliases";

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
  if (/^ARREGLO\s*\$?\s*\d+/.test(p)) return false;
  if (/^VENTA\s+DE\s+HOY/.test(p)) return false;
  if (p.includes("MODIFICACION DE PRECIO")) return false;
  if (p.includes("VENTA 2026")) return false;
  if (/^PRECIO\s*\$?\s*\d+/.test(p)) return false;
  if (p.startsWith("TOTAL")) return false;
  if (p.includes("PRODUCTO") || p.includes("ESPECIALIDAD")) return false;
  if (looksLikeCalendarHeader(p)) return false;
  if (/^\d+$/.test(p)) return false;
  return true;
}

function isValidInventoryProduct(value, row = []) {
  if (isValidProduct(value, row)) return true;
  const product = norm(value);
  return product.includes("SEMANA SANTA") && !isDateLikeValue(value) && !isCalendarRow(row);
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
  const isoMatch = raw.match(/^(\d{4})[/-](\d{1,2})[/-](\d{1,2})$/);
  if (isoMatch) {
    const [, year, month, day] = isoMatch.map(Number);
    const date = new Date(year, month - 1, day);
    return date.getFullYear() === year && date.getMonth() === month - 1 && date.getDate() === day ? date : null;
  }
  const localMatch = raw.match(/^(\d{1,2})[/-](\d{1,2})[/-](\d{2}|\d{4})$/);
  if (localMatch) {
    const day = Number(localMatch[1]);
    const month = Number(localMatch[2]);
    const yearValue = Number(localMatch[3]);
    const year = yearValue < 100 ? 2000 + yearValue : yearValue;
    const date = new Date(year, month - 1, day);
    return date.getFullYear() === year && date.getMonth() === month - 1 && date.getDate() === day ? date : null;
  }
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
  const candidateYears = [...new Set([currentYear - 2, currentYear - 1, currentYear, currentYear + 1, 2024, 2025, 2026, 2027])];
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

function monthKeyFromDate(date) {
  const d = parseDateCell(date);
  if (!d) return "";
  return `${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, "0")}`;
}

function monthKeyFromRecord(record) {
  return monthKeyFromDate(record?.fecha) || "";
}

function filterVentasByMonth(ventas, monthKey, mode = "exclude") {
  if (!monthKey) return ventas;
  return ventas.filter((record) => {
    const key = monthKeyFromRecord(record);
    if (!key) return mode === "exclude";
    return mode === "exclude" ? key !== monthKey : key === monthKey;
  });
}

function filterVentasBeforeMonth(ventas, monthKey) {
  if (!monthKey) return ventas;
  return ventas.filter((record) => {
    const key = monthKeyFromRecord(record);
    return key && key < monthKey;
  });
}

function precisionScore(forecast, actual) {
  if (!Number.isFinite(actual) || actual <= 0) return null;
  return Math.max(0, (1 - Math.abs(forecast - actual) / actual) * 100);
}

function aggregateDailySalesRows(records) {
  const map = new Map();
  for (const item of records) {
    if (item.monthlyTotal) continue;
    if (!isValidProduct(item.producto)) continue;
    const product = normalizeProduct(item.producto);
    if (isSliceProduct(product)) continue;
    const fechaKey = dateKey(item.fecha);
    if (!fechaKey) continue;
    const key = `${product}|${fechaKey}`;
    map.set(key, (map.get(key) || 0) + toNumber(item.cantidad));
  }
  return map;
}

function aggregateMonthlySalesByProduct(records) {
  const map = new Map();
  for (const item of records) {
    const product = normalizeProduct(item.producto);
    if (!product || isSliceProduct(product)) continue;
    map.set(product, (map.get(product) || 0) + toNumber(item.cantidad));
  }
  return map;
}

function buildWeekdayRow(productWeekdayAverages) {
  return {
    promedioLunes: Number(productWeekdayAverages.get(1) || 0),
    promedioMartes: Number(productWeekdayAverages.get(2) || 0),
    promedioMiercoles: Number(productWeekdayAverages.get(3) || 0),
    promedioJueves: Number(productWeekdayAverages.get(4) || 0),
    promedioViernes: Number(productWeekdayAverages.get(5) || 0),
    promedioSabado: Number(productWeekdayAverages.get(6) || 0),
    promedioDomingo: Number(productWeekdayAverages.get(0) || 0),
  };
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

const STOCK_SHEET_CANDIDATES = [
  "TOTAL A TENER SUC.(EXIST.+DIST)",
  "STOCK DE SUCURSALES",
  "EXIST. SUCURSALES Y RESTANTE CF",
];

function findStockSheet(workbook) {
  for (const candidate of STOCK_SHEET_CANDIDATES) {
    const normalizedCandidate = norm(candidate);
    const match = workbook.SheetNames.find(
      (name) => norm(name) === normalizedCandidate || norm(name).includes(normalizedCandidate)
    );
    if (match) return workbook.Sheets[match];
  }
  return workbook.Sheets[workbook.SheetNames[0]];
}

function parseStock(workbook) {
  const rows = XLSX.utils.sheet_to_json(findStockSheet(workbook), { header: 1, defval: "" });
  let headerIndex = -1;
  let productCol = 0;
  let stockCol = 1;

  for (let i = 0; i < Math.min(rows.length, 12); i++) {
    const row = rows[i].map(norm);
    const productIndex = row.findIndex((cell) => cell.includes("PRODUCTO"));
    const exactStockIndex = row.findIndex((cell) => cell === "STOCK");
    const fallbackTotalIndex = row.findIndex(
      (cell, index) =>
        index > 0 &&
        (cell === "TOTAL" ||
          (cell.includes("TOTAL") && !cell.includes("SUC") && !cell.includes("GRAL") && !cell.includes("GENERAL")))
    );
    const resolvedStockIndex = exactStockIndex >= 0 ? exactStockIndex : fallbackTotalIndex;
    if (productIndex >= 0 && resolvedStockIndex >= 0) {
      headerIndex = i;
      productCol = productIndex;
      stockCol = resolvedStockIndex;
      break;
    }
  }

  const parsed = [];
  const start = headerIndex >= 0 ? headerIndex + 1 : 0;
  for (let i = start; i < rows.length; i++) {
    if (!isValidInventoryProduct(rows[i][productCol], rows[i])) continue;
    const productoOriginal = String(rows[i][productCol] ?? "").trim();
    const product = normalizeProduct(productoOriginal);
    const stock = toNumber(rows[i][stockCol]);
    parsed.push({ producto: product, productoOriginal, stock, orden: parsed.length + 1 });
  }
  return parsed;
}

function parseExistencias(workbook) {
  const sheetName = workbook.SheetNames.find((name) => norm(name).includes("EXISTENCIA EN SUCURSALES"));
  const sheet = sheetName ? workbook.Sheets[sheetName] : workbook.Sheets[workbook.SheetNames[0]];
  const rows = XLSX.utils.sheet_to_json(sheet, { header: 1, defval: "" });

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
    const suma = row.findIndex(
      (x) => (x.includes("SUMA") && x.includes("SUC")) || (x.includes("SUCURSALES") && (x.includes("C.F") || x.includes("CF")))
    );
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
    if (!isValidInventoryProduct(rows[i][productCol], rows[i])) continue;
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

function inferYearHintFromFileName(fileName = "", fallbackYear = 2026) {
  const yearMatch = norm(fileName).match(/20\d{2}/);
  return yearMatch ? Number(yearMatch[0]) : fallbackYear;
}

function inferMonthHintFromFileName(fileName = "", fallbackYear = 2026) {
  const text = norm(fileName);
  const aliases = [
    ["ENERO", 0],
    ["FEBRERO", 1],
    ["FEEEBRERO", 1],
    ["MARZO", 2],
    ["ABRIL", 3],
    ["MAYO", 4],
    ["JUNIO", 5],
    ["JULIO", 6],
    ["AGOSTO", 7],
    ["SEPTIEMBRE", 8],
    ["OCTUBRE", 9],
    ["NOVIEMBRE", 10],
    ["DICIEMBRE", 11],
  ];
  const monthIndexes = [...new Set(aliases.filter(([alias]) => text.includes(alias)).map(([, monthIndex]) => monthIndex))];
  if (monthIndexes.length !== 1) return null;
  return { year: inferYearHintFromFileName(fileName, fallbackYear), monthIndex: monthIndexes[0] };
}

function parseMonthlyDailySheets(workbook, type = "ventas", monthHint = null) {
  const out = [];
  const skip = new Set(["RESUMEN", "REPORTE", "TOTAL", "TOTALES", "CONCENTRADO"]);
  for (const sheetName of workbook.SheetNames) {
    if (skip.has(norm(sheetName))) continue;

    const sheet = workbook.Sheets[sheetName];
    const rows = XLSX.utils.sheet_to_json(sheet, { header: 1, defval: "" });
    if (!rows.length) continue;

    const parsedSheetDate = parseDateCell(sheetName);
    const dayMatch = String(sheetName).match(/(?:^|\D)(\d{1,2})(?:\D|$)/);
    const sheetDay = dayMatch ? Number(dayMatch[1]) : null;
    const sheetDate = parsedSheetDate ||
      (monthHint && sheetDay >= 1 && sheetDay <= 31
        ? new Date(monthHint.year, monthHint.monthIndex, sheetDay)
        : null);

    let headerIndex = -1;
    let productCol = -1;
    let qtyCol = -1;
    let amountCol = -1;
    let dateCol = -1;
    let branchCol = -1;
    let clientCol = -1;
    let reasonCol = -1;

    for (let i = 0; i < Math.min(rows.length, 12); i++) {
      const row = rows[i].map(norm);
      const p = row.findIndex((x) => x.includes("PRODUCTO") || x.includes("DESCRIPCION") || x.includes("ARTICULO"));
      const q = row.findIndex((x) => x.includes("CANT") || x.includes("VENTA") || x.includes("PIEZAS") || x.includes("UNIDADES"));
      const a = row.findIndex((x) => x.includes("IMPORTE") || x.includes("TOTAL"));
      const d = row.findIndex((x) => x === "FECHA" || x.includes("FECHA") || x === "DIA");
      const b = row.findIndex(
        (x) => x.includes("SUCURSAL") || x.includes("TIENDA") || x === "CANAL" || x === "ZONA"
      );
      const c = row.findIndex((x) => x.includes("CLIENTE") || x === "CUSTOMER");
      const m = row.findIndex((x) => x.includes("MOTIVO") || x.includes("CAUSA") || x.includes("TIPO BAJA"));
      if (p >= 0 && q >= 0) {
        headerIndex = i;
        productCol = p;
        qtyCol = q;
        amountCol = a;
        dateCol = d;
        branchCol = b;
        clientCol = c;
        reasonCol = m;
        break;
      }
    }

    if (headerIndex < 0 && rows[0]?.length >= 2) {
      headerIndex = 0;
      productCol = 0;
      qtyCol = 1;
      amountCol = 2;
    }

    // A generic sheet without dates is a monthly summary and must be handled by the fallback parser.
    if (headerIndex < 0 || (dateCol < 0 && !sheetDate)) continue;

    for (let i = headerIndex + 1; i < rows.length; i++) {
      if (!isValidProduct(rows[i][productCol], rows[i])) continue;
      let fecha = dateCol >= 0 ? parseDateCell(rows[i][dateCol]) : sheetDate;
      if (!fecha && dateCol >= 0 && monthHint) {
        const day = parseDayNumber(rows[i][dateCol]);
        if (day) fecha = new Date(monthHint.year, monthHint.monthIndex, day);
      }
      if (!fecha) continue;
      const productoOriginal = String(rows[i][productCol] ?? "").trim();
      const product = normalizeProduct(productoOriginal);
      const rawCantidad = rows[i][qtyCol];
      if (type === "ventas" && String(rawCantidad ?? "").trim() === "") continue;
      const cantidad = toNumber(rawCantidad);
      const importe = amountCol >= 0 ? toNumber(rows[i][amountCol]) : 0;
      if (type !== "ventas" && cantidad === 0 && importe === 0) continue;
      const sucursal = branchCol >= 0 ? String(rows[i][branchCol] ?? "").trim() : "";
      const cliente = clientCol >= 0 ? String(rows[i][clientCol] ?? "").trim() : "";
      const motivo = reasonCol >= 0 ? String(rows[i][reasonCol] ?? "").trim() : "";
      out.push({ fecha, producto: product, productoOriginal, cantidad, importe, sucursal, canal: sucursal, cliente, motivo, tipo: type });
    }
  }
  return out;
}

function parseBajasReport(workbook) {
  const sheetName = workbook.SheetNames.find((name) => norm(name) === "REPORTE");
  if (!sheetName) return [];
  const rows = XLSX.utils.sheet_to_json(workbook.Sheets[sheetName], { header: 1, defval: "" });
  let headerIndex = -1;
  let productCol = -1;
  let qtyCol = -1;
  let dateCol = -1;
  let branchCol = -1;
  let reasonCol = -1;

  for (let i = 0; i < Math.min(rows.length, 8); i += 1) {
    const row = rows[i].map(norm);
    const product = row.findIndex((value) => value.includes("PRODUCTO"));
    const quantity = row.findIndex((value) => value.includes("CANT"));
    const date = row.findIndex((value) => value.includes("FECHA"));
    if (product < 0 || quantity < 0 || date < 0) continue;
    headerIndex = i;
    productCol = product;
    qtyCol = quantity;
    dateCol = date;
    branchCol = row.findIndex((value) => value.includes("SUCURSAL"));
    reasonCol = row.findIndex((value) => value.includes("MOTIVO"));
    break;
  }

  if (headerIndex < 0) return [];
  const parsed = [];
  for (let i = headerIndex + 1; i < rows.length; i += 1) {
    if (!isValidInventoryProduct(rows[i][productCol], rows[i])) continue;
    const fecha = parseDateCell(rows[i][dateCol]);
    const cantidad = toNumber(rows[i][qtyCol]);
    if (!fecha || cantidad === 0) continue;
    const productoOriginal = String(rows[i][productCol] ?? "").trim();
    const sucursal = branchCol >= 0 ? String(rows[i][branchCol] ?? "").trim() : "";
    parsed.push({
      fecha,
      producto: normalizeProduct(productoOriginal),
      productoOriginal,
      cantidad,
      importe: 0,
      sucursal,
      canal: sucursal,
      cliente: "",
      motivo: reasonCol >= 0 ? String(rows[i][reasonCol] ?? "").trim() : "",
      tipo: "bajas",
    });
  }
  return parsed;
}

function parseWideSales(workbook, type = "ventas", monthHint = null, yearHint = 2026, fileName = "") {
  const parsed = [];
  const skip = new Set(["RESUMEN", "REPORTE", "TOTAL", "TOTALES", "CONCENTRADO", "HOJA1"]);

  for (const sheetName of workbook.SheetNames) {
    if (skip.has(norm(sheetName))) continue;

    const sheet = workbook.Sheets[sheetName];
    const sheetBranch = /\b(SUC|SUCURSAL|TIENDA)\b/.test(norm(sheetName)) ? String(sheetName).trim() : "";
    const rows = XLSX.utils.sheet_to_json(sheet, { header: 1, defval: "" });
    if (rows.length < 4) continue;

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

    if (bestWeekdayScore <= 0 || dateHeaderIndex < 0) continue;

    const weekdayHeaders = rows[weekdayHeaderIndex] || [];
    const dateHeaders = rows[dateHeaderIndex] || [];
    let sheetMonthHint = inferMonthHintFromFileName(sheetName, yearHint);
    const fileText = norm(fileName);
    if (fileText.includes("MAYO") && fileText.includes("JUNIO") && norm(sheetName).includes("JULIO")) {
      // La fuente histórica combinada tiene la hoja de junio etiquetada como JULIO 2026.
      sheetMonthHint = { year: yearHint, monthIndex: 5 };
    }
    const inferredMonth = sheetMonthHint || monthHint || inferMonthYearFromWideHeaders(weekdayHeaders, dateHeaders);
    const inferredMonthDays = inferredMonth ? new Date(inferredMonth.year, inferredMonth.monthIndex + 1, 0).getDate() : null;
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
        if (inferredMonthDays && dayNumber && dayNumber > inferredMonthDays) continue;
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
          sucursal: sheetBranch,
          canal: sheetBranch,
          cliente: "",
          weekday: weekdayHeader,
          tipo: type,
        });
      }
    }
  }
  return parsed;
}

function parseSalesOrReturns(workbook, type, fileName = "") {
  const yearHint = inferYearHintFromFileName(fileName);
  const monthHint = inferMonthHintFromFileName(fileName, yearHint);
  if (type === "bajas") {
    const report = parseBajasReport(workbook);
    if (report.length > 0) return report;
  }
  const wide = parseWideSales(workbook, type, monthHint, yearHint, fileName);
  if (wide.length > 0) return wide;
  const bySheets = parseMonthlyDailySheets(workbook, type, monthHint);
  if (bySheets.length > 0) return bySheets;
  return parseMonthlySummaryWorkbook(workbook, monthHint);
}

function parseProductionReal(workbook) {
  const rows = rowsFromFirstSheet(workbook);
  let headerIndex = -1;
  let productCol = 0;
  let qtyCol = 1;
  let dateCol = -1;
  let shiftCol = -1;

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
    const t = row.findIndex((x) => x.includes("TURNO") || x === "SHIFT");
    if (p >= 0 && q >= 0 && p !== q) {
      headerIndex = i;
      productCol = p;
      qtyCol = q;
      dateCol = d;
      shiftCol = t;
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
    const turno = shiftCol >= 0 ? String(rows[i][shiftCol] ?? "").trim() : "";
    records.push({
      producto: product,
      productoOriginal,
      cantidad,
      fecha,
      fechaKey: fecha ? dateKey(fecha) : "",
      turno,
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

function parseMonthlySummaryWorkbook(workbook, monthHint = null) {
  const sheetName = workbook.SheetNames.find((name) => !/REPORTE/i.test(name)) || workbook.SheetNames[0];
  const rows = XLSX.utils.sheet_to_json(workbook.Sheets[sheetName], { header: 1, defval: "" });
  const headerIndex = rows.findIndex((row) => {
    const values = row.map(norm);
    return values.some(
      (value) => value.includes("CANT") || value.includes("VENTA") || value.includes("PIEZAS") || value.includes("UNIDADES")
    ) && values.some((value) => value.includes("PRODUCTO") || value.includes("ETIQUETAS"));
  });
  const start = headerIndex >= 0 ? headerIndex + 1 : 1;
  const header = headerIndex >= 0 ? rows[headerIndex].map(norm) : [];
  const quantityCol = header.findIndex(
    (value) => value.includes("CANT") || value.includes("VENTA") || value.includes("PIEZAS") || value.includes("UNIDADES") || value.includes("SUMA")
  );
  const productCol = header.findIndex((value) => value.includes("PRODUCTO") || value.includes("ETIQUETAS"));
  const resolvedQuantityCol = quantityCol >= 0 ? quantityCol : 0;
  const resolvedProductCol = productCol >= 0 ? productCol : 1;
  const map = new Map();

  for (let index = start; index < rows.length; index += 1) {
    const original = String(rows[index][resolvedProductCol] ?? "").trim();
    if (!isValidProduct(original)) continue;
    const product = normalizeProduct(original);
    if (isSliceProduct(product)) continue;
    const quantity = toNumber(rows[index][resolvedQuantityCol]);
    map.set(product, (map.get(product) || 0) + quantity);
  }

  const fecha = monthHint ? new Date(monthHint.year, monthHint.monthIndex, 1) : null;
  const monthDays = monthHint ? new Date(monthHint.year, monthHint.monthIndex + 1, 0).getDate() : 0;
  return [...map.entries()].map(([producto, cantidad]) => ({
    fecha,
    producto,
    cantidad,
    monthlyTotal: true,
    monthDays,
  }));
}

function parseBajasSummaryWorkbook(workbook) {
  const sheetName = workbook.SheetNames.find((name) => norm(name).includes("BAJAS") && norm(name) !== "REPORTE");
  if (!sheetName) return [];
  const rows = XLSX.utils.sheet_to_json(workbook.Sheets[sheetName], { header: 1, defval: "" });
  const headerIndex = rows.findIndex((row) => {
    const values = row.map(norm);
    return values.some((value) => value.includes("ETIQUETAS") || value.includes("PRODUCTO")) &&
      values.some((value) => value.includes("CANT") || value.includes("SUMA"));
  });
  if (headerIndex < 0) return [];
  const header = rows[headerIndex].map(norm);
  const productCol = header.findIndex((value) => value.includes("ETIQUETAS") || value.includes("PRODUCTO"));
  const quantityCol = header.findIndex((value) => value.includes("CANT") || value.includes("SUMA"));
  const map = new Map();
  for (let index = headerIndex + 1; index < rows.length; index += 1) {
    const original = String(rows[index][productCol] ?? "").trim();
    if (!isValidInventoryProduct(original) || /^ERICK?\b/.test(norm(original))) continue;
    const product = normalizeProduct(original);
    const quantity = toNumber(rows[index][quantityCol]);
    if (quantity === 0) continue;
    map.set(product, (map.get(product) || 0) + quantity);
  }
  return [...map.entries()].map(([producto, cantidad]) => ({ producto, cantidad, monthlyTotal: true }));
}

async function parseMonthlySummaryFile(file) {
  if (!file) return [];
  const workbook = await readWorkbook(file);
  return parseMonthlySummaryWorkbook(workbook);
}

async function parseBajasSummaryFile(file) {
  if (!file) return [];
  const workbook = await readWorkbook(file);
  return parseBajasSummaryWorkbook(workbook);
}

function groupByProduct(records) {
  const map = new Map();
  for (const r of records) {
    const product = normalizeProduct(r.producto);
    if (isSliceProduct(product) || isPromotionalProduct(product)) continue;
    const current = map.get(product) || [];
    current.push(r);
    map.set(product, current);
  }
  return map;
}

function filterIncompleteHistoricalMonths(records) {
  const coverage = new Map();
  for (const row of records) {
    const date = parseDateCell(row.fecha);
    if (!date) continue;
    const key = monthKeyFromDate(date);
    const current = coverage.get(key) || { days: new Set(), monthlyTotal: false, daysInMonth: 0 };
    current.days.add(date.getDate());
    current.monthlyTotal ||= Boolean(row.monthlyTotal);
    current.daysInMonth = new Date(date.getFullYear(), date.getMonth() + 1, 0).getDate();
    coverage.set(key, current);
  }
  const incompleteMonths = new Set(
    [...coverage.entries()]
      .filter(([, value]) => !value.monthlyTotal && value.days.size > 1 && value.days.size / value.daysInMonth < 0.7)
      .map(([key]) => key)
  );
  return incompleteMonths.size
    ? records.filter((row) => !incompleteMonths.has(monthKeyFromRecord(row)))
    : records;
}

function fillCompleteZeroMonths(records, completeMonths) {
  if (!records.length || !completeMonths.length) return records;
  const observedMonths = new Set(records.map((row) => monthKeyFromRecord(row)).filter(Boolean));
  const firstObservedMonth = [...observedMonths].sort()[0];
  if (!firstObservedMonth) return records;
  const product = normalizeProduct(records[0].producto);
  const missingZeros = completeMonths
    .filter((month) => month >= firstObservedMonth && !observedMonths.has(month))
    .map((month) => {
      const [year, monthNumber] = month.split("-").map(Number);
      return {
        fecha: new Date(year, monthNumber - 1, 1),
        producto: product,
        cantidad: 0,
        monthlyTotal: true,
        monthDays: new Date(year, monthNumber, 0).getDate(),
        inferredZeroMonth: true,
      };
    });
  return missingZeros.length ? [...records, ...missingZeros] : records;
}

function calculateForecast({
  stockRows,
  historicalVentas,
  bajas,
  existencias,
  realProduction,
  selectedMonth,
  dailyBufferPct,
  modelVersion = FORECAST_MODEL_VERSION,
}) {
  const usableHistoricalVentas = filterIncompleteHistoricalMonths(historicalVentas);
  const completeHistoricalMonths = [...new Set(
    usableHistoricalVentas.map((row) => monthKeyFromRecord(row)).filter(Boolean)
  )].sort();
  const ventasByProduct = groupByProduct(usableHistoricalVentas);
  const bajasByProduct = groupByProduct(bajas);
  const existMap = new Map(existencias.map((e) => [e.producto, e]));
  const realMap = new Map(aggregateProductionRows(realProduction).map((e) => [e.producto, e.cantidad]));
  const monthDates = datesForMonth(selectedMonth);

  return stockRows.filter((s) => !isSliceProduct(s.producto) && !isPromotionalProduct(s.producto)).map((s) => {
    const v = fillCompleteZeroMonths(ventasByProduct.get(s.producto) || [], completeHistoricalMonths);
    const b = bajasByProduct.get(s.producto) || [];
    const forecastModel = calculateForecastModelForVersion(v, selectedMonth, s.producto, modelVersion);
    const weekdayRow = buildWeekdayRow(forecastModel.averages);

    let pronosticoVenta = 0;
    let colchonOperativo = 0;
    for (const date of monthDates) {
      const pronosticoDia = getWeekdayAverage(weekdayRow, weekdayLabel(date.getDay()));
      pronosticoVenta += pronosticoDia;
      colchonOperativo += pronosticoDia * (dailyBufferPct / 100);
    }

    const values = v.map((x) => x.cantidad);
    const promedioHistorico = values.length ? values.reduce((a, n) => a + n, 0) / values.length : 0;
    const promedioDiario =
      monthDates.length > 0 ? pronosticoVenta / monthDates.length : promedioHistorico;
    const registrosHistoricos = v.length;

    const bajasTotal = b.reduce((a, n) => a + n.cantidad, 0);
    const ventasTotal = v.reduce((a, n) => a + n.cantidad, 0);
    const tasaBajas = ventasTotal > 0 ? bajasTotal / ventasTotal : 0;
    const bajasEsperadas = pronosticoVenta * tasaBajas;
    const baseConColchon = pronosticoVenta + colchonOperativo;
    const produccionSugerida = getProduccionSugerida(s.producto, baseConColchon);

    const ex = existMap.get(s.producto) || { totalSuc: 0, cf: 0, sumaSucCf: 0 };
    const sumaSucCf = ex.sumaSucCf || ex.totalSuc + ex.cf;
    const inventarioObjetivo = s.stock;
    const produccionBalanceada = (inventarioObjetivo - sumaSucCf + produccionSugerida) / 2;
    const produccionRecomendada = Math.max(0, getProduccionSugerida(s.producto, produccionBalanceada));
    const hasRealData = realMap.has(s.producto);
    const produccionReal = hasRealData ? realMap.get(s.producto) : 0;
    const diferenciaReal = produccionReal - produccionSugerida;
    const precision =
      hasRealData && produccionReal > 0
        ? precisionScore(produccionSugerida, produccionReal)
        : null;

    const confianza =
      registrosHistoricos >= 14
        ? 90
        : registrosHistoricos >= 7
          ? 75
          : registrosHistoricos > 0
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
      promedioHistorico,
      promedioDiario,
      registrosHistoricos,
      ...weekdayRow,
      demandaPronosticada: pronosticoVenta,
      pronosticoVenta,
      tasaBajas,
      bajasEsperadas,
      colchonOperativo,
      baseConColchon,
      reglaOperativa: getReglaOperativaLabel(s.producto, baseConColchon),
      produccionSugerida,
      inventarioObjetivo,
      produccionBalanceada,
      totalSuc: ex.totalSuc || 0,
      cf: ex.cf || 0,
      sumaSucCf,
      produccionRecomendada,
      tendenciaAplicada: forecastModel.trend,
      mesesUsados: forecastModel.recentMonths.join(", "),
      metodoPronostico: forecastModel.method,
      mesValidacionModelo: forecastModel.backtestMonth,
      realValidacionModelo: forecastModel.backtestActual,
      pronosticoValidacionModelo: forecastModel.backtestForecast,
      errorValidacionModelo: forecastModel.backtestError,
      produccionReal,
      hasRealData,
      diferenciaReal,
      precision,
      confianza,
      estatus,
    };
  });
}

function clamp(value, min, max) {
  return Math.min(max, Math.max(min, value));
}

function previousMonthKey(monthKey) {
  const [year, month] = String(monthKey).split("-").map(Number);
  if (!year || !month) return "";
  const previous = new Date(year, month - 2, 1);
  return `${previous.getFullYear()}-${String(previous.getMonth() + 1).padStart(2, "0")}`;
}

function sameMonthPreviousYear(monthKey) {
  const [year, month] = String(monthKey).split("-").map(Number);
  return year && month ? `${year - 1}-${String(month).padStart(2, "0")}` : "";
}

function buildMonthlyForecastData(records) {
  const monthlyData = new Map();
  for (const record of records) {
    const monthKey = monthKeyFromRecord(record);
    if (!monthKey) continue;
    const monthData = monthlyData.get(monthKey) || {
      total: 0,
      valuesByDate: new Map(),
      syntheticDays: 0,
    };
    if (record.monthlyTotal && record.monthDays) {
      monthData.total += toNumber(record.cantidad);
      monthData.syntheticDays = Math.max(monthData.syntheticDays, record.monthDays);
    } else {
      const date = parseDateCell(record.fecha);
      if (!date) continue;
      const key = dateKey(date);
      monthData.valuesByDate.set(key, (monthData.valuesByDate.get(key) || 0) + toNumber(record.cantidad));
    }
    monthlyData.set(monthKey, monthData);
  }

  for (const [monthKey, monthData] of monthlyData.entries()) {
    if (monthData.syntheticDays) {
      const [year, month] = monthKey.split("-").map(Number);
      const dailyTotal = [...monthData.valuesByDate.values()].reduce((sum, value) => sum + value, 0);
      if (monthData.valuesByDate.size > 1 && dailyTotal > 0) {
        // El mes trae total mensual y detalle diario: se conserva la forma por dia de semana
        // del detalle y solo se ajusta su nivel al total declarado.
        const levelFactor = monthData.total / dailyTotal;
        for (const [key, value] of monthData.valuesByDate.entries()) {
          monthData.valuesByDate.set(key, value * levelFactor);
        }
      } else {
        const dailyValue = monthData.total / monthData.syntheticDays;
        for (let day = 1; day <= monthData.syntheticDays; day += 1) {
          monthData.valuesByDate.set(dateKey(new Date(year, month - 1, day)), dailyValue);
        }
      }
    } else {
      monthData.total = [...monthData.valuesByDate.values()].reduce((sum, value) => sum + value, 0);
    }
    monthData.dailyRate = monthData.valuesByDate.size ? monthData.total / monthData.valuesByDate.size : 0;
    monthData.weekdays = new Map();
    for (const [key, value] of monthData.valuesByDate.entries()) {
      const weekday = parseDateCell(`${key}T12:00:00`)?.getDay();
      if (weekday === undefined) continue;
      const bucket = monthData.weekdays.get(weekday) || { total: 0, count: 0 };
      bucket.total += value;
      bucket.count += 1;
      monthData.weekdays.set(weekday, bucket);
    }
  }
  return monthlyData;
}

function uniformWeekdayAverages(dailyValue) {
  const averages = new Map();
  WEEKDAYS.forEach((day) => averages.set(day.index, Math.max(0, dailyValue || 0)));
  return averages;
}

function weightedWeekdayAverages(monthlyData, monthKeys, weights) {
  const averages = new Map();
  for (const weekday of WEEKDAYS.map((day) => day.index)) {
    let total = 0;
    let usedWeight = 0;
    monthKeys.forEach((monthKey, index) => {
      const bucket = monthlyData.get(monthKey)?.weekdays?.get(weekday);
      if (!bucket?.count) return;
      total += (bucket.total / bucket.count) * weights[index];
      usedWeight += weights[index];
    });
    if (usedWeight) averages.set(weekday, total / usedWeight);
  }
  return averages;
}

function forecastTotalFromAverages(averages, targetMonth) {
  return datesForMonth(targetMonth).reduce((sum, date) => sum + (averages.get(date.getDay()) || 0), 0);
}

function scaleForecastAverages(averages, factor) {
  return new Map([...averages.entries()].map(([weekday, value]) => [weekday, Math.max(0, value * factor)]));
}

function blendForecastAverages(primary, secondary, primaryWeight) {
  return new Map(
    WEEKDAYS.map((day) => [
      day.index,
      Math.max(0, (primary.get(day.index) || 0) * primaryWeight + (secondary.get(day.index) || 0) * (1 - primaryWeight)),
    ])
  );
}

function median(values) {
  const sorted = values.filter(Number.isFinite).sort((a, b) => a - b);
  if (!sorted.length) return 0;
  const middle = Math.floor(sorted.length / 2);
  return sorted.length % 2 ? sorted[middle] : (sorted[middle - 1] + sorted[middle]) / 2;
}

function buildForecastCandidates(monthlyData, targetMonth) {
  const historicalMonths = [...monthlyData.keys()].filter((monthKey) => monthKey < targetMonth).sort();
  const recentMonths = historicalMonths.slice(-3);
  const latestMonth = recentMonths.at(-1);
  const targetDays = datesForMonth(targetMonth).length;
  const candidates = new Map();
  const addCandidate = (name, averages, sourceMonths = recentMonths) => {
    if (!averages?.size) return;
    candidates.set(name, {
      averages,
      total: forecastTotalFromAverages(averages, targetMonth),
      sourceMonths,
    });
  };

  if (!latestMonth) {
    addCandidate("Sin histórico", uniformWeekdayAverages(0), []);
    return candidates;
  }

  const rates = recentMonths.map((monthKey) => monthlyData.get(monthKey)?.dailyRate || 0);
  const totals = recentMonths.map((monthKey) => monthlyData.get(monthKey)?.total || 0);
  addCandidate("Último mes por día", uniformWeekdayAverages(rates.at(-1)), [latestMonth]);
  addCandidate("Último total mensual", uniformWeekdayAverages(totals.at(-1) / Math.max(1, targetDays)), [latestMonth]);
  if (recentMonths.length >= 2) {
    addCandidate(
      "Promedio 2 meses por día",
      uniformWeekdayAverages(rates.at(-1) * 0.7 + rates.at(-2) * 0.3),
      recentMonths.slice(-2)
    );
    addCandidate(
      "Promedio total 2 meses",
      uniformWeekdayAverages((totals.at(-1) * 0.7 + totals.at(-2) * 0.3) / Math.max(1, targetDays)),
      recentMonths.slice(-2)
    );
    addCandidate(
      "Día de semana ponderado",
      weightedWeekdayAverages(monthlyData, recentMonths.slice(-2), [0.35, 0.65]),
      recentMonths.slice(-2)
    );
  }
  if (recentMonths.length >= 3) {
    addCandidate(
      "Promedio 3 meses por día",
      uniformWeekdayAverages(rates[0] * 0.2 + rates[1] * 0.3 + rates[2] * 0.5),
      recentMonths
    );
    const trendRate = clamp(rates[2] + (rates[2] - rates[1]) * 0.5, rates[2] * 0.8, rates[2] * 1.2);
    addCandidate("Tendencia reciente", uniformWeekdayAverages(trendRate), recentMonths.slice(-2));
    addCandidate("Mediana 3 meses", uniformWeekdayAverages(median(rates)), recentMonths);
    addCandidate(
      "Promedio total 3 meses",
      uniformWeekdayAverages((totals[0] * 0.2 + totals[1] * 0.3 + totals[2] * 0.5) / Math.max(1, targetDays)),
      recentMonths
    );
  }
  addCandidate(
    "Día de semana último mes",
    weightedWeekdayAverages(monthlyData, [latestMonth], [1]),
    [latestMonth]
  );

  const priorYearMonth = sameMonthPreviousYear(targetMonth);
  if (monthlyData.has(priorYearMonth)) {
    const previousTargetMonth = previousMonthKey(targetMonth);
    const previousYearReference = sameMonthPreviousYear(previousTargetMonth);
    const currentPreviousTotal = monthlyData.get(previousTargetMonth)?.total || 0;
    const priorPreviousTotal = monthlyData.get(previousYearReference)?.total || 0;
    const growth = currentPreviousTotal > 0 && priorPreviousTotal > 0
      ? clamp(currentPreviousTotal / priorPreviousTotal, 0.8, 1.2)
      : 1;
    const seasonalBase = weightedWeekdayAverages(monthlyData, [priorYearMonth], [1]);
    const adjustedSeasonal = scaleForecastAverages(seasonalBase, growth);
    addCandidate(
      "Mismo mes año anterior",
      adjustedSeasonal,
      [priorYearMonth, previousTargetMonth, previousYearReference].filter(Boolean)
    );
    const recentWeights = recentMonths.length === 1 ? [1] : recentMonths.length === 2 ? [0.35, 0.65] : [0.2, 0.3, 0.5];
    const recentBase = weightedWeekdayAverages(monthlyData, recentMonths, recentWeights);
    const seasonalTotal = forecastTotalFromAverages(adjustedSeasonal, targetMonth);
    const recentTotal = forecastTotalFromAverages(recentBase, targetMonth);
    const referencesDiffer = Math.max(seasonalTotal, recentTotal) > 0 &&
      Math.abs(seasonalTotal - recentTotal) / Math.max(seasonalTotal, recentTotal) > 0.5;
    addCandidate(
      "Estacional-reciente",
      blendForecastAverages(adjustedSeasonal, recentBase, referencesDiffer ? 0.5 : 0.9),
      [...new Set([priorYearMonth, ...recentMonths])]
    );
  }

  return candidates;
}

function calculateForecastModelLegacy(records, selectedMonth, useLatestAvailableBacktest = true) {
  const monthlyData = buildMonthlyForecastData(records);
  const historicalMonths = [...monthlyData.keys()].filter((month) => month < selectedMonth).sort();
  const previousCalendarMonth = previousMonthKey(selectedMonth);
  const backtestMonth = !useLatestAvailableBacktest || monthlyData.has(previousCalendarMonth)
    ? previousCalendarMonth
    : historicalMonths.at(-1) || previousCalendarMonth;
  const backtestActual = monthlyData.get(backtestMonth)?.total || 0;
  const backtestCandidates = buildForecastCandidates(monthlyData, backtestMonth);
  let selectedMethod = "Último mes por día";
  let backtestError = Infinity;
  for (const [method, candidate] of backtestCandidates.entries()) {
    const error = Math.abs(backtestActual - candidate.total);
    if (error < backtestError) {
      selectedMethod = method;
      backtestError = error;
    }
  }

  const targetCandidates = buildForecastCandidates(monthlyData, selectedMonth);
  const selected = targetCandidates.get(selectedMethod) ||
    targetCandidates.get("Promedio 3 meses por día") ||
    targetCandidates.values().next().value ||
    { averages: uniformWeekdayAverages(0), total: 0, sourceMonths: [] };
  const previousPrediction = backtestCandidates.get(selectedMethod)?.total || 0;
  const calibration = backtestActual > 0 && previousPrediction > 0
    ? clamp(backtestActual / previousPrediction, 0.85, 1.15)
    : 1;
  const averages = scaleForecastAverages(selected.averages, calibration);

  return {
    averages,
    trend: calibration,
    recentMonths: selected.sourceMonths,
    method: selectedMethod,
    backtestMonth,
    backtestActual,
    backtestForecast: previousPrediction,
    backtestError: Number.isFinite(backtestError) ? backtestError : 0,
  };
}

function calculateForecastModelSeasonal(records, selectedMonth, seasonalWeight) {
  const product = normalizeProduct(records[0]?.producto || "");
  const useConservativeBacktest = /\b(GALLETA|BOLLO|PAN)\b/.test(product);
  const legacy = calculateForecastModelLegacy(records, selectedMonth, !useConservativeBacktest);
  if (useConservativeBacktest) return legacy;
  const candidates = buildForecastCandidates(buildMonthlyForecastData(records), selectedMonth);
  const seasonal = candidates.get("Estacional-reciente") || candidates.get("Mismo mes año anterior");
  const recent = candidates.get("Último mes por día");
  if (!seasonal || !recent || recent.total <= 0 || Math.abs(seasonal.total - recent.total) / recent.total < 0.08) {
    return legacy;
  }
  return {
    ...legacy,
    averages: blendForecastAverages(seasonal.averages, legacy.averages, seasonalWeight),
    method: `Resguardo estacional ${Math.round(seasonalWeight * 100)}%`,
    recentMonths: [...new Set([...legacy.recentMonths, ...seasonal.sourceMonths])],
  };
}

// El peso estacional fijo falla cuando la forma del año anterior contradice la tendencia
// reciente. Aqui se elige por producto probando varios pesos contra los ultimos meses,
// usando solo informacion anterior a cada mes de validacion.
function calculateForecastModelSeasonalAdaptive(records, selectedMonth, candidateWeights = [0, 0.25, 0.5, 0.75, 1]) {
  const monthlyData = buildMonthlyForecastData(records);
  const historicalMonths = [...monthlyData.keys()].filter((month) => month < selectedMonth).sort();
  const validationMonths = historicalMonths.slice(-3);
  const fallbackWeight = ["Otros", "Mini medianos"].includes(productCategory(records[0]?.producto || "")) ? 0.5 : 0.75;
  if (validationMonths.length < 2) return calculateForecastModelSeasonal(records, selectedMonth, fallbackWeight);

  const scores = new Map(candidateWeights.map((weight) => [weight, { error: 0, scale: 0 }]));
  validationMonths.forEach((validationMonth, index) => {
    const actual = monthlyData.get(validationMonth)?.total || 0;
    const priorRecords = records.filter((row) => {
      const key = monthKeyFromRecord(row);
      return key && key < validationMonth;
    });
    if (!priorRecords.length) return;
    const recencyWeight = 1 + index * 0.5;
    for (const candidateWeight of candidateWeights) {
      const model = calculateForecastModelSeasonal(priorRecords, validationMonth, candidateWeight);
      const total = forecastTotalFromAverages(model.averages, validationMonth);
      const score = scores.get(candidateWeight);
      score.error += Math.abs(actual - total) * recencyWeight;
      score.scale += Math.max(actual, 1) * recencyWeight;
    }
  });

  let bestWeight = fallbackWeight;
  let bestScore = Infinity;
  for (const [candidateWeight, score] of scores.entries()) {
    if (!score.scale) continue;
    const relativeError = score.error / score.scale;
    if (relativeError < bestScore) {
      bestScore = relativeError;
      bestWeight = candidateWeight;
    }
  }

  const model = calculateForecastModelSeasonal(records, selectedMonth, bestWeight);
  return { ...model, method: `Estacional adaptativo ${Math.round(bestWeight * 100)}%` };
}

function calculateForecastModelForVersion(records, selectedMonth, product, modelVersion) {
  const version = String(modelVersion || FORECAST_MODEL_VERSION);
  if (version === "seasonalAdaptive") return calculateForecastModelSeasonalAdaptive(records, selectedMonth);
  if (version === "categorySeasonal") {
    const seasonalWeight = ["Otros", "Mini medianos"].includes(productCategory(product)) ? 0.5 : 0.75;
    return calculateForecastModelSeasonal(records, selectedMonth, seasonalWeight);
  }
  if (version.startsWith("seasonal")) {
    return calculateForecastModelSeasonal(records, selectedMonth, Number(version.replace("seasonal", "")) / 100);
  }
  if (version === "rolling") return calculateForecastModel(records, selectedMonth);
  return calculateForecastModelLegacy(records, selectedMonth);
}

function calculateForecastModel(records, selectedMonth) {
  const monthlyData = buildMonthlyForecastData(records);
  const targetCandidates = buildForecastCandidates(monthlyData, selectedMonth);
  const targetMethods = [...targetCandidates.keys()].filter((method) => method !== "Sin histórico");
  if (!targetMethods.length) {
    return calculateForecastModelLegacy(records, selectedMonth);
  }

  const historicalMonths = [...monthlyData.keys()]
    .filter((monthKey) => monthKey < selectedMonth)
    .sort();
  const validationMonths = historicalMonths.slice(-4);
  const evaluations = new Map(targetMethods.map((method) => [method, {
    weightedError: 0,
    weightedScale: 0,
    observations: 0,
    predictions: [],
  }]));

  validationMonths.forEach((validationMonth, index) => {
    const actual = monthlyData.get(validationMonth)?.total || 0;
    const candidates = buildForecastCandidates(monthlyData, validationMonth);
    const weight = 1 + index * 0.25;
    for (const method of targetMethods) {
      const candidate = candidates.get(method);
      if (!candidate) continue;
      const stats = evaluations.get(method);
      stats.weightedError += Math.abs(actual - candidate.total) * weight;
      stats.weightedScale += Math.max(actual, 1) * weight;
      stats.observations += 1;
      stats.predictions.push({ month: validationMonth, actual, forecast: candidate.total, weight });
    }
  });

  const requiredObservations = validationMonths.length >= 2 ? 2 : 1;
  let selectedMethod = "";
  let selectedScore = Infinity;
  for (const method of targetMethods) {
    const stats = evaluations.get(method);
    if (stats.observations < requiredObservations) continue;
    const normalizedError = stats.weightedScale > 0 ? stats.weightedError / stats.weightedScale : Infinity;
    const coveragePenalty = Math.max(0, validationMonths.length - stats.observations) * 0.05;
    const score = normalizedError + coveragePenalty;
    if (score < selectedScore) {
      selectedMethod = method;
      selectedScore = score;
    }
  }

  if (!selectedMethod) {
    const legacy = calculateForecastModelLegacy(records, selectedMonth);
    selectedMethod = targetCandidates.has(legacy.method)
      ? legacy.method
      : targetCandidates.has("Promedio 3 meses por día")
        ? "Promedio 3 meses por día"
        : targetMethods[0];
  }

  const selected = targetCandidates.get(selectedMethod);
  const selectedEvaluation = evaluations.get(selectedMethod);
  const usableRatios = (selectedEvaluation?.predictions || [])
    .filter((row) => row.actual > 0 && row.forecast > 0)
    .slice(-3)
    .map((row) => ({ ratio: clamp(row.actual / row.forecast, 0.8, 1.2), weight: row.weight }));
  const ratioWeight = usableRatios.reduce((sum, row) => sum + row.weight, 0);
  const calibration = ratioWeight > 0
    ? clamp(usableRatios.reduce((sum, row) => sum + row.ratio * row.weight, 0) / ratioWeight, 0.85, 1.15)
    : 1;
  const averages = scaleForecastAverages(selected.averages, calibration);
  const latestValidation = selectedEvaluation?.predictions?.at(-1) || null;

  return {
    averages,
    trend: calibration,
    recentMonths: selected.sourceMonths,
    method: selectedMethod,
    backtestMonth: latestValidation?.month || previousMonthKey(selectedMonth),
    backtestActual: latestValidation?.actual || 0,
    backtestForecast: latestValidation?.forecast || 0,
    backtestError: latestValidation ? Math.abs(latestValidation.actual - latestValidation.forecast) : 0,
    validationMonths: selectedEvaluation?.predictions?.map((row) => row.month) || [],
    validationWape: Number.isFinite(selectedScore) ? selectedScore : null,
  };
}

function calculateDailyForecast({ monthlyRows, ventasReales, realProduction, selectedMonth, dailyBufferPct }) {
  const realDailyMap = aggregateDailyProductionRows(realProduction);
  const salesDailyMap = aggregateDailySalesRows(ventasReales);
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
      const hasVentaReal = salesDailyMap.has(realKey);
      const ventaRealDia = hasVentaReal ? salesDailyMap.get(realKey) : null;
      const diferenciaVenta = hasVentaReal ? ventaRealDia - pronosticoVentaDia : null;
      const precisionVenta = precisionScore(pronosticoVentaDia, ventaRealDia);

      let estatus = "Sin dato real";
      if (hasRealData && diferenciaPiezas < 0) estatus = "Riesgo faltante";
      else if (hasRealData && diferenciaPiezas > 0) estatus = "Sobreproduccion";
      else if (hasRealData) estatus = "Dentro de rango";

      let estatusVenta = "Sin dato real";
      if (hasVentaReal) {
        if (precisionVenta !== null && precisionVenta < 80) estatusVenta = "Revisar";
        else if (diferenciaVenta < 0) estatusVenta = "Riesgo faltante";
        else if (diferenciaVenta > 0) estatusVenta = "Sobreproduccion";
        else estatusVenta = "Dentro de rango";
      }

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
        ventaRealDia,
        hasVentaReal,
        diferenciaVenta,
        precisionVenta,
        estatusVenta,
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
      ? precisionScore(produccionSugeridaMensual, produccionRealMensual) ?? 0
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

function summarizeSalesValidation(dailyRows) {
  const rowsWithReal = dailyRows.filter((row) => row.hasVentaReal);
  const pronosticoTotal = rowsWithReal.reduce((sum, row) => sum + row.pronosticoVentaDia, 0);
  const ventaRealTotal = rowsWithReal.reduce((sum, row) => sum + row.ventaRealDia, 0);
  const diferenciaTotal = ventaRealTotal - pronosticoTotal;
  const precisionGlobal = precisionScore(pronosticoTotal, ventaRealTotal) ?? 0;

  return {
    diasConReal: rowsWithReal.length,
    pronosticoTotal,
    ventaRealTotal,
    diferenciaTotal,
    precisionGlobal,
  };
}

function productCategory(product) {
  const value = normalizeProduct(product);
  if (value.includes("GELATINA")) return "Gelatinas";
  if (value.includes("GALLETA")) return "Galletas";
  if (/\b(GDE|GRANDE)\b/.test(value)) return "Pasteles grandes";
  if (/\b(MED|MEDIANO)\b/.test(value) && !value.includes("MINI")) return "Pasteles medianos";
  if (/\b(CH|CHICO)\b/.test(value)) return "Pasteles chicos";
  if (value.includes("MINI")) return "Mini medianos";
  if (value.includes("BOLLO") || /\bPAN\b/.test(value)) return "Pan";
  return "Otros";
}

function weekStartKey(value) {
  const date = parseDateCell(value);
  if (!date) return "";
  const start = new Date(date);
  start.setDate(start.getDate() - ((start.getDay() + 6) % 7));
  return dateKey(start);
}

function getProgressStatus(actual, expected, hasData) {
  if (!hasData) return { label: "Sin información", className: "muted", tone: "" };
  if (expected <= 0) {
    return actual > 0
      ? { label: "Sin base", className: "warn", tone: "warn" }
      : { label: "En objetivo", className: "ok", tone: "ok" };
  }
  const deviationPct = ((actual - expected) / expected) * 100;
  if (Math.abs(deviationPct) <= 10) return { label: "En objetivo", className: "ok", tone: "ok" };
  if (Math.abs(deviationPct) <= 20) return { label: "Atención", className: "warn", tone: "warn" };
  return { label: "Crítico", className: "danger", tone: "danger" };
}

function buildWeeklyProgress(dailyRows, selectedMonth) {
  const allDateKeys = [...new Set(dailyRows.map((row) => row.fecha).filter(Boolean))].sort();
  const loadedDateKeys = [...new Set(
    dailyRows.filter((row) => row.hasVentaReal).map((row) => row.fecha).filter(Boolean)
  )].sort();
  const firstRealDate = loadedDateKeys[0] || "";
  const cutoffDate = loadedDateKeys.at(-1) || "";
  const comparableDates = new Set(
    firstRealDate && cutoffDate
      ? allDateKeys.filter((key) => key >= firstRealDate && key <= cutoffDate)
      : []
  );
  const loadedDates = new Set(loadedDateKeys);
  const rowsByWeek = new Map();

  for (const row of dailyRows) {
    const key = weekStartKey(row.fecha);
    if (!key) continue;
    const current = rowsByWeek.get(key) || [];
    current.push(row);
    rowsByWeek.set(key, current);
  }

  const summarizeRows = (rows, periodComparableDates) => {
    const products = new Map();
    for (const row of rows) {
      const current = products.get(row.producto) || {
        producto: row.producto,
        categoria: productCategory(row.producto),
        pronosticoPeriodo: 0,
        pronosticoCorte: 0,
        ventaReal: 0,
      };
      current.pronosticoPeriodo += row.pronosticoVentaDia;
      if (periodComparableDates.has(row.fecha)) {
        current.pronosticoCorte += row.pronosticoVentaDia;
        current.ventaReal += row.hasVentaReal ? row.ventaRealDia : 0;
      }
      products.set(row.producto, current);
    }

    const hasData = periodComparableDates.size > 0;
    const productRows = [...products.values()].map((row) => {
      const diferencia = row.ventaReal - row.pronosticoCorte;
      const cumplimiento = row.pronosticoCorte > 0 ? (row.ventaReal / row.pronosticoCorte) * 100 : null;
      return {
        ...row,
        diferencia,
        cumplimiento,
        proyeccionPeriodo: row.ventaReal + Math.max(0, row.pronosticoPeriodo - row.pronosticoCorte),
        status: getProgressStatus(row.ventaReal, row.pronosticoCorte, hasData),
      };
    }).sort((a, b) => Math.abs(b.diferencia) - Math.abs(a.diferencia));

    const categoryMap = new Map();
    for (const row of productRows) {
      const current = categoryMap.get(row.categoria) || {
        categoria: row.categoria,
        productos: 0,
        pronosticoPeriodo: 0,
        pronosticoCorte: 0,
        ventaReal: 0,
      };
      current.productos += 1;
      current.pronosticoPeriodo += row.pronosticoPeriodo;
      current.pronosticoCorte += row.pronosticoCorte;
      current.ventaReal += row.ventaReal;
      categoryMap.set(row.categoria, current);
    }
    const categories = [...categoryMap.values()].map((row) => ({
      ...row,
      diferencia: row.ventaReal - row.pronosticoCorte,
      cumplimiento: row.pronosticoCorte > 0 ? (row.ventaReal / row.pronosticoCorte) * 100 : null,
      proyeccionPeriodo: row.ventaReal + Math.max(0, row.pronosticoPeriodo - row.pronosticoCorte),
      status: getProgressStatus(row.ventaReal, row.pronosticoCorte, hasData),
    })).sort((a, b) => Math.abs(b.diferencia) - Math.abs(a.diferencia));

    const pronosticoPeriodo = productRows.reduce((sum, row) => sum + row.pronosticoPeriodo, 0);
    const pronosticoCorte = productRows.reduce((sum, row) => sum + row.pronosticoCorte, 0);
    const ventaReal = productRows.reduce((sum, row) => sum + row.ventaReal, 0);
    return {
      pronosticoPeriodo,
      pronosticoCorte,
      ventaReal,
      diferencia: ventaReal - pronosticoCorte,
      cumplimiento: pronosticoCorte > 0 ? (ventaReal / pronosticoCorte) * 100 : null,
      proyeccionPeriodo: ventaReal + Math.max(0, pronosticoPeriodo - pronosticoCorte),
      status: getProgressStatus(ventaReal, pronosticoCorte, hasData),
      products: productRows,
      categories,
    };
  };

  const weeks = [...rowsByWeek.entries()].sort(([a], [b]) => a.localeCompare(b)).map(([key, rows], index) => {
    const weekDates = [...new Set(rows.map((row) => row.fecha))].sort();
    const periodComparableDates = new Set(weekDates.filter((date) => comparableDates.has(date)));
    const loadedDays = weekDates.filter((date) => loadedDates.has(date)).length;
    return {
      key,
      label: `Semana ${index + 1} · ${displayDate(weekDates[0])} al ${displayDate(weekDates.at(-1))}`,
      firstDate: weekDates[0] || "",
      lastDate: weekDates.at(-1) || "",
      comparedDays: periodComparableDates.size,
      loadedDays,
      coveragePct: periodComparableDates.size ? (loadedDays / periodComparableDates.size) * 100 : 0,
      ...summarizeRows(rows, periodComparableDates),
    };
  });

  const month = {
    ...summarizeRows(dailyRows, comparableDates),
    firstRealDate,
    cutoffDate,
    comparedDays: comparableDates.size,
    loadedDays: loadedDates.size,
    coveragePct: comparableDates.size ? (loadedDates.size / comparableDates.size) * 100 : 0,
  };
  month.projectedDifference = month.proyeccionPeriodo - month.pronosticoPeriodo;
  month.projectedStatus = getProgressStatus(month.proyeccionPeriodo, month.pronosticoPeriodo, Boolean(cutoffDate));

  const todayKey = dateKey(new Date());
  const referenceDate = cutoffDate || (todayKey.startsWith(`${selectedMonth}-`) ? todayKey : allDateKeys[0]);
  const suggestedWeekKey = weeks.find((week) => referenceDate >= week.firstDate && referenceDate <= week.lastDate)?.key || weeks[0]?.key || "";

  return {
    weeks,
    month,
    suggestedWeekKey,
    hasRealData: Boolean(cutoffDate),
  };
}

function buildMonthlyCloseSummary({ forecastRows, salesRows, productionRows }) {
  const aggregate = (records) => {
    const map = new Map();
    for (const row of records) {
      const product = normalizeProduct(row.producto);
      if (!isValidProduct(product) || isSliceProduct(product) || isPromotionalProduct(product)) continue;
      map.set(product, (map.get(product) || 0) + toNumber(row.cantidad));
    }
    return map;
  };
  const salesMap = aggregate(salesRows);
  const productionMap = aggregate(productionRows);
  const forecastProducts = new Set(forecastRows.map((row) => normalizeProduct(row.producto)));
  const salesLoaded = salesRows.length > 0;
  const productionLoaded = productionRows.length > 0;
  const rows = forecastRows.map((forecast) => {
    const producto = normalizeProduct(forecast.producto);
    const pronostico = toNumber(forecast.pronosticoVenta);
    const ventaReal = salesMap.get(producto) || 0;
    const producido = productionMap.get(producto) || 0;
    const diferenciaPronostico = ventaReal - pronostico;
    const diferenciaProduccion = producido - ventaReal;
    return {
      producto,
      categoria: productCategory(producto),
      pronostico,
      ventaReal,
      producido,
      diferenciaPronostico,
      diferenciaProduccion,
      errorAbsoluto: Math.abs(diferenciaPronostico),
      cumplimiento: pronostico > 0 ? (ventaReal / pronostico) * 100 : null,
      status: getProgressStatus(ventaReal, pronostico, salesLoaded),
      productionStatus: getProgressStatus(producido, ventaReal, salesLoaded && productionLoaded),
    };
  }).sort((a, b) => b.errorAbsoluto - a.errorAbsoluto);

  const summarize = (detailRows) => {
    const pronostico = detailRows.reduce((sum, row) => sum + row.pronostico, 0);
    const ventaReal = detailRows.reduce((sum, row) => sum + row.ventaReal, 0);
    const producido = detailRows.reduce((sum, row) => sum + row.producido, 0);
    const absoluteError = detailRows.reduce((sum, row) => sum + row.errorAbsoluto, 0);
    return {
      productos: detailRows.length,
      pronostico,
      ventaReal,
      producido,
      diferenciaPronostico: ventaReal - pronostico,
      diferenciaProduccion: producido - ventaReal,
      cumplimiento: pronostico > 0 ? (ventaReal / pronostico) * 100 : null,
      wape: salesLoaded && ventaReal > 0 ? (absoluteError / ventaReal) * 100 : null,
      mae: salesLoaded && detailRows.length ? absoluteError / detailRows.length : null,
      dentro15: salesLoaded ? detailRows.filter((row) => row.errorAbsoluto <= 15).length : 0,
      status: getProgressStatus(ventaReal, pronostico, salesLoaded),
      productionStatus: getProgressStatus(producido, ventaReal, salesLoaded && productionLoaded),
    };
  };

  const categoryMap = new Map();
  for (const row of rows) {
    const current = categoryMap.get(row.categoria) || [];
    current.push(row);
    categoryMap.set(row.categoria, current);
  }
  const categories = [...categoryMap.entries()].map(([categoria, detailRows]) => ({
    categoria,
    ...summarize(detailRows),
  })).sort((a, b) => b.ventaReal - a.ventaReal);
  const unmatchedRows = (sourceMap) => [...sourceMap.entries()]
    .filter(([product, quantity]) => quantity > 0 && !forecastProducts.has(product))
    .map(([producto, cantidad]) => ({ producto, cantidad }))
    .sort((a, b) => b.cantidad - a.cantidad);
  const unmatchedSales = unmatchedRows(salesMap);
  const unmatchedProduction = unmatchedRows(productionMap);

  return {
    salesLoaded,
    productionLoaded,
    summary: summarize(rows),
    rows,
    categories,
    unmatchedSales,
    unmatchedProduction,
    unmatchedSalesTotal: unmatchedSales.reduce((sum, row) => sum + row.cantidad, 0),
    unmatchedProductionTotal: unmatchedProduction.reduce((sum, row) => sum + row.cantidad, 0),
  };
}

function buildProductValidationSummary(dailyRows, forecastRows) {
  const byProduct = new Map();

  for (const row of dailyRows) {
    const current = byProduct.get(row.producto) || {
      producto: row.producto,
      pronosticoMensual: 0,
      ventaRealMensual: 0,
      diasConReal: 0,
      diasFueraRango: 0,
    };
    current.pronosticoMensual += row.pronosticoVentaDia;
    if (row.hasVentaReal) {
      current.ventaRealMensual += row.ventaRealDia;
      current.diasConReal += 1;
      if (row.precisionVenta !== null && row.precisionVenta < 80) current.diasFueraRango += 1;
    }
    byProduct.set(row.producto, current);
  }

  return [...byProduct.values()]
    .map((row) => {
      const forecast = forecastRows.find((item) => item.producto === row.producto);
      const diferencia = row.ventaRealMensual - row.pronosticoMensual;
      const precision = precisionScore(row.pronosticoMensual, row.ventaRealMensual);
      const errorPct =
        row.ventaRealMensual > 0 ? (Math.abs(diferencia) / row.ventaRealMensual) * 100 : null;
      let estatus = "Sin dato real";
      if (row.diasConReal > 0) {
        if (precision !== null && precision >= 90) estatus = "Dentro de rango";
        else if (precision !== null && precision >= 80) estatus = "Revisar";
        else estatus = "Riesgo faltante";
      }

      return {
        ...row,
        registrosHistoricos: forecast?.registrosHistoricos || 0,
        diferencia,
        precision,
        errorPct,
        estatus,
      };
    })
    .sort((a, b) => a.producto.localeCompare(b.producto, "es"));
}

function buildValidationAlerts(productSummary, homologationRows, historicalVentas, selectedMonth) {
  const alerts = [];

  for (const row of homologationRows.filter((item) => item.status === "Pendiente")) {
    alerts.push({
      tipo: "Homologación",
      producto: row.product,
      detalle: "Producto sin nombre oficial en catálogo",
      severidad: "Alta",
    });
  }

  for (const row of productSummary) {
    if (row.estatus === "Sin dato real") {
      alerts.push({
        tipo: "Sin venta real",
        producto: row.producto,
        detalle: `No hay ventas reales en ${selectedMonth}`,
        severidad: "Media",
      });
    } else if (row.precision !== null && row.precision < 70) {
      alerts.push({
        tipo: "Precisión baja",
        producto: row.producto,
        detalle: `Precisión ${formatPercent(row.precision, 1)}`,
        severidad: "Alta",
      });
    } else if (row.registrosHistoricos < 4) {
      alerts.push({
        tipo: "Poco histórico",
        producto: row.producto,
        detalle: `Solo ${row.registrosHistoricos} registros históricos`,
        severidad: "Media",
      });
    }
  }

  if (!historicalVentas.length) {
    alerts.unshift({
      tipo: "Histórico vacío",
      producto: "-",
      detalle: `No hay ventas históricas fuera de ${selectedMonth}`,
      severidad: "Alta",
    });
  }

  return alerts;
}

function sumProductRecords(records, product, monthKey = "") {
  return records
    .filter((record) => normalizeProduct(record.producto) === product && (!monthKey || monthKeyFromRecord(record) === monthKey))
    .reduce((sum, record) => sum + toNumber(record.cantidad), 0);
}

function weekdayAveragesForRecords(records) {
  const buckets = new Map();
  for (const record of records) {
    const weekday = recordWeekday(record);
    if (weekday === null) continue;
    const bucket = buckets.get(weekday) || { total: 0, count: 0 };
    bucket.total += toNumber(record.cantidad);
    bucket.count += 1;
    buckets.set(weekday, bucket);
  }
  return new Map([...buckets.entries()].map(([weekday, bucket]) => [weekday, bucket.count ? bucket.total / bucket.count : 0]));
}

function forecastFromHistoricalRecords(records, product, targetMonth) {
  const averages = weekdayAveragesForRecords(
    records.filter((record) => normalizeProduct(record.producto) === product)
  );
  return datesForMonth(targetMonth).reduce((sum, date) => sum + (averages.get(date.getDay()) || 0), 0);
}

function buildHistoricalValidationRows({ ventas, producedMay, producedJune, bajasJune, bajasJuly, stockRows }) {
  const stockOrder = new Map(stockRows.map((row) => [normalizeProduct(row.producto), row.orden]));
  const productSet = new Set([
    ...ventas.map((row) => normalizeProduct(row.producto)),
    ...producedMay.map((row) => normalizeProduct(row.producto)),
    ...producedJune.map((row) => normalizeProduct(row.producto)),
    ...bajasJune.map((row) => normalizeProduct(row.producto)),
    ...bajasJuly.map((row) => normalizeProduct(row.producto)),
  ]);
  const historicalRecords = ventas.filter((row) => ["2026-05", "2026-06"].includes(monthKeyFromRecord(row)));
  const mayRecords = ventas.filter((row) => monthKeyFromRecord(row) === "2026-05");

  return [...productSet]
    .filter(Boolean)
    .map((product) => {
      const ventaMayo = sumProductRecords(ventas, product, "2026-05");
      const ventaJunio = sumProductRecords(ventas, product, "2026-06");
      const producidoMayo = producedMay.find((row) => normalizeProduct(row.producto) === product)?.cantidad || 0;
      const producidoJunio = producedJune.find((row) => normalizeProduct(row.producto) === product)?.cantidad || 0;
      const bajasJunioCantidad = bajasJune.find((row) => normalizeProduct(row.producto) === product)?.cantidad || 0;
      const bajasJulioCantidad = bajasJuly.find((row) => normalizeProduct(row.producto) === product)?.cantidad || 0;
      const pronosticoJunio = forecastFromHistoricalRecords(mayRecords, product, "2026-06");
      const precisionJunio = precisionScore(pronosticoJunio, ventaJunio);
      const pronosticoJulio = forecastFromHistoricalRecords(historicalRecords, product, "2026-07");
      const tasaBajas = ventaJunio > 0 ? bajasJunioCantidad / ventaJunio : 0;
      const bajasEsperadasJulio = pronosticoJulio * tasaBajas;
      const margen = pronosticoJulio * 0.1;

      return {
        producto: stockRows.find((row) => normalizeProduct(row.producto) === product)?.producto || product,
        ventaMayo,
        producidoMayo,
        diferenciaMayo: producidoMayo - ventaMayo,
        ventaJunio,
        producidoJunio,
        bajasJunio: bajasJunioCantidad,
        demandaAjustadaJunio: ventaJunio + bajasJunioCantidad,
        saldoJunio: producidoJunio - ventaJunio - bajasJunioCantidad,
        pronosticoJunio,
        precisionJunio,
        tasaBajas,
        bajasEsperadasJulio,
        bajasJulio: bajasJulioCantidad,
        promedioDiarioHistorico: pronosticoJulio / 31,
        pronosticoJulio,
        margenSeguridad: margen,
        produccionSugeridaBase: getProduccionSugerida(product, pronosticoJulio + margen),
        produccionSugeridaAjustada: getProduccionSugerida(product, pronosticoJulio + margen + bajasEsperadasJulio),
      };
    })
    .sort((a, b) => (stockOrder.get(normalizeProduct(a.producto)) || 99999) - (stockOrder.get(normalizeProduct(b.producto)) || 99999) || a.producto.localeCompare(b.producto, "es"));
}

function buildOperationalForecastScenario(forecastRows, marginPct = OPERATIONAL_MARGIN_PCT) {
  const safeMarginPct = Math.max(0, Number(marginPct) || 0);
  return forecastRows.map((row) => {
    const baseForecast = Number(row.pronosticoVenta || 0);
    const marginPieces = baseForecast * safeMarginPct / 100;
    return {
      producto: row.producto,
      categoria: productCategory(row.producto),
      orden: row.orden,
      pronosticoBase: baseForecast,
      margenOperativoPct: safeMarginPct,
      margenOperativoPiezas: marginPieces,
      pronosticoOperativo: baseForecast + marginPieces,
      metodoPronostico: row.metodoPronostico,
      tendenciaAplicada: row.tendenciaAplicada,
      mesesUsados: row.mesesUsados,
    };
  });
}

const MONTHLY_REVIEW_STATUSES = ["ACTIVO", "BAJA", "BAJO PEDIDO", "ESTACIONAL"];

function historicalMonthlyStats(records, product) {
  const productRecords = records.filter(
    (record) => normalizeProduct(record.producto) === normalizeProduct(product)
  );
  const monthlyData = buildMonthlyForecastData(productRecords);
  const values = [...monthlyData.entries()]
    .sort(([a], [b]) => a.localeCompare(b))
    .slice(-6)
    .map(([, value]) => value.total);
  if (!values.length) return { months: 0, average: 0, volatility: null };
  const average = values.reduce((sum, value) => sum + value, 0) / values.length;
  const variance = values.reduce((sum, value) => sum + (value - average) ** 2, 0) / values.length;
  return {
    months: values.length,
    average,
    volatility: average > 0 ? Math.sqrt(variance) / average : null,
  };
}

function buildMonthlyReviewRows({
  sourceRows,
  forecastRows,
  historicalVentas,
  loadedExistencias,
  inputs = {},
}) {
  const forecastByProduct = new Map(forecastRows.map((row) => [normalizeProduct(row.producto), row]));
  const existenceByProduct = new Map(
    loadedExistencias.map((row) => [
      normalizeProduct(row.producto),
      toNumber(row.sumaSucCf || row.totalSuc + row.cf),
    ])
  );

  return sourceRows.map((sourceRow) => {
    const product = normalizeProduct(sourceRow.producto);
    const forecastRow = forecastByProduct.get(product) || {};
    const input = inputs[product] || {};
    const productStatus = MONTHLY_REVIEW_STATUSES.includes(input.status) ? input.status : "ACTIVO";
    const baseForecast = toNumber(sourceRow.pronosticoBase ?? sourceRow.pronosticoVenta);
    const baseOperational = toNumber(
      sourceRow.pronosticoOperativo ?? sourceRow.produccionSugerida ?? baseForecast * (1 + OPERATIONAL_MARGIN_PCT / 100)
    );
    const hasLoadedInventory = existenceByProduct.has(product);
    const inventory = input.inventoryOverride === null || input.inventoryOverride === undefined || input.inventoryOverride === ""
      ? (hasLoadedInventory ? existenceByProduct.get(product) : 0)
      : Math.max(0, toNumber(input.inventoryOverride));
    const stats = historicalMonthlyStats(historicalVentas, product);
    const reasons = [];
    let severity = "ok";
    let marginPct = 8;
    let proposed = baseOperational;

    if (stats.volatility === null || stats.months < 2) {
      marginPct = 15;
      reasons.push("Historial insuficiente: revisar manualmente.");
      severity = "warn";
    } else if (stats.volatility >= 0.5) {
      marginPct = 15;
      reasons.push(`Variación histórica alta (${(stats.volatility * 100).toFixed(0)}%).`);
      severity = "danger";
    } else if (stats.volatility >= 0.3) {
      marginPct = 12;
      reasons.push(`Variación histórica media (${(stats.volatility * 100).toFixed(0)}%).`);
      severity = "warn";
    } else {
      reasons.push(`Demanda estable: colchón local de ${marginPct}%.`);
    }

    if (productStatus === "BAJA") {
      proposed = 0;
      marginPct = 0;
      reasons.unshift("Producto marcado como BAJA: no producir.");
      severity = "danger";
    } else if (productStatus === "BAJO PEDIDO") {
      proposed = 0;
      marginPct = 0;
      reasons.unshift("Producto BAJO PEDIDO: excluir del plan regular.");
      severity = "warn";
    } else if (productStatus === "ESTACIONAL") {
      proposed = baseOperational;
      marginPct = OPERATIONAL_MARGIN_PCT;
      reasons.unshift("Producto ESTACIONAL: requiere confirmar su temporada; no se ajustó automáticamente.");
      severity = "warn";
    } else {
      proposed = Math.max(0, getProduccionSugerida(product, baseForecast * (1 + marginPct / 100) - inventory));
      if (hasLoadedInventory || input.inventoryOverride !== null && input.inventoryOverride !== undefined && input.inventoryOverride !== "") {
        reasons.push(`Se descontaron ${formatNumber(inventory, 0)} piezas de existencia.`);
      } else {
        reasons.push("Sin existencias capturadas: la propuesta puede estar sobrestimada.");
        if (severity === "ok") severity = "warn";
      }
    }

    if (Math.abs(toNumber(sourceRow.tendenciaAplicada ?? forecastRow.tendenciaAplicada)) >= 0.15) {
      reasons.push("Cambio de tendencia atípico: validar con operación.");
      if (severity === "ok") severity = "warn";
    }

    return {
      producto: sourceRow.producto,
      categoria: sourceRow.categoria || productCategory(product),
      orden: sourceRow.orden ?? forecastRow.orden,
      status: productStatus,
      note: String(input.note || ""),
      inventory,
      inventoryOverride: input.inventoryOverride ?? null,
      hasLoadedInventory,
      baseForecast,
      baseOperational,
      marginPct,
      proposed,
      difference: proposed - baseOperational,
      decision: ["accepted", "rejected"].includes(input.decision) ? input.decision : "pending",
      reasons,
      severity,
      historicalMonths: stats.months,
      volatility: stats.volatility,
    };
  }).sort((a, b) => (a.orden || 99999) - (b.orden || 99999) || a.producto.localeCompare(b.producto, "es"));
}

function exportMonthlyReview({ rows, review, selectedMonth, sourceVersion }) {
  const accepted = rows.filter((row) => row.decision === "accepted");
  const rejected = rows.filter((row) => row.decision === "rejected");
  const pending = rows.filter((row) => row.decision === "pending");
  const finalTotal = rows.reduce(
    (sum, row) => sum + (row.decision === "accepted" ? row.proposed : row.baseOperational),
    0
  );
  const summary = [
    { Indicador: "Mes", Valor: selectedMonth },
    { Indicador: "Estado", Valor: review.state === "approved" ? "Aprobada" : "Borrador" },
    { Indicador: "Versión de revisión", Valor: review.version || "Sin guardar" },
    { Indicador: "Versión del pronóstico congelado", Valor: sourceVersion || "Sin congelar" },
    { Indicador: "Pronóstico estadístico", Valor: Number(rows.reduce((sum, row) => sum + row.baseForecast, 0).toFixed(2)) },
    { Indicador: "Plan operativo base", Valor: Number(rows.reduce((sum, row) => sum + row.baseOperational, 0).toFixed(2)) },
    { Indicador: "Plan después de decisiones", Valor: Number(finalTotal.toFixed(2)) },
    { Indicador: "Aceptadas / rechazadas / pendientes", Valor: `${accepted.length} / ${rejected.length} / ${pending.length}` },
    { Indicador: "Nota general", Valor: review.generalNote || "" },
  ];
  const detail = rows.map((row) => ({
    Producto: row.producto,
    Categoria: row.categoria,
    Estatus: row.status,
    Existencias: row.inventory,
    "Pronóstico estadístico": Number(row.baseForecast.toFixed(2)),
    "Plan operativo base": Number(row.baseOperational.toFixed(2)),
    "Colchón recomendado %": row.marginPct,
    "Propuesta local": Number(row.proposed.toFixed(2)),
    Diferencia: Number(row.difference.toFixed(2)),
    Decisión: row.decision === "accepted" ? "Aceptada" : row.decision === "rejected" ? "Rechazada" : "Pendiente",
    "Plan resultante": Number((row.decision === "accepted" ? row.proposed : row.baseOperational).toFixed(2)),
    "Volatilidad histórica %": row.volatility === null ? "" : Number((row.volatility * 100).toFixed(1)),
    Motivos: row.reasons.join(" "),
    Nota: row.note,
  }));
  const methodology = [
    { Regla: "Separación", Detalle: "La revisión nunca modifica el pronóstico congelado; produce un plan operativo separado." },
    { Regla: "Sin look-ahead", Detalle: "Solo usa información anterior al mes objetivo, estatus, notas y existencias capturadas." },
    { Regla: "Volatilidad", Detalle: "Colchón de 8%, 12% o 15% según la variación de los últimos seis meses disponibles." },
    { Regla: "Existencias", Detalle: "Las existencias se descuentan del pronóstico con colchón antes de aplicar la regla de producción." },
    { Regla: "Estatus", Detalle: "BAJA y BAJO PEDIDO salen del plan regular. ESTACIONAL exige confirmación manual." },
  ];
  const workbook = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(workbook, XLSX.utils.json_to_sheet(summary), "Resumen");
  XLSX.utils.book_append_sheet(workbook, XLSX.utils.json_to_sheet(detail), "Propuestas");
  XLSX.utils.book_append_sheet(workbook, XLSX.utils.json_to_sheet(methodology), "Metodologia");
  XLSX.writeFile(workbook, `revision_asistida_${selectedMonth || "mes"}.xlsx`);
}

function exportFrozenForecast({ rows, selectedMonth, frozenAt, snapshotVersion, workspaceVersion }) {
  const baseTotal = rows.reduce((sum, row) => sum + row.pronosticoBase, 0);
  const marginTotal = rows.reduce((sum, row) => sum + row.margenOperativoPiezas, 0);
  const operationalTotal = rows.reduce((sum, row) => sum + row.pronosticoOperativo, 0);
  const summary = [
    { Indicador: "Mes pronosticado", Valor: selectedMonth },
    { Indicador: "Fecha de congelamiento", Valor: new Date(frozenAt).toLocaleString("es-MX") },
    { Indicador: "Version congelada", Valor: snapshotVersion },
    { Indicador: "Version del respaldo fuente", Valor: workspaceVersion || "Sin version" },
    { Indicador: "Modelo estadistico", Valor: FORECAST_MODEL_VERSION },
    { Indicador: "Productos", Valor: rows.length },
    { Indicador: "Pronostico estadistico", Valor: Number(baseTotal.toFixed(2)) },
    { Indicador: `Margen operativo ${OPERATIONAL_MARGIN_PCT}%`, Valor: Number(marginTotal.toFixed(2)) },
    { Indicador: `Escenario operativo +${OPERATIONAL_MARGIN_PCT}%`, Valor: Number(operationalTotal.toFixed(2)) },
    { Indicador: "Control", Valor: "El margen operativo es un escenario separado y no modifica el pronostico estadistico." },
  ];
  const products = rows.map((row) => ({
    Producto: row.producto,
    Categoria: row.categoria,
    "Pronostico estadistico": Number(row.pronosticoBase.toFixed(2)),
    [`Margen ${row.margenOperativoPct}% piezas`]: Number(row.margenOperativoPiezas.toFixed(2)),
    [`Escenario operativo +${row.margenOperativoPct}%`]: Number(row.pronosticoOperativo.toFixed(2)),
    "Metodo seleccionado": row.metodoPronostico,
    "Factor calibracion": `${(Number(row.tendenciaAplicada || 0) * 100).toFixed(1)}%`,
    "Meses usados": row.mesesUsados,
  }));
  const methodology = [
    { Concepto: "Pronostico estadistico", Detalle: "Resultado del modelo sin ventas del mes objetivo." },
    { Concepto: `Escenario +${OPERATIONAL_MARGIN_PCT}%`, Detalle: `Pronostico estadistico multiplicado por ${(1 + OPERATIONAL_MARGIN_PCT / 100).toFixed(2)}.` },
    { Concepto: "Congelamiento", Detalle: "Los valores se guardan por producto y no deben modificarse al evaluar el cierre." },
    { Concepto: "Comparacion", Detalle: "Al cierre se deben medir por separado el modelo estadistico y el escenario operativo." },
  ];
  const workbook = XLSX.utils.book_new();
  const appendSheet = (data, name, widths) => {
    const sheet = XLSX.utils.json_to_sheet(data);
    sheet["!cols"] = widths.map((wch) => ({ wch }));
    sheet["!autofilter"] = { ref: sheet["!ref"] };
    XLSX.utils.book_append_sheet(workbook, sheet, name);
  };
  appendSheet(summary, "Resumen", [38, 88]);
  appendSheet(products, "Por producto", [38, 24, 24, 20, 26, 30, 20, 32]);
  appendSheet(methodology, "Metodologia", [28, 100]);
  XLSX.writeFile(workbook, `pronostico_congelado_${selectedMonth || "mes"}.xlsx`);
}

function exportHistoricalValidation(rows) {
  const summary = [
    { Indicador: "Productos analizados", Valor: rows.length },
    { Indicador: "Precision promedio junio", Valor: `${(rows.filter((row) => row.precisionJunio !== null).reduce((sum, row) => sum + row.precisionJunio, 0) / Math.max(1, rows.filter((row) => row.precisionJunio !== null).length)).toFixed(1)}%` },
    { Indicador: "Bajas junio registradas", Valor: rows.reduce((sum, row) => sum + row.bajasJunio, 0) },
    { Indicador: "Bajas julio registradas", Valor: rows.reduce((sum, row) => sum + row.bajasJulio, 0) },
    { Indicador: "Produccion julio base", Valor: rows.reduce((sum, row) => sum + row.produccionSugeridaBase, 0) },
    { Indicador: "Produccion julio ajustada por bajas", Valor: rows.reduce((sum, row) => sum + row.produccionSugeridaAjustada, 0) },
  ];
  const detail = rows.map((row) => ({
    Producto: row.producto,
    "Venta mayo": Number(row.ventaMayo.toFixed(2)),
    "Producido mayo": Number(row.producidoMayo.toFixed(2)),
    "Diferencia mayo": Number(row.diferenciaMayo.toFixed(2)),
    "Venta junio": Number(row.ventaJunio.toFixed(2)),
    "Producido junio": Number(row.producidoJunio.toFixed(2)),
    "Bajas junio": Number(row.bajasJunio.toFixed(2)),
    "Demanda ajustada junio": Number(row.demandaAjustadaJunio.toFixed(2)),
    "Saldo junio": Number(row.saldoJunio.toFixed(2)),
    "Pronostico junio usando mayo": Number(row.pronosticoJunio.toFixed(2)),
    "Precision prueba junio": row.precisionJunio === null ? "Sin dato" : `${row.precisionJunio.toFixed(1)}%`,
    "Tasa bajas junio": `${(row.tasaBajas * 100).toFixed(1)}%`,
    "Bajas esperadas julio": Number(row.bajasEsperadasJulio.toFixed(2)),
    "Pronostico julio": Number(row.pronosticoJulio.toFixed(2)),
    "Margen de seguridad": Number(row.margenSeguridad.toFixed(2)),
    "Produccion julio base": row.produccionSugeridaBase,
    "Produccion julio ajustada por bajas": row.produccionSugeridaAjustada,
    "Bajas julio registradas": row.bajasJulio,
  }));
  const methodology = [
    { Concepto: "Demanda ajustada junio", Formula: "Venta junio + Bajas junio" },
    { Concepto: "Saldo junio", Formula: "Producido junio - Venta junio - Bajas junio" },
    { Concepto: "Tasa de bajas", Formula: "Bajas junio / Venta junio" },
    { Concepto: "Bajas esperadas julio", Formula: "Pronostico julio * Tasa de bajas" },
    { Concepto: "Produccion julio base", Formula: "Pronostico julio + margen de seguridad" },
    { Concepto: "Produccion julio ajustada", Formula: "Pronostico julio + bajas esperadas + margen" },
  ];
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(summary), "Resumen bajas");
  XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(detail), "Bajas y ajuste");
  XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(methodology), "Metodologia bajas");
  XLSX.writeFile(wb, "validacion_historica_ventas_producido_bajas.xlsx");
}

function exportToExcel(rows, summary) {
  const detalle = rows.map((r) => ({
    Producto: r.producto,
    "Promedio historico": Number(r.promedioHistorico.toFixed(2)),
    "Pronostico venta": Number(r.pronosticoVenta.toFixed(2)),
    "Metodo seleccionado": r.metodoPronostico,
    "Factor calibracion": `${(r.tendenciaAplicada * 100).toFixed(1)}%`,
    "Meses usados": r.mesesUsados,
    "Mes validacion modelo": r.mesValidacionModelo,
    "Real validacion modelo": Number(r.realValidacionModelo.toFixed(2)),
    "Pronostico validacion modelo": Number(r.pronosticoValidacionModelo.toFixed(2)),
    "Error validacion modelo": Number(r.errorValidacionModelo.toFixed(2)),
    "Margen de seguridad": Number(r.colchonOperativo.toFixed(2)),
    "Base con margen de seguridad": Number((r.baseConColchon || 0).toFixed(2)),
    "Regla operativa": r.reglaOperativa,
    "Produccion sugerida": r.produccionSugerida,
    "Produccion balanceada": Number((r.produccionBalanceada || 0).toFixed(2)),
    "Produccion recomendada": r.produccionRecomendada,
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
    { Indicador: "Margen de seguridad mensual", Valor: Number(summary.colchonDiarioMensual.toFixed(2)) },
    { Indicador: "Base con margen mensual", Valor: Number(summary.baseConColchonMensual.toFixed(2)) },
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
    "Margen de seguridad diario": Number(row.colchonDiario.toFixed(2)),
    "Base con margen de seguridad": Number((row.baseConColchonDia || 0).toFixed(2)),
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

function exportWeeklyProgress(progress, selectedWeek, selectedMonth) {
  const summary = [
    { Indicador: "Mes", Valor: selectedMonth },
    { Indicador: "Semana seleccionada", Valor: selectedWeek?.label || "Sin semana" },
    { Indicador: "Corte de venta real", Valor: progress.month.cutoffDate || "Sin venta diaria" },
    { Indicador: "Pronostico mensual", Valor: Number(progress.month.pronosticoPeriodo.toFixed(2)) },
    { Indicador: "Pronostico acumulado al corte", Valor: Number(progress.month.pronosticoCorte.toFixed(2)) },
    { Indicador: "Venta real acumulada", Valor: Number(progress.month.ventaReal.toFixed(2)) },
    { Indicador: "Cumplimiento acumulado", Valor: progress.month.cumplimiento === null ? "Sin dato" : `${progress.month.cumplimiento.toFixed(1)}%` },
    { Indicador: "Proyeccion de cierre", Valor: Number(progress.month.proyeccionPeriodo.toFixed(2)) },
    { Indicador: "Diferencia proyectada", Valor: Number(progress.month.projectedDifference.toFixed(2)) },
    { Indicador: "Cobertura de fechas", Valor: `${progress.month.coveragePct.toFixed(1)}%` },
  ];
  const weeks = progress.weeks.map((week) => ({
    Semana: week.label,
    "Pronostico completo": Number(week.pronosticoPeriodo.toFixed(2)),
    "Pronostico al corte": Number(week.pronosticoCorte.toFixed(2)),
    "Venta real": Number(week.ventaReal.toFixed(2)),
    Diferencia: Number(week.diferencia.toFixed(2)),
    "Cumplimiento %": week.cumplimiento === null ? "" : Number(week.cumplimiento.toFixed(1)),
    "Proyeccion semanal": Number(week.proyeccionPeriodo.toFixed(2)),
    "Dias comparados": week.comparedDays,
    "Cobertura %": Number(week.coveragePct.toFixed(1)),
    Estado: week.status.label,
  }));
  const products = (selectedWeek?.products || []).map((row) => ({
    Producto: row.producto,
    Categoria: row.categoria,
    "Pronostico semana": Number(row.pronosticoPeriodo.toFixed(2)),
    "Pronostico al corte": Number(row.pronosticoCorte.toFixed(2)),
    "Venta real": Number(row.ventaReal.toFixed(2)),
    Diferencia: Number(row.diferencia.toFixed(2)),
    "Cumplimiento %": row.cumplimiento === null ? "" : Number(row.cumplimiento.toFixed(1)),
    "Proyeccion semanal": Number(row.proyeccionPeriodo.toFixed(2)),
    Estado: row.status.label,
  }));
  const categories = (selectedWeek?.categories || []).map((row) => ({
    Categoria: row.categoria,
    Productos: row.productos,
    "Pronostico semana": Number(row.pronosticoPeriodo.toFixed(2)),
    "Pronostico al corte": Number(row.pronosticoCorte.toFixed(2)),
    "Venta real": Number(row.ventaReal.toFixed(2)),
    Diferencia: Number(row.diferencia.toFixed(2)),
    "Cumplimiento %": row.cumplimiento === null ? "" : Number(row.cumplimiento.toFixed(1)),
    "Proyeccion semanal": Number(row.proyeccionPeriodo.toFixed(2)),
    Estado: row.status.label,
  }));

  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(summary), "Resumen semanal");
  XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(weeks), "Semanas");
  XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(products), "Productos");
  XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(categories), "Categorias");
  XLSX.writeFile(wb, `avance_semanal_${selectedMonth || "mes"}.xlsx`);
}

function exportMonthlyClose(close, selectedMonth) {
  const valueOrBlank = (value, digits = 2) => Number.isFinite(value) ? Number(value.toFixed(digits)) : "Sin dato";
  const summary = [
    { Indicador: "Mes", Valor: selectedMonth },
    { Indicador: "Productos comparados", Valor: close.summary.productos },
    { Indicador: "Pronostico de venta", Valor: valueOrBlank(close.summary.pronostico) },
    { Indicador: "Venta real", Valor: valueOrBlank(close.summary.ventaReal) },
    { Indicador: "Diferencia real vs pronostico", Valor: valueOrBlank(close.summary.diferenciaPronostico) },
    { Indicador: "Cumplimiento", Valor: close.summary.cumplimiento === null ? "Sin dato" : `${close.summary.cumplimiento.toFixed(1)}%` },
    { Indicador: "WAPE", Valor: close.summary.wape === null ? "Sin dato" : `${close.summary.wape.toFixed(2)}%` },
    { Indicador: "MAE", Valor: valueOrBlank(close.summary.mae) },
    { Indicador: "Productos dentro de +/-15", Valor: close.summary.dentro15 },
    { Indicador: "Produccion real", Valor: valueOrBlank(close.summary.producido) },
    { Indicador: "Producido menos vendido", Valor: valueOrBlank(close.summary.diferenciaProduccion) },
    { Indicador: "Venta fuera del catalogo regular", Valor: valueOrBlank(close.unmatchedSalesTotal) },
    { Indicador: "Produccion fuera del catalogo regular", Valor: valueOrBlank(close.unmatchedProductionTotal) },
  ];
  const products = close.rows.map((row) => ({
    Producto: row.producto,
    Categoria: row.categoria,
    Pronostico: valueOrBlank(row.pronostico),
    "Venta real": valueOrBlank(row.ventaReal),
    "Diferencia pronostico": valueOrBlank(row.diferenciaPronostico),
    "Error absoluto": valueOrBlank(row.errorAbsoluto),
    "Cumplimiento %": valueOrBlank(row.cumplimiento, 1),
    Producido: valueOrBlank(row.producido),
    "Producido menos vendido": valueOrBlank(row.diferenciaProduccion),
    Estado: row.status.label,
  }));
  const categories = close.categories.map((row) => ({
    Categoria: row.categoria,
    Productos: row.productos,
    Pronostico: valueOrBlank(row.pronostico),
    "Venta real": valueOrBlank(row.ventaReal),
    "Diferencia pronostico": valueOrBlank(row.diferenciaPronostico),
    "WAPE %": valueOrBlank(row.wape, 2),
    "MAE": valueOrBlank(row.mae),
    Producido: valueOrBlank(row.producido),
    "Producido menos vendido": valueOrBlank(row.diferenciaProduccion),
  }));
  const unmatchedSales = close.unmatchedSales.map((row) => ({ Producto: row.producto, "Venta fuera de catalogo": row.cantidad }));
  const unmatchedProduction = close.unmatchedProduction.map((row) => ({ Producto: row.producto, "Produccion fuera de catalogo": row.cantidad }));

  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(summary), "Resumen cierre");
  XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(products), "Por producto");
  XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(categories), "Por categoria");
  XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(unmatchedSales), "Venta no comparable");
  XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(unmatchedProduction), "Producido no comparable");
  XLSX.writeFile(wb, `cierre_mensual_${selectedMonth || "mes"}.xlsx`);
}

function exportValidationToExcel({
  selectedMonth,
  dailyBufferPct,
  salesSummary,
  productSummary,
  dailyRows,
  forecastRows,
  homologationRows,
  historicalVentas,
  alerts,
}) {
  const resumen = [
    { Indicador: "Mes validado", Valor: selectedMonth },
    { Indicador: "Margen de seguridad", Valor: `${dailyBufferPct}%` },
    { Indicador: "Pronostico venta total", Valor: Number(salesSummary.pronosticoTotal.toFixed(2)) },
    { Indicador: "Venta real total", Valor: Number(salesSummary.ventaRealTotal.toFixed(2)) },
    { Indicador: "Diferencia total", Valor: Number(salesSummary.diferenciaTotal.toFixed(2)) },
    { Indicador: "Precision global", Valor: `${salesSummary.precisionGlobal.toFixed(1)}%` },
    { Indicador: "Dias con venta real", Valor: salesSummary.diasConReal },
    { Indicador: "Productos analizados", Valor: productSummary.length },
    {
      Indicador: "Productos con precision < 80%",
      Valor: productSummary.filter((row) => row.precision !== null && row.precision < 80).length,
    },
  ];

  const porProducto = productSummary.map((row) => ({
    Producto: row.producto,
    "Pronostico mensual": Number(row.pronosticoMensual.toFixed(2)),
    "Venta real mensual": Number(row.ventaRealMensual.toFixed(2)),
    Diferencia: Number(row.diferencia.toFixed(2)),
    "Error %": row.errorPct === null ? "" : Number(row.errorPct.toFixed(1)),
    "Precision %": row.precision === null ? "" : Number(row.precision.toFixed(1)),
    "Dias con real": row.diasConReal,
    "Registros historicos": row.registrosHistoricos,
    Estatus: STATUS_META[row.estatus]?.label || row.estatus,
  }));

  const realVsPronostico = dailyRows
    .filter((row) => row.hasVentaReal)
    .map((row) => ({
      Fecha: row.fechaDisplay,
      Dia: row.dia,
      Producto: row.producto,
      "Pronostico venta": Number(row.pronosticoVentaDia.toFixed(2)),
      "Venta real": Number(row.ventaRealDia.toFixed(2)),
      Diferencia: Number(row.diferenciaVenta.toFixed(2)),
      "Precision %": row.precisionVenta === null ? "" : Number(row.precisionVenta.toFixed(1)),
      Estatus: STATUS_META[row.estatusVenta]?.label || row.estatusVenta,
    }));

  const promedios = forecastRows.map((row) => ({
    Producto: row.producto,
    Lunes: Number(row.promedioLunes.toFixed(2)),
    Martes: Number(row.promedioMartes.toFixed(2)),
    Miercoles: Number(row.promedioMiercoles.toFixed(2)),
    Jueves: Number(row.promedioJueves.toFixed(2)),
    Viernes: Number(row.promedioViernes.toFixed(2)),
    Sabado: Number(row.promedioSabado.toFixed(2)),
    Domingo: Number(row.promedioDomingo.toFixed(2)),
    "Registros historicos": row.registrosHistoricos,
  }));

  const homologacion = homologationRows.map((row) => ({
    "Producto leido": row.product,
    "Nombre original": row.originalNames.join(", "),
    Origen: row.sources.join(", "),
    Registros: row.count,
    "Producto oficial": row.official || "",
    Estado: row.status,
  }));

  const alertas = alerts.map((row) => ({
    Tipo: row.tipo,
    Producto: row.producto,
    Detalle: row.detalle,
    Severidad: row.severidad,
  }));

  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(resumen), "Resumen ejecutivo");
  XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(porProducto), "Por producto");
  XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(realVsPronostico), "Real vs pronostico");
  XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(promedios), "Promedios semana");
  if (homologacion.length) {
    XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(homologacion), "Homologacion");
  }
  if (alertas.length) {
    XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(alertas), "Alertas");
  }
  XLSX.writeFile(wb, `validacion_pronostico_${selectedMonth || "mes"}.xlsx`);
}

function UploadBox({ title, description, onFile, fileName, required, accept = ".xlsx,.xls", multiple = false }) {
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
        <input
          type="file"
          accept={accept}
          multiple={multiple}
          onChange={(e) => onFile(multiple ? Array.from(e.target.files || []) : e.target.files?.[0])}
        />
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

function summarizeOperationalRows(rows, dimensionValues) {
  const dailyRows = [];
  const dimensions = new Set();
  const grouped = new Set();
  let monthlyTotals = 0;
  let invalid = 0;
  let repeated = 0;
  for (const row of rows) {
    if (row.monthlyTotal) {
      monthlyTotals += 1;
      continue;
    }
    const fecha = dateKey(row.fecha);
    const product = normalizeProduct(row.producto);
    if (!fecha || !product || !Number.isFinite(Number(row.cantidad))) {
      invalid += 1;
      continue;
    }
    const values = dimensionValues(row).map((value) => String(value || "").trim());
    values.filter(Boolean).forEach((value) => dimensions.add(value));
    const key = [fecha, product, ...values.map(norm)].join("|");
    if (grouped.has(key)) repeated += 1;
    else grouped.add(key);
    dailyRows.push(row);
  }
  const dates = dailyRows.map((row) => dateKey(row.fecha)).filter(Boolean).sort();
  return {
    recognized: rows.length,
    daily: dailyRows.length,
    monthlyTotals,
    invalid,
    repeated,
    dimensions: dimensions.size,
    firstDate: dates[0] || "",
    lastDate: dates[dates.length - 1] || "",
  };
}

function mergeRemoteRecords(localRows, remoteRows, keyForRow) {
  const remoteKeys = new Set(remoteRows.map(keyForRow));
  return [...localRows.filter((row) => !remoteKeys.has(keyForRow(row))), ...remoteRows];
}

function salesRecordKey(row) {
  return [dateKey(row.fecha), normalizeProduct(row.producto || row.producto_codigo), norm(row.sucursal || row.canal), norm(row.cliente)].join("|");
}

function productionRecordKey(row) {
  return [dateKey(row.fecha), normalizeProduct(row.producto || row.producto_codigo), norm(row.turno)].join("|");
}

function wasteRecordKey(row) {
  return [dateKey(row.fecha), normalizeProduct(row.producto || row.producto_codigo), norm(row.sucursal || row.canal), norm(row.motivo)].join("|");
}

function snapshotOperationalRows(rows, keyForRow, persistedKeys = new Set()) {
  return rows.filter((row) => row.monthlyTotal || !dateKey(row.fecha) || (!row.databaseSynced && !persistedKeys.has(keyForRow(row))));
}

function OperationalImportPanel({
  title,
  description,
  dimensionLabel,
  pendingRows,
  preview,
  status,
  importing,
  canSave,
  onImport,
}) {
  if (!pendingRows.length && !status) return null;
  return (
    <section className="sales-import-section">
      <div className="section-heading compact-heading">
        <div>
          <span className="eyebrow">Carga masiva diaria</span>
          <h3>{title}</h3>
          <p>{description}</p>
        </div>
        {pendingRows.length > 0 && (
          <button className="primary" type="button" onClick={onImport} disabled={!canSave || importing || preview.daily === 0}>
            <Database size={18} /> {importing ? "Guardando..." : "Guardar en la base"}
          </button>
        )}
      </div>
      {pendingRows.length > 0 && (
        <div className="sales-import-grid">
          <div><span>Filas reconocidas</span><strong>{formatNumber(preview.recognized)}</strong></div>
          <div><span>Registros diarios</span><strong>{formatNumber(preview.daily)}</strong></div>
          <div><span>Filas a consolidar</span><strong>{formatNumber(preview.repeated)}</strong></div>
          <div><span>Totales mensuales excluidos</span><strong>{formatNumber(preview.monthlyTotals)}</strong></div>
          <div><span>{dimensionLabel}</span><strong>{formatNumber(preview.dimensions)}</strong></div>
          <div><span>Rango diario</span><strong>{preview.firstDate ? `${preview.firstDate} a ${preview.lastDate}` : "Sin fechas"}</strong></div>
        </div>
      )}
      <p className={`sales-import-status ${status.startsWith("No ") ? "error" : ""}`}>{status}</p>
      {preview.monthlyTotals > 0 && pendingRows.length > 0 && (
        <p className="sales-import-note">Los totales mensuales no se convierten en registros diarios.</p>
      )}
    </section>
  );
}

function AccessScreen({ needsSetup, loading, error, onSubmit }) {
  const [nombre, setNombre] = useState("");
  const [usuario, setUsuario] = useState("");
  const [password, setPassword] = useState("");
  const [setupKey, setSetupKey] = useState("");

  function submit(event) {
    event.preventDefault();
    onSubmit({ nombre, usuario, password, setupKey });
  }

  return (
    <main className="access-page">
      <section className="access-card">
        <div className="access-mark"><ShieldCheck size={26} /></div>
        <span className="eyebrow">Archivo Maestro</span>
        <h1>{needsSetup ? "Crear administrador" : "Acceso operativo"}</h1>
        <p>
          {needsSetup
            ? "Configura la primera cuenta para proteger los datos y activar el historial."
            : "Inicia sesión para consultar y respaldar la planeación."}
        </p>
        <form onSubmit={submit} className="access-form">
          {needsSetup && (
            <label>
              Nombre
              <input value={nombre} onChange={(event) => setNombre(event.target.value)} autoComplete="name" required />
            </label>
          )}
          <label>
            Usuario
            <input value={usuario} onChange={(event) => setUsuario(event.target.value)} autoComplete="username" required />
          </label>
          <label>
            Contraseña
            <input
              type="password"
              value={password}
              onChange={(event) => setPassword(event.target.value)}
              autoComplete={needsSetup ? "new-password" : "current-password"}
              minLength={needsSetup ? 10 : undefined}
              required
            />
          </label>
          {needsSetup && (
            <label>
              Clave de instalación
              <input type="password" value={setupKey} onChange={(event) => setSetupKey(event.target.value)} required />
            </label>
          )}
          {error && <div className="access-error">{error}</div>}
          <button className="primary access-submit" type="submit" disabled={loading}>
            <UserRound size={18} /> {loading ? "Validando..." : needsSetup ? "Crear cuenta" : "Iniciar sesión"}
          </button>
        </form>
      </section>
    </main>
  );
}

function Dashboard({ session, onLogout }) {
  const [stockRows, setStockRows] = useState([]);
  const [ventas, setVentas] = useState([]);
  const [ventasValidacion, setVentasValidacion] = useState([]);
  const [bajas, setBajas] = useState([]);
  const [existencias, setExistencias] = useState([]);
  const [realProduction, setRealProduction] = useState([]);
  const [producedMay, setProducedMay] = useState([]);
  const [producedJune, setProducedJune] = useState([]);
  const [bajasJune, setBajasJune] = useState([]);
  const [bajasJuly, setBajasJuly] = useState([]);
  const [monthlyCloseSales, setMonthlyCloseSales] = useState([]);
  const [monthlyCloseProduction, setMonthlyCloseProduction] = useState([]);
  const [monthlyClosePeriod, setMonthlyClosePeriod] = useState("");
  const [files, setFiles] = useState({});
  const [query, setQuery] = useState("");
  const [showMissingReal, setShowMissingReal] = useState(false);
  const [selectedMonth, setSelectedMonth] = useState(defaultMonthValue());
  const [selectedMonthTouched, setSelectedMonthTouched] = useState(false);
  const [dailyBufferPct, setDailyBufferPct] = useState(10);
  const [dailyDateFilter, setDailyDateFilter] = useState("");
  const [dailyProductQuery, setDailyProductQuery] = useState("");
  const [dailyWeekdayFilter, setDailyWeekdayFilter] = useState("");
  const [selectedWeekKey, setSelectedWeekKey] = useState("");
  const [onlyDailyShortage, setOnlyDailyShortage] = useState(false);
  const [onlyDailyOverproduction, setOnlyDailyOverproduction] = useState(false);
  const [onlySalesMismatch, setOnlySalesMismatch] = useState(false);
  const [validationProduct, setValidationProduct] = useState("");
  const [productAliases, setProductAliases] = useState(loadStoredProductAliases);
  const [cloudStatus, setCloudStatus] = useState("Buscando respaldo...");
  const [cloudSaving, setCloudSaving] = useState(false);
  const [forecastFreezing, setForecastFreezing] = useState(false);
  const [monthlyReview, setMonthlyReview] = useState({
    period: "",
    state: "draft",
    generalNote: "",
    inputs: {},
    version: null,
  });
  const [monthlyReviewSource, setMonthlyReviewSource] = useState(null);
  const [monthlyReviewLoading, setMonthlyReviewLoading] = useState(false);
  const [monthlyReviewSaving, setMonthlyReviewSaving] = useState(false);
  const [monthlyReviewQuery, setMonthlyReviewQuery] = useState("");
  const [monthlyReviewFilter, setMonthlyReviewFilter] = useState("all");
  const [hasUnsavedChanges, setHasUnsavedChanges] = useState(false);
  const [lastBackup, setLastBackup] = useState(null);
  const [showUserAdmin, setShowUserAdmin] = useState(false);
  const [users, setUsers] = useState([]);
  const [userStatus, setUserStatus] = useState("");
  const [userForm, setUserForm] = useState({ nombre: "", usuario: "", password: "", rol: "operador" });
  const [pendingSalesImport, setPendingSalesImport] = useState([]);
  const [salesImportFiles, setSalesImportFiles] = useState([]);
  const [salesImporting, setSalesImporting] = useState(false);
  const [salesImportStatus, setSalesImportStatus] = useState("");
  const [pendingProductionImport, setPendingProductionImport] = useState([]);
  const [productionImportFile, setProductionImportFile] = useState("");
  const [productionImporting, setProductionImporting] = useState(false);
  const [productionImportStatus, setProductionImportStatus] = useState("");
  const [pendingWasteImport, setPendingWasteImport] = useState([]);
  const [wasteImportFile, setWasteImportFile] = useState("");
  const [wasteImporting, setWasteImporting] = useState(false);
  const [wasteImportStatus, setWasteImportStatus] = useState("");
  const [toast, setToast] = useState(null);

  const salesImportPreview = useMemo(() => {
    const dailyRows = [];
    let monthlyTotals = 0;
    let invalid = 0;
    const branches = new Set();
    const grouped = new Set();
    let repeated = 0;
    for (const row of pendingSalesImport) {
      if (row.monthlyTotal) {
        monthlyTotals += 1;
        continue;
      }
      const fecha = dateKey(row.fecha);
      const product = normalizeProduct(row.producto);
      if (!fecha || !product || !Number.isFinite(Number(row.cantidad))) {
        invalid += 1;
        continue;
      }
      const branch = String(row.sucursal || row.canal || "").trim();
      if (branch) branches.add(branch);
      const key = `${fecha}|${product}|${norm(branch)}|${norm(row.cliente)}`;
      if (grouped.has(key)) repeated += 1;
      else grouped.add(key);
      dailyRows.push(row);
    }
    const dates = dailyRows.map((row) => dateKey(row.fecha)).filter(Boolean).sort();
    return {
      recognized: pendingSalesImport.length,
      daily: dailyRows.length,
      monthlyTotals,
      invalid,
      repeated,
      branches: branches.size,
      firstDate: dates[0] || "",
      lastDate: dates[dates.length - 1] || "",
    };
  }, [pendingSalesImport]);
  const productionImportPreview = useMemo(
    () => summarizeOperationalRows(pendingProductionImport, (row) => [row.turno]),
    [pendingProductionImport]
  );
  const wasteImportPreview = useMemo(
    () => summarizeOperationalRows(pendingWasteImport, (row) => [row.sucursal || row.canal, row.motivo]),
    [pendingWasteImport]
  );

  useEffect(() => {
    if (!toast) return undefined;
    const timer = window.setTimeout(() => setToast(null), 4500);
    return () => window.clearTimeout(timer);
  }, [toast]);

  useEffect(() => {
    let active = true;
    async function restoreData() {
      let data = {};
      let snapshot = null;
      try {
        const response = await apiRequest("/api/snapshots/workspace", { token: session.token });
        snapshot = response.snapshot;
        data = snapshot.contenido || {};
      } catch (error) {
        if (error.status !== 404) throw error;
      }

      const [salesSync, productionSync, wasteSync] = await Promise.all([
        apiRequestAllRows("/api/ventas/sync", { token: session.token }).catch(() => ({ rows: [] })),
        apiRequestAllRows("/api/produccion-real/sync", { token: session.token }).catch(() => ({ rows: [] })),
        apiRequestAllRows("/api/bajas/sync", { token: session.token }).catch(() => ({ rows: [] })),
      ]);
      if (!active) return;

      const remoteSales = (salesSync.rows || []).map((row) => ({
        fecha: row.fecha,
        producto: normalizeProduct(row.producto_codigo),
        productoOriginal: row.producto_nombre,
        cantidad: Number(row.cantidad),
        importe: Number(row.importe || 0),
        sucursal: row.canal || "",
        canal: row.canal || "",
        cliente: row.cliente || "",
        tipo: "ventas",
        sourceFile: "Base de datos",
        databaseSynced: true,
      }));
      const remoteProduction = (productionSync.rows || []).map((row) => ({
        fecha: row.fecha,
        fechaKey: row.fecha,
        producto: normalizeProduct(row.producto_codigo),
        productoOriginal: row.producto_nombre,
        cantidad: Number(row.cantidad),
        turno: row.turno || "",
        sourceFile: "Base de datos",
        databaseSynced: true,
      }));
      const remoteWaste = (wasteSync.rows || []).map((row) => ({
        fecha: row.fecha,
        producto: normalizeProduct(row.producto_codigo),
        productoOriginal: row.producto_nombre,
        cantidad: Number(row.cantidad),
        sucursal: row.sucursal || "",
        canal: row.sucursal || "",
        motivo: row.motivo || "",
        tipo: "bajas",
        sourceFile: "Base de datos",
        databaseSynced: true,
      }));
      const localSales = Array.isArray(data.ventas) ? data.ventas : [];
      const localProduction = Array.isArray(data.realProduction) ? data.realProduction : [];
      const localWaste = Array.isArray(data.bajas) ? data.bajas : [];
      setStockRows(Array.isArray(data.stockRows) ? data.stockRows : []);
      setVentas(mergeRemoteRecords(localSales, remoteSales, salesRecordKey));
      setVentasValidacion(Array.isArray(data.ventasValidacion) ? data.ventasValidacion : []);
      setBajas(mergeRemoteRecords(localWaste, remoteWaste, wasteRecordKey));
      setExistencias(Array.isArray(data.existencias) ? data.existencias : []);
      setRealProduction(mergeRemoteRecords(localProduction, remoteProduction, productionRecordKey));
      setProducedMay(Array.isArray(data.producedMay) ? data.producedMay : []);
      setProducedJune(Array.isArray(data.producedJune) ? data.producedJune : []);
      setBajasJune(Array.isArray(data.bajasJune) ? data.bajasJune : []);
      setBajasJuly(Array.isArray(data.bajasJuly) ? data.bajasJuly : []);
      setMonthlyCloseSales(Array.isArray(data.monthlyCloseSales) ? data.monthlyCloseSales : []);
      setMonthlyCloseProduction(Array.isArray(data.monthlyCloseProduction) ? data.monthlyCloseProduction : []);
      setMonthlyClosePeriod(String(data.monthlyClosePeriod || ""));
      setFiles(data.files && typeof data.files === "object" ? data.files : {});
      setProductAliases(data.productAliases && typeof data.productAliases === "object" ? data.productAliases : {});
      if (data.selectedMonth) setSelectedMonth(data.selectedMonth);
      if (Number.isFinite(Number(data.dailyBufferPct))) setDailyBufferPct(Number(data.dailyBufferPct));
      setLastBackup(snapshot);
      setHasUnsavedChanges(false);
      setCloudStatus(
        `${snapshot ? `Respaldo v${snapshot.version} restaurado. ` : ""}Base sincronizada: ${remoteSales.length} ventas, ${remoteProduction.length} producciones y ${remoteWaste.length} bajas.`
      );
    }
    restoreData().catch((error) => {
      if (!active) return;
      setCloudStatus(`No se pudo restaurar: ${error.message}`);
    });
    return () => {
      active = false;
    };
  }, [session.token]);

  useEffect(() => {
    if (!selectedMonth) return undefined;
    let active = true;
    setMonthlyReviewLoading(true);
    setMonthlyReviewSource(null);
    setMonthlyReview({
      period: selectedMonth,
      state: "draft",
      generalNote: "",
      inputs: {},
      version: null,
    });

    async function loadMonthlyReview() {
      const [frozenResult, reviewResult] = await Promise.all([
        apiRequest(`/api/snapshots/forecast-frozen?periodo=${encodeURIComponent(selectedMonth)}`, {
          token: session.token,
        }).catch((error) => {
          if (error.status === 404) return null;
          throw error;
        }),
        apiRequest(`/api/snapshots/monthly-review?periodo=${encodeURIComponent(selectedMonth)}`, {
          token: session.token,
        }).catch((error) => {
          if (error.status === 404) return null;
          throw error;
        }),
      ]);
      if (!active) return;
      setMonthlyReviewSource(frozenResult?.snapshot || null);
      const saved = reviewResult?.snapshot;
      const content = saved?.contenido;
      if (saved && content && content.period === selectedMonth) {
        setMonthlyReview({
          period: selectedMonth,
          state: content.state === "approved" ? "approved" : "draft",
          generalNote: String(content.generalNote || ""),
          inputs: content.inputs && typeof content.inputs === "object" ? content.inputs : {},
          version: saved.version,
          savedAt: content.savedAt || saved.created_at,
          approvedAt: content.approvedAt || null,
          sourceFrozenVersion: content.sourceFrozenVersion || null,
        });
      }
    }

    loadMonthlyReview()
      .catch((error) => {
        if (!active) return;
        if (error.status === 401) onLogout();
        else setToast({ tone: "error", message: `No se pudo cargar la revisión mensual: ${error.message}` });
      })
      .finally(() => {
        if (active) setMonthlyReviewLoading(false);
      });
    return () => {
      active = false;
    };
  }, [selectedMonth, session.token]);

  async function saveWorkspace({ silent = false, persistedKeys = {} } = {}) {
    setCloudSaving(true);
    if (!silent) setCloudStatus("Guardando respaldo...");
    try {
      const contenido = {
        schemaVersion: 2,
        savedAt: new Date().toISOString(),
        selectedMonth,
        dailyBufferPct,
        stockRows,
        ventas: snapshotOperationalRows(ventas, salesRecordKey, persistedKeys.sales),
        ventasValidacion,
        bajas: snapshotOperationalRows(bajas, wasteRecordKey, persistedKeys.waste),
        existencias,
        realProduction: snapshotOperationalRows(realProduction, productionRecordKey, persistedKeys.production),
        producedMay,
        producedJune,
        bajasJune,
        bajasJuly,
        monthlyCloseSales,
        monthlyCloseProduction,
        monthlyClosePeriod,
        files,
        productAliases,
      };
      const payloadBytes = new TextEncoder().encode(JSON.stringify(contenido)).byteLength;
      if (payloadBytes > MAX_SNAPSHOT_BYTES) {
        throw new Error("El respaldo aún contiene demasiados datos sin sincronizar. Guarda primero las ventas, producción y bajas en la base de datos.");
      }
      const response = await apiRequest("/api/snapshots/workspace", {
        token: session.token,
        method: "POST",
        body: {
          periodo: "global",
          archivos: files,
          contenido,
        },
      });
      const saved = { version: response.version, created_at: new Date().toISOString() };
      setLastBackup(saved);
      setHasUnsavedChanges(false);
      setCloudStatus(`Respaldo v${response.version} guardado correctamente`);
      if (!silent) setToast({ tone: "success", message: `Respaldo v${response.version} guardado.` });
      return { ok: true, version: response.version };
    } catch (error) {
      if (error.status === 401) onLogout();
      else setCloudStatus(`No se pudo guardar: ${error.message}`);
      if (!silent) setToast({ tone: "error", message: `No se pudo guardar el respaldo: ${error.message}` });
      return { ok: false, error };
    } finally {
      setCloudSaving(false);
    }
  }

  async function toggleUsers() {
    const next = !showUserAdmin;
    setShowUserAdmin(next);
    if (!next || session.user.rol !== "admin") return;
    setUserStatus("Cargando usuarios...");
    try {
      const response = await apiRequest("/api/auth/users", { token: session.token });
      setUsers(response.rows || []);
      setUserStatus("");
    } catch (error) {
      setUserStatus(error.message);
    }
  }

  async function createUser(event) {
    event.preventDefault();
    setUserStatus("Creando usuario...");
    try {
      await apiRequest("/api/auth/users", {
        token: session.token,
        method: "POST",
        body: userForm,
      });
      const response = await apiRequest("/api/auth/users", { token: session.token });
      setUsers(response.rows || []);
      setUserForm({ nombre: "", usuario: "", password: "", rol: "operador" });
      setUserStatus("Usuario creado correctamente");
    } catch (error) {
      setUserStatus(error.message);
    }
  }

  async function handleFile(file, parser, key, setter) {
    if (!file) return;
    const wb = await readWorkbook(file);
    setter(parser(wb));
    setFiles((f) => ({ ...f, [key]: file.name }));
    setHasUnsavedChanges(true);
  }

  async function handleSalesFiles(selectedFiles) {
    const filesToRead = Array.isArray(selectedFiles) ? selectedFiles : [selectedFiles];
    const loadedNames = new Set(String(files.ventas || "").split(", ").filter(Boolean));
    const validFiles = filesToRead.filter(Boolean);
    if (!validFiles.length) return;

    const hasCanonicalJuneClose = validFiles.some((file) => {
      const hint = inferMonthHintFromFileName(file.name);
      return hint?.monthIndex === 5 && !norm(file.name).includes("MAYO");
    });
    const parsedFiles = await Promise.all(
      validFiles.map(async (file) => {
        const workbook = await readWorkbook(file);
        let rows = parseSalesOrReturns(workbook, "ventas", file.name);
        if (hasCanonicalJuneClose && norm(file.name).includes("MAYO") && norm(file.name).includes("JUNIO")) {
          rows = rows.filter((row) => dateKey(row.fecha).slice(0, 7) !== "2026-06");
        }
        return rows.map((row) => ({ ...row, sourceFile: file.name }));
      })
    );
    const parsed = parsedFiles.flat();
    setVentas((current) => {
      const selectedNames = new Set(validFiles.map((file) => file.name));
      const withoutReplacedFiles = current.filter((row) => {
        if (row.sourceFile && selectedNames.has(row.sourceFile)) return false;
        if (hasCanonicalJuneClose && norm(row.sourceFile).includes("MAYO") && norm(row.sourceFile).includes("JUNIO")) {
          return dateKey(row.fecha).slice(0, 7) !== "2026-06";
        }
        return true;
      });
      const rowsToAppend = parsedFiles.flatMap((rows, index) => {
        const fileName = validFiles[index].name;
        const hasTaggedRows = current.some((row) => row.sourceFile === fileName);
        return loadedNames.has(fileName) && !hasTaggedRows ? [] : rows;
      });
      return [...withoutReplacedFiles, ...rowsToAppend];
    });
    setPendingSalesImport(parsed);
    setSalesImportFiles(validFiles.map((file) => file.name));
    setSalesImportStatus(parsed.length ? "Revisa el resumen antes de guardar en la base." : "No se reconocieron ventas en los archivos.");
    setFiles((current) => ({
      ...current,
      ventas: [...new Set([...loadedNames, ...validFiles.map((file) => file.name)])].join(", "),
    }));
    setHasUnsavedChanges(true);
  }

  async function importSalesToDatabase() {
    if (!pendingSalesImport.length) return;
    setSalesImporting(true);
    setSalesImportStatus("Validando y guardando ventas...");
    try {
      const mappedRows = pendingSalesImport.map((row) => {
        const product = resolveOfficialProduct(row.producto, productAliases, officialProducts);
        return {
          fecha: dateKey(row.fecha),
          producto_codigo: product,
          producto_nombre: product,
          cantidad: Number(row.cantidad),
          importe: Number.isFinite(Number(row.importe)) ? Number(row.importe) : null,
          sucursal: row.sucursal || row.canal || "",
          cliente: row.cliente || "",
          monthlyTotal: Boolean(row.monthlyTotal),
        };
      });
      const rows = consolidateSalesRowsForUpload(mappedRows);
      if (!rows.length) throw new Error("El archivo no contiene ventas diarias para guardar en la base de datos");
      const response = await uploadRowsInBatches({
        endpoint: "/api/ventas/importar",
        bodyKey: "ventas",
        rows,
        archivo: salesImportFiles.join(", "),
        token: session.token,
        onProgress: (batch, total) => setSalesImportStatus(`Guardando ventas: lote ${batch} de ${total}...`),
      });
      const importedKeys = new Set(rows.map(salesRecordKey));
      setVentas((current) => current.map((row) => importedKeys.has(salesRecordKey(row)) ? { ...row, databaseSynced: true } : row));
      setPendingSalesImport([]);
      setSalesImportStatus("");
      const backup = await saveWorkspace({ silent: true, persistedKeys: { sales: importedKeys } });
      setToast({
        tone: backup.ok ? "success" : "warning",
        message: backup.ok
          ? `Ventas guardadas: ${response.inserted} nuevas, ${response.updated} actualizadas. Respaldo v${backup.version} creado.`
          : `Ventas guardadas en la base de datos, pero el respaldo automático falló.`,
      });
    } catch (error) {
      if (error.status === 401) onLogout();
      else setSalesImportStatus(error.message);
    } finally {
      setSalesImporting(false);
    }
  }

  async function handleProductionReal(file) {
    if (!file) return;
    const parsed = (await parseProductionRealFile(file)).map((row) => ({ ...row, sourceFile: file.name }));
    setRealProduction(parsed);
    setPendingProductionImport(parsed);
    setProductionImportFile(file.name);
    setProductionImportStatus(parsed.length ? "Revisa la producción antes de guardarla en la base." : "No se reconocieron registros de producción.");
    setFiles((f) => ({ ...f, real: file.name }));
    setHasUnsavedChanges(true);
  }

  async function handleWasteFile(file) {
    if (!file) return;
    const workbook = await readWorkbook(file);
    const parsed = parseSalesOrReturns(workbook, "bajas", file.name).map((row) => ({ ...row, sourceFile: file.name }));
    setBajas(parsed);
    setPendingWasteImport(parsed);
    setWasteImportFile(file.name);
    setWasteImportStatus(parsed.length ? "Revisa las bajas antes de guardarlas en la base." : "No se reconocieron registros de bajas.");
    setFiles((current) => ({ ...current, bajas: file.name }));
    setHasUnsavedChanges(true);
  }

  async function importOperationalToDatabase(type) {
    const isProduction = type === "produccion";
    const pending = isProduction ? pendingProductionImport : pendingWasteImport;
    if (!pending.length) return;
    const setImporting = isProduction ? setProductionImporting : setWasteImporting;
    const setStatus = isProduction ? setProductionImportStatus : setWasteImportStatus;
    const clearPending = isProduction ? setPendingProductionImport : setPendingWasteImport;
    setImporting(true);
    setStatus(`Validando y guardando ${isProduction ? "producción" : "bajas"}...`);
    try {
      const mappedRows = pending.map((row) => {
        const product = resolveOfficialProduct(row.producto, productAliases, officialProducts);
        return {
          fecha: dateKey(row.fecha),
          producto_codigo: product,
          producto_nombre: product,
          cantidad: Number(row.cantidad),
          turno: row.turno || "",
          sucursal: row.sucursal || row.canal || "",
          motivo: row.motivo || "",
          monthlyTotal: Boolean(row.monthlyTotal),
        };
      });
      const rows = consolidateOperationalRowsForUpload(mappedRows, isProduction);
      if (!rows.length) throw new Error(`El archivo no contiene registros diarios de ${isProduction ? "producción" : "bajas"}`);
      const endpoint = isProduction ? "/api/produccion-real/importar" : "/api/bajas/importar";
      const bodyKey = isProduction ? "produccion" : "bajas";
      const response = await uploadRowsInBatches({
        endpoint,
        bodyKey,
        rows,
        archivo: isProduction ? productionImportFile : wasteImportFile,
        token: session.token,
        onProgress: (batch, total) => setStatus(`Guardando ${isProduction ? "producción" : "bajas"}: lote ${batch} de ${total}...`),
      });
      const keyForRow = isProduction ? productionRecordKey : wasteRecordKey;
      const importedKeys = new Set(rows.map(keyForRow));
      if (isProduction) {
        setRealProduction((current) => current.map((row) => importedKeys.has(productionRecordKey(row)) ? { ...row, databaseSynced: true } : row));
      } else {
        setBajas((current) => current.map((row) => importedKeys.has(wasteRecordKey(row)) ? { ...row, databaseSynced: true } : row));
      }
      clearPending([]);
      setStatus("");
      const backup = await saveWorkspace({
        silent: true,
        persistedKeys: isProduction ? { production: importedKeys } : { waste: importedKeys },
      });
      setToast({
        tone: backup.ok ? "success" : "warning",
        message: backup.ok
          ? `${isProduction ? "Producción" : "Bajas"} guardadas: ${response.inserted} nuevas, ${response.updated} actualizadas. Respaldo v${backup.version} creado.`
          : `${isProduction ? "Producción" : "Bajas"} guardadas en la base de datos, pero el respaldo automático falló.`,
      });
    } catch (error) {
      if (error.status === 401) onLogout();
      else setStatus(error.message);
    } finally {
      setImporting(false);
    }
  }

  async function handleMonthlySummaryFile(file, key, setter, parser = parseMonthlySummaryFile) {
    if (!file) return;
    const parsed = await parser(file);
    setter(parsed);
    setFiles((current) => ({ ...current, [key]: file.name }));
    setHasUnsavedChanges(true);
  }

  async function handleMonthlyCloseFile(file, type) {
    if (!file) return;
    const parsed = await parseMonthlySummaryFile(file);
    const hint = inferMonthHintFromFileName(file.name);
    const period = hint
      ? `${hint.year}-${String(hint.monthIndex + 1).padStart(2, "0")}`
      : selectedMonth;
    if (period !== monthlyClosePeriod) {
      setMonthlyCloseSales([]);
      setMonthlyCloseProduction([]);
    }
    if (type === "sales") setMonthlyCloseSales(parsed);
    else setMonthlyCloseProduction(parsed);
    setMonthlyClosePeriod(period);
    if (period && period !== selectedMonth) {
      setSelectedMonth(period);
      setSelectedMonthTouched(true);
    }
    setFiles((current) => ({
      ...current,
      [type === "sales" ? "monthlyCloseSales" : "monthlyCloseProduction"]: file.name,
    }));
    setHasUnsavedChanges(true);
    setToast({
      tone: "success",
      message: `${type === "sales" ? "Ventas" : "Producción"} de cierre reconocida: ${parsed.length} productos en ${period}.`,
    });
  }

  function saveProductAlias(alias, official) {
    const aliasKey = normalizeProduct(alias);
    const officialProduct = normalizeProduct(official);
    if (!aliasKey) return;

    setProductAliases((current) => {
      const next = { ...current };
      if (officialProduct) next[aliasKey] = officialProduct;
      else delete next[aliasKey];
      return next;
    });
    setHasUnsavedChanges(true);
  }

  useEffect(() => {
    try {
      localStorage.setItem(PRODUCT_ALIAS_STORAGE_KEY, JSON.stringify(productAliases));
    } catch {
      // Si el navegador bloquea localStorage, la app debe seguir funcionando.
    }
  }, [productAliases]);

  const officialProducts = useMemo(() => getOfficialProducts(stockRows), [stockRows]);

  const effectiveVentas = useMemo(
    () => applyProductAliases(ventas, productAliases, officialProducts),
    [ventas, productAliases, officialProducts]
  );
  const effectiveVentasValidacion = useMemo(
    () => applyProductAliases(ventasValidacion, productAliases, officialProducts),
    [ventasValidacion, productAliases, officialProducts]
  );
  const effectiveBajas = useMemo(
    () => applyProductAliases(bajas, productAliases, officialProducts),
    [bajas, productAliases, officialProducts]
  );
  const effectiveExistencias = useMemo(
    () => applyProductAliases(existencias, productAliases, officialProducts),
    [existencias, productAliases, officialProducts]
  );
  const effectiveRealProduction = useMemo(
    () => applyProductAliases(realProduction, productAliases, officialProducts),
    [realProduction, productAliases, officialProducts]
  );
  const effectiveMonthlyCloseSales = useMemo(
    () => applyProductAliases(monthlyCloseSales, productAliases, officialProducts),
    [monthlyCloseSales, productAliases, officialProducts]
  );
  const effectiveMonthlyCloseProduction = useMemo(
    () => applyProductAliases(monthlyCloseProduction, productAliases, officialProducts),
    [monthlyCloseProduction, productAliases, officialProducts]
  );

  const homologationRows = useMemo(
    () =>
      buildHomologationRows({
        ventas,
        bajas,
        existencias,
        realProduction,
        productAliases,
        officialProducts,
      }),
    [ventas, bajas, existencias, realProduction, productAliases, officialProducts]
  );

  const pendingHomologationCount = homologationRows.filter((row) => row.status === "Pendiente").length;

  const historicalValidationRows = useMemo(
    () =>
      buildHistoricalValidationRows({
        ventas: effectiveVentas,
        producedMay,
        producedJune,
        bajasJune,
        bajasJuly,
        stockRows,
      }),
    [effectiveVentas, producedMay, producedJune, bajasJune, bajasJuly, stockRows]
  );

  const historicalValidationSummary = useMemo(() => {
    const withPrecision = historicalValidationRows.filter((row) => row.precisionJunio !== null);
    return {
      products: historicalValidationRows.length,
      precision: withPrecision.length
        ? withPrecision.reduce((sum, row) => sum + row.precisionJunio, 0) / withPrecision.length
        : 0,
      bajasJune: historicalValidationRows.reduce((sum, row) => sum + row.bajasJunio, 0),
      bajasJuly: historicalValidationRows.reduce((sum, row) => sum + row.bajasJulio, 0),
      base: historicalValidationRows.reduce((sum, row) => sum + row.produccionSugeridaBase, 0),
      adjusted: historicalValidationRows.reduce((sum, row) => sum + row.produccionSugeridaAjustada, 0),
    };
  }, [historicalValidationRows]);

  const historicalVentas = useMemo(
    () => filterVentasBeforeMonth(effectiveVentas, selectedMonth),
    [effectiveVentas, selectedMonth]
  );

  const ventasRealesMes = useMemo(() => {
    if (effectiveVentasValidacion.length) {
      const scoped = effectiveVentasValidacion.filter((record) => {
        const monthKey = monthKeyFromRecord(record);
        return !monthKey || monthKey === selectedMonth;
      });
      return scoped.length ? scoped : effectiveVentasValidacion;
    }
    return filterVentasByMonth(effectiveVentas, selectedMonth, "include");
  }, [effectiveVentas, effectiveVentasValidacion, selectedMonth]);

  const historicalMonthKeys = useMemo(() => {
    const keys = new Set(historicalVentas.map((record) => monthKeyFromRecord(record)).filter(Boolean));
    return [...keys].sort();
  }, [historicalVentas]);

  const forecast = useMemo(
    () =>
      calculateForecast({
        stockRows,
        historicalVentas,
        bajas: effectiveBajas,
        existencias: effectiveExistencias,
        realProduction: effectiveRealProduction,
        selectedMonth,
        dailyBufferPct,
      }),
    [
      stockRows,
      historicalVentas,
      effectiveBajas,
      effectiveExistencias,
      effectiveRealProduction,
      selectedMonth,
      dailyBufferPct,
    ]
  );
  const operationalScenario = useMemo(
    () => buildOperationalForecastScenario(forecast),
    [forecast]
  );
  const operationalScenarioTotal = useMemo(
    () => operationalScenario.reduce((sum, row) => sum + row.pronosticoOperativo, 0),
    [operationalScenario]
  );
  const monthlyReviewSourceRows = useMemo(
    () => monthlyReviewSource?.contenido?.rows?.length ? monthlyReviewSource.contenido.rows : operationalScenario,
    [monthlyReviewSource, operationalScenario]
  );
  const monthlyReviewRows = useMemo(
    () => buildMonthlyReviewRows({
      sourceRows: monthlyReviewSourceRows,
      forecastRows: forecast,
      historicalVentas,
      loadedExistencias: effectiveExistencias,
      inputs: monthlyReview.inputs,
    }),
    [monthlyReviewSourceRows, forecast, historicalVentas, effectiveExistencias, monthlyReview.inputs]
  );
  const filteredMonthlyReviewRows = useMemo(
    () => monthlyReviewRows.filter((row) => {
      if (monthlyReviewQuery && !row.producto.includes(normalizeProduct(monthlyReviewQuery))) return false;
      if (monthlyReviewFilter === "alerts" && row.severity === "ok") return false;
      if (monthlyReviewFilter === "pending" && row.decision !== "pending") return false;
      if (monthlyReviewFilter === "adjusted" && Math.abs(row.difference) < 0.01) return false;
      return true;
    }),
    [monthlyReviewRows, monthlyReviewQuery, monthlyReviewFilter]
  );
  const monthlyReviewSummary = useMemo(() => {
    const baseTotal = monthlyReviewRows.reduce((sum, row) => sum + row.baseOperational, 0);
    const proposedTotal = monthlyReviewRows.reduce((sum, row) => sum + row.proposed, 0);
    const finalTotal = monthlyReviewRows.reduce(
      (sum, row) => sum + (row.decision === "accepted" ? row.proposed : row.baseOperational),
      0
    );
    return {
      baseTotal,
      proposedTotal,
      finalTotal,
      alerts: monthlyReviewRows.filter((row) => row.severity !== "ok").length,
      accepted: monthlyReviewRows.filter((row) => row.decision === "accepted").length,
      rejected: monthlyReviewRows.filter((row) => row.decision === "rejected").length,
      pending: monthlyReviewRows.filter((row) => row.decision === "pending").length,
    };
  }, [monthlyReviewRows]);
  const monthlyReviewSourceMatches = !monthlyReview.sourceFrozenVersion ||
    monthlyReview.sourceFrozenVersion === monthlyReviewSource?.version;

  function updateMonthlyReview(patch) {
    setMonthlyReview((current) => {
      const startsNewRevision = current.state === "approved" ||
        Boolean(current.sourceFrozenVersion && current.sourceFrozenVersion !== monthlyReviewSource?.version);
      return {
        ...current,
        ...patch,
        period: selectedMonth,
        state: startsNewRevision ? "draft" : current.state,
        version: startsNewRevision ? null : current.version,
        parentVersion: startsNewRevision ? current.version : current.parentVersion,
        approvedAt: startsNewRevision ? null : current.approvedAt,
        sourceFrozenVersion: startsNewRevision ? monthlyReviewSource?.version || null : current.sourceFrozenVersion,
      };
    });
  }

  function updateMonthlyReviewItem(product, patch) {
    const key = normalizeProduct(product);
    setMonthlyReview((current) => {
      const startsNewRevision = current.state === "approved" ||
        Boolean(current.sourceFrozenVersion && current.sourceFrozenVersion !== monthlyReviewSource?.version);
      return {
        ...current,
        period: selectedMonth,
        state: startsNewRevision ? "draft" : current.state,
        version: startsNewRevision ? null : current.version,
        parentVersion: startsNewRevision ? current.version : current.parentVersion,
        approvedAt: startsNewRevision ? null : current.approvedAt,
        sourceFrozenVersion: startsNewRevision ? monthlyReviewSource?.version || null : current.sourceFrozenVersion,
        inputs: {
          ...current.inputs,
          [key]: { ...(current.inputs[key] || {}), ...patch },
        },
      };
    });
  }

  function decideVisibleMonthlyReview(decision) {
    setMonthlyReview((current) => {
      const startsNewRevision = current.state === "approved" ||
        Boolean(current.sourceFrozenVersion && current.sourceFrozenVersion !== monthlyReviewSource?.version);
      const inputs = { ...current.inputs };
      for (const row of filteredMonthlyReviewRows) {
        const key = normalizeProduct(row.producto);
        inputs[key] = { ...(inputs[key] || {}), decision };
      }
      return {
        ...current,
        period: selectedMonth,
        state: startsNewRevision ? "draft" : current.state,
        version: startsNewRevision ? null : current.version,
        parentVersion: startsNewRevision ? current.version : current.parentVersion,
        approvedAt: startsNewRevision ? null : current.approvedAt,
        sourceFrozenVersion: startsNewRevision ? monthlyReviewSource?.version || null : current.sourceFrozenVersion,
        inputs,
      };
    });
  }

  async function saveMonthlyReview(targetState = "draft") {
    if (!canSave || monthlyReviewSaving || !monthlyReviewSource) return;
    if (targetState === "approved" && session.user.rol !== "admin") return;
    if (targetState === "approved" && monthlyReviewSummary.pending > 0) {
      setToast({ tone: "warning", message: "Decide todas las propuestas antes de aprobar la revisión." });
      return;
    }
    if (!monthlyReviewSourceMatches) {
      setToast({ tone: "warning", message: "El pronóstico congelado cambió. Recarga la revisión antes de guardarla." });
      return;
    }
    setMonthlyReviewSaving(true);
    try {
      const savedAt = new Date().toISOString();
      const approvedAt = targetState === "approved" ? savedAt : null;
      const response = await apiRequest("/api/snapshots/monthly-review", {
        token: session.token,
        method: "POST",
        body: {
          periodo: selectedMonth,
          archivos: {},
          contenido: {
            schemaVersion: 1,
            period: selectedMonth,
            state: targetState,
            savedAt,
            approvedAt,
            sourceFrozenId: monthlyReviewSource.id,
            sourceFrozenVersion: monthlyReviewSource.version,
            sourceModelVersion: monthlyReviewSource.contenido?.modelVersion || FORECAST_MODEL_VERSION,
            parentVersion: monthlyReview.parentVersion || monthlyReview.version || null,
            generalNote: monthlyReview.generalNote,
            inputs: monthlyReview.inputs,
            summary: monthlyReviewSummary,
            rows: monthlyReviewRows,
          },
        },
      });
      setMonthlyReview((current) => ({
        ...current,
        state: targetState,
        version: response.version,
        savedAt,
        approvedAt,
        sourceFrozenVersion: monthlyReviewSource.version,
      }));
      setToast({
        tone: "success",
        message: targetState === "approved"
          ? `Revisión ${selectedMonth} aprobada en versión ${response.version}.`
          : `Revisión ${selectedMonth} guardada en versión ${response.version}.`,
      });
    } catch (error) {
      if (error.status === 401) onLogout();
      else setToast({ tone: "error", message: `No se pudo guardar la revisión: ${error.message}` });
    } finally {
      setMonthlyReviewSaving(false);
    }
  }

  async function freezeAndExportForecast() {
    if (!operationalScenario.length || forecastFreezing) return;
    setForecastFreezing(true);
    try {
      let workspaceVersion = lastBackup?.version || null;
      if (hasUnsavedChanges || !lastBackup) {
        const saved = await saveWorkspace({ silent: true });
        if (!saved.ok) throw saved.error;
        workspaceVersion = saved.version;
      }
      const frozenAt = new Date().toISOString();
      const frozenContent = {
        schemaVersion: 1,
        frozenAt,
        selectedMonth,
        modelVersion: FORECAST_MODEL_VERSION,
        operationalMarginPct: OPERATIONAL_MARGIN_PCT,
        source: { workspaceVersion, files },
        rows: operationalScenario,
      };
      const response = await apiRequest("/api/snapshots/forecast-frozen", {
        token: session.token,
        method: "POST",
        body: {
          periodo: selectedMonth,
          archivos: files,
          contenido: frozenContent,
        },
      });
      setMonthlyReviewSource({
        id: response.id,
        version: response.version,
        periodo: selectedMonth,
        contenido: frozenContent,
      });
      exportFrozenForecast({
        rows: operationalScenario,
        selectedMonth,
        frozenAt,
        snapshotVersion: response.version,
        workspaceVersion,
      });
      setToast({
        tone: "success",
        message: `Pronóstico ${selectedMonth} congelado en versión ${response.version} y exportado.`,
      });
    } catch (error) {
      if (error.status === 401) onLogout();
      else setToast({ tone: "error", message: `No se pudo congelar el pronóstico: ${error.message}` });
    } finally {
      setForecastFreezing(false);
    }
  }

  const monthlyCloseMatchesSelectedMonth = Boolean(monthlyClosePeriod) && monthlyClosePeriod === selectedMonth;
  const monthlyClose = useMemo(
    () => buildMonthlyCloseSummary({
      forecastRows: forecast,
      salesRows: monthlyCloseMatchesSelectedMonth ? effectiveMonthlyCloseSales : [],
      productionRows: monthlyCloseMatchesSelectedMonth ? effectiveMonthlyCloseProduction : [],
    }),
    [forecast, effectiveMonthlyCloseSales, effectiveMonthlyCloseProduction, monthlyCloseMatchesSelectedMonth]
  );
  const priorityMonthlyCloseProducts = monthlyClose.rows.slice(0, 15);

  const comparableForecast = showMissingReal ? forecast : forecast.filter((r) => r.hasRealData);
  const filtered = comparableForecast.filter((r) => r.producto.includes(norm(query)));

  const dailyRows = useMemo(
    () =>
      calculateDailyForecast({
        monthlyRows: forecast,
        ventasReales: ventasRealesMes,
        realProduction: effectiveRealProduction,
        selectedMonth,
        dailyBufferPct,
      }),
    [forecast, ventasRealesMes, effectiveRealProduction, selectedMonth, dailyBufferPct]
  );
  const filteredDailyRows = dailyRows.filter((row) => {
    if (dailyDateFilter && row.fecha !== dailyDateFilter) return false;
    if (dailyProductQuery && !row.producto.includes(norm(dailyProductQuery))) return false;
    if (dailyWeekdayFilter !== "" && row.weekday !== Number(dailyWeekdayFilter)) return false;
    if (onlyDailyShortage && row.estatus !== "Riesgo faltante") return false;
    if (onlyDailyOverproduction && row.estatus !== "Sobreproduccion") return false;
    if (onlySalesMismatch && (!row.hasVentaReal || row.estatusVenta === "Dentro de rango")) return false;
    return true;
  });

  const dailySummary = useMemo(() => summarizeDailyMonth(dailyRows), [dailyRows]);
  const weeklyProgress = useMemo(
    () => buildWeeklyProgress(dailyRows, selectedMonth),
    [dailyRows, selectedMonth]
  );
  useEffect(() => {
    if (weeklyProgress.weeks.some((week) => week.key === selectedWeekKey)) return;
    setSelectedWeekKey(weeklyProgress.suggestedWeekKey);
  }, [weeklyProgress, selectedWeekKey]);
  const selectedWeek = weeklyProgress.weeks.find((week) => week.key === selectedWeekKey) ||
    weeklyProgress.weeks.find((week) => week.key === weeklyProgress.suggestedWeekKey) ||
    weeklyProgress.weeks[0] || null;
  const priorityWeeklyProducts = selectedWeek?.products.slice(0, 12) || [];
  const salesValidationSummary = useMemo(() => summarizeSalesValidation(dailyRows), [dailyRows]);
  const productValidationSummary = useMemo(
    () => buildProductValidationSummary(dailyRows, forecast),
    [dailyRows, forecast]
  );
  const validationAlerts = useMemo(
    () =>
      buildValidationAlerts(
        productValidationSummary,
        homologationRows,
        historicalVentas,
        selectedMonth
      ),
    [productValidationSummary, homologationRows, historicalVentas, selectedMonth]
  );
  const hasSalesValidation = salesValidationSummary.diasConReal > 0;

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
  const validationSales = historicalVentas
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
    const precisionEjecutiva =
      totalReal > 0 ? precisionScore(totalPronosticada, totalReal) ?? 0 : 0;
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

  const shouldShowHomologation = homologationRows.length > 0;

  const loadedFileItems = [
    { label: "Ventas históricas", loaded: Boolean(files.ventas || ventas.length) },
    { label: "Stock fijo", loaded: Boolean(files.stock || stockRows.length) },
    { label: "Producción real", loaded: Boolean(files.real || realProduction.length) },
    { label: "Bajas/devoluciones", loaded: Boolean(files.bajas || bajas.length) },
    { label: "Existencias", loaded: Boolean(files.existencias || existencias.length) },
  ];
  const canSave = session.user.rol === "admin" || session.user.rol === "operador";

  return (
    <div className="app">
      <main className="main">
        <header className="top">
          <div>
            <span className="eyebrow">Archivo Maestro</span>
            <h2>Producción diaria sugerida</h2>
            <p>Selecciona por producto el método con menor error histórico y lo distribuye por día de semana.</p>
          </div>
          <div className="top-actions">
            <div className="session-user">
              <UserRound size={17} />
              <span>{session.user.nombre}</span>
              <small>{session.user.rol}</small>
            </div>
            <button
              className="primary"
              type="button"
              onClick={() => saveWorkspace()}
              disabled={!canSave || cloudSaving || (!hasUnsavedChanges && Boolean(lastBackup))}
            >
              <Save size={18} /> {cloudSaving ? "Guardando..." : hasUnsavedChanges ? "Guardar respaldo" : "Respaldo guardado"}
            </button>
            <button className="secondary" type="button" onClick={() => exportToExcel(filtered, summary)} disabled={!forecast.length}>
              <Download size={18} /> Exportar
            </button>
            <button
              className="secondary"
              type="button"
              onClick={freezeAndExportForecast}
              disabled={!canSave || cloudSaving || forecastFreezing || !forecast.length}
              title="Guarda una versión inmutable por producto y descarga el mismo pronóstico en Excel"
            >
              <ShieldCheck size={18} /> {forecastFreezing ? "Congelando..." : "Congelar mes"}
            </button>
            {session.user.rol === "admin" && (
              <button className="secondary" type="button" onClick={toggleUsers}>
                <UserRound size={18} /> Usuarios
              </button>
            )}
            <button className="icon-button" type="button" onClick={onLogout} title="Cerrar sesión" aria-label="Cerrar sesión">
              <LogOut size={18} />
            </button>
          </div>
        </header>

        {toast && (
          <div className={`app-toast ${toast.tone}`} role="status" aria-live="polite">
            <CheckCircle2 size={18} />
            <span>{toast.message}</span>
            <button type="button" onClick={() => setToast(null)} aria-label="Cerrar confirmación">×</button>
          </div>
        )}

        {showUserAdmin && session.user.rol === "admin" && (
          <section className="user-admin-section">
            <div className="section-heading compact-heading">
              <div>
                <span className="eyebrow">Control de acceso</span>
                <h3>Usuarios del sistema</h3>
                <p>Los operadores pueden guardar información; consulta solo puede visualizar y exportar.</p>
              </div>
              <strong>{formatNumber(users.length)} usuarios</strong>
            </div>
            <div className="user-admin-grid">
              <form className="user-create-form" onSubmit={createUser}>
                <label>
                  Nombre
                  <input value={userForm.nombre} onChange={(event) => setUserForm((value) => ({ ...value, nombre: event.target.value }))} required />
                </label>
                <label>
                  Usuario
                  <input value={userForm.usuario} onChange={(event) => setUserForm((value) => ({ ...value, usuario: event.target.value }))} required />
                </label>
                <label>
                  Contraseña inicial
                  <input type="password" minLength="10" value={userForm.password} onChange={(event) => setUserForm((value) => ({ ...value, password: event.target.value }))} required />
                </label>
                <label>
                  Rol
                  <select value={userForm.rol} onChange={(event) => setUserForm((value) => ({ ...value, rol: event.target.value }))}>
                    <option value="operador">Operador</option>
                    <option value="consulta">Consulta</option>
                    <option value="admin">Administrador</option>
                  </select>
                </label>
                <button className="primary" type="submit"><UserRound size={17} /> Crear usuario</button>
                {userStatus && <small className="user-status">{userStatus}</small>}
              </form>
              <div className="user-list">
                {users.map((user) => (
                  <div className="user-list-row" key={user.id}>
                    <div><strong>{user.nombre}</strong><span>@{user.usuario}</span></div>
                    <span className="pill muted">{user.rol}</span>
                    <small>{user.ultimo_acceso ? new Date(user.ultimo_acceso).toLocaleString("es-MX") : "Sin acceso"}</small>
                  </div>
                ))}
                {!users.length && <div className="empty">No hay usuarios para mostrar.</div>}
              </div>
            </div>
          </section>
        )}

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
              caption={`Margen de seguridad: ${dailyBufferPct}%`}
            />
          </section>

          <div className="forecast-concepts" role="note">
            <p><strong>Pronóstico de venta:</strong> cantidad estimada que se espera vender.</p>
            <p><strong>Producción sugerida:</strong> cantidad recomendada a producir después de aplicar margen de seguridad y regla de múltiplos de 5.</p>
            <p><strong>Escenario operativo +{OPERATIONAL_MARGIN_PCT}%:</strong> {formatNumber(operationalScenarioTotal, 0)} piezas; se conserva separado del pronóstico estadístico.</p>
          </div>
        </section>

        <details className="weekly-progress-section compact-analytics-section">
          <summary className="compact-analytics-summary">
            <span className="compact-analytics-icon"><CalendarRange size={19} /></span>
            <div>
              <span className="eyebrow">Seguimiento contra pronóstico</span>
              <strong>Avance semanal</strong>
              <small>Venta real, cumplimiento y proyección mensual</small>
            </div>
            <span className={`pill ${weeklyProgress.hasRealData ? weeklyProgress.month.status.className : "muted"}`}>
              {weeklyProgress.hasRealData && weeklyProgress.month.cumplimiento !== null
                ? `${formatPercent(weeklyProgress.month.cumplimiento, 1)} al corte`
                : "Sin ventas diarias"}
            </span>
          </summary>

          <div className="compact-analytics-content">
            <div className="compact-analytics-toolbar">
              <p>Compara la venta real cargada contra el pronóstico original y proyecta el cierre sin modificarlo.</p>
            <div className="weekly-actions">
              <label>
                Semana
                <select value={selectedWeek?.key || ""} onChange={(event) => setSelectedWeekKey(event.target.value)}>
                  {!weeklyProgress.weeks.length && <option value="">Sin semanas</option>}
                  {weeklyProgress.weeks.map((week) => (
                    <option value={week.key} key={week.key}>{week.label}</option>
                  ))}
                </select>
              </label>
              <button
                className="primary"
                type="button"
                onClick={() => exportWeeklyProgress(weeklyProgress, selectedWeek, selectedMonth)}
                disabled={!selectedWeek || !weeklyProgress.hasRealData}
              >
                <Download size={18} /> Exportar avance
              </button>
            </div>
          </div>

          <p className={`weekly-data-note ${weeklyProgress.hasRealData && weeklyProgress.month.coveragePct >= 90 ? "success" : "warning"}`}>
            {weeklyProgress.hasRealData
              ? `Venta real del ${displayDate(weeklyProgress.month.firstRealDate)} al ${displayDate(weeklyProgress.month.cutoffDate)} · cobertura de fechas ${formatPercent(weeklyProgress.month.coveragePct, 0)}.`
              : "Aún no hay ventas diarias del mes seleccionado. El avance se activará al sincronizar o importar ventas reales."}
          </p>

          {weeklyProgress.hasRealData && <>
          <section className="weekly-kpis">
            <KpiCard
              icon={CalendarRange}
              label="Venta real de la semana"
              value={selectedWeek?.comparedDays ? formatNumber(selectedWeek.ventaReal, 0) : "Sin datos"}
              caption={selectedWeek?.comparedDays ? `${selectedWeek.comparedDays} días comparados` : "Esperando venta diaria"}
            />
            <KpiCard
              icon={Target}
              label="Pronóstico al corte"
              value={selectedWeek?.comparedDays ? formatNumber(selectedWeek.pronosticoCorte, 0) : "Sin datos"}
              caption={selectedWeek ? `Semana completa: ${formatNumber(selectedWeek.pronosticoPeriodo, 0)}` : "Sin semana"}
            />
            <KpiCard
              icon={BarChart3}
              label="Cumplimiento semanal"
              value={selectedWeek?.cumplimiento === null || selectedWeek?.cumplimiento === undefined ? "Sin datos" : formatPercent(selectedWeek.cumplimiento, 1)}
              caption={selectedWeek?.status.label || "Sin información"}
              tone={selectedWeek?.status.tone}
            />
            <KpiCard
              icon={TrendingUp}
              label="Proyección de cierre mensual"
              value={weeklyProgress.hasRealData ? formatNumber(weeklyProgress.month.proyeccionPeriodo, 0) : "Sin datos"}
              caption={weeklyProgress.hasRealData
                ? `${weeklyProgress.month.projectedDifference >= 0 ? "+" : ""}${formatNumber(weeklyProgress.month.projectedDifference, 0)} vs. pronóstico mensual`
                : `Pronóstico mensual: ${formatNumber(weeklyProgress.month.pronosticoPeriodo, 0)}`}
              tone={weeklyProgress.month.projectedStatus.tone}
            />
          </section>

          <div className="weekly-timeline" aria-label="Resumen de semanas del mes">
            {weeklyProgress.weeks.map((week) => (
              <button
                type="button"
                className={`weekly-step ${selectedWeek?.key === week.key ? "active" : ""}`}
                onClick={() => setSelectedWeekKey(week.key)}
                key={week.key}
              >
                <span>{week.label.split(" · ")[0]}</span>
                <strong>{week.comparedDays ? formatPercent(week.cumplimiento, 0) : "Pendiente"}</strong>
                <small>{week.comparedDays ? `${formatNumber(week.ventaReal)} / ${formatNumber(week.pronosticoCorte)}` : week.label.split(" · ")[1]}</small>
                <i className={`weekly-status-dot ${week.status.className}`} />
              </button>
            ))}
          </div>

          {selectedWeek?.comparedDays > 0 && (
            <div className="weekly-detail-grid">
              <section className="table-card weekly-product-card">
                <div className="weekly-table-title">
                  <div>
                    <span className="eyebrow">Prioridad de revisión</span>
                    <h4>Productos con mayor desviación</h4>
                  </div>
                  <small>Primeros {priorityWeeklyProducts.length}</small>
                </div>
                <table className="weekly-table">
                  <thead>
                    <tr>
                      <th>Producto</th>
                      <th>Pronóstico al corte</th>
                      <th>Venta real</th>
                      <th>Diferencia</th>
                      <th>Cumplimiento</th>
                      <th>Proyección semanal</th>
                      <th>Estado</th>
                    </tr>
                  </thead>
                  <tbody>
                    {priorityWeeklyProducts.map((row) => (
                      <tr key={row.producto}>
                        <td className="strong">{row.producto}</td>
                        <td>{formatNumber(row.pronosticoCorte, 1)}</td>
                        <td>{formatNumber(row.ventaReal, 1)}</td>
                        <td className={row.diferencia < 0 ? "negative" : "positive"}>
                          {row.diferencia > 0 ? "+" : ""}{formatNumber(row.diferencia, 1)}
                        </td>
                        <td>{row.cumplimiento === null ? "Sin dato" : formatPercent(row.cumplimiento, 1)}</td>
                        <td>{formatNumber(row.proyeccionPeriodo, 1)}</td>
                        <td><span className={`pill ${row.status.className}`}>{row.status.label}</span></td>
                      </tr>
                    ))}
                  </tbody>
                </table>
              </section>

              <section className="table-card weekly-category-card">
                <div className="weekly-table-title">
                  <div>
                    <span className="eyebrow">Lectura por familia</span>
                    <h4>Categorías</h4>
                  </div>
                </div>
                <table className="weekly-table">
                  <thead>
                    <tr>
                      <th>Categoría</th>
                      <th>Real</th>
                      <th>Pronóstico</th>
                      <th>Diferencia</th>
                      <th>Estado</th>
                    </tr>
                  </thead>
                  <tbody>
                    {selectedWeek.categories.map((row) => (
                      <tr key={row.categoria}>
                        <td className="strong">{row.categoria}</td>
                        <td>{formatNumber(row.ventaReal, 0)}</td>
                        <td>{formatNumber(row.pronosticoCorte, 0)}</td>
                        <td className={row.diferencia < 0 ? "negative" : "positive"}>
                          {row.diferencia > 0 ? "+" : ""}{formatNumber(row.diferencia, 0)}
                        </td>
                        <td><span className={`pill ${row.status.className}`}>{row.status.label}</span></td>
                      </tr>
                    ))}
                  </tbody>
                </table>
              </section>
            </div>
          )}
          </>}
          </div>
        </details>

        <details className="monthly-review-section compact-analytics-section">
          <summary className="compact-analytics-summary">
            <span className="compact-analytics-icon"><ShieldCheck size={19} /></span>
            <div>
              <span className="eyebrow">Decisión operativa explicable</span>
              <strong>Revisión mensual asistida</strong>
              <small>Estatus, existencias y alertas sin modificar el pronóstico</small>
            </div>
            <span className={`pill ${monthlyReview.state === "approved" ? "ok" : monthlyReviewSummary.alerts ? "warn" : "muted"}`}>
              {monthlyReviewLoading
                ? "Cargando"
                : monthlyReview.state === "approved"
                  ? `Aprobada v${monthlyReview.version}`
                  : `${monthlyReviewSummary.alerts} alertas`}
            </span>
          </summary>

          <div className="compact-analytics-content">
            <div className="compact-analytics-toolbar monthly-review-toolbar">
              <p>
                El motor local usa únicamente historial anterior, estatus y existencias. Sus propuestas forman
                un plan operativo separado del pronóstico estadístico.
              </p>
              <div className="monthly-review-actions">
                <button
                  className="secondary"
                  type="button"
                  onClick={() => exportMonthlyReview({
                    rows: monthlyReviewRows,
                    review: monthlyReview,
                    selectedMonth,
                    sourceVersion: monthlyReviewSource?.version,
                  })}
                  disabled={!monthlyReviewRows.length}
                >
                  <Download size={17} /> Exportar revisión
                </button>
                <button
                  className="secondary"
                  type="button"
                  onClick={() => saveMonthlyReview("draft")}
                  disabled={!canSave || !monthlyReviewSource || monthlyReviewSaving}
                >
                  <Save size={17} /> {monthlyReviewSaving ? "Guardando..." : "Guardar borrador"}
                </button>
                {session.user.rol === "admin" && (
                  <button
                    className="primary"
                    type="button"
                    onClick={() => saveMonthlyReview("approved")}
                    disabled={!monthlyReviewSource || monthlyReviewSaving || monthlyReviewSummary.pending > 0}
                    title={monthlyReviewSummary.pending ? "Decide todas las propuestas antes de aprobar" : "Aprobar plan operativo"}
                  >
                    <CheckCircle2 size={17} /> Aprobar plan
                  </button>
                )}
              </div>
            </div>

            <p className={`monthly-review-note ${monthlyReviewSource ? "success" : "warning"}`}>
              {monthlyReviewSource
                ? `Fuente inmutable: pronóstico congelado ${selectedMonth}, versión ${monthlyReviewSource.version}. ${monthlyReview.state === "approved" ? "Editar cualquier dato abrirá un nuevo borrador." : "El pronóstico original no será modificado."}`
                : `Aún no existe un pronóstico congelado para ${selectedMonth}. Puedes revisar propuestas provisionales, pero debes usar “Congelar mes” antes de guardarlas.`}
            </p>

            <section className="monthly-review-kpis">
              <KpiCard
                icon={Target}
                label="Plan operativo base"
                value={formatNumber(monthlyReviewSummary.baseTotal, 0)}
                caption={`Pronóstico congelado +${OPERATIONAL_MARGIN_PCT}%`}
              />
              <KpiCard
                icon={TrendingUp}
                label="Propuesta del motor"
                value={formatNumber(monthlyReviewSummary.proposedTotal, 0)}
                caption={`${monthlyReviewSummary.proposedTotal - monthlyReviewSummary.baseTotal >= 0 ? "+" : ""}${formatNumber(monthlyReviewSummary.proposedTotal - monthlyReviewSummary.baseTotal, 0)} piezas`}
              />
              <KpiCard
                icon={PackageCheck}
                label="Plan según decisiones"
                value={formatNumber(monthlyReviewSummary.finalTotal, 0)}
                caption={`${monthlyReviewSummary.accepted} aceptadas · ${monthlyReviewSummary.rejected} rechazadas`}
              />
              <KpiCard
                icon={ShieldCheck}
                label="Pendientes / alertas"
                value={`${monthlyReviewSummary.pending} / ${monthlyReviewSummary.alerts}`}
                caption="La aprobación exige resolver todas"
                tone={monthlyReviewSummary.pending ? "warn" : ""}
              />
            </section>

            <label className="monthly-review-general-note">
              Nota general del mes
              <textarea
                value={monthlyReview.generalNote}
                onChange={(event) => updateMonthlyReview({ generalNote: event.target.value })}
                placeholder="Ejemplo: apertura de sucursales, campaña, cambio de horario o decisión comercial."
                disabled={!canSave}
                rows="2"
              />
            </label>

            <div className="monthly-review-controls">
              <label className="search">
                <Search size={17} />
                <input
                  type="search"
                  value={monthlyReviewQuery}
                  onChange={(event) => setMonthlyReviewQuery(event.target.value)}
                  placeholder="Buscar producto"
                />
              </label>
              <label>
                Mostrar
                <select value={monthlyReviewFilter} onChange={(event) => setMonthlyReviewFilter(event.target.value)}>
                  <option value="all">Todos</option>
                  <option value="alerts">Solo alertas</option>
                  <option value="pending">Solo pendientes</option>
                  <option value="adjusted">Con ajuste propuesto</option>
                </select>
              </label>
              <span>{filteredMonthlyReviewRows.length} productos visibles</span>
              <button
                className="secondary"
                type="button"
                onClick={() => decideVisibleMonthlyReview("rejected")}
                disabled={!canSave || !filteredMonthlyReviewRows.length}
              >
                Rechazar visibles
              </button>
              <button
                className="primary"
                type="button"
                onClick={() => decideVisibleMonthlyReview("accepted")}
                disabled={!canSave || !filteredMonthlyReviewRows.length}
              >
                Aceptar visibles
              </button>
            </div>

            <section className="table-card monthly-review-table-card">
              <table className="monthly-review-table">
                <thead>
                  <tr>
                    <th>Producto</th>
                    <th>Estatus</th>
                    <th>Existencias</th>
                    <th>Pronóstico</th>
                    <th>Plan base</th>
                    <th>Propuesta</th>
                    <th>Diferencia</th>
                    <th>Motivos</th>
                    <th>Nota</th>
                    <th>Decisión</th>
                  </tr>
                </thead>
                <tbody>
                  {filteredMonthlyReviewRows.map((row) => (
                    <tr className={`monthly-review-row ${row.severity}`} key={row.producto}>
                      <td>
                        <strong>{row.producto}</strong>
                        <small>{row.categoria}</small>
                      </td>
                      <td>
                        <select
                          value={row.status}
                          onChange={(event) => updateMonthlyReviewItem(row.producto, { status: event.target.value })}
                          disabled={!canSave}
                        >
                          {MONTHLY_REVIEW_STATUSES.map((status) => <option value={status} key={status}>{status}</option>)}
                        </select>
                      </td>
                      <td>
                        <input
                          className="monthly-review-number"
                          type="number"
                          min="0"
                          step="1"
                          value={row.inventoryOverride ?? row.inventory}
                          onChange={(event) => updateMonthlyReviewItem(row.producto, { inventoryOverride: event.target.value })}
                          disabled={!canSave}
                          title={row.hasLoadedInventory && row.inventoryOverride === null ? "Valor cargado desde existencias" : "Valor capturado en la revisión"}
                        />
                      </td>
                      <td>{formatNumber(row.baseForecast, 0)}</td>
                      <td>{formatNumber(row.baseOperational, 0)}</td>
                      <td>
                        <strong>{formatNumber(row.proposed, 0)}</strong>
                        <small>colchón {row.marginPct}%</small>
                      </td>
                      <td className={row.difference < 0 ? "negative" : row.difference > 0 ? "positive" : ""}>
                        {row.difference > 0 ? "+" : ""}{formatNumber(row.difference, 0)}
                      </td>
                      <td>
                        <span className={`pill ${row.severity}`}>{row.severity === "danger" ? "Crítica" : row.severity === "warn" ? "Revisar" : "Estable"}</span>
                        <small className="monthly-review-reasons">{row.reasons.join(" ")}</small>
                      </td>
                      <td>
                        <input
                          className="monthly-review-note-input"
                          value={row.note}
                          onChange={(event) => updateMonthlyReviewItem(row.producto, { note: event.target.value })}
                          placeholder="Contexto operativo"
                          disabled={!canSave}
                        />
                      </td>
                      <td>
                        <div className="monthly-review-decision">
                          <button
                            type="button"
                            className={row.decision === "accepted" ? "accepted" : ""}
                            onClick={() => updateMonthlyReviewItem(row.producto, { decision: "accepted" })}
                            disabled={!canSave}
                            title="Aceptar propuesta"
                          >
                            Sí
                          </button>
                          <button
                            type="button"
                            className={row.decision === "rejected" ? "rejected" : ""}
                            onClick={() => updateMonthlyReviewItem(row.producto, { decision: "rejected" })}
                            disabled={!canSave}
                            title="Conservar plan base"
                          >
                            No
                          </button>
                        </div>
                      </td>
                    </tr>
                  ))}
                  {!filteredMonthlyReviewRows.length && (
                    <tr><td colSpan="10" className="empty">No hay productos para este filtro.</td></tr>
                  )}
                </tbody>
              </table>
            </section>
          </div>
        </details>

        <details className="monthly-close-section compact-analytics-section">
          <summary className="compact-analytics-summary">
            <span className="compact-analytics-icon"><FileSpreadsheet size={19} /></span>
            <div>
              <span className="eyebrow">Resultado definitivo del periodo</span>
              <strong>Cierre mensual</strong>
              <small>Venta, pronóstico, producción y error por producto</small>
            </div>
            <span className={`pill ${monthlyClose.salesLoaded ? monthlyClose.summary.status.className : "muted"}`}>
              {monthlyClose.salesLoaded && monthlyClose.summary.cumplimiento !== null
                ? `${formatPercent(monthlyClose.summary.cumplimiento, 1)} · WAPE ${formatPercent(monthlyClose.summary.wape, 1)}`
                : "Cargar cierre"}
            </span>
          </summary>

          <div className="compact-analytics-content">
          <div className="compact-analytics-toolbar">
            <p>Carga resúmenes por producto para comparar venta, pronóstico y producción sin convertirlos en registros diarios.</p>
            <button
              className="primary"
              type="button"
              onClick={() => exportMonthlyClose(monthlyClose, selectedMonth)}
              disabled={!monthlyClose.salesLoaded}
            >
              <Download size={18} /> Exportar cierre
            </button>
          </div>

          <div className="monthly-close-upload-grid">
            <UploadBox
              title="Ventas del cierre"
              description="Resumen mensual con Producto y Cantidad Total. Reemplaza el cierre anterior del mismo mes."
              required
              onFile={(file) => handleMonthlyCloseFile(file, "sales")}
              fileName={files.monthlyCloseSales}
            />
            <UploadBox
              title="Producción del cierre"
              description="Resumen mensual con Cantidad y Producto. No necesita fecha diaria."
              onFile={(file) => handleMonthlyCloseFile(file, "production")}
              fileName={files.monthlyCloseProduction}
            />
          </div>

          <p className={`monthly-close-note ${monthlyClose.salesLoaded ? "success" : "warning"}`}>
            {!monthlyCloseMatchesSelectedMonth && monthlyClosePeriod
              ? `El cierre cargado corresponde a ${monthlyClosePeriod}. Selecciona ese mes o carga los archivos de ${selectedMonth}.`
              : monthlyClose.salesLoaded
                ? `Cierre ${selectedMonth}: ${formatNumber(monthlyClose.summary.productos)} productos comparados contra el pronóstico vigente.`
                : "Carga las ventas mensuales para activar WAPE, MAE y cumplimiento. La producción es complementaria."}
          </p>

          {monthlyClose.salesLoaded && <section className="monthly-close-kpis">
            <KpiCard
              icon={Target}
              label="Pronóstico del cierre"
              value={formatNumber(monthlyClose.summary.pronostico, 0)}
              caption={`${formatNumber(monthlyClose.summary.productos)} productos regulares`}
            />
            <KpiCard
              icon={BarChart3}
              label="Venta real"
              value={monthlyClose.salesLoaded ? formatNumber(monthlyClose.summary.ventaReal, 0) : "Sin datos"}
              caption={monthlyClose.salesLoaded
                ? `${monthlyClose.summary.diferenciaPronostico >= 0 ? "+" : ""}${formatNumber(monthlyClose.summary.diferenciaPronostico, 0)} vs. pronóstico`
                : "Esperando cierre de ventas"}
              tone={monthlyClose.summary.status.tone}
            />
            <KpiCard
              icon={TrendingUp}
              label="Cumplimiento"
              value={monthlyClose.summary.cumplimiento === null ? "Sin datos" : formatPercent(monthlyClose.summary.cumplimiento, 1)}
              caption={monthlyClose.summary.status.label}
              tone={monthlyClose.summary.status.tone}
            />
            <KpiCard
              icon={ShieldCheck}
              label="WAPE / MAE"
              value={monthlyClose.summary.wape === null ? "Sin datos" : `${formatPercent(monthlyClose.summary.wape, 2)} / ${formatNumber(monthlyClose.summary.mae, 2)}`}
              caption={monthlyClose.salesLoaded ? `${monthlyClose.summary.dentro15} productos dentro de ±15` : "Error por producto"}
            />
            <KpiCard
              icon={PackageCheck}
              label="Producción real"
              value={monthlyClose.productionLoaded ? formatNumber(monthlyClose.summary.producido, 0) : "Sin datos"}
              caption={monthlyClose.productionLoaded && monthlyClose.salesLoaded
                ? `${monthlyClose.summary.diferenciaProduccion >= 0 ? "+" : ""}${formatNumber(monthlyClose.summary.diferenciaProduccion, 0)} producido menos vendido`
                : "Producción mensual opcional"}
              tone={monthlyClose.summary.productionStatus.tone}
            />
          </section>}

          {monthlyClose.salesLoaded && (
            <>
              <div className="monthly-close-detail-grid">
                <section className="table-card monthly-close-product-card">
                  <div className="weekly-table-title">
                    <div>
                      <span className="eyebrow">Mayor impacto en el error</span>
                      <h4>Productos prioritarios</h4>
                    </div>
                    <small>Primeros {priorityMonthlyCloseProducts.length}</small>
                  </div>
                  <table className="monthly-close-table">
                    <thead>
                      <tr>
                        <th>Producto</th>
                        <th>Pronóstico</th>
                        <th>Venta</th>
                        <th>Diferencia</th>
                        <th>Error absoluto</th>
                        <th>Producido</th>
                        <th>Prod. - venta</th>
                        <th>Estado</th>
                      </tr>
                    </thead>
                    <tbody>
                      {priorityMonthlyCloseProducts.map((row) => (
                        <tr key={row.producto}>
                          <td className="strong">{row.producto}</td>
                          <td>{formatNumber(row.pronostico, 1)}</td>
                          <td>{formatNumber(row.ventaReal, 1)}</td>
                          <td className={row.diferenciaPronostico < 0 ? "negative" : "positive"}>
                            {row.diferenciaPronostico > 0 ? "+" : ""}{formatNumber(row.diferenciaPronostico, 1)}
                          </td>
                          <td>{formatNumber(row.errorAbsoluto, 1)}</td>
                          <td>{monthlyClose.productionLoaded ? formatNumber(row.producido, 1) : "-"}</td>
                          <td>{monthlyClose.productionLoaded ? `${row.diferenciaProduccion > 0 ? "+" : ""}${formatNumber(row.diferenciaProduccion, 1)}` : "-"}</td>
                          <td><span className={`pill ${row.status.className}`}>{row.status.label}</span></td>
                        </tr>
                      ))}
                    </tbody>
                  </table>
                </section>

                <section className="table-card monthly-close-category-card">
                  <div className="weekly-table-title">
                    <div>
                      <span className="eyebrow">Resultado por familia</span>
                      <h4>Categorías</h4>
                    </div>
                  </div>
                  <table className="monthly-close-table">
                    <thead>
                      <tr>
                        <th>Categoría</th>
                        <th>Pronóstico</th>
                        <th>Venta</th>
                        <th>WAPE</th>
                        <th>Prod. - venta</th>
                      </tr>
                    </thead>
                    <tbody>
                      {monthlyClose.categories.map((row) => (
                        <tr key={row.categoria}>
                          <td className="strong">{row.categoria}</td>
                          <td>{formatNumber(row.pronostico, 0)}</td>
                          <td>{formatNumber(row.ventaReal, 0)}</td>
                          <td>{row.wape === null ? "Sin dato" : formatPercent(row.wape, 1)}</td>
                          <td className={row.diferenciaProduccion < 0 ? "negative" : "positive"}>
                            {monthlyClose.productionLoaded ? `${row.diferenciaProduccion > 0 ? "+" : ""}${formatNumber(row.diferenciaProduccion, 0)}` : "-"}
                          </td>
                        </tr>
                      ))}
                    </tbody>
                  </table>
                </section>
              </div>

              <div className="monthly-close-reconciliation">
                <div>
                  <span>Venta regular comparable</span>
                  <strong>{formatNumber(monthlyClose.summary.ventaReal, 0)}</strong>
                </div>
                <div>
                  <span>Venta fuera del catálogo regular</span>
                  <strong>{formatNumber(monthlyClose.unmatchedSalesTotal, 0)}</strong>
                  <small>{monthlyClose.unmatchedSales.length} productos</small>
                </div>
                <div>
                  <span>Producción fuera del catálogo regular</span>
                  <strong>{monthlyClose.productionLoaded ? formatNumber(monthlyClose.unmatchedProductionTotal, 0) : "Sin datos"}</strong>
                  <small>{monthlyClose.productionLoaded ? `${monthlyClose.unmatchedProduction.length} productos` : "Carga opcional"}</small>
                </div>
              </div>
            </>
          )}
          </div>
        </details>

        <section className="loaded-files-section">
          <div className="section-heading compact-heading">
            <div>
              <span className="eyebrow">Preparación de datos</span>
              <h3>Archivos cargados</h3>
              <p>Carga stock fijo y ventas para generar el pronóstico. Producción real, bajas y existencias son complementarios.</p>
            </div>
            <div className="loaded-context">
              <span>Mes: {selectedMonth || "Sin mes"}</span>
              <span>Margen: {dailyBufferPct}%</span>
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
          <div className={`cloud-status ${cloudStatus.startsWith("No se pudo") ? "error" : ""}`}>
            <Database size={17} />
            <span>{cloudStatus}</span>
            {lastBackup?.created_at && <small>Último respaldo: {new Date(lastBackup.created_at).toLocaleString("es-MX")}</small>}
          </div>
        </section>

        {shouldShowHomologation && <details className="homologation-section compact-homologation advanced-details">
          <summary className="advanced-summary">
            <div>
              <span className="eyebrow">Catálogo maestro</span>
              <strong>Homologación de productos</strong>
              <small>{formatNumber(pendingHomologationCount)} nombres pendientes de revisar</small>
            </div>
            <span className="advanced-count">{formatNumber(officialProducts.length)} oficiales</span>
          </summary>

          <section className="table-card homologation-table-card advanced-details-content">
            <table className="homologation-table">
              <thead>
                <tr>
                  <th>Producto leído</th>
                  <th>Nombre original</th>
                  <th>Origen</th>
                  <th>Registros</th>
                  <th>Producto oficial</th>
                  <th>Estado</th>
                </tr>
              </thead>
              <tbody>
                {homologationRows.map((row) => (
                  <tr key={row.product}>
                    <td>{row.product}</td>
                    <td>{row.originalNames.join(", ")}</td>
                    <td>{row.sources.join(", ")}</td>
                    <td>{formatNumber(row.count)}</td>
                    <td>
                      <select value={row.official} onChange={(e) => saveProductAlias(row.product, e.target.value)}>
                        <option value="">Seleccionar producto</option>
                        {officialProducts.map((product) => (
                          <option value={product} key={product}>
                            {product}
                          </option>
                        ))}
                      </select>
                    </td>
                    <td>
                      <span className={`pill ${row.status === "Pendiente" ? "warn" : row.status === "Manual" ? "ok" : "muted"}`}>
                        {row.status}
                      </span>
                    </td>
                  </tr>
                ))}
              </tbody>
            </table>
          </section>
        </details>}

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
            description="Selecciona juntos los históricos de 2024, 2025 y 2026."
            required
            multiple
            onFile={handleSalesFiles}
            fileName={files.ventas}
          />
          <UploadBox
            title="Bajas"
            description="Merma, devoluciones o bajas por producto."
            onFile={handleWasteFile}
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

        {(pendingSalesImport.length > 0 || salesImportStatus) && (
          <section className="sales-import-section">
            <div className="section-heading compact-heading">
              <div>
                <span className="eyebrow">Carga masiva diaria</span>
                <h3>Validación antes de guardar ventas</h3>
                <p>Consolida fecha, producto y sucursal; una segunda carga actualiza el registro en lugar de duplicarlo.</p>
              </div>
              {pendingSalesImport.length > 0 && (
                <button
                  className="primary"
                  type="button"
                  onClick={importSalesToDatabase}
                  disabled={!canSave || salesImporting || salesImportPreview.daily === 0}
                >
                  <Database size={18} /> {salesImporting ? "Guardando..." : "Guardar ventas en la base"}
                </button>
              )}
            </div>

            {pendingSalesImport.length > 0 && (
              <div className="sales-import-grid">
                <div><span>Filas reconocidas</span><strong>{formatNumber(salesImportPreview.recognized)}</strong></div>
                <div><span>Ventas diarias</span><strong>{formatNumber(salesImportPreview.daily)}</strong></div>
                <div><span>Filas a consolidar</span><strong>{formatNumber(salesImportPreview.repeated)}</strong></div>
                <div><span>Totales mensuales excluidos</span><strong>{formatNumber(salesImportPreview.monthlyTotals)}</strong></div>
                <div><span>Sucursales identificadas</span><strong>{formatNumber(salesImportPreview.branches)}</strong></div>
                <div>
                  <span>Rango diario</span>
                  <strong>{salesImportPreview.firstDate ? `${salesImportPreview.firstDate} a ${salesImportPreview.lastDate}` : "Sin fechas"}</strong>
                </div>
              </div>
            )}

            <p className={`sales-import-status ${salesImportStatus.startsWith("No ") ? "error" : ""}`}>
              {salesImportStatus}
            </p>
            {salesImportPreview.monthlyTotals > 0 && pendingSalesImport.length > 0 && (
              <p className="sales-import-note">
                Los totales mensuales permanecen disponibles para el pronóstico, pero no se guardan como ventas diarias.
              </p>
            )}
          </section>
        )}

        <OperationalImportPanel
          title="Validación antes de guardar producción real"
          description="Consolida por fecha, producto y turno; las cargas repetidas actualizan el registro existente."
          dimensionLabel="Turnos identificados"
          pendingRows={pendingProductionImport}
          preview={productionImportPreview}
          status={productionImportStatus}
          importing={productionImporting}
          canSave={canSave}
          onImport={() => importOperationalToDatabase("produccion")}
        />

        <OperationalImportPanel
          title="Validación antes de guardar bajas"
          description="Consolida por fecha, producto, sucursal y motivo sin convertir totales mensuales en bajas diarias."
          dimensionLabel="Sucursales y motivos"
          pendingRows={pendingWasteImport}
          preview={wasteImportPreview}
          status={wasteImportStatus}
          importing={wasteImporting}
          canSave={canSave}
          onImport={() => importOperationalToDatabase("bajas")}
        />

        <details className="historical-validation-section advanced-details">
          <summary className="advanced-summary historical-validation-heading">
            <div>
              <span className="eyebrow">Validación histórica</span>
              <strong>Ventas, producido y bajas</strong>
              <small>Comparativo mayo-junio y referencia de julio</small>
            </div>
            <span className="advanced-count">{historicalValidationRows.length ? `${formatNumber(historicalValidationSummary.products)} productos` : "Opcional"}</span>
          </summary>

          <div className="advanced-details-content historical-validation-content">
              <section className="historical-validation-uploads">
                <UploadBox
                  title="Producido mayo"
                  description="Resumen mensual por producto."
                  onFile={(file) => handleMonthlySummaryFile(file, "producedMay", setProducedMay)}
                  fileName={files.producedMay}
                />
                <UploadBox
                  title="Producido junio"
                  description="Resumen mensual por producto."
                  onFile={(file) => handleMonthlySummaryFile(file, "producedJune", setProducedJune)}
                  fileName={files.producedJune}
                />
                <UploadBox
                  title="Bajas junio"
                  description="Hoja BAJAS ERICK."
                  onFile={(file) => handleMonthlySummaryFile(file, "bajasJune", setBajasJune, parseBajasSummaryFile)}
                  fileName={files.bajasJune}
                />
                <UploadBox
                  title="Bajas julio"
                  description="Referencia real, puede ser corte parcial."
                  onFile={(file) => handleMonthlySummaryFile(file, "bajasJuly", setBajasJuly, parseBajasSummaryFile)}
                  fileName={files.bajasJuly}
                />
              </section>

              <div className="historical-validation-kpis">
                <KpiCard icon={PackageCheck} label="Productos evaluados" value={formatNumber(historicalValidationSummary.products)} caption="Ventas, producido y bajas" />
                <KpiCard icon={Target} label="Precisión promedio junio" value={formatPercent(historicalValidationSummary.precision, 1)} caption="Mayo pronostica junio" />
                <KpiCard icon={Database} label="Bajas junio" value={formatNumber(historicalValidationSummary.bajasJune)} caption="Hoja BAJAS ERICK" />
                <KpiCard icon={ShieldCheck} label="Producción julio ajustada" value={formatNumber(historicalValidationSummary.adjusted)} caption="Incluye bajas esperadas" />
              </div>

              <section className="table-card historical-validation-table-card">
                <table className="historical-validation-table">
                  <thead>
                    <tr>
                      <th>Producto</th>
                      <th>Venta junio</th>
                      <th>Producido junio</th>
                      <th>Bajas junio</th>
                      <th>Demanda ajustada</th>
                      <th>Saldo junio</th>
                      <th>Tasa bajas</th>
                      <th>Pronóstico junio</th>
                      <th>Precisión</th>
                      <th>Pronóstico julio</th>
                      <th>Bajas esperadas</th>
                      <th>Producción julio base</th>
                      <th>Producción ajustada</th>
                    </tr>
                  </thead>
                  <tbody>
                    {historicalValidationRows.map((row) => (
                      <tr key={row.producto}>
                        <td>{row.producto}</td>
                        <td>{formatNumber(row.ventaJunio)}</td>
                        <td>{formatNumber(row.producidoJunio)}</td>
                        <td>{formatNumber(row.bajasJunio)}</td>
                        <td>{formatNumber(row.demandaAjustadaJunio)}</td>
                        <td>{formatNumber(row.saldoJunio)}</td>
                        <td>{formatPercent(row.tasaBajas * 100, 1)}</td>
                        <td>{formatNumber(row.pronosticoJunio, 2)}</td>
                        <td>{row.precisionJunio === null ? "Sin dato" : formatPercent(row.precisionJunio, 1)}</td>
                        <td>{formatNumber(row.pronosticoJulio, 2)}</td>
                        <td>{formatNumber(row.bajasEsperadasJulio, 2)}</td>
                        <td className="strong">{formatNumber(row.produccionSugeridaBase)}</td>
                        <td className="strong">{formatNumber(row.produccionSugeridaAjustada)}</td>
                      </tr>
                    ))}
                  </tbody>
                </table>
                {!historicalValidationRows.length && (
                  <div className="empty">Carga ventas, producido y bajas para generar la validación histórica.</div>
                )}
              </section>

              <div className="historical-validation-actions">
                <p>La producción ajustada es un escenario informativo: agrega las bajas esperadas a la producción base.</p>
                <button className="primary" type="button" onClick={() => exportHistoricalValidation(historicalValidationRows)} disabled={!historicalValidationRows.length}>
                  <Download size={18} /> Exportar validación a Excel
                </button>
              </div>
          </div>
        </details>

        <section className="controls">
          <div className="search">
            <Search size={18} />
            <input placeholder="Buscar producto..." value={query} onChange={(e) => setQuery(e.target.value)} />
          </div>
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
              <p>Usa backtesting, estacionalidad y comportamiento reciente sin leer el mes objetivo ni meses futuros.</p>
              <strong className="row-counter">{formatNumber(dailyRows.length)} filas diarias generadas</strong>
              {files.ventas && (
                <p className={`real-validation-message ${historicalVentas.length ? "success" : "warning"}`}>
                  {!effectiveVentas.length
                    ? "El archivo de ventas se cargó, pero no se reconocieron registros. Revisa las columnas Fecha, Producto y Cantidad."
                    : historicalVentas.length
                      ? `${formatNumber(historicalVentas.length)} registros históricos reconocidos de ${historicalMonthKeys.join(", ")}.`
                      : `Se reconocieron ${formatNumber(effectiveVentas.length)} registros, pero ninguno tiene fecha anterior a ${selectedMonth}. Revisa el mes del archivo.`}
                </p>
              )}
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
                  setHasUnsavedChanges(true);
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
              Margen de seguridad %
              <input
                min="0"
                type="number"
                value={dailyBufferPct}
                onChange={(e) => {
                  setDailyBufferPct(Math.max(0, Number(e.target.value)));
                  setHasUnsavedChanges(true);
                }}
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
                  <th>Margen de seguridad</th>
                  <th>Base con margen</th>
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
            <p>El pronóstico elige el método con menor error en el mes anterior y aplica una calibración limitada.</p>
            <p>La producción sugerida aplica el margen de seguridad.</p>
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
                   caption={`${validationForecast.metodoPronostico} · factor ${formatPercent(validationForecast.tendenciaAplicada * 100, 1)}`}
                 />
                <KpiCard
                  icon={ShieldCheck}
                  label="Producción sugerida mensual"
                  value={formatNumber(validationSummary.produccionSugeridaMensual)}
                  caption={`Regla operativa con ${dailyBufferPct}% de margen`}
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
                    <p>Cada fila muestra margen de seguridad, base con margen y la regla operativa aplicada.</p>
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
                        <th>Margen aplicado</th>
                        <th>Base con margen</th>
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

      </main>
    </div>
  );
}

function App() {
  const [session, setSession] = useState(loadStoredSession);
  const [needsSetup, setNeedsSetup] = useState(false);
  const [authChecking, setAuthChecking] = useState(true);
  const [authSubmitting, setAuthSubmitting] = useState(false);
  const [authError, setAuthError] = useState("");

  useEffect(() => {
    let active = true;
    const request = session
      ? apiRequest("/api/auth/me", { token: session.token })
      : apiRequest("/api/auth/status");
    request
      .then((response) => {
        if (!active) return;
        if (session && response.user) {
          const nextSession = { ...session, user: response.user };
          setSession(nextSession);
          localStorage.setItem(SESSION_STORAGE_KEY, JSON.stringify(nextSession));
        } else {
          setNeedsSetup(Boolean(response.needsSetup));
        }
        setAuthError("");
      })
      .catch((error) => {
        if (!active) return;
        if (session && error.status === 401) {
          localStorage.removeItem(SESSION_STORAGE_KEY);
          setSession(null);
        }
        setAuthError(`No se pudo conectar con el servidor: ${error.message}`);
      })
      .finally(() => {
        if (active) setAuthChecking(false);
      });
    return () => {
      active = false;
    };
  }, [session?.token]);

  async function authenticate(credentials) {
    setAuthSubmitting(true);
    setAuthError("");
    try {
      const response = await apiRequest(needsSetup ? "/api/auth/setup" : "/api/auth/login", {
        method: "POST",
        body: credentials,
      });
      const nextSession = { token: response.token, user: response.user };
      localStorage.setItem(SESSION_STORAGE_KEY, JSON.stringify(nextSession));
      setNeedsSetup(false);
      setSession(nextSession);
    } catch (error) {
      setAuthError(error.message);
    } finally {
      setAuthSubmitting(false);
    }
  }

  async function logout() {
    const current = session;
    localStorage.removeItem(SESSION_STORAGE_KEY);
    setSession(null);
    setNeedsSetup(false);
    setAuthError("");
    if (current?.token) {
      apiRequest("/api/auth/logout", { token: current.token, method: "POST" }).catch(() => {});
    }
    try {
      const response = await apiRequest("/api/auth/status");
      setNeedsSetup(Boolean(response.needsSetup));
    } catch (error) {
      setAuthError(`No se pudo conectar con el servidor: ${error.message}`);
    }
  }

  if (authChecking) {
    return (
      <main className="access-page">
        <section className="access-card access-loading">
          <div className="access-mark"><Database size={26} /></div>
          <h1>Conectando datos</h1>
          <p>Validando la sesión y el respaldo operativo...</p>
        </section>
      </main>
    );
  }

  if (!session) {
    return <AccessScreen needsSetup={needsSetup} loading={authSubmitting} error={authError} onSubmit={authenticate} />;
  }

  return <Dashboard session={session} onLogout={logout} />;
}

export { buildMonthlyCloseSummary, buildOperationalForecastScenario, buildWeeklyProgress, calculateForecast, consolidateOperationalRowsForUpload, consolidateSalesRowsForUpload, filterVentasBeforeMonth, parseBajasSummaryWorkbook, parseExistencias, parseProductionReal, parseSalesOrReturns, parseStock };

if (typeof document !== "undefined") {
  createRoot(document.getElementById("root")).render(<App />);
}
