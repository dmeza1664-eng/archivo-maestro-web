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
  Megaphone,
  Warehouse,
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
const MIN_SALES_DAILY_COVERAGE = 0.7;
// Con 2 meses seguidos del año en curso el modelo ya tiene trayectoria
// y las reglas de calendario #20–#22 siguen activas. El año anterior
// solo se abre si además el año en curso aún no tiene un mes cerrado.
const COLD_START_MIN_RECENT_MONTHS = 2;
// Fracción del ajuste por error del mes anterior (modelo base de validación).
const CALIBRATION_SHRINK = 0.5;
const API_PAGE_SIZE = 4000;
const API_UPLOAD_BATCH_SIZE = 1500;
const MAX_SNAPSHOT_BYTES = 3.5 * 1024 * 1024;

function loadStoredSession() {
  try {
    const raw = sessionStorage.getItem(SESSION_STORAGE_KEY) || localStorage.getItem(SESSION_STORAGE_KEY);
    const value = JSON.parse(raw || "null");
    if (!value?.user) return null;
    const session = { token: API_URL ? value.token || "" : "", user: value.user };
    writeStoredSession(session);
    return session.token || !API_URL ? session : null;
  } catch {
    return null;
  }
}

function writeStoredSession(session) {
  if (!session?.user) return;
  const stored = API_URL ? { token: session.token, user: session.user } : { user: session.user };
  try {
    sessionStorage.setItem(SESSION_STORAGE_KEY, JSON.stringify(stored));
    localStorage.setItem(SESSION_STORAGE_KEY, JSON.stringify(stored));
  } catch {
    // Si el navegador bloquea el storage, la sesión vive en memoria y en la cookie HttpOnly.
  }
}

function clearStoredSession() {
  sessionStorage.removeItem(SESSION_STORAGE_KEY);
  localStorage.removeItem(SESSION_STORAGE_KEY);
}

async function apiRequest(path, { token, method = "GET", body } = {}) {
  const response = await fetch(`${API_URL}${path}`, {
    method,
    credentials: "include",
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

async function fetchOperationalSync(path, token) {
  try {
    const response = await apiRequestAllRows(path, { token });
    return { ok: true, rows: response.rows || [], error: "", status: 200 };
  } catch (error) {
    return {
      ok: false,
      rows: [],
      error: error.message || "Error de sincronización",
      status: error.status || 0,
    };
  }
}

function isDatabaseSyncComplete(sync) {
  return Boolean(sync?.sales?.ok && sync?.production?.ok && sync?.waste?.ok);
}

function databaseSyncStatusText(snapshot, sync) {
  const backup = snapshot ? `Respaldo v${snapshot.version} restaurado. ` : "";
  const parts = [
    sync.sales.ok ? `${sync.sales.count} ventas` : `ventas no sincronizadas (${sync.sales.error})`,
    sync.production.ok ? `${sync.production.count} producciones` : `producción no sincronizada (${sync.production.error})`,
    sync.waste.ok ? `${sync.waste.count} bajas` : `bajas no sincronizadas (${sync.waste.error})`,
  ];
  if (isDatabaseSyncComplete(sync)) {
    return `${backup}Base sincronizada: ${parts[0]}, ${parts[1]} y ${parts[2]}.`;
  }
  return `${backup}Sincronización incompleta: ${parts.join("; ")}.`;
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

async function uploadRowsInBatches({
  endpoint,
  bodyKey,
  rows,
  archivo,
  token,
  onProgress,
  importRunId,
  startBatch = 1,
}) {
  const batches = [];
  for (let offset = 0; offset < rows.length; offset += API_UPLOAD_BATCH_SIZE) {
    batches.push(rows.slice(offset, offset + API_UPLOAD_BATCH_SIZE));
  }
  const runId = importRunId || (globalThis.crypto?.randomUUID ? globalThis.crypto.randomUUID() : `run-${Date.now()}`);
  const totals = {
    received: 0,
    valid: 0,
    rejected: 0,
    consolidated: 0,
    duplicatesInFile: 0,
    inserted: 0,
    updated: 0,
    issues: [],
    importRunId: runId,
    completedBatch: Math.max(0, startBatch - 1),
    totalBatches: batches.length,
  };
  const firstBatch = Math.max(1, startBatch);
  try {
    for (let index = firstBatch - 1; index < batches.length; index += 1) {
      onProgress?.(index + 1, batches.length);
      const response = await apiRequest(endpoint, {
        token,
        method: "POST",
        body: {
          archivo: batches.length > 1 ? `${archivo} [lote ${index + 1}/${batches.length}]` : archivo,
          importRunId: runId,
          batchIndex: index + 1,
          batchTotal: batches.length,
          [bodyKey]: batches[index],
        },
      });
      for (const field of ["received", "valid", "rejected", "consolidated", "duplicatesInFile", "inserted", "updated"]) {
        totals[field] += Number(response[field] || 0);
      }
      totals.issues.push(...(response.issues || []).slice(0, Math.max(0, 25 - totals.issues.length)));
      totals.completedBatch = index + 1;
    }
  } catch (error) {
    error.importProgress = { ...totals };
    const failedBatch = totals.completedBatch + 1;
    error.message = totals.completedBatch
      ? `Lotes 1 a ${totals.completedBatch} de ${batches.length} ya están en la base. Falló el lote ${failedBatch}: ${error.message}`
      : `Falló el lote 1 de ${batches.length}: ${error.message}`;
    throw error;
  }
  return totals;
}

function accumulateImportRun(previous, progress) {
  const completed = Number(progress?.completedBatch || 0);
  return {
    importRunId: progress?.importRunId || previous?.importRunId || null,
    nextBatch: completed + 1,
    totalBatches: Number(progress?.totalBatches || previous?.totalBatches || 0),
    inserted: Number(previous?.inserted || 0) + Number(progress?.inserted || 0),
    updated: Number(previous?.updated || 0) + Number(progress?.updated || 0),
  };
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
    // El catalogo escribe CHESSECAKE y CHEESE CAKE; las ventas mezclan ambas.
    .replace(/\bCHEESE\s+CAKE\b/g, "CHESSECAKE")
    .replace(/CHE{1,2}S{1,2}ECAKE/g, "CHESSECAKE")
    .replace(/\bNUTELLA\b/g, "NUTELA")
    .replace(/\bM\s*&\s*M\b/g, "M & M")
    .replace(/\bM\s+Y\s+M\b/g, "M & M")
    .replace(/\bMYM\b/g, "M & M")
    .replace(/\bINDIVIDUAL\b/g, "IND")
    .replace(/\bTRES\s+LECHES?\b/g, "3 LECHES")
    .replace(/\b3\s+LECHES?\b/g, "3 LECHES")
    .replace(/\bGRANDE\b/g, "GDE")
    .replace(/\bMEDIANO\b/g, "MED")
    .replace(/\bCHICO\b/g, "CH")
    // "GELATINA DE PINA" vs "GELATINA PINA"; no se toca PAY/PASTEL.
    .replace(/\bDE\b/g, " ");
  const compact = normalized.replace(/[^A-Z0-9]/g, "");
  if (compact === "PINAGDE" || compact === "PINAGRANDE") return "PINA GDE";
  if (compact === "MMGDE" || compact === "MYMGDE") return "M & M GDE";
  if (compact === "MMMED" || compact === "MYMMED") return "M & M MED";
  if (compact === "MMCH" || compact === "MYMCH") return "M & M CH";
  return normalized.replace(/\s+/g, " ").trim();
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

function lookupBuiltinProductAlias(product) {
  const normalized = normalizeProduct(product);
  if (!normalized) return "";
  const raw = norm(product).replace(/[.,/\\_-]+/g, " ").replace(/\s+/g, " ").trim();
  const alias = BUILTIN_PRODUCT_ALIASES[product]
    || BUILTIN_PRODUCT_ALIASES[raw]
    || BUILTIN_PRODUCT_ALIASES[normalized]
    || "";
  return alias ? normalizeProduct(alias) : "";
}

function findOfficialProduct(product, officialProducts) {
  const normalized = normalizeProduct(product);
  if (!normalized) return "";
  const builtin = lookupBuiltinProductAlias(product);
  if (builtin && officialProducts.includes(builtin)) return builtin;
  if (officialProducts.includes(normalized)) return normalized;

  const matchKey = productMatchKey(normalized);
  const officialByKey = officialProducts.find((official) => productMatchKey(official) === matchKey);
  return officialByKey || "";
}

function resolveOfficialProduct(product, productAliases, officialProducts) {
  const normalized = normalizeProduct(product);
  if (!normalized) return "";
  const builtin = lookupBuiltinProductAlias(product);
  if (builtin && officialProducts.includes(builtin)) return builtin;
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

// Etiquetas de precio en el nombre ($35, $45) marcan SKUs de venta irregular
// o por evento; no se excluyen del catálogo, pero el pronóstico se vuelve más
// conservador en applyCatalogOutlierCleanup.
function isPriceTaggedProduct(value) {
  return /\$\s*\d+/.test(norm(value));
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
const ACTIVE_PROMOS_STORAGE_KEY = "archivoMaestroActivePromos";
// Grafías de Pepes/catálogo que normalizeProduct no alcanza por sí solo
// cuando el mapeo entra por código o por el nombre crudo de 2024.
const BUILTIN_PRODUCT_ALIASES = {
  "CHEESECAKE MMD CORAZON": "CHESSECAKE MMD CORAZON",
};
const PROMO_DURATION_PRESETS = [
  { value: "hoy", label: "Hoy" },
  { value: "3dias", label: "3 días" },
  { value: "hasta_desactivar", label: "Hasta desactivar" },
];

function todayKey(now = new Date()) {
  return dateKey(now);
}

function defaultInventoryDate(selectedMonth, now = new Date()) {
  const today = todayKey(now);
  if (selectedMonth && today.startsWith(selectedMonth)) return today;
  if (/^\d{4}-(0[1-9]|1[0-2])$/.test(String(selectedMonth || ""))) return `${selectedMonth}-01`;
  return today;
}

function isPlausibleIsoDate(value) {
  const key = dateKey(value);
  if (!/^\d{4}-\d{2}-\d{2}$/.test(key)) return false;
  const year = Number(key.slice(0, 4));
  return year >= 2000 && year <= 2100;
}

function createPromoId() {
  if (typeof crypto !== "undefined" && typeof crypto.randomUUID === "function") {
    return `promo-${crypto.randomUUID()}`;
  }
  return `promo-${Date.now()}-${Math.random().toString(36).slice(2, 8)}`;
}

function resolvePromoEndDate(promo) {
  const start = dateKey(promo?.startDate);
  const preset = String(promo?.durationPreset || "hasta_desactivar");
  if (!start) return dateKey(promo?.endDate) || "";
  if (preset === "hoy") return start;
  if (preset === "3dias") {
    const startDate = parseDateCell(start);
    return startDate ? dateKey(addDays(startDate, 2)) : "";
  }
  if (preset === "hasta_desactivar") return "";
  return dateKey(promo?.endDate) || "";
}

function normalizeActivePromo(raw, { today = todayKey() } = {}) {
  if (!raw || typeof raw !== "object") return null;
  const producto = normalizeProduct(raw.producto || raw.product || "");
  if (!producto) return null;
  const startDate = dateKey(raw.startDate || raw.inicio || today) || today;
  const durationPreset = PROMO_DURATION_PRESETS.some((item) => item.value === raw.durationPreset)
    ? raw.durationPreset
    : raw.endDate || raw.fin
      ? "rango"
      : "hasta_desactivar";
  const endDate = resolvePromoEndDate({
    startDate,
    durationPreset,
    endDate: raw.endDate || raw.fin || "",
  });
  const multiplierRaw = Number(raw.multiplier ?? raw.multiplicador ?? 1);
  const multiplier = Number.isFinite(multiplierRaw) && multiplierRaw > 0 ? multiplierRaw : 1;
  const extraRaw = Number(raw.extraPiecesPerDay ?? raw.piezasExtraDia ?? 0);
  const extraPiecesPerDay = Number.isFinite(extraRaw) ? Math.max(0, extraRaw) : 0;
  return {
    id: String(raw.id || createPromoId()),
    producto,
    startDate,
    endDate,
    durationPreset: durationPreset === "rango" ? "rango" : durationPreset,
    multiplier,
    extraPiecesPerDay,
    note: String(raw.note || raw.nota || "").trim(),
    active: raw.active !== false,
    createdAt: raw.createdAt || new Date().toISOString(),
    updatedAt: raw.updatedAt || raw.createdAt || new Date().toISOString(),
    deactivatedAt: raw.active === false ? raw.deactivatedAt || new Date().toISOString() : null,
  };
}

function sanitizeActivePromos(value, options = {}) {
  if (!Array.isArray(value)) return [];
  const seen = new Set();
  return value
    .map((item) => normalizeActivePromo(item, options))
    .filter((promo) => {
      if (!promo || seen.has(promo.id)) return false;
      seen.add(promo.id);
      return true;
    });
}

function loadStoredActivePromos() {
  try {
    return sanitizeActivePromos(JSON.parse(localStorage.getItem(ACTIVE_PROMOS_STORAGE_KEY) || "[]"));
  } catch {
    return [];
  }
}

function productsMatch(left, right) {
  return Boolean(left) && normalizeProduct(left) === normalizeProduct(right);
}

function promoOverlapsMonth(promo, selectedMonth) {
  if (!promo?.active || !/^\d{4}-(0[1-9]|1[0-2])$/.test(String(selectedMonth || ""))) return false;
  const start = dateKey(promo.startDate);
  const end = resolvePromoEndDate(promo);
  const monthStart = `${selectedMonth}-01`;
  const [year, month] = selectedMonth.split("-").map(Number);
  const monthEnd = dateKey(new Date(year, month, 0));
  if (start && start > monthEnd) return false;
  if (end && end < monthStart) return false;
  return true;
}

function isPromoActiveOnDate(promo, date) {
  if (!promo?.active) return false;
  const key = dateKey(date);
  if (!key) return false;
  const start = dateKey(promo.startDate);
  const end = resolvePromoEndDate(promo);
  if (start && key < start) return false;
  if (end && key > end) return false;
  return true;
}

function findActivePromoForProduct(promos, product, date) {
  const matches = (promos || []).filter(
    (promo) => productsMatch(promo.producto, product) && isPromoActiveOnDate(promo, date)
  );
  return matches.at(-1) || null;
}

function findActivePromoForProductInMonth(promos, product, selectedMonth) {
  const matches = (promos || []).filter(
    (promo) => productsMatch(promo.producto, product) && promoOverlapsMonth(promo, selectedMonth)
  );
  return matches.at(-1) || null;
}

function isPromoListedAsActive(promo, today = todayKey()) {
  if (!promo?.active) return false;
  const end = resolvePromoEndDate(promo);
  return !end || end >= today;
}

function formatPromoUpliftLabel(promo) {
  if (!promo) return "";
  const parts = [];
  if (Number(promo.multiplier) > 0 && Number(promo.multiplier) !== 1) {
    parts.push(`×${Number(promo.multiplier).toLocaleString("es-MX", { maximumFractionDigits: 2 })}`);
  }
  if (Number(promo.extraPiecesPerDay) > 0) {
    parts.push(`+${Number(promo.extraPiecesPerDay)} pzas/día`);
  }
  if (!parts.length) parts.push("sin apagar pronóstico");
  return parts.join(" · ");
}

function formatPromoWindowLabel(promo) {
  if (!promo) return "";
  const start = displayDate(promo.startDate);
  const end = resolvePromoEndDate(promo);
  if (!end) return `${start} → hasta desactivar`;
  if (end === dateKey(promo.startDate)) return start;
  return `${start} → ${displayDate(end)}`;
}

function applyPromoUpliftToQuantity(product, baseQuantity, promo) {
  const base = Math.max(0, Number(baseQuantity) || 0);
  if (!promo) return getProduccionSugerida(product, base);
  const multiplier = Number(promo.multiplier);
  const safeMultiplier = Number.isFinite(multiplier) && multiplier > 0 ? multiplier : 1;
  const extra = Math.max(0, Number(promo.extraPiecesPerDay) || 0);
  return getProduccionSugerida(product, base * safeMultiplier + extra);
}

function dailyBranchStockKey(row) {
  return [dateKey(row?.fecha), normalizeProduct(row?.producto), norm(row?.sucursal)].join("|");
}

function normalizeDailyBranchStockRow(raw) {
  if (!raw || typeof raw !== "object") return null;
  const fecha = dateKey(raw.fecha);
  const productoOriginal = String(raw.productoOriginal || raw.producto || "").trim();
  const producto = normalizeProduct(raw.producto || productoOriginal);
  const sucursal = String(raw.sucursal || raw.tienda || raw.canal || "").trim();
  if (!fecha || !producto || !isUsableBranchLocation(sucursal)) return null;
  const cantidad = Math.max(0, Math.round(toNumber(raw.cantidad ?? raw.stock ?? raw.piezas)));
  return {
    fecha,
    sucursal,
    producto,
    productoOriginal: productoOriginal || producto,
    cantidad,
  };
}

function sanitizeDailyBranchStock(rows) {
  if (!Array.isArray(rows)) return [];
  const byKey = new Map();
  for (const raw of rows) {
    const row = normalizeDailyBranchStockRow(raw);
    if (!row) continue;
    byKey.set(dailyBranchStockKey(row), row);
  }
  return [...byKey.values()].sort((left, right) =>
    left.fecha.localeCompare(right.fecha)
    || left.producto.localeCompare(right.producto, "es")
    || left.sucursal.localeCompare(right.sucursal, "es")
  );
}

function upsertDailyBranchStock(existing, incoming) {
  return sanitizeDailyBranchStock([...(existing || []), ...(incoming || [])]);
}

function removeDailyBranchStockKey(rows, key) {
  return sanitizeDailyBranchStock(rows).filter((row) => dailyBranchStockKey(row) !== key);
}

function sumDailyBranchStockByProductDate(rows) {
  const totals = new Map();
  for (const row of sanitizeDailyBranchStock(rows)) {
    const key = `${row.fecha}|${row.producto}`;
    totals.set(key, (totals.get(key) || 0) + row.cantidad);
  }
  return totals;
}

// El pedido de planta ya trajo lote y promo. Aquí se resta lo que la sucursal
// ya tiene: importa no sobrerepartir, no volver a redondear a lote de pastel.
function applyDailyBranchStockToPlantSuggestion(baseQuantity, stockOnHand) {
  return applyInventoryToProductionSuggestion(baseQuantity, stockOnHand, 0);
}

// A producir: el target del día es el bruto (promo + lote). Se resta lo que
// ya está en sucursales y lo que queda en cuarto frío. Sin rearmar lote.
function applyInventoryToProductionSuggestion(baseQuantity, stockSucursales = 0, cuartoFrio = 0) {
  const base = Math.max(0, Number(baseQuantity) || 0);
  const stock = Math.max(0, Number(stockSucursales) || 0);
  const cold = Math.max(0, Number(cuartoFrio) || 0);
  return Math.max(0, Math.round(base - stock - cold));
}

function dailyColdRoomKey(row) {
  return [dateKey(row?.fecha), normalizeProduct(row?.producto)].join("|");
}

function normalizeDailyColdRoomRow(raw) {
  if (!raw || typeof raw !== "object") return null;
  const fecha = dateKey(raw.fecha);
  const productoOriginal = String(raw.productoOriginal || raw.producto || "").trim();
  const producto = normalizeProduct(raw.producto || productoOriginal);
  if (!fecha || !producto) return null;
  const cantidad = Math.max(0, Math.round(toNumber(
    raw.cantidad ?? raw.cuartoFrio ?? raw.cf ?? raw.restante ?? raw.stock
  )));
  return {
    fecha,
    producto,
    productoOriginal: productoOriginal || producto,
    cantidad,
  };
}

function sanitizeDailyColdRoom(rows) {
  if (!Array.isArray(rows)) return [];
  const byKey = new Map();
  for (const raw of rows) {
    const row = normalizeDailyColdRoomRow(raw);
    if (!row) continue;
    byKey.set(dailyColdRoomKey(row), row);
  }
  return [...byKey.values()].sort((left, right) =>
    left.fecha.localeCompare(right.fecha) || left.producto.localeCompare(right.producto, "es")
  );
}

function upsertDailyColdRoom(existing, incoming) {
  return sanitizeDailyColdRoom([...(existing || []), ...(incoming || [])]);
}

function removeDailyColdRoomKey(rows, key) {
  return sanitizeDailyColdRoom(rows).filter((row) => dailyColdRoomKey(row) !== key);
}

function mapDailyColdRoomByProductDate(rows) {
  const totals = new Map();
  for (const row of sanitizeDailyColdRoom(rows)) {
    totals.set(dailyColdRoomKey(row), row.cantidad);
  }
  return totals;
}

function collectSucursales({ ventas = [], bajas = [], dailyBranchStock = [], extra = [] } = {}) {
  const names = new Set();
  for (const row of [...ventas, ...bajas, ...dailyBranchStock]) {
    const name = String(row?.sucursal || row?.canal || row?.tienda || "").trim();
    if (name) names.add(name);
  }
  for (const name of extra) {
    const cleaned = String(name || "").trim();
    if (cleaned) names.add(cleaned);
  }
  return [...names].sort((left, right) => left.localeCompare(right, "es"));
}

function emptyPromoForm(today = todayKey()) {
  return {
    id: "",
    producto: "",
    startDate: today,
    durationPreset: "hoy",
    multiplier: 1.3,
    extraPiecesPerDay: 0,
    note: "",
  };
}

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

function lastDateOfMonth(year, monthIndex) {
  return dateKey(new Date(year, monthIndex + 1, 0));
}

function inventoryCutoffStatus(cutoffDate, selectedMonth) {
  const cutoff = dateKey(cutoffDate);
  if (!cutoff || !/^\d{4}-(0[1-9]|1[0-2])$/.test(String(selectedMonth || ""))) {
    return { status: "missing", cutoff: cutoff || "", windowStart: "", windowEnd: "" };
  }
  const [year, month] = selectedMonth.split("-").map(Number);
  const windowStart = dateKey(new Date(year, month - 2, 1));
  const windowEnd = dateKey(new Date(year, month, 0));
  return {
    status: cutoff >= windowStart && cutoff <= windowEnd ? "fresh" : "stale",
    cutoff,
    windowStart,
    windowEnd,
  };
}

function productionDateKeyForDemand(date, monthKeySet) {
  const key = dateKey(date);
  if (date.getDay() !== 0) return key;
  const previous = new Date(date);
  previous.setDate(previous.getDate() - 1);
  const previousKey = dateKey(previous);
  if (monthKeySet.has(previousKey)) return previousKey;
  const laterSaturday = [...monthKeySet]
    .filter((candidate) => candidate > key)
    .find((candidate) => parseDateCell(`${candidate}T12:00:00`)?.getDay() === 6);
  return laterSaturday || key;
}

function allocateIntegerTotal(weights, total) {
  const safeTotal = Math.max(0, Math.round(Number(total) || 0));
  const slots = weights.map((weight) => Math.max(0, Number(weight) || 0));
  if (!slots.length || safeTotal === 0) return slots.map(() => 0);
  const weightSum = slots.reduce((sum, weight) => sum + weight, 0);
  if (weightSum <= 0) {
    const allocated = slots.map(() => 0);
    allocated[0] = safeTotal;
    return allocated;
  }
  const raw = slots.map((weight) => (weight / weightSum) * safeTotal);
  const floors = raw.map((value) => Math.floor(value));
  let remain = safeTotal - floors.reduce((sum, value) => sum + value, 0);
  const order = raw
    .map((value, index) => ({ index, frac: value - Math.floor(value) }))
    .sort((a, b) => b.frac - a.frac || a.index - b.index);
  for (let step = 0; step < remain; step += 1) {
    floors[order[step % order.length].index] += 1;
  }
  return floors;
}

function allocateDailyProduction(product, productionWeights, monthlySuggested) {
  if (isOperationalCakeProduct(product)) {
    return productionWeights.map((weight) => getProduccionSugerida(product, weight));
  }
  return allocateIntegerTotal(productionWeights, monthlySuggested);
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

function matchSheetName(workbook, candidate) {
  const normalizedCandidate = norm(candidate);
  return workbook.SheetNames.find(
    (name) => norm(name) === normalizedCandidate || norm(name).includes(normalizedCandidate)
  );
}

function findStockSheetName(workbook) {
  for (const candidate of STOCK_SHEET_CANDIDATES) {
    const match = matchSheetName(workbook, candidate);
    if (match) return match;
  }
  return workbook.SheetNames[0];
}

function findStockSheet(workbook) {
  return workbook.Sheets[findStockSheetName(workbook)];
}

function parseStock(workbook) {
  return parseStockSheet(findStockSheet(workbook));
}

// El catalogo de productos sale de una sola hoja del stock ideal. El archivo
// guarda varias hojas que son fotos de fechas distintas, asi que cambiar de
// hoja cambia el universo del pronostico. Eso ya paso el 2026-08-21 y dejo
// fuera los 16 productos de temporada sin que nada avisara.
function assessStockSheetSelection(workbook) {
  const chosenSheet = findStockSheetName(workbook);
  if (!chosenSheet) return { chosenSheet: "", products: 0, alternatives: [], missingTotal: 0, message: "" };

  const chosenProducts = new Set(parseStockSheet(workbook.Sheets[chosenSheet]).map((row) => row.producto));
  const alternatives = [];
  for (const candidate of STOCK_SHEET_CANDIDATES) {
    const name = matchSheetName(workbook, candidate);
    if (!name || name === chosenSheet) continue;
    const missing = parseStockSheet(workbook.Sheets[name])
      .map((row) => row.producto)
      .filter((product) => !chosenProducts.has(product));
    if (missing.length) alternatives.push({ sheet: name, missing });
  }

  const missingProducts = [...new Set(alternatives.flatMap((item) => item.missing))];
  const sample = missingProducts.slice(0, 3).join(", ");
  return {
    chosenSheet,
    products: chosenProducts.size,
    alternatives,
    missingTotal: missingProducts.length,
    message: missingProducts.length
      ? `El catálogo se tomó de la hoja "${chosenSheet}" con ${chosenProducts.size} productos. Otras hojas del archivo traen ${missingProducts.length} productos que esta no incluye (${sample}${missingProducts.length > 3 ? ", entre otros" : ""}). Confirma que sea la hoja correcta: el universo del pronóstico depende de esta elección.`
      : "",
  };
}

function parseStockSheet(sheet) {
  const rows = XLSX.utils.sheet_to_json(sheet, { header: 1, defval: "" });
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

function inferInventoryCutoffDate(workbook, fileName = "") {
  const sheetName = workbook.SheetNames.find((name) => norm(name).includes("EXISTENCIA EN SUCURSALES")) || workbook.SheetNames[0];
  const sheet = workbook.Sheets[sheetName];
  const rows = sheet ? XLSX.utils.sheet_to_json(sheet, { header: 1, defval: "" }) : [];
  for (let i = 0; i < Math.min(rows.length, 20); i++) {
    for (const cell of (rows[i] || []).slice(0, 12)) {
      const parsed = parseDateCell(cell);
      if (parsed) return { date: dateKey(parsed), source: "celda del archivo" };
    }
  }
  const fromName = parseDateCell(String(fileName || "").replace(/\.[^.]+$/, "").replace(/[_-]+/g, " "));
  if (fromName) return { date: dateKey(fromName), source: "nombre de archivo" };
  const hint = inferMonthHintFromFileName(fileName);
  if (hint) return { date: lastDateOfMonth(hint.year, hint.monthIndex), source: "mes del archivo" };
  return { date: "", source: "" };
}

function looksLikeStockQtyHeader(value) {
  const p = norm(value);
  return ["CANTIDAD", "CANT", "STOCK", "PIEZAS", "PZAS", "EXISTENCIA", "INVENTARIO", "QTY", "UNIDADES"].some(
    (token) => p === token || p.startsWith(`${token} `)
  );
}

function looksLikeDateHeader(value) {
  const p = norm(value);
  return p === "FECHA" || p === "DATE" || p === "DIA" || p.startsWith("FECHA");
}

function looksLikeBranchHeader(value) {
  const p = norm(value);
  return p === "SUCURSAL" || p === "TIENDA" || p === "BRANCH" || p === "CANAL" || p.includes("SUCURSAL");
}

function looksLikeProductHeader(value) {
  const p = norm(value);
  return p === "PRODUCTO" || p === "SKU" || p === "PRODUCT" || p.startsWith("PRODUCTO");
}

function looksLikeBareRestanteHeader(value) {
  const p = norm(value);
  return p === "RESTANTE" || p.startsWith("RESTANTE ");
}

function looksLikeTargetOrIdealName(value) {
  const p = norm(value);
  if (!p) return false;
  return p.includes("A TENER") || p.includes("IDEAL") || p.includes("OBJETIVO") || p.includes("TARGET");
}

function looksLikeRegionalRollupName(value) {
  const p = norm(value);
  if (!p) return false;
  return p.includes("LOCALES") || p.includes("FORANEAS") || p.includes("FORANEOS");
}

// Totales, metas y recortes RAIZ (TOTAL A TENER, Suma suc+CF, TOTAL LOCALES…)
// no son sucursales. Si se sube el workbook completo no deben colarse al inventario.
function looksLikeRollupOrTargetName(value) {
  const p = norm(value);
  if (!p) return false;
  if (looksLikeTargetOrIdealName(p) || looksLikeRegionalRollupName(p)) return true;
  if (p.includes("SUMA")) return true;
  if (p.includes("GRAL") || p.includes("GENERAL")) return true;
  if (p === "TOTAL" || p.startsWith("TOTAL ") || p.includes("TOTAL")) return true;
  if (p.includes("STOCK") && p.includes("SUCURSAL")) return true;
  return false;
}

function looksLikeColdRoomHeader(value, context = {}) {
  const p = norm(value);
  if (!p) return false;
  if (["CF", "C.F", "C.F.", "C F"].includes(p)) return true;
  if (p.includes("CUARTO")) return true;
  if (p.includes("RESTANTE") && (p.includes("CF") || p.includes("C.F") || p.includes("FRIO"))) return true;
  if (p.includes("DISPONIBLE") && p.includes("PLANTA")) return true;
  if (context.treatRestanteAsCold && looksLikeBareRestanteHeader(p)) return true;
  return false;
}

function looksLikeMixedSucursalesCfSheet(name) {
  const p = norm(name);
  if (!p) return false;
  const hasSucursales = p.includes("SUCURSAL") || p.includes("EXIST");
  const hasColdHint = p.includes("CF") || p.includes("C.F") || p.includes("CUARTO") || p.includes("FRIO") || p.includes("RESTANTE");
  return hasSucursales && hasColdHint;
}

function looksLikeColdRoomLocation(value) {
  const p = norm(value);
  if (!p || looksLikeMixedSucursalesCfSheet(p)) return false;
  if (looksLikeColdRoomHeader(p)) return true;
  // Hoja RESTANTE (sin "CF" en el nombre) es el restante de planta / cuarto frío.
  if (p === "RESTANTE" || (p.startsWith("RESTANTE") && !p.includes("SUC"))) return true;
  return false;
}

function sheetHasColdRoomContext(sheetName) {
  const p = norm(sheetName);
  if (!p) return false;
  if (looksLikeColdRoomLocation(sheetName)) return true;
  if (p.includes("RESTANTE") && (p.includes("CF") || p.includes("C.F") || p.includes("FRIO") || p.includes("CUARTO"))) return true;
  if ((p.includes("CF") || p.includes("C.F") || p.includes("CUARTO") || p.includes("FRIO")) && !p.includes("SUCURSAL")) return true;
  return false;
}

function looksLikeTotalSucursalesHeader(value) {
  const p = norm(value);
  if (!p || p.includes("SUMA")) return false;
  if (looksLikeTargetOrIdealName(p) || looksLikeRegionalRollupName(p)) return false;
  if (p === "TOTAL SUCURSALES" || p === "TOTAL SUC" || p === "TOTAL SUC.") return true;
  return p.includes("TOTAL") && p.includes("SUC");
}

function isReservedDailyStockHeader(value) {
  const p = norm(value);
  if (!p) return true;
  if (looksLikeStockQtyHeader(p) || looksLikeDateHeader(p) || looksLikeBranchHeader(p) || looksLikeProductHeader(p)) return true;
  if (looksLikeColdRoomHeader(p) || looksLikeBareRestanteHeader(p) || looksLikeTotalSucursalesHeader(p)) return true;
  if (looksLikeRollupOrTargetName(p)) return true;
  return p === "CF" || p === "C.F." || p.includes("CUARTO") || p.includes("RESTANTE");
}

function isGenericSheetName(name) {
  const p = norm(name);
  return !p || /^HOJA\s*\d*$/.test(p) || /^SHEET\s*\d*$/.test(p) || p === "INVENTARIO" || p === "STOCK" || p === "DATOS";
}

function looksLikeDateOnlySheetName(name) {
  const raw = String(name || "").trim();
  if (!raw) return false;
  return /^\d{4}[/-]\d{1,2}[/-]\d{1,2}$/.test(raw) || /^\d{1,2}[/-]\d{1,2}[/-]\d{2,4}$/.test(raw);
}

function looksLikeLayoutOrCatalogSheet(name) {
  const p = norm(name);
  if (!p || isGenericSheetName(name)) return true;
  if (p === "RAIZ" || p.includes("CATALOGO")) return true;
  if (p.includes("EXIST") && (p.includes("SUCURSAL") || p.includes("CF") || p.includes("RESTANTE"))) return true;
  if (p.includes("STOCK") && p.includes("SUCURSAL")) return true;
  if (looksLikeDateOnlySheetName(name)) return true;
  return false;
}

function sheetNameUsableAsBranch(name) {
  if (!String(name || "").trim()) return false;
  if (isGenericSheetName(name) || looksLikeColdRoomLocation(name)) return false;
  if (looksLikeRollupOrTargetName(name) || looksLikeLayoutOrCatalogSheet(name)) return false;
  return true;
}

function isUsableBranchLocation(value) {
  const sucursal = String(value ?? "").trim();
  if (!sucursal) return false;
  if (looksLikeColdRoomLocation(sucursal) || looksLikeBareRestanteHeader(sucursal)) return false;
  if (looksLikeRollupOrTargetName(sucursal)) return false;
  if (looksLikeStockQtyHeader(sucursal) || looksLikeDateHeader(sucursal) || looksLikeProductHeader(sucursal)) return false;
  const p = norm(sucursal);
  if (p === "SUCURSAL" || p === "TIENDA" || p === "BRANCH" || p === "CANAL") return false;
  return true;
}

function inferDateFromSheetRows(rows, sheetName = "", fallbackDate = "") {
  for (const row of (rows || []).slice(0, 8)) {
    for (const cell of (row || []).slice(0, 10)) {
      const parsed = parseDateCell(cell);
      if (parsed) return dateKey(parsed);
    }
  }
  const fromSheet = parseDateCell(sheetName) || parseDateCell(String(sheetName || "").replace(/[_-]+/g, " "));
  if (fromSheet) return dateKey(fromSheet);
  return dateKey(fallbackDate);
}

function parseDailyInventory(workbook, fallbackDate = "") {
  if (!workbook?.SheetNames?.length) return { branchStock: [], coldRoom: [] };
  const branchParsed = [];
  const coldParsed = [];

  for (const sheetName of workbook.SheetNames) {
    const sheet = workbook.Sheets[sheetName];
    const rows = sheet ? XLSX.utils.sheet_to_json(sheet, { header: 1, defval: "" }) : [];
    if (!rows.length) continue;

    let headerIndex = -1;
    let productCol = -1;
    let qtyCol = -1;
    let dateCol = -1;
    let branchCol = -1;
    let totalSucCol = -1;
    const coldCols = [];
    const wideBranchCols = [];

    for (let i = 0; i < Math.min(rows.length, 12); i++) {
      const row = rows[i] || [];
      const productIdx = row.findIndex((cell) => looksLikeProductHeader(cell));
      if (productIdx < 0) continue;
      const qtyIdx = row.findIndex((cell) => looksLikeStockQtyHeader(cell));
      const dateIdx = row.findIndex((cell) => looksLikeDateHeader(cell));
      const branchIdx = row.findIndex((cell) => looksLikeBranchHeader(cell));
      let totalIdx = row.findIndex((cell) => looksLikeTotalSucursalesHeader(cell));
      const hasBareRestante = row.some((cell, index) => index !== productIdx && looksLikeBareRestanteHeader(cell));
      const treatRestanteAsCold = sheetHasColdRoomContext(sheetName)
        || looksLikeMixedSucursalesCfSheet(sheetName)
        || (totalIdx >= 0 && hasBareRestante);
      const coldIdxs = row
        .map((cell, index) => ({ cell, index }))
        .filter(({ cell, index }) => index !== productIdx && looksLikeColdRoomHeader(cell, { treatRestanteAsCold }));
      if (totalIdx < 0 && looksLikeMixedSucursalesCfSheet(sheetName) && qtyIdx >= 0 && coldIdxs.length) {
        totalIdx = qtyIdx;
      }
      const branchNameCols = row
        .map((cell, index) => ({ cell, index }))
        .filter(({ cell, index }) => index !== productIdx && String(cell ?? "").trim() && !isReservedDailyStockHeader(cell) && isUsableBranchLocation(cell));

      const hasLong = qtyIdx >= 0 && branchIdx >= 0;
      const hasWide = branchNameCols.length >= 1;
      const hasQtySheet = qtyIdx >= 0 && (sheetNameUsableAsBranch(sheetName) || looksLikeColdRoomLocation(sheetName));
      const hasCold = coldIdxs.length >= 1;
      const hasTotalOnly = totalIdx >= 0 && !hasWide && !hasLong;

      if (hasLong || hasWide || hasQtySheet || hasCold || hasTotalOnly) {
        headerIndex = i;
        productCol = productIdx;
        qtyCol = qtyIdx;
        dateCol = dateIdx;
        branchCol = branchIdx;
        totalSucCol = totalIdx;
        coldCols.push(...coldIdxs.map(({ index }) => index));
        if (hasWide) {
          wideBranchCols.push(...branchNameCols.map(({ cell, index }) => ({
            index,
            sucursal: String(cell ?? "").trim(),
          })));
        }
        break;
      }
    }

    if (headerIndex < 0) continue;
    const sheetFallbackDate = inferDateFromSheetRows(rows.slice(0, headerIndex + 1), sheetName, fallbackDate);
    const sheetBranch = sheetNameUsableAsBranch(sheetName) ? String(sheetName).trim() : "";
    const sheetIsCold = looksLikeColdRoomLocation(sheetName);

    for (let i = headerIndex + 1; i < rows.length; i++) {
      const row = rows[i] || [];
      if (!isValidInventoryProduct(row[productCol], row)) continue;
      const productoOriginal = String(row[productCol] ?? "").trim();
      const producto = normalizeProduct(productoOriginal);
      const rowDate = dateCol >= 0 ? dateKey(parseDateCell(row[dateCol])) : "";
      const fecha = rowDate || sheetFallbackDate;
      if (!fecha) continue;

      let emittedBranch = false;

      if (wideBranchCols.length) {
        for (const branch of wideBranchCols) {
          if (!isUsableBranchLocation(branch.sucursal)) continue;
          const raw = row[branch.index];
          if (String(raw ?? "").trim() === "") continue;
          branchParsed.push({
            fecha,
            sucursal: branch.sucursal,
            producto,
            productoOriginal,
            cantidad: toNumber(raw),
          });
          emittedBranch = true;
        }
      } else if (
        qtyCol >= 0
        && qtyCol !== totalSucCol
        && !coldCols.includes(qtyCol)
        && String(row[qtyCol] ?? "").trim() !== ""
      ) {
        const sucursal = branchCol >= 0 ? String(row[branchCol] ?? "").trim() : sheetBranch;
        if (sucursal && looksLikeColdRoomLocation(sucursal)) {
          coldParsed.push({
            fecha,
            producto,
            productoOriginal,
            cantidad: toNumber(row[qtyCol]),
          });
        } else if (isUsableBranchLocation(sucursal)) {
          branchParsed.push({
            fecha,
            sucursal,
            producto,
            productoOriginal,
            cantidad: toNumber(row[qtyCol]),
          });
          emittedBranch = true;
        } else if (sheetIsCold) {
          coldParsed.push({
            fecha,
            producto,
            productoOriginal,
            cantidad: toNumber(row[qtyCol]),
          });
        }
      }

      if (!emittedBranch && !wideBranchCols.length && totalSucCol >= 0 && String(row[totalSucCol] ?? "").trim() !== "") {
        branchParsed.push({
          fecha,
          sucursal: "Sucursales",
          producto,
          productoOriginal,
          cantidad: toNumber(row[totalSucCol]),
        });
      }

      if (coldCols.length) {
        let coldQty = null;
        for (const col of coldCols) {
          const raw = row[col];
          if (String(raw ?? "").trim() === "") continue;
          coldQty = (coldQty ?? 0) + toNumber(raw);
        }
        if (coldQty !== null) {
          coldParsed.push({
            fecha,
            producto,
            productoOriginal,
            cantidad: coldQty,
          });
        }
      }
    }
  }

  return {
    branchStock: sanitizeDailyBranchStock(branchParsed),
    coldRoom: sanitizeDailyColdRoom(coldParsed),
  };
}

function parseDailyBranchStock(workbook, fallbackDate = "") {
  return parseDailyInventory(workbook, fallbackDate).branchStock;
}

function parseDailyColdRoom(workbook, fallbackDate = "") {
  return parseDailyInventory(workbook, fallbackDate).coldRoom;
}

function exportDailyBranchStockTemplate(date, sucursales, products) {
  const fecha = dateKey(date) || todayKey();
  const branches = (sucursales || []).filter(Boolean);
  const usableBranches = branches.length ? branches : ["Sucursal 1"];
  const catalog = (products || []).filter(Boolean);
  const rows = catalog.length
    ? catalog.flatMap((producto) => usableBranches.map((sucursal) => ({
      Fecha: fecha,
      Sucursal: sucursal,
      Producto: producto,
      Cantidad: "",
    })))
    : [{ Fecha: fecha, Sucursal: usableBranches[0], Producto: "", Cantidad: "" }];
  const coldRows = catalog.length
    ? catalog.map((producto) => ({ Fecha: fecha, Producto: producto, "Cuarto frio": "" }))
    : [{ Fecha: fecha, Producto: "", "Cuarto frio": "" }];
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(rows), "Inventario diario");
  XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(coldRows), "Cuarto frio");
  XLSX.writeFile(wb, `plantilla_inventario_sucursales_${fecha}.xlsx`);
}

function inferYearHintFromFileName(fileName = "", fallbackYear = 2026) {
  const yearMatch = norm(fileName).match(/20\d{2}/);
  return yearMatch ? Number(yearMatch[0]) : fallbackYear;
}

const MONTH_FILE_ALIASES = [
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

function monthKeyFromYearIndex(year, monthIndex) {
  return `${year}-${String(monthIndex + 1).padStart(2, "0")}`;
}

function monthsNamedInFileName(fileName = "", fallbackYear = 2026) {
  const text = norm(fileName);
  const year = inferYearHintFromFileName(fileName, fallbackYear);
  const monthIndexes = [...new Set(MONTH_FILE_ALIASES.filter(([alias]) => text.includes(alias)).map(([, monthIndex]) => monthIndex))];
  return monthIndexes.map((monthIndex) => monthKeyFromYearIndex(year, monthIndex));
}

function inferMonthHintFromFileName(fileName = "", fallbackYear = 2026) {
  const named = monthsNamedInFileName(fileName, fallbackYear);
  if (named.length !== 1) return null;
  const [year, month] = named[0].split("-").map(Number);
  return { year, monthIndex: month - 1 };
}

function recordMonthKey(row) {
  return dateKey(row.fecha).slice(0, 7);
}

function isDatabaseSalesSource(sourceFile) {
  return !sourceFile || sourceFile === "Base de datos";
}

function sourceRankForMonth(fileName, monthKey) {
  const named = monthsNamedInFileName(fileName);
  if (named.length === 1 && named[0] === monthKey) return 3;
  if (!named.length) return 2;
  if (named.includes(monthKey)) return 1;
  return 0;
}

function sourceKindForMonth(rows, monthKey) {
  let daily = false;
  let close = false;
  for (const row of rows || []) {
    if (recordMonthKey(row) !== monthKey) continue;
    if (row.monthlyTotal) close = true;
    else daily = true;
  }
  if (daily && close) return "both";
  if (daily) return "daily";
  if (close) return "close";
  return "empty";
}

function pickPreferredSourceName(candidates, monthKey, names) {
  return [...candidates].sort((left, right) => {
    const rankDiff = sourceRankForMonth(right, monthKey) - sourceRankForMonth(left, monthKey);
    return rankDiff || names.indexOf(right) - names.indexOf(left);
  })[0];
}

function describeSourceDecision(decision) {
  const month = displayMonthLabel(decision.month);
  if (decision.strategy === "keep-daily-and-close") {
    const dailyKept = (decision.kept || []).filter((name) => name !== decision.winner);
    const dailyText = dailyKept.length ? ` y se conservó el diario de ${dailyKept.join(", ")}` : "";
    const omittedText = decision.omitted?.length ? ` Se omitió ${decision.omitted.join(", ")}.` : "";
    return `${month}: se usó el cierre de ${decision.winner}${dailyText}.${omittedText}`;
  }
  const omittedText = decision.omitted?.length ? `; se omitió de ${decision.omitted.join(", ")}` : "";
  return `${month}: se usó ${decision.winner}${omittedText}.`;
}

function displayMonthLabel(monthKey) {
  const [year, month] = String(monthKey || "").split("-").map(Number);
  if (!year || !month) return monthKey || "";
  const label = new Date(year, month - 1, 1).toLocaleDateString("es-MX", { month: "long", year: "numeric" });
  return label.charAt(0).toUpperCase() + label.slice(1);
}

function shortMonthLabel(monthKey) {
  const [year, month] = String(monthKey || "").split("-").map(Number);
  if (!year || !month) return monthKey || "";
  const label = new Date(year, month - 1, 1).toLocaleDateString("es-MX", { month: "short" }).replace(".", "");
  return label.charAt(0).toUpperCase() + label.slice(1);
}

// Regla: si un mes trae Excel diario y cierre dedicado, NO son excluyentes.
// El cierre gana el total (monthlyTotal); el diario se conserva para la forma
// (día de semana e impulso GDE de 2ª quincena). Un override sí es exclusivo.
function resolveCanonicalMonthSources(entries, overrides = {}) {
  const names = entries.map((entry) => entry.name);
  const byName = new Map(entries.map((entry) => [entry.name, { name: entry.name, rows: [...(entry.rows || [])] }]));
  const months = new Set();
  for (const entry of byName.values()) {
    for (const row of entry.rows) {
      const key = recordMonthKey(row);
      if (/^\d{4}-\d{2}$/.test(key)) months.add(key);
    }
  }
  const decisions = [];
  for (const month of [...months].sort()) {
    const candidates = names.filter((name) => byName.get(name).rows.some((row) => recordMonthKey(row) === month));
    if (candidates.length < 2) continue;
    const override = overrides[month];
    const kinds = Object.fromEntries(candidates.map((name) => [name, sourceKindForMonth(byName.get(name).rows, month)]));

    const applyExclusive = (winner, strategy) => {
      const omitted = candidates.filter((name) => name !== winner);
      for (const name of omitted) {
        const entry = byName.get(name);
        entry.rows = entry.rows.filter((row) => recordMonthKey(row) !== month);
      }
      decisions.push({ month, winner, omitted, strategy, kept: [winner], kinds });
    };

    if (candidates.includes(override)) {
      applyExclusive(override, "override");
      continue;
    }

    const dailySources = candidates.filter((name) => kinds[name] === "daily" || kinds[name] === "both");
    const closeSources = candidates.filter((name) => kinds[name] === "close" || kinds[name] === "both");
    if (dailySources.length && closeSources.length) {
      const dailyKeep = pickPreferredSourceName(dailySources, month, names);
      const closeKeep = pickPreferredSourceName(closeSources, month, names);
      const keep = new Set([dailyKeep, closeKeep].filter(Boolean));
      const omitted = [];
      for (const name of candidates) {
        if (keep.has(name)) continue;
        const entry = byName.get(name);
        entry.rows = entry.rows.filter((row) => recordMonthKey(row) !== month);
        omitted.push(name);
      }
      if (dailyKeep && closeKeep && dailyKeep !== closeKeep) {
        const dailyEntry = byName.get(dailyKeep);
        dailyEntry.rows = dailyEntry.rows.filter((row) => recordMonthKey(row) !== month || !row.monthlyTotal);
        const closeEntry = byName.get(closeKeep);
        closeEntry.rows = closeEntry.rows.filter((row) => recordMonthKey(row) !== month || row.monthlyTotal);
      }
      decisions.push({
        month,
        winner: closeKeep,
        omitted,
        strategy: "keep-daily-and-close",
        kept: [...keep],
        kinds,
      });
      continue;
    }

    applyExclusive(pickPreferredSourceName(candidates, month, names), "same-kind-winner");
  }
  return { entries: names.map((name) => byName.get(name)), decisions };
}

function shiftMonthKey(monthKey, delta) {
  const [year, month] = String(monthKey).split("-").map(Number);
  const shifted = new Date(year, month - 1 + delta, 1);
  return monthKeyFromYearIndex(shifted.getFullYear(), shifted.getMonth());
}

function inferSheetMonthHint(fileName, sheetName, yearHint) {
  let sheetMonthHint = inferMonthHintFromFileName(sheetName, yearHint);
  const named = monthsNamedInFileName(fileName, yearHint);
  if (named.length < 2 || !sheetMonthHint) return sheetMonthHint;
  const lastNamed = named[named.length - 1];
  const sheetKey = monthKeyFromYearIndex(sheetMonthHint.year, sheetMonthHint.monthIndex);
  if (sheetKey !== shiftMonthKey(lastNamed, 1)) return sheetMonthHint;
  const [year, month] = lastNamed.split("-").map(Number);
  return { year, monthIndex: month - 1 };
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
    if (/^ERICK?\b/.test(norm(productoOriginal))) continue;
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
    let sheetMonthHint = inferSheetMonthHint(fileName, sheetName, yearHint);
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

// Empata ventas y catálogo por nombre oficial o alias ya homologado.
// Evita el 0/0 silencioso en WAPE cuando la venta usa otra grafía.
function collectProductRecords(byProduct, product, original, officialProducts = []) {
  const official = findOfficialProduct(product, officialProducts)
    || findOfficialProduct(original, officialProducts)
    || "";
  const wanted = new Set(
    [product, original, official, lookupBuiltinProductAlias(product), lookupBuiltinProductAlias(original)]
      .filter(Boolean)
      .map((name) => normalizeProduct(name))
      .filter(Boolean)
  );
  const merged = [];
  for (const [key, rows] of byProduct || []) {
    if (wanted.has(normalizeProduct(key))) merged.push(...rows);
  }
  return merged;
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
      .filter(([, value]) => !value.monthlyTotal && value.days.size > 1 && value.days.size / value.daysInMonth < MIN_SALES_DAILY_COVERAGE)
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

function monthHasHistoryRecord(records, monthKey) {
  if (!monthKey) return false;
  return (records || []).some((row) => monthKeyFromRecord(row) === monthKey);
}

function countSameYearConsecutiveRecentMonths(records, selectedMonth) {
  const year = String(selectedMonth || "").slice(0, 4);
  if (!year) return 0;
  let count = 0;
  let cursor = previousMonthKey(selectedMonth);
  while (cursor && cursor.slice(0, 4) === year) {
    if (!monthHasHistoryRecord(records, cursor)) break;
    count += 1;
    cursor = previousMonthKey(cursor);
  }
  return count;
}

function yearHasClosedMonthBefore(selectedMonth) {
  const previous = previousMonthKey(selectedMonth);
  return Boolean(previous && previous.slice(0, 4) === String(selectedMonth || "").slice(0, 4));
}

function priorYearActualRecords(records, selectedMonth) {
  const year = String(selectedMonth || "").slice(0, 4);
  return (records || []).filter((row) => {
    const key = monthKeyFromRecord(row);
    if (!key || key >= selectedMonth || row.inferredZeroMonth) return false;
    return key.slice(0, 4) < year && toNumber(row.cantidad) > 0;
  });
}

function sameYearObservedRecords(records, selectedMonth) {
  const year = String(selectedMonth || "").slice(0, 4);
  return (records || []).filter((row) => {
    const key = monthKeyFromRecord(row);
    return key && key.slice(0, 4) === year && key < selectedMonth;
  });
}

function hasPriorYearHistory(records, selectedMonth) {
  const year = String(selectedMonth || "").slice(0, 4);
  return (records || []).some((row) => {
    const key = monthKeyFromRecord(row);
    if (!key || key >= selectedMonth || row.inferredZeroMonth) return false;
    return key.slice(0, 4) < year && toNumber(row.cantidad) > 0;
  });
}

function sameYearRecordsBefore(records, selectedMonth) {
  const year = String(selectedMonth || "").slice(0, 4);
  return (records || []).filter((row) => {
    const key = monthKeyFromRecord(row);
    return key && key.slice(0, 4) === year && key < selectedMonth;
  });
}

function priorYearEquivalentMonths(monthlyData, selectedMonth) {
  const priorMonth = sameMonthPreviousYear(selectedMonth);
  const priorTotal = monthTotalFromData(monthlyData, priorMonth);
  if (priorTotal > 0) {
    return { monthKeys: [priorMonth], total: priorTotal, usedSameMonth: true };
  }
  const priorYear = String(Number(String(selectedMonth || "").slice(0, 4)) - 1);
  if (!Number.isFinite(Number(priorYear))) return { monthKeys: [], total: 0, usedSameMonth: false };
  const lastMonths = [...monthlyData.keys()]
    .filter((key) => key.startsWith(`${priorYear}-`) && key < selectedMonth)
    .sort()
    .filter((key) => monthTotalFromData(monthlyData, key) > 0)
    .slice(-3);
  if (!lastMonths.length) return { monthKeys: [], total: 0, usedSameMonth: false };
  const total = lastMonths.reduce((sum, key) => sum + monthTotalFromData(monthlyData, key), 0) / lastMonths.length;
  return { monthKeys: lastMonths, total, usedSameMonth: false };
}

function buildColdStartPriorYearModel(records, selectedMonth) {
  const monthlyData = buildMonthlyForecastData(records || []);
  const equivalent = priorYearEquivalentMonths(monthlyData, selectedMonth);
  const growth = computeAnnualGrowthFactor(monthlyData, selectedMonth, 3);
  const targetDays = Math.max(1, datesForMonth(selectedMonth).length);
  if (!(equivalent.total > 0)) {
    return {
      averages: uniformWeekdayAverages(0),
      trend: 1,
      recentMonths: [],
      method: "Arranque en frío: sin año anterior",
      backtestMonth: previousMonthKey(selectedMonth),
      backtestActual: 0,
      backtestForecast: 0,
      backtestError: 0,
    };
  }

  const sourceKey = equivalent.monthKeys.at(-1);
  const sourceShape = weightedWeekdayAverages(monthlyData, [sourceKey], [1]);
  const shapedTotal = forecastTotalFromAverages(sourceShape, selectedMonth);
  const targetTotal = equivalent.total * growth;
  const averages = shapedTotal > 0
    ? scaleForecastAverages(sourceShape, targetTotal / shapedTotal)
    : uniformWeekdayAverages(targetTotal / targetDays);
  const baseLabel = equivalent.usedSameMonth
    ? "Arranque en frío: mismo mes año anterior"
    : "Arranque en frío: promedio últimos meses año anterior";
  const method = Math.abs(growth - 1) > 0.001
    ? `${baseLabel} · nivel ${growth.toFixed(2)}`
    : baseLabel;

  return {
    averages,
    trend: growth,
    recentMonths: equivalent.monthKeys,
    method,
    backtestMonth: previousMonthKey(selectedMonth),
    backtestActual: monthTotalFromData(monthlyData, previousMonthKey(selectedMonth)),
    backtestForecast: 0,
    backtestError: 0,
  };
}

function forecastHidesPriorYearMonths(records, selectedMonth) {
  // Si no hay año anterior, el modelo debe ser bit-idéntico a main.
  // Solo se ocultan meses de un calendario previo cuando ya hay un mes
  // cerrado del año en curso (si no, enero en frío sí puede leerlos).
  return yearHasClosedMonthBefore(selectedMonth) && hasPriorYearHistory(records, selectedMonth);
}

function forecastAllowsYearOverYear(records, selectedMonth) {
  if (forecastHidesPriorYearMonths(records, selectedMonth)) return false;
  const recentCount = countSameYearConsecutiveRecentMonths(records, selectedMonth);
  return recentCount === 0 && hasPriorYearHistory(records, selectedMonth);
}

function monthlyDataWithoutPriorYear(monthlyData, targetMonth) {
  const year = String(targetMonth || "").slice(0, 4);
  const filtered = new Map();
  for (const [key, value] of monthlyData || []) {
    if (String(key).slice(0, 4) >= year) filtered.set(key, value);
  }
  return filtered;
}

function prepareProductForecastHistory(observed, selectedMonth, completeHistoricalMonths) {
  const beforeTarget = (completeHistoricalMonths || []).filter((month) => month < selectedMonth);
  const hidePriorYear = forecastHidesPriorYearMonths(observed, selectedMonth);
  if (!hidePriorYear) {
    const records = fillCompleteZeroMonths(observed || [], beforeTarget);
    return {
      mode: forecastAllowsYearOverYear(observed, selectedMonth) ? "prior-year" : "recent",
      recentCount: countSameYearConsecutiveRecentMonths(records, selectedMonth),
      allowYearOverYear: true,
      records,
    };
  }

  const year = String(selectedMonth || "").slice(0, 4);
  const sameYearComplete = beforeTarget.filter((month) => month.slice(0, 4) === year);
  const sameYearObserved = sameYearObservedRecords(observed, selectedMonth);
  const sameYearFilled = fillCompleteZeroMonths(sameYearObserved, sameYearComplete);
  return {
    mode: "recent",
    recentCount: countSameYearConsecutiveRecentMonths(sameYearFilled, selectedMonth),
    allowYearOverYear: false,
    records: [...sameYearFilled, ...priorYearActualRecords(observed, selectedMonth)],
  };
}

// Arranque en frío con año anterior (#23): enero lee el año previo porque
// aún no hay un mes cerrado del año en curso. Dos casos en los que copiar
// o escalar ese año no es honesto (solo datos anteriores al mes):
//  1) El producto fue intermitente el año anterior (vendió en menos de la
//     mitad de sus meses cerrados) y el MISMO mes de ese año no vendió: no
//     se rellena con meses vecinos (p. ej. febrero de San Valentín); queda 0.
//  2) El producto se apagó al cierre del año anterior: el promedio de sus
//     últimos 2 meses cerrados es menor al 25% del mismo mes de ese año; el
//     pronóstico no pasa de ese nivel de cierre.
// Con un mes cerrado del año en curso o sin año anterior no hace nada, así
// que feb–dic y el walk-forward sin año previo quedan idénticos a main.
const COLD_START_DORMANT_CLOSE_RATIO = 0.25;

function applyColdStartDormantPriorYearGuard(model, records, selectedMonth) {
  if (!model?.averages || yearHasClosedMonthBefore(selectedMonth)) return model;
  if (!hasPriorYearHistory(records, selectedMonth)) return model;
  const monthlyData = buildMonthlyForecastData(records || []);
  const priorYear = String(Number(String(selectedMonth || "").slice(0, 4)) - 1);
  const priorMonths = [...monthlyData.keys()]
    .filter((key) => key.startsWith(`${priorYear}-`) && key < selectedMonth)
    .sort();
  if (!priorMonths.length) return model;
  const total = forecastTotalFromAverages(model.averages, selectedMonth);
  if (!(total > 0)) return model;
  const sameMonthTotal = monthTotalFromData(monthlyData, sameMonthPreviousYear(selectedMonth));
  const activeMonths = priorMonths.filter((key) => monthTotalFromData(monthlyData, key) > 0).length;
  if (!(sameMonthTotal > 0) && activeMonths * 2 < priorMonths.length) {
    return {
      ...model,
      averages: uniformWeekdayAverages(0),
      trend: 0,
      method: `${model.method || "Modelo"} · sin venta el mismo mes del año anterior (intermitente)`,
    };
  }
  const closeMonths = priorMonths.slice(-2);
  const closeLevel = closeMonths.reduce((sum, key) => sum + monthTotalFromData(monthlyData, key), 0) / closeMonths.length;
  if (sameMonthTotal > 0 && closeLevel < sameMonthTotal * COLD_START_DORMANT_CLOSE_RATIO && total > closeLevel) {
    const factor = closeLevel / total;
    return {
      ...model,
      averages: scaleForecastAverages(model.averages, factor),
      trend: (Number(model.trend) || 1) * factor,
      method: `${model.method || "Modelo"} · tope al nivel de cierre del año anterior`,
    };
  }
  return model;
}

// Índice del año anterior en Madres (mayo) y Navidad (diciembre). Con el
// año en curso ya abierto, #23 oculta el año anterior y el pronóstico se
// queda corto en estos dos picos. Si el producto trae el mes del evento y
// el mes previo del año anterior, y el mes previo del año en curso (todos
// > 40 piezas), se estima evento = mes previo actual × (evento / mes previo
// del año anterior) con razón acotada a 0.7–1.6, y el pronóstico sube la
// mitad del camino hacia ese nivel. Solo sube; nunca usa el mes pronosticado.
// Se sostiene en 2024 (con 2023) y en 2025. Junio no entra: empeora 2024.
const EVENT_INDEX_MONTHS = new Set(["madres", "navidad"]);
const EVENT_INDEX_BLEND = 0.5;

function applyEventPriorYearIndex(model, records, selectedMonth) {
  if (!model?.averages) return model;
  const event = calendarEventForMonth(selectedMonth);
  if (!event || !EVENT_INDEX_MONTHS.has(event.id)) return model;
  const monthlyData = buildMonthlyForecastData(records || []);
  const previousMonth = previousMonthKey(selectedMonth);
  const priorEvent = monthTotalFromData(monthlyData, sameMonthPreviousYear(selectedMonth));
  const priorPrevious = monthTotalFromData(monthlyData, sameMonthPreviousYear(previousMonth));
  const currentPrevious = monthTotalFromData(monthlyData, previousMonth);
  if (!(priorEvent > 40 && priorPrevious > 40 && currentPrevious > 40)) return model;
  const total = forecastTotalFromAverages(model.averages, selectedMonth);
  if (!(total > 0)) return model;
  const target = currentPrevious * clamp(priorEvent / priorPrevious, 0.7, 1.6);
  if (!(target > total)) return model;
  const factor = (total + (target - total) * EVENT_INDEX_BLEND) / total;
  return {
    ...model,
    averages: scaleForecastAverages(model.averages, factor),
    trend: (Number(model.trend) || 1) * factor,
    method: `${model.method || "Modelo"} · índice ${event.label} año anterior`,
  };
}

// Producto de pulso con antecedente (#25). El producto vendió el mes pasado
// pero casi nada en los dos meses anteriores (<= 10% del mes pasado). Si el
// año anterior hizo el mismo pulso (vendió en el mes previo y en el mes
// objetivo cayó a <= 10%), el pronóstico baja a la mitad: p. ej. productos de
// San Valentín en marzo o de Madres en junio. No se apaga del todo porque el
// antecedente es de un solo año y el producto puede quedarse (jun-2026 siguió
// vendiendo tras Madres). No actúa si la Cuaresma cae distinto en el mes
// objetivo que el año anterior (más de 7 días de diferencia), porque la
// Semana Santa cambia de mes. Solo usa meses anteriores al pronosticado; sin
// año anterior no hace nada.
const PULSE_FADE_RATIO = 0.1;
const PULSE_FADE_MIN_UNITS = 20;
const PULSE_FADE_LENT_TOLERANCE_DAYS = 7;
const PULSE_FADE_FACTOR = 0.5;

function easterSundayUTC(year) {
  const a = year % 19;
  const b = Math.floor(year / 100);
  const c = year % 100;
  const d = Math.floor(b / 4);
  const e = b % 4;
  const f = Math.floor((b + 8) / 25);
  const g = Math.floor((b - f + 1) / 3);
  const h = (19 * a + b - d - g + 15) % 30;
  const i = Math.floor(c / 4);
  const k = c % 4;
  const l = (32 + 2 * e + 2 * i - h - k) % 7;
  const m = Math.floor((a + 11 * h + 22 * l) / 451);
  const month = Math.floor((h + l - 7 * m + 114) / 31);
  const day = ((h + l - 7 * m + 114) % 31) + 1;
  return Date.UTC(year, month - 1, day);
}

// Días de Cuaresma (miércoles de ceniza a sábado de gloria) dentro del mes.
function lentDaysInMonth(monthKey) {
  const [year, month] = String(monthKey || "").split("-").map(Number);
  if (!year || !month) return 0;
  const DAY = 86400000;
  const easter = easterSundayUTC(year);
  const start = easter - 46 * DAY;
  const end = easter - DAY;
  const monthStart = Date.UTC(year, month - 1, 1);
  const monthEnd = Date.UTC(year, month, 1) - DAY;
  const from = Math.max(start, monthStart);
  const to = Math.min(end, monthEnd);
  return to < from ? 0 : Math.round((to - from) / DAY) + 1;
}

function applyPriorYearPulseFade(model, records, selectedMonth) {
  if (!model?.averages) return model;
  const total = forecastTotalFromAverages(model.averages, selectedMonth);
  if (!(total > 0)) return model;
  const priorSameMonth = sameMonthPreviousYear(selectedMonth);
  if (!priorSameMonth) return model;
  if (Math.abs(lentDaysInMonth(selectedMonth) - lentDaysInMonth(priorSameMonth)) > PULSE_FADE_LENT_TOLERANCE_DAYS) return model;
  const monthlyData = buildMonthlyForecastData(records || []);
  const m1 = previousMonthKey(selectedMonth);
  const m2 = previousMonthKey(m1);
  const m3 = previousMonthKey(m2);
  const last = monthTotalFromData(monthlyData, m1);
  if (!(last >= PULSE_FADE_MIN_UNITS)) return model;
  if (monthTotalFromData(monthlyData, m2) > last * PULSE_FADE_RATIO) return model;
  if (monthTotalFromData(monthlyData, m3) > last * PULSE_FADE_RATIO) return model;
  const priorPulse = monthTotalFromData(monthlyData, sameMonthPreviousYear(m1));
  const priorAfter = monthTotalFromData(monthlyData, priorSameMonth);
  if (!(priorPulse >= PULSE_FADE_MIN_UNITS) || priorAfter > priorPulse * PULSE_FADE_RATIO) return model;
  return {
    ...model,
    averages: scaleForecastAverages(model.averages, PULSE_FADE_FACTOR),
    trend: (Number(model.trend) || 1) * PULSE_FADE_FACTOR,
    method: `${model.method || "Modelo"} · pulso: el año anterior se apagó tras el mismo pulso (x${PULSE_FADE_FACTOR})`,
  };
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
  activePromos = [],
}) {
  const usableHistoricalVentas = filterIncompleteHistoricalMonths(historicalVentas);
  const completeHistoricalMonths = [...new Set(
    usableHistoricalVentas.map((row) => monthKeyFromRecord(row)).filter(Boolean)
  )].sort();
  const officialProducts = getOfficialProducts(stockRows);
  const ventasByProduct = groupByProduct(usableHistoricalVentas);
  const bajasByProduct = groupByProduct(bajas);
  const existMap = new Map(existencias.map((e) => [e.producto, e]));
  const realMap = new Map(aggregateProductionRows(realProduction).map((e) => [e.producto, e.cantidad]));
  const monthDates = datesForMonth(selectedMonth);

  return stockRows.filter((s) => !isSliceProduct(s.producto) && !isPromotionalProduct(s.producto)).map((s) => {
    const product = normalizeProduct(s.producto);
    const observed = collectProductRecords(ventasByProduct, product, s.producto, officialProducts);
    const history = prepareProductForecastHistory(observed, selectedMonth, completeHistoricalMonths);
    const v = history.records;
    const b = collectProductRecords(bajasByProduct, product, s.producto, officialProducts);
    const activePromo = findActivePromoForProductInMonth(activePromos, s.producto || product, selectedMonth);
    const rawForecastModel = applyPriorYearPulseFade(
      applyEventPriorYearIndex(
        applyColdStartDormantPriorYearGuard(
          calculateForecastModelForVersion(v, selectedMonth, product, modelVersion),
          v,
          selectedMonth
        ),
        v,
        selectedMonth
      ),
      v,
      selectedMonth
    );
    const forecastModel = activePromo
      ? {
          ...rawForecastModel,
          method: `${rawForecastModel.method || "Modelo"} · promo activa: sin limpieza de catálogo`,
          catalogCleanup: "omitida por promo activa",
        }
      : applyCatalogOutlierCleanup(
          rawForecastModel,
          v,
          selectedMonth,
          s.producto || product
        );
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

    const ex = existMap.get(product) || existMap.get(s.producto) || { totalSuc: 0, cf: 0, sumaSucCf: 0 };
    const sumaSucCf = ex.sumaSucCf || ex.totalSuc + ex.cf;
    const inventarioObjetivo = s.stock;
    const produccionBalanceada = (inventarioObjetivo - sumaSucCf + produccionSugerida) / 2;
    const produccionRecomendada = Math.max(0, getProduccionSugerida(s.producto, produccionBalanceada));
    const hasRealData = realMap.has(product) || realMap.has(s.producto);
    const produccionReal = hasRealData ? (realMap.get(product) ?? realMap.get(s.producto) ?? 0) : 0;
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
      catalogCleanup: forecastModel.catalogCleanup || "",
      promoActiva: Boolean(activePromo),
      promoEtiqueta: formatPromoUpliftLabel(activePromo),
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

function nextMonthKey(monthKey) {
  const [year, month] = String(monthKey).split("-").map(Number);
  if (!year || !month) return "";
  const next = new Date(year, month, 1);
  return `${next.getFullYear()}-${String(next.getMonth() + 1).padStart(2, "0")}`;
}

function monthTotalFromData(monthlyData, monthKey) {
  return monthlyData.get(monthKey)?.total || 0;
}

// Julio 2025 cayó vs sus vecinos y julio 2026 subió. Copiar el mismo mes del
// año anterior arrastra esa caída atípica. Si el mes análogo está claro por
// debajo de sus vecinos, se usa el nivel de esos vecinos y se conserva la
// forma semanal del mes original. Si el mes análogo no existe, se toma el
// vecino más cercano del año anterior.
// Pasteles grandes: umbral 8% (FRUTAS 2025 cayó ~8.5% y no disparaba el 10%).
function resolvePriorYearSeasonal(monthlyData, targetMonth, options = {}) {
  const priorYearMonth = sameMonthPreviousYear(targetMonth);
  if (!priorYearMonth) return null;
  const dipRatio = Number.isFinite(Number(options.dipRatio)) ? Number(options.dipRatio) : 0.9;
  const priorTotal = monthTotalFromData(monthlyData, priorYearMonth);
  const neighborKeys = [
    previousMonthKey(previousMonthKey(priorYearMonth)),
    previousMonthKey(priorYearMonth),
    nextMonthKey(priorYearMonth),
    nextMonthKey(nextMonthKey(priorYearMonth)),
  ].filter((key) => key && key < targetMonth);
  const neighbors = neighborKeys
    .map((key) => ({ key, total: monthTotalFromData(monthlyData, key) }))
    .filter((row) => row.total > 0);
  const neighborMedian = neighbors.length ? median(neighbors.map((row) => row.total)) : 0;
  const previousOfPriorTotal = monthTotalFromData(monthlyData, previousMonthKey(priorYearMonth));
  const isDip = priorTotal > 0 && neighbors.length >= 1 && (
    (neighborMedian > 0 && priorTotal < neighborMedian * dipRatio) ||
    (previousOfPriorTotal > 0 && priorTotal < previousOfPriorTotal * dipRatio)
  );

  if (priorTotal > 0 && (!isDip || priorYearLooksLikeUnconfirmedFade(monthlyData, targetMonth))) {
    return {
      monthKey: priorYearMonth,
      levelFactor: 1,
      sourceMonths: [priorYearMonth],
      usedProxy: false,
      usedDipCorrection: false,
    };
  }

  if (isDip && neighborMedian > 0 && priorTotal > 0) {
    return {
      monthKey: priorYearMonth,
      levelFactor: neighborMedian / priorTotal,
      sourceMonths: [priorYearMonth, ...neighbors.map((row) => row.key)],
      usedProxy: false,
      usedDipCorrection: true,
    };
  }

  const previousNeighbor = neighbors.find((row) => row.key === previousMonthKey(priorYearMonth));
  const nextNeighbor = neighbors.find((row) => row.key === nextMonthKey(priorYearMonth));
  const proxy = nextNeighbor || previousNeighbor || neighbors.at(-1);
  if (!proxy) return null;
  return {
    monthKey: proxy.key,
    levelFactor: 1,
    sourceMonths: [proxy.key],
    usedProxy: true,
    usedDipCorrection: false,
  };
}

function daysInMonthKey(monthKey) {
  const [year, month] = String(monthKey).split("-").map(Number);
  if (!year || !month) return 0;
  return new Date(year, month, 0).getDate();
}

function salesCoverageStatusLabel(status) {
  if (status === "complete") return "Diario + cierre";
  if (status === "daily-only") return "Solo diario";
  if (status === "close-only") return "Solo cierre";
  if (status === "close-and-partial-daily") return "Cierre y diario incompleto";
  if (status === "partial-daily") return "Diario incompleto";
  return "Sin ventas";
}

function emptySalesMonthCoverage(monthKey) {
  return {
    monthKey,
    daysInMonth: daysInMonthKey(monthKey),
    dailyDays: 0,
    dailyTotal: 0,
    closeTotal: 0,
    dailyRows: 0,
    closeRows: 0,
    hasDaily: false,
    hasClose: false,
    hasPartialDaily: false,
    status: "empty",
  };
}

function buildSalesMonthCoverage(records) {
  const coverage = new Map();
  for (const row of records || []) {
    const monthKey = monthKeyFromRecord(row);
    if (!monthKey) continue;
    const current = coverage.get(monthKey) || {
      dailyDays: new Set(),
      dailyTotal: 0,
      closeTotal: 0,
      dailyRows: 0,
      closeRows: 0,
    };
    if (row.inferredZeroMonth) {
      coverage.set(monthKey, current);
      continue;
    }
    if (row.monthlyTotal) {
      current.closeRows += 1;
      current.closeTotal += toNumber(row.cantidad);
    } else {
      const date = parseDateCell(row.fecha);
      if (date) {
        current.dailyDays.add(date.getDate());
        current.dailyTotal += toNumber(row.cantidad);
        current.dailyRows += 1;
      }
    }
    coverage.set(monthKey, current);
  }

  return [...coverage.keys()].sort().map((monthKey) => {
    const value = coverage.get(monthKey);
    const daysInMonth = daysInMonthKey(monthKey);
    const dailyDays = value.dailyDays.size;
    const hasClose = value.closeRows > 0;
    const hasDaily = dailyDays > 1 && (!daysInMonth || dailyDays / daysInMonth >= MIN_SALES_DAILY_COVERAGE);
    const hasPartialDaily = dailyDays > 0 && !hasDaily;
    let status = "empty";
    if (hasDaily && hasClose) status = "complete";
    else if (hasDaily) status = "daily-only";
    else if (hasClose && hasPartialDaily) status = "close-and-partial-daily";
    else if (hasClose) status = "close-only";
    else if (hasPartialDaily) status = "partial-daily";
    return {
      monthKey,
      daysInMonth,
      dailyDays,
      dailyTotal: value.dailyTotal,
      closeTotal: value.closeTotal,
      dailyRows: value.dailyRows,
      closeRows: value.closeRows,
      hasDaily,
      hasClose,
      hasPartialDaily,
      status,
    };
  });
}

function monthCoverageByKey(coverageRows, monthKey) {
  if (!monthKey) return emptySalesMonthCoverage("");
  return coverageRows.find((row) => row.monthKey === monthKey) || emptySalesMonthCoverage(monthKey);
}

function selectSalesCoverageForDisplay(coverageRows, selectedMonth) {
  const previous = previousMonthKey(selectedMonth);
  const recent = new Set();
  let cursor = previous;
  for (let index = 0; index < 6 && cursor; index += 1) {
    recent.add(cursor);
    cursor = previousMonthKey(cursor);
  }
  const keys = new Set([previous, selectedMonth].filter(Boolean));
  for (const row of coverageRows || []) {
    if (recent.has(row.monthKey) && row.status !== "complete") keys.add(row.monthKey);
  }
  return [...keys].sort().map((key) => monthCoverageByKey(coverageRows, key));
}

function assessForecastFreezeReadiness({
  selectedMonth,
  coverageRows = [],
  databaseSync = null,
  alreadyFrozen = false,
  isAdmin = false,
  capturedStatuses = 0,
  catalogCount = 0,
} = {}) {
  const blockers = [];
  const warnings = [];
  const previous = previousMonthKey(selectedMonth);
  const previousCoverage = monthCoverageByKey(coverageRows, previous);
  const targetCoverage = monthCoverageByKey(coverageRows, selectedMonth);

  if (!databaseSync) {
    blockers.push({ code: "sync-pending", message: "Espera a que termine la sincronización con la base." });
  } else if (!isDatabaseSyncComplete(databaseSync)) {
    blockers.push({ code: "sync-incomplete", message: "No se puede congelar con ventas, producción o bajas incompletas." });
  }

  if (alreadyFrozen && !isAdmin) {
    blockers.push({
      code: "already-frozen",
      message: "Este mes ya está congelado. Solo un administrador puede emitir otra versión.",
    });
  }

  if (targetCoverage.dailyRows > 0 || targetCoverage.closeRows > 0) {
    blockers.push({
      code: "target-has-sales",
      message: `No se puede congelar ${displayMonthLabel(selectedMonth)}: ya hay ventas de ese mes. El pronóstico debe congelarse antes de conocer el resultado.`,
    });
  }

  if (previous) {
    const previousLabel = displayMonthLabel(previous);
    if (!previousCoverage.hasClose && !previousCoverage.hasDaily && !previousCoverage.hasPartialDaily) {
      blockers.push({
        code: "previous-missing",
        message: `No hay ventas de ${previousLabel}. Carga el cierre mensual y el detalle diario.`,
      });
    } else {
      if (!previousCoverage.hasClose) {
        blockers.push({
          code: "previous-missing-close",
          message: `Falta el cierre mensual de ${previousLabel}. Cárgalo junto con el detalle diario.`,
        });
      }
      if (!previousCoverage.hasDaily) {
        warnings.push({
          code: "previous-missing-daily",
          message: previousCoverage.hasClose
            ? `El cierre de ${previousLabel} ya está. No hay Excel diario de ese mes: el pronóstico reparte el total y no distingue sábado de martes. No bloquea septiembre.`
            : previousCoverage.hasPartialDaily
              ? `El detalle diario de ${previousLabel} está incompleto (${previousCoverage.dailyDays} de ${previousCoverage.daysInMonth} días). El pronóstico sigue calculándose.`
              : `Falta el detalle diario de ${previousLabel}. El pronóstico puede seguir, pero no habrá forma por día de semana.`,
        });
      }
    }
  }

  if (catalogCount > 0 && capturedStatuses <= 0) {
    warnings.push({
      code: "missing-status",
      message: "Ningún producto tiene estatus (ACTIVO, BAJA, ESTACIONAL). Se puede congelar, pero las bajas no se aplicarán hasta la revisión.",
    });
  } else if (catalogCount > 0 && capturedStatuses < catalogCount) {
    warnings.push({
      code: "partial-status",
      message: `Estatus capturado en ${capturedStatuses} de ${catalogCount} productos. El resto se tratará como ACTIVO.`,
    });
  }

  return {
    canFreeze: blockers.length === 0,
    blockers,
    warnings,
    previousMonth: previousCoverage,
    targetMonth: targetCoverage,
    previousMonthKey: previous,
  };
}

function sameMonthPreviousYear(monthKey) {
  const [year, month] = String(monthKey).split("-").map(Number);
  return year && month ? `${year - 1}-${String(month).padStart(2, "0")}` : "";
}

function computeAnnualGrowthFactor(monthlyData, targetMonth, lookbackMonths = 1, options = {}) {
  const ratios = [];
  let cursor = previousMonthKey(targetMonth);
  const limit = Math.max(1, Number(lookbackMonths) || 1);
  for (let index = 0; index < limit && cursor; index += 1) {
    const priorYear = sameMonthPreviousYear(cursor);
    const current = monthlyData.get(cursor)?.total || 0;
    const previous = monthlyData.get(priorYear)?.total || 0;
    if (current > 0 && previous >= 10) ratios.push(current / previous);
    cursor = previousMonthKey(cursor);
  }
  if (!ratios.length) return 1;
  const central = ratios.length === 1 ? ratios[0] : median(ratios);
  const strongGrowthMonths = ratios.filter((ratio) => ratio >= 1.12).length;
  const upper = strongGrowthMonths >= 2 ? 1.28 : 1.2;
  let growth = clamp(central, 0.8, upper);
  // Pasteles GDE: un mayo/junio flojo no debe recortar el julio estacional
  // (M & M 2026: YoY 0.86 sobre un julio 2025 sano de 308).
  const keepDecline = Number(options.dampenDecline);
  if (Number.isFinite(keepDecline) && keepDecline >= 0 && keepDecline < 1 && growth < 1) {
    growth = 1 - (1 - growth) * keepDecline;
  }
  return growth;
}

function monthHalfTotals(monthData) {
  let first = 0;
  let second = 0;
  let firstDays = 0;
  let secondDays = 0;
  for (const [key, value] of monthData?.valuesByDate || []) {
    const day = parseDateCell(`${key}T12:00:00`)?.getDate();
    if (!day) continue;
    const qty = Number(value) || 0;
    if (day <= 15) {
      first += qty;
      firstDays += 1;
    } else {
      second += qty;
      secondDays += 1;
    }
  }
  return { first, second, firstDays, secondDays };
}

// Si el mismo mes del año anterior bajó y el siguiente siguió abajo, es baja
// de temporada (PAY DE FRESA), no un hueco atípico. No se le aplica impulso.
function priorYearLooksLikeSeasonalFade(monthlyData, targetMonth) {
  const prior = sameMonthPreviousYear(targetMonth);
  if (!prior) return false;
  const priorTotal = monthTotalFromData(monthlyData, prior);
  const prevTotal = monthTotalFromData(monthlyData, previousMonthKey(prior));
  const nextTotal = monthTotalFromData(monthlyData, nextMonthKey(prior));
  if (priorTotal <= 0 || prevTotal <= 0) return false;
  return priorTotal < prevTotal * 0.85 && nextTotal > 0 && nextTotal < prevTotal * 0.9;
}

// Caída fuerte sin mes siguiente en muestra: GELATINA FRESA $150 ago 2025
// (80 vs jul 420) no es un dip atípico a "corregir" hacia el vecino.
function priorYearLooksLikeUnconfirmedFade(monthlyData, targetMonth) {
  const prior = sameMonthPreviousYear(targetMonth);
  if (!prior) return false;
  const priorTotal = monthTotalFromData(monthlyData, prior);
  const prevTotal = monthTotalFromData(monthlyData, previousMonthKey(prior));
  const nextTotal = monthTotalFromData(monthlyData, nextMonthKey(prior));
  return priorTotal > 0 && prevTotal > 0 && nextTotal <= 0 && priorTotal < prevTotal * 0.5;
}

// Impulso por SKU: si el último mes completo con diario ya corre más fuerte
// en la segunda quincena que en la primera, el mes siguiente suele subir.
// Junio 2026 GDE: segunda/primera ~1.12–1.34 y julio quedó corto. Agosto no
// tiene diario, así que no se mueve. Mayo (Día de las Madres) no se usa.
function computeRecentMomentumFactor(monthlyData, targetMonth, options = {}) {
  const latest = previousMonthKey(targetMonth);
  if (!latest) return 1;
  if (String(latest).endsWith("-05")) return 1;
  const monthData = monthlyData.get(latest);
  if (!monthData || monthData.filledFromMonthlyTotal) return 1;
  if ((monthData.valuesByDate?.size || 0) < 20) return 1;
  const { first, second, firstDays, secondDays } = monthHalfTotals(monthData);
  const minHalf = Number(options.minHalfTotal) || 20;
  if (firstDays < 8 || secondDays < 8 || first < minHalf || second < minHalf) return 1;
  const ratio = second / first;
  const trigger = Number(options.momentumTrigger) || 1.12;
  if (!(ratio >= trigger)) return 1;
  if (priorYearLooksLikeSeasonalFade(monthlyData, targetMonth)) return 1;
  const strength = Number.isFinite(Number(options.momentumStrength)) ? Number(options.momentumStrength) : 0.4;
  const cap = Number.isFinite(Number(options.momentumCap)) ? Number(options.momentumCap) : 1.12;
  return clamp(1 + (ratio - 1) * strength, 1, cap);
}

function applyRecentMomentum(model, records, selectedMonth, options = {}) {
  const monthlyData = buildMonthlyForecastData(records);
  const factor = computeRecentMomentumFactor(monthlyData, selectedMonth, options);
  if (!(factor > 1.001) || !model?.averages) return model;
  const liftPct = Math.round((factor - 1) * 100);
  return {
    ...model,
    averages: scaleForecastAverages(model.averages, factor),
    trend: (model.trend || 1) * factor,
    method: `${model.method || "Modelo"} · impulso reciente ${liftPct}%`,
  };
}

// Mayo sin el mismo mes del año anterior copia abril y, si abril quedó
// por encima, la calibración (0.85–1.15) recorta todavía más. En minis y
// gelatinas eso cae justo en Día de las Madres, que sube frente a abril.
// Primero se quita el recorte. El impulso de nivel va aparte y solo si
// abril no venía ya alto. Con mayo del año pasado, la estacionalidad trae el evento.
function isCakeSizeCategory(category) {
  return category === "Pasteles grandes" || category === "Pasteles medianos" || category === "Pasteles chicos";
}

function isLargeGelatina(product) {
  const value = normalizeProduct(product);
  return value.includes("GELATINA") && /\bGDE\b/.test(value) && !isPriceTaggedProduct(product);
}

function liftColdStartMadresCalibration(model, records, selectedMonth, product) {
  if (!model?.averages) return model;
  if (calendarEventForMonth(selectedMonth)?.id !== "madres") return model;
  const category = productCategory(product);
  if (category !== "Mini medianos" && category !== "Gelatinas" && !isCakeSizeCategory(category)) return model;
  const monthlyData = buildMonthlyForecastData(records || []);
  const priorYearTotal = monthTotalFromData(monthlyData, sameMonthPreviousYear(selectedMonth));
  const recentCount = countSameYearConsecutiveRecentMonths(records || [], selectedMonth);
  if (priorYearTotal > 40 && recentCount < COLD_START_MIN_RECENT_MONTHS) return model;
  const trend = Number(model.trend);
  if (!Number.isFinite(trend) || trend <= 0 || trend >= 0.995) return model;
  return {
    ...model,
    averages: scaleForecastAverages(model.averages, 1 / trend),
    trend: 1,
    method: `${model.method || "Modelo"} · sin recorte pre-Madres`,
  };
}

// Domingo de Pascua (algoritmo gregoriano). Semana Santa cae en ese mes
// o, si el Viernes Santo queda en el mes anterior, en ambos.
function easterSunday(year) {
  const a = year % 19;
  const b = Math.floor(year / 100);
  const c = year % 100;
  const d = Math.floor(b / 4);
  const e = b % 4;
  const f = Math.floor((b + 8) / 25);
  const g = Math.floor((b - f + 1) / 3);
  const h = (19 * a + b - d - g + 15) % 30;
  const i = Math.floor(c / 4);
  const k = c % 4;
  const l = (32 + 2 * e + 2 * i - h - k) % 7;
  const m = Math.floor((a + 11 * h + 22 * l) / 451);
  const month = Math.floor((h + l - 7 * m + 114) / 31);
  const day = ((h + l - 7 * m + 114) % 31) + 1;
  return { year, month, day };
}

function monthContainsSemanaSanta(monthKey) {
  const [year, month] = String(monthKey || "").split("-").map(Number);
  if (!year || !month) return false;
  const easter = easterSunday(year);
  if (easter.month === month) return true;
  const easterDate = new Date(year, easter.month - 1, easter.day);
  const goodFriday = new Date(easterDate);
  goodFriday.setDate(easterDate.getDate() - 2);
  return goodFriday.getMonth() + 1 === month;
}

function isPetitTresLeches(product) {
  const value = normalizeProduct(product);
  return /\bPETIT\b/.test(value) && /\b3 LECHES\b/.test(value);
}

function isIndividualGelatina(product) {
  const value = normalizeProduct(product);
  return value.includes("GELATINA") && /\bIND\b/.test(value);
}

// "1/4 KG" y "1 2 KG DE GALLETA" quedan como "1 4 KG GALLETA" / "1 2 KG GALLETA".
function isKiloGalleta(product) {
  const value = normalizeProduct(product);
  return value.includes("GALLETA") && /\b1\s*[24]\s*KG\b/.test(value);
}

function recentLevelBeforeTarget(records, selectedMonth) {
  const monthlyData = buildMonthlyForecastData(records || []);
  const history = [...monthlyData.keys()].filter((month) => month < selectedMonth).sort();
  const year = String(selectedMonth || "").slice(0, 4);
  const allObserved = history.filter((month) => !monthIsInferredZero(records, month));
  const sameYear = allObserved.filter((month) => month.slice(0, 4) === year);
  const observed = sameYear.length ? sameYear : allObserved;
  const last = observed.length ? (monthlyData.get(observed.at(-1))?.total || 0) : 0;
  const prior = observed
    .slice(-4, -1)
    .map((month) => monthlyData.get(month)?.total || 0)
    .filter((value) => value > 0.5);
  const priorMed = prior.length ? median(prior) : last;
  return { monthlyData, last, priorMed };
}

// Arranque en frío (sin el mismo mes del año anterior): el modelo copia el
// hombro y se queda corto en el evento. No inventa cierres. No aplica si
// el mes previo ya venía alto, ni a gelatinas con hueco ($150), ni a petit
// decorado. Sep–Nov no es ninguno de estos meses.
function coldStartEventUplift(records, selectedMonth, product) {
  const priorYear = monthTotalFromData(
    buildMonthlyForecastData(records || []),
    sameMonthPreviousYear(selectedMonth)
  );
  const recentCount = countSameYearConsecutiveRecentMonths(records || [], selectedMonth);
  if (priorYear > 40 && recentCount < COLD_START_MIN_RECENT_MONTHS) return null;
  const { last, priorMed } = recentLevelBeforeTarget(records, selectedMonth);
  if (!(last > 40) || !(priorMed > 0)) return null;
  const event = calendarEventForMonth(selectedMonth);
  const category = productCategory(product);

  if (event?.id === "madres" && (category === "Mini medianos" || isIndividualGelatina(product))) {
    if (last > priorMed * 1.12) return null;
    return { factor: 1.18, label: "impulso frío Día de las Madres" };
  }

  // Mayo sin año anterior: pasteles GDE/MED/CH también suben por el 10 de mayo.
  // Mismo factor que minis (1.18), in-sample y por debajo del ~1.4× visto en
  // MOKA/FRUTAS del informe Pepes. Sep–Nov no es mayo.
  if (event?.id === "madres" && isCakeSizeCategory(category)) {
    if (last > priorMed * 1.12) return null;
    return { factor: 1.18, label: "impulso frío pastel Madres" };
  }

  // Junio sin el mismo mes del año anterior: el arrastre de Madres baja mayo
  // y el mini se queda corto en Día del Padre. No pasa del mayo observado.
  // Gelatinas y petit no entran: en junio ya no dominaban el error. Sep–Nov
  // no es junio.
  if (event?.id === "padre" && category === "Mini medianos") {
    return { factor: 1.1, label: "impulso frío Día del Padre", capAtLast: true };
  }

  if (monthContainsSemanaSanta(selectedMonth) && isPetitTresLeches(product)) {
    if (last > priorMed * 1.25) return null;
    return { factor: 1.5, label: "impulso frío Semana Santa" };
  }

  if (isKiloGalleta(product) && (event?.id === "madres" || event?.id === "navidad")) {
    if (last > priorMed * 1.2) return null;
    const label = event?.id === "madres" ? "impulso frío kilo Madres" : "impulso frío kilo Navidad";
    return { factor: 1.65, label };
  }

  // Diciembre sin año anterior: Navidad/fin de año sube pasteles grandes y
  // gelatinas GDE. No MED (PINA MED cae), ni petit (3 leches pinero cae),
  // ni SKU con precio. Factor 1.18, in-sample, mismo tope de hombro.
  if (event?.id === "navidad" && (category === "Pasteles grandes" || isLargeGelatina(product))) {
    if (last > priorMed * 1.12) return null;
    const label = category === "Pasteles grandes"
      ? "impulso frío pastel Navidad"
      : "impulso frío gelatina Navidad";
    return { factor: 1.18, label };
  }

  return null;
}

function applyColdStartEventUplift(model, records, selectedMonth, product) {
  if (!model?.averages) return model;
  const uplift = coldStartEventUplift(records, selectedMonth, product);
  if (!uplift) return model;
  const modelTotal = forecastTotalFromAverages(model.averages, selectedMonth);
  if (!(modelTotal > 40)) return model;
  let factor = uplift.factor;
  if (uplift.capAtLast) {
    const { last } = recentLevelBeforeTarget(records, selectedMonth);
    if (!(last > modelTotal * 1.02)) return model;
    factor = Math.min(factor, last / modelTotal);
  }
  if (!(factor > 1.001)) return model;
  return {
    ...model,
    averages: scaleForecastAverages(model.averages, factor),
    trend: (model.trend || 1) * factor,
    method: `${model.method || "Modelo"} · ${uplift.label}`,
  };
}

// Limpieza de outliers de catálogo: no inventar volumen para SKUs dormidos,
// amortiguar caídas hacia cero y evitar extrapolar picos de productos con
// etiqueta de precio / venta intermitente (p. ej. PALETA GALLETA $35).
// No toca reglas de lote de pasteles ni el selector estacional/GDE.
function monthIsInferredZero(records, monthKey) {
  if (!monthKey) return false;
  const rows = (records || []).filter((row) => monthKeyFromRecord(row) === monthKey);
  return rows.length > 0 && rows.every((row) => row.inferredZeroMonth);
}

// Petit, cheesecake y 3 leches se venden todo el año. Si el archivo
// empieza en marzo, ese primer mes es su nivel normal (~400–450), no un
// estreno tipo CAJITA FELIZ (599 y al mes siguiente 37).
function isYearRoundDessertLine(product) {
  const value = normalizeProduct(product);
  return /\b(PETIT|CHESSECAKE|3 LECHES)\b/.test(value);
}

// Pico de un solo mes sin el mismo mes del año anterior: CAJITA FELIZ
// (abril 599 → mayo 37), PETIT 3 LECHES (abril 895 → mayo 525) y el kilo
// de galleta en Día de las Madres (mayo ~620 → junio ~375). No se copia
// entero al mes siguiente. Una subida de pastel ~1.3× (Madres sobre el
// mes previo) no entra: el umbral es 1.85× la mediana reciente.
function unsupportedRecentSpikeCap({ modelTotal, last, observedHistory, monthlyData, selectedMonth, product, ignorePriorYear = false }) {
  if (!(modelTotal > 20) || !(last > 80) || !observedHistory?.length) return null;
  // Si el mes que estamos pronosticando ya vendió parecido el año pasado,
  // el nivel lo sostiene la estacionalidad y no es un pico suelto.
  const targetLastYear = ignorePriorYear
    ? 0
    : monthTotalFromData(monthlyData, sameMonthPreviousYear(selectedMonth));
  if (targetLastYear > modelTotal * 0.75) return null;

  const prior = observedHistory
    .slice(-6, -1)
    .map((month) => monthlyData.get(month)?.total || 0);
  const priorPositive = prior.filter((value) => value > 0.5);

  if (!priorPositive.length) {
    // Sin línea base solo se recorta un estreno enorme de "Otros"
    // (CAJITA FELIZ). Gelatinas, kilos de galleta, pasteles y las líneas
    // de todo el año (petit / cheesecake / 3 leches) arrancan en su nivel
    // normal y no se apagan.
    if (productCategory(product) !== "Otros" || last < 400) return null;
    if (isYearRoundDessertLine(product)) return null;
    // No es estreno si el mes calendario anterior al último ya vendió al menos
    // la mitad (dato del año anterior, previo al mes pronosticado). En febrero,
    // con el año abierto, solo se ve enero y un producto de todo el año
    // (TARTALETA, VASO ARROZ, JERICALLA) caía al 35% de enero.
    const beforeLast = monthTotalFromData(monthlyData, previousMonthKey(observedHistory.at(-1)));
    if (beforeLast >= last * 0.5) return null;
    if (!(modelTotal > last * 0.5)) return null;
    return last * 0.35;
  }

  const priorMed = median(priorPositive);
  if (!(priorMed > 0)) return null;
  // PETIT 3 LECHES en Semana Santa sube ~1.6× (no llega al 1.85× de un
  // estreno). El mes siguiente no es el evento: no copiar ese hombro.
  const lastMonth = observedHistory.at(-1);
  const petitHolidayShoulder = isPetitTresLeches(product)
    && monthContainsSemanaSanta(lastMonth)
    && !monthContainsSemanaSanta(selectedMonth);
  const spikeRatio = petitHolidayShoulder ? 1.45 : 1.85;
  const spikeGap = petitHolidayShoulder ? 40 : 80;
  const spiked = last > Math.max(priorMed * spikeRatio, priorMed + spikeGap);
  if (!spiked || !(modelTotal > Math.max(priorMed * 1.35, 40))) return null;
  return Math.max(priorMed * 1.25, Math.min(last * 0.55, priorMed * 1.9));
}

function applyCatalogOutlierCleanup(model, records, selectedMonth, product) {
  if (!model?.averages) return model;
  const modelTotal = forecastTotalFromAverages(model.averages, selectedMonth);
  if (!(modelTotal > 0.5)) return model;

  const monthlyData = buildMonthlyForecastData(records || []);
  const history = [...monthlyData.keys()].filter((month) => month < selectedMonth).sort();
  // Los ceros inferidos (mes cargado para otro SKU, este no apareció / alias
  // partido) no son una baja real. Si se cuentan, CHEESECAKE / NUTELA / M & M
  // MED se marcan intermitentes y dominan el WAPE.
  const allObserved = history.filter((month) => !monthIsInferredZero(records, month));
  const recentCount = countSameYearConsecutiveRecentMonths(records || [], selectedMonth);
  const targetYear = String(selectedMonth || "").slice(0, 4);
  const sameYearObserved = allObserved.filter((month) => month.slice(0, 4) === targetYear);
  const hidePriorYear = forecastHidesPriorYearMonths(records || [], selectedMonth);
  const observedHistory = hidePriorYear && sameYearObserved.length
    ? sameYearObserved
    : allObserved;
  if (!allObserved.length) {
    return {
      ...model,
      averages: scaleForecastAverages(model.averages, 0),
      trend: 0,
      method: `${model.method || "Modelo"} · limpieza catálogo: sin evidencia de venta`,
      catalogCleanup: "sin evidencia de venta",
    };
  }

  const recent3 = observedHistory.slice(-3).map((month) => monthlyData.get(month)?.total || 0);
  const recent6 = observedHistory.slice(-6).map((month) => monthlyData.get(month)?.total || 0);
  const last = recent3.at(-1) ?? 0;
  const prev = recent3.length >= 2 ? recent3.at(-2) : null;
  const mean6 = recent6.reduce((sum, value) => sum + value, 0) / Math.max(recent6.length, 1);
  const variance6 =
    recent6.reduce((sum, value) => sum + (value - mean6) ** 2, 0) / Math.max(recent6.length, 1);
  const cv6 = mean6 > 0 ? Math.sqrt(variance6) / mean6 : 0;
  const zeroRate6 = recent6.filter((value) => value <= 0.5).length / Math.max(recent6.length, 1);
  const median6 = median(recent6);
  const priceTagged = isPriceTaggedProduct(product);
  const regularCake = isOperationalCakeProduct(product);
  const nearZero = (value) => value <= 0.5;
  // GELATINA FRESA $150: mayo/junio en 0 y julio vuelve (420 en 2025, 438 en 2026).
  // No es baja: es reactivación del mismo mes del año anterior.
  const priorYearTotal = monthTotalFromData(monthlyData, sameMonthPreviousYear(selectedMonth));
  const seasonalReactivation = priorYearTotal > 40 && nearZero(last) && (prev == null || nearZero(prev));
  const intermittent = !regularCake && !seasonalReactivation && (zeroRate6 >= 0.4 || (cv6 >= 1.2 && mean6 < 350));

  let targetTotal = modelTotal;
  let reason = "";

  if (recent3.length >= 2 && nearZero(last) && nearZero(prev) && !seasonalReactivation) {
    targetTotal = 0;
    reason = "sin venta en 2 meses";
  } else if (nearZero(last) && modelTotal > 10 && !seasonalReactivation) {
    const softCap = prev != null && prev > 0 ? Math.min(prev * 0.25, 25) : 0;
    targetTotal = Math.min(modelTotal, softCap);
    reason = "último mes en cero";
  } else if (
    prev != null &&
    prev > 25 &&
    last < prev * 0.45 &&
    last <= 40 &&
    modelTotal > Math.max(last * 1.4, 15)
  ) {
    targetTotal = Math.min(modelTotal, Math.max(last * 1.15, 0));
    reason = "demanda colapsando";
  } else if (
    prev != null &&
    prev > 80 &&
    last < prev * 0.35 &&
    modelTotal > Math.max(last * 1.25, 20)
  ) {
    // Pico tipo promo/evento (CAJITA FELIZ en abril) seguido de caída:
    // no arrastrar el mes siguiente con el nivel estacional del año anterior.
    targetTotal = Math.min(modelTotal, Math.max(last * 0.55, median6));
    reason = "resguardo post-pico";
  }

  if ((priceTagged || intermittent) && targetTotal > 0.5) {
    const prior = observedHistory.slice(-6, -1).map((month) => monthlyData.get(month)?.total || 0);
    const priorPositive = prior.filter((value) => value > 0.5);
    const priorMed = priorPositive.length ? median(priorPositive) : median(prior);
    const isSpike =
      priorMed > 0 &&
      last > Math.max(priorMed * 2.4, priorMed + 60) &&
      last > 80;
    if (isSpike && modelTotal > Math.max(priorMed * 1.4, 30)) {
      const spikeCap = Math.max(priorMed * 1.25, last * 0.35);
      if (spikeCap < targetTotal) {
        targetTotal = spikeCap;
        reason = reason || "no extrapolar pico intermitente";
      }
    }
    if (intermittent && median6 < targetTotal * 0.55) {
      const intermittentCap = Math.max(median6, last * 0.4);
      if (intermittentCap < targetTotal) {
        targetTotal = intermittentCap;
        reason = reason || "venta intermitente";
      }
    }
  }

  const unsupportedCap = unsupportedRecentSpikeCap({
    modelTotal: targetTotal,
    last,
    observedHistory,
    monthlyData,
    selectedMonth,
    product,
    ignorePriorYear: hidePriorYear,
  });
  if (unsupportedCap != null && unsupportedCap < targetTotal) {
    targetTotal = unsupportedCap;
    reason = reason || "no extrapolar pico sin soporte";
  }

  if (!(targetTotal < modelTotal * 0.98)) return model;
  const factor = modelTotal > 0 ? Math.max(0, targetTotal / modelTotal) : 0;
  return {
    ...model,
    averages: scaleForecastAverages(model.averages, factor),
    trend: (model.trend || 1) * factor,
    method: `${model.method || "Modelo"} · limpieza catálogo: ${reason || "resguardo"}`,
    catalogCleanup: reason || "resguardo",
  };
}

function summarizeForecastAccuracy(rows) {
  const actual = rows.reduce((sum, row) => sum + row.actual, 0);
  const forecast = rows.reduce((sum, row) => sum + row.forecast, 0);
  const absoluteError = rows.reduce((sum, row) => sum + Math.abs(row.actual - row.forecast), 0);
  return {
    products: rows.length,
    actual,
    forecast,
    absoluteError,
    wape: actual > 0 ? (absoluteError / actual) * 100 : null,
    mae: rows.length ? absoluteError / rows.length : null,
    inside15: rows.filter((row) => Math.abs(row.actual - row.forecast) <= 15).length,
  };
}

function backtestAbsoluteError(row) {
  if (Number.isFinite(row?.absoluteError)) return Number(row.absoluteError);
  if (Number.isFinite(row?.wape) && Number(row?.actual) > 0) {
    return (Number(row.wape) / 100) * Number(row.actual);
  }
  return 0;
}

function weightedWapeFromBacktests(backtests = []) {
  const usable = (backtests || []).filter((row) => Number(row?.actual) > 0);
  const actual = usable.reduce((sum, row) => sum + toNumber(row.actual), 0);
  const absoluteError = usable.reduce((sum, row) => sum + backtestAbsoluteError(row), 0);
  return actual > 0 ? (absoluteError / actual) * 100 : null;
}

function describeAccuracyWindow(backtests = []) {
  const months = [...new Set((backtests || []).map((row) => row.month).filter(Boolean))].sort();
  if (!months.length) return "";
  if (months.length === 1) return displayMonthLabel(months[0]);
  return `${shortMonthLabel(months[0])}–${shortMonthLabel(months[months.length - 1])}`;
}

function actualFromMap(actualMap, product, original) {
  if (!actualMap || typeof actualMap.entries !== "function") return null;
  let sum = 0;
  let found = false;
  const productNorm = normalizeProduct(product);
  const originalNorm = original ? normalizeProduct(original) : "";
  for (const [key, value] of actualMap.entries()) {
    const keyNorm = normalizeProduct(key);
    if (keyNorm === productNorm || key === product || (original && (key === original || keyNorm === originalNorm))) {
      sum += toNumber(value);
      found = true;
    }
  }
  return found ? sum : null;
}

function analyzeForecastProductErrors(forecastRows, actualMap = new Map(), { topN = 20 } = {}) {
  const rows = (forecastRows || []).map((row) => {
    const product = normalizeProduct(row.producto || row.product || "");
    const forecast = toNumber(row.pronosticoVenta ?? row.forecast);
    const mappedActual = actualFromMap(actualMap, product, row.producto);
    const actual = mappedActual != null ? mappedActual : toNumber(row.actual);
    const absoluteError = Math.abs(actual - forecast);
    return {
      producto: row.producto || product,
      categoria: productCategory(row.producto || product),
      metodo: row.metodoPronostico || row.method || "",
      actual,
      forecast,
      error: actual - forecast,
      absoluteError,
      wape: actual > 0 ? (absoluteError / actual) * 100 : forecast > 0 ? 100 : 0,
    };
  });
  const summary = summarizeForecastAccuracy(rows.map((row) => ({ actual: row.actual, forecast: row.forecast })));
  const totalAbs = rows.reduce((sum, row) => sum + row.absoluteError, 0);
  const totalActual = summary.actual;
  const ranked = rows
    .map((row) => ({
      ...row,
      errorShare: totalAbs > 0 ? row.absoluteError / totalAbs : 0,
      volumeShare: totalActual > 0 ? row.actual / totalActual : 0,
    }))
    .sort((a, b) => b.absoluteError - a.absoluteError || a.producto.localeCompare(b.producto, "es"));
  return {
    ...summary,
    topErrors: ranked.slice(0, topN),
    rows: ranked,
  };
}

function forecastAccuracyTone(wape) {
  if (!Number.isFinite(wape)) return "muted";
  if (wape <= 12) return "ok";
  if (wape <= 18) return "warn";
  return "blocked";
}

function buildForecastHealth({
  stockRows = [],
  ventas = [],
  selectedMonth = "",
  dailyBufferPct = 10,
  currentForecastRows = null,
} = {}) {
  const catalog = stockRows.filter((row) => !isSliceProduct(row.producto) && !isPromotionalProduct(row.producto));
  const months = [...new Set(ventas.map((row) => monthKeyFromRecord(row)).filter(Boolean))].sort();
  const coverage = buildSalesMonthCoverage(ventas);
  const omitted = coverage.filter((row) => row.status === "partial-daily");
  const checks = [];

  if (!catalog.length) {
    checks.push({
      code: "missing-stock",
      level: "error",
      message: "Carga el stock ideal. Sin catálogo no se puede asegurar el pronóstico de planta.",
    });
  }
  if (!ventas.length) {
    checks.push({
      code: "missing-sales",
      level: "error",
      message: "Carga ventas históricas. Sin ellas el pronóstico queda en cero.",
    });
  }

  const priorYear = sameMonthPreviousYear(selectedMonth);
  const hasPriorYear = Boolean(priorYear) && months.includes(priorYear);
  if (catalog.length && ventas.length && priorYear && !hasPriorYear) {
    checks.push({
      code: "missing-year-ago",
      level: "warning",
      message: `No hay ventas de ${displayMonthLabel(priorYear)}. Sin el mismo mes del año anterior no se usa estacionalidad.`,
    });
  }

  if (omitted.length) {
    checks.push({
      code: "incomplete-months",
      level: "warning",
      message: `Meses omitidos por cobertura diaria menor a 70%: ${omitted.map((row) => row.monthKey).join(", ")}.`,
    });
  }

  const closedMonths = months.filter((month) => selectedMonth && month < selectedMonth);
  const backtests = [];
  if (catalog.length) {
    for (const hideMonth of closedMonths.slice(-3)) {
      const historical = filterVentasBeforeMonth(ventas, hideMonth);
      if (!historical.length) continue;
      const forecastRows = calculateForecast({
        stockRows: catalog,
        historicalVentas: historical,
        bajas: [],
        existencias: [],
        realProduction: [],
        selectedMonth: hideMonth,
        dailyBufferPct,
      });
      const actualMap = new Map();
      for (const row of ventas) {
        if (monthKeyFromRecord(row) !== hideMonth) continue;
        const product = normalizeProduct(row.producto);
        actualMap.set(product, (actualMap.get(product) || 0) + toNumber(row.cantidad));
      }
      const comparison = forecastRows.map((row) => ({
        producto: row.producto,
        metodoPronostico: row.metodoPronostico,
        forecast: toNumber(row.pronosticoVenta),
        actual: actualMap.get(normalizeProduct(row.producto)) || 0,
      }));
      const productErrors = analyzeForecastProductErrors(comparison, actualMap, { topN: 12 });
      backtests.push({
        month: hideMonth,
        historyMonths: [...new Set(historical.map((row) => monthKeyFromRecord(row)).filter(Boolean))].sort(),
        ...summarizeForecastAccuracy(comparison),
        topErrors: productErrors.topErrors,
      });
    }
  }

  const currentTotal = Array.isArray(currentForecastRows)
    ? currentForecastRows.reduce((sum, row) => sum + toNumber(row.pronosticoVenta), 0)
    : catalog.length && selectedMonth && filterVentasBeforeMonth(ventas, selectedMonth).length
      ? calculateForecast({
          stockRows: catalog,
          historicalVentas: filterVentasBeforeMonth(ventas, selectedMonth),
          bajas: [],
          existencias: [],
          realProduction: [],
          selectedMonth,
          dailyBufferPct,
        }).reduce((sum, row) => sum + toNumber(row.pronosticoVenta), 0)
      : 0;

  if (catalog.length && ventas.length && selectedMonth && currentTotal <= 0) {
    checks.push({
      code: "zero-forecast",
      level: "error",
      message: `Hay datos cargados pero el pronóstico de ${displayMonthLabel(selectedMonth)} quedó en cero.`,
    });
  } else if (currentTotal > 0) {
    checks.push({
      code: "forecast-ready",
      level: "ok",
      message: `Pronóstico ${displayMonthLabel(selectedMonth)}: modelo activo sobre ${catalog.length} productos del stock.`,
    });
  }

  const latestBacktest = backtests.at(-1) || null;
  const ready = currentTotal > 0 && catalog.length > 0;
  const healthy = ready && (latestBacktest == null || (latestBacktest.wape !== null && latestBacktest.wape <= 18));

  return {
    ready,
    healthy,
    catalogCount: catalog.length,
    months,
    hasPriorYear,
    currentTotal,
    backtests,
    latestBacktest,
    wapeTone: forecastAccuracyTone(latestBacktest?.wape ?? null),
    checks,
    omittedMonths: omitted.map((row) => row.monthKey),
  };
}

function forecastTuningOptions(modelVersion = "") {
  const version = String(modelVersion || "");
  if (version === "csG1S50") return { growthLookback: 1, seasonalSplit: 0.5 };
  if (version === "csG6S50") return { growthLookback: 6, seasonalSplit: 0.5 };
  if (version === "csG3S25") return { growthLookback: 3, seasonalSplit: 0.25 };
  if (version === "csG3S20") return { growthLookback: 3, seasonalSplit: 0.2 };
  if (version === "csG1S25") return { growthLookback: 1, seasonalSplit: 0.25 };
  return { growthLookback: 3, seasonalSplit: 0.5 };
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
        monthData.filledFromMonthlyTotal = true;
      }
    } else {
      monthData.total = [...monthData.valuesByDate.values()].reduce((sum, value) => sum + value, 0);
    }
    rebuildMonthWeekdays(monthData);
  }

  applyInheritedWeekdayShape(monthlyData);
  return monthlyData;
}

function rebuildMonthWeekdays(monthData) {
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

function weekdayShapeStrength(monthData) {
  const averages = [...(monthData.weekdays?.values() || [])]
    .filter((bucket) => bucket.count)
    .map((bucket) => bucket.total / bucket.count);
  if (averages.length < 4) return 0;
  const max = Math.max(...averages);
  const min = Math.min(...averages.filter((value) => value > 0));
  if (!(max > 0) || !Number.isFinite(min)) return 0;
  return (max - min) / max;
}

function applyInheritedWeekdayShape(monthlyData) {
  const donors = [...monthlyData.entries()]
    .filter(([, data]) => !data.filledFromMonthlyTotal && weekdayShapeStrength(data) >= 0.08)
    .map(([key]) => key)
    .sort();
  if (!donors.length) return;

  for (const [monthKey, monthData] of monthlyData.entries()) {
    if (!monthData.filledFromMonthlyTotal || !monthData.total || !monthData.syntheticDays) continue;
    const donorKey = [...donors].reverse().find((key) => key < monthKey) || donors.find((key) => key > monthKey);
    if (!donorKey) continue;
    const donor = monthlyData.get(donorKey);
    const donorWeekdayAvg = new Map();
    for (const [weekday, bucket] of donor.weekdays.entries()) {
      if (bucket.count) donorWeekdayAvg.set(weekday, bucket.total / bucket.count);
    }
    const donorMean = [...donorWeekdayAvg.values()].reduce((sum, value) => sum + value, 0) / Math.max(1, donorWeekdayAvg.size);
    if (donorMean <= 0) continue;
    const [year, month] = monthKey.split("-").map(Number);
    const shaped = new Map();
    for (let day = 1; day <= monthData.syntheticDays; day += 1) {
      const date = new Date(year, month - 1, day);
      shaped.set(dateKey(date), donorWeekdayAvg.get(date.getDay()) || donorMean);
    }
    const shapedTotal = [...shaped.values()].reduce((sum, value) => sum + value, 0);
    const factor = shapedTotal > 0 ? monthData.total / shapedTotal : 0;
    monthData.valuesByDate = new Map([...shaped.entries()].map(([key, value]) => [key, value * factor]));
    monthData.inheritedWeekdayShapeFrom = donorKey;
    rebuildMonthWeekdays(monthData);
  }
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

function thirdSundayOfJune(year) {
  const first = new Date(year, 5, 1);
  const firstSunday = 1 + ((7 - first.getDay()) % 7);
  return firstSunday + 14;
}

// Ventanas validadas en CONTROL_MODELO: Madres ~8-11 mayo; Padre vie-lun
// alrededor del tercer domingo de junio. El incremento se queda en el mes
// del evento; no debe inflar el "último mes" del mes siguiente.
function calendarEventForMonth(monthKey) {
  const [year, month] = String(monthKey || "").split("-").map(Number);
  if (!year || !month) return null;
  if (month === 5) {
    return { id: "madres", label: "Día de las Madres", upliftShare: 0.9, peakDays: [8, 9, 10, 11] };
  }
  if (month === 6) {
    const sunday = thirdSundayOfJune(year);
    return { id: "padre", label: "Día del Padre", upliftShare: 0.35, peakDays: [sunday - 2, sunday - 1, sunday, sunday + 1] };
  }
  if (month === 12) {
    return { id: "navidad", label: "Navidad / fin de año", upliftShare: 0.7, peakDays: [12, 24, 25, 31] };
  }
  return null;
}

function scaleMonthDataForCarryover(monthData, scale) {
  if (!monthData || !(scale < 0.999)) return monthData;
  const weekdays = new Map();
  for (const [weekday, bucket] of monthData.weekdays || []) {
    weekdays.set(weekday, { total: (bucket.total || 0) * scale, count: bucket.count });
  }
  return {
    ...monthData,
    total: (monthData.total || 0) * scale,
    dailyRate: (monthData.dailyRate || 0) * scale,
    weekdays,
    eventCarryoverScale: scale,
  };
}

function computeEventCarryoverScale(monthlyData, sourceMonth, targetMonth) {
  const event = calendarEventForMonth(sourceMonth);
  if (!event) return 1;
  if (calendarEventForMonth(targetMonth)?.id === event.id) return 1;
  const sourceTotal = monthTotalFromData(monthlyData, sourceMonth);
  if (!(sourceTotal > 0)) return 1;
  const neighborKeys = [
    previousMonthKey(previousMonthKey(sourceMonth)),
    previousMonthKey(sourceMonth),
    nextMonthKey(sourceMonth),
    nextMonthKey(nextMonthKey(sourceMonth)),
  ].filter((key) => key && key < targetMonth && !calendarEventForMonth(key));
  const neighbors = neighborKeys
    .map((key) => monthTotalFromData(monthlyData, key))
    .filter((total) => total > 0);
  const neighborMed = neighbors.length ? median(neighbors) : 0;
  if (!(neighborMed > 0) || sourceTotal <= neighborMed * 1.08) return 1;
  return clamp(neighborMed / sourceTotal, 0.72, 1);
}

// El impulso de 2ª quincena solo vale para el mes siguiente. Si junio
// aceleró y julio saltó, agosto no debe copiar ese salto como piso nuevo
// cuando el mismo mes del año anterior ya estaba en el nivel pre-impulso.
// M & M GDE julio 324 ≈ agosto 2025 324: es el piso estacional, no un pico.
function computeImpulseCarryoverScale(monthlyData, sourceMonth, targetMonth) {
  if (!sourceMonth || !targetMonth) return 1;
  if (nextMonthKey(sourceMonth) !== targetMonth) return 1;
  if (calendarEventForMonth(sourceMonth)) return 1;

  const accelMonth = previousMonthKey(sourceMonth);
  if (!accelMonth || String(accelMonth).endsWith("-05")) return 1;
  const accelData = monthlyData.get(accelMonth);
  if (!accelData || accelData.filledFromMonthlyTotal) return 1;
  if ((accelData.valuesByDate?.size || 0) < 20) return 1;
  const { first, second, firstDays, secondDays } = monthHalfTotals(accelData);
  if (firstDays < 8 || secondDays < 8 || first < 20 || second < 20) return 1;
  if (!(second / first >= 1.1)) return 1;

  const sourceTotal = monthTotalFromData(monthlyData, sourceMonth);
  const accelTotal = monthTotalFromData(monthlyData, accelMonth);
  const targetPriorYear = monthTotalFromData(monthlyData, sameMonthPreviousYear(targetMonth));
  if (!(sourceTotal > 0) || !(accelTotal > 0)) return 1;
  if (sourceTotal <= accelTotal * 1.08) return 1;
  if (targetPriorYear > 0 && sourceTotal <= targetPriorYear * 1.08) return 1;

  const baseline = targetPriorYear > 0
    ? median([accelTotal, targetPriorYear].filter((value) => value > 0))
    : accelTotal;
  if (!(baseline > 0) || sourceTotal <= baseline * 1.08) return 1;
  return clamp(baseline / sourceTotal, 0.82, 1);
}

function monthlyDataWithoutEventCarryover(monthlyData, targetMonth, forecastOptions = {}) {
  const source = forecastOptions.allowYearOverYear === false
    ? monthlyDataWithoutPriorYear(monthlyData, targetMonth)
    : monthlyData;
  const adjusted = new Map();
  for (const [monthKey, monthData] of source.entries()) {
    if (monthKey >= targetMonth) {
      adjusted.set(monthKey, monthData);
      continue;
    }
    const eventScale = computeEventCarryoverScale(source, monthKey, targetMonth);
    const impulseScale = computeImpulseCarryoverScale(source, monthKey, targetMonth);
    adjusted.set(monthKey, scaleMonthDataForCarryover(monthData, Math.min(eventScale, impulseScale)));
  }
  return adjusted;
}

function buildForecastCandidates(monthlyData, targetMonth, forecastOptions = {}) {
  const growthLookback = Math.max(1, Number(forecastOptions.growthLookback) || 1);
  const seasonalSplit = Number.isFinite(Number(forecastOptions.seasonalSplit))
    ? Number(forecastOptions.seasonalSplit)
    : 0.5;
  const historicalMonths = [...monthlyData.keys()].filter((monthKey) => monthKey < targetMonth).sort();
  const targetYear = String(targetMonth || "").slice(0, 4);
  const recentPool = forecastOptions.allowYearOverYear === false
    ? historicalMonths.filter((monthKey) => monthKey.slice(0, 4) === targetYear)
    : historicalMonths;
  const recentMonths = recentPool.slice(-3);
  const latestMonth = recentMonths.at(-1);
  const targetDays = datesForMonth(targetMonth).length;
  // Nivel reciente: si mayo (Madres) es un pico, no copiarlo a junio.
  // Tampoco copiar a agosto un julio que solo saltó por el impulso de junio.
  // Crecimiento YoY y "mismo mes año anterior" siguen en crudo.
  const hidePriorYear = forecastOptions.allowYearOverYear === false;
  const seasonalData = hidePriorYear
    ? monthlyDataWithoutPriorYear(monthlyData, targetMonth)
    : monthlyData;
  const recentData = monthlyDataWithoutEventCarryover(monthlyData, targetMonth, forecastOptions);
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

  const rates = recentMonths.map((monthKey) => recentData.get(monthKey)?.dailyRate || 0);
  const totals = recentMonths.map((monthKey) => recentData.get(monthKey)?.total || 0);
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
      weightedWeekdayAverages(recentData, recentMonths.slice(-2), [0.35, 0.65]),
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
    weightedWeekdayAverages(recentData, [latestMonth], [1]),
    [latestMonth]
  );

  const seasonalRef = resolvePriorYearSeasonal(seasonalData, targetMonth, forecastOptions);
  if (seasonalRef) {
    const previousTargetMonth = previousMonthKey(targetMonth);
    const previousYearReference = hidePriorYear ? "" : sameMonthPreviousYear(previousTargetMonth);
    const growth = hidePriorYear
      ? 1
      : computeAnnualGrowthFactor(monthlyData, targetMonth, growthLookback, forecastOptions);
    const seasonalBase = weightedWeekdayAverages(seasonalData, [seasonalRef.monthKey], [1]);
    const adjustedSeasonal = scaleForecastAverages(seasonalBase, growth * seasonalRef.levelFactor);
    addCandidate(
      "Mismo mes año anterior",
      adjustedSeasonal,
      [...new Set([seasonalRef.monthKey, ...seasonalRef.sourceMonths, previousTargetMonth, previousYearReference].filter(Boolean))]
    );
    const recentWeights = recentMonths.length === 1 ? [1] : recentMonths.length === 2 ? [0.35, 0.65] : [0.2, 0.3, 0.5];
    const recentBase = weightedWeekdayAverages(recentData, recentMonths, recentWeights);
    const seasonalTotal = forecastTotalFromAverages(adjustedSeasonal, targetMonth);
    const recentTotal = forecastTotalFromAverages(recentBase, targetMonth);
    const referencesDiffer = Math.max(seasonalTotal, recentTotal) > 0 &&
      Math.abs(seasonalTotal - recentTotal) / Math.max(seasonalTotal, recentTotal) > seasonalSplit;
    let seasonalShare = referencesDiffer ? 0.5 : 0.9;
    // Si lo reciente ya corre por encima del año anterior, no arrastrar el
    // pronóstico hacia una baja que este año no se está cumpliendo.
    if (recentTotal > seasonalTotal && seasonalTotal > 0) {
      const lift = (recentTotal - seasonalTotal) / recentTotal;
      seasonalShare = referencesDiffer
        ? clamp(0.5 - lift * 0.3, 0.35, 0.5)
        : clamp(0.9 - lift, 0.45, 0.9);
    }
    addCandidate(
      "Estacional-reciente",
      blendForecastAverages(adjustedSeasonal, recentBase, seasonalShare),
      [...new Set([seasonalRef.monthKey, ...seasonalRef.sourceMonths, ...recentMonths])]
    );
  }

  return candidates;
}

function calculateForecastModelLegacy(records, selectedMonth, useLatestAvailableBacktest = true, forecastOptions = {}) {
  const monthlyData = buildMonthlyForecastData(records);
  const historicalMonths = [...monthlyData.keys()].filter((month) => month < selectedMonth).sort();
  const previousCalendarMonth = previousMonthKey(selectedMonth);
  const backtestMonth = !useLatestAvailableBacktest || monthlyData.has(previousCalendarMonth)
    ? previousCalendarMonth
    : historicalMonths.at(-1) || previousCalendarMonth;
  const backtestActual = monthlyData.get(backtestMonth)?.total || 0;
  const backtestCandidates = buildForecastCandidates(monthlyData, backtestMonth, forecastOptions);
  let selectedMethod = "Último mes por día";
  let backtestError = Infinity;
  for (const [method, candidate] of backtestCandidates.entries()) {
    const error = Math.abs(backtestActual - candidate.total);
    if (error < backtestError) {
      selectedMethod = method;
      backtestError = error;
    }
  }

  const targetCandidates = buildForecastCandidates(monthlyData, selectedMonth, forecastOptions);
  const fallbackMethod = [...targetCandidates.keys()].find((method) => method !== "Sin histórico") || selectedMethod;
  const resolvedMethod = targetCandidates.has(selectedMethod) ? selectedMethod : fallbackMethod;
  const selected = targetCandidates.get(resolvedMethod) || {
    averages: uniformWeekdayAverages(0),
    total: 0,
    sourceMonths: [],
  };
  const previousPrediction = backtestCandidates.get(selectedMethod)?.total || 0;
  const rawCalibration = backtestActual > 0 && previousPrediction > 0
    ? clamp(backtestActual / previousPrediction, 0.85, 1.15)
    : 1;
  // Media calibración: corregir el 100% del error del mes anterior persigue
  // el ruido (doble persecución de tendencia en minis). Se aplica la mitad,
  // salvo que el mes de validación sea un evento (Madres, Padre, Navidad o
  // Semana Santa), donde el error sí trae información del nivel.
  // Validado fuera de muestra en 2024 (con 2023): 14.58 -> 14.23 Ene–Dic.
  const validationIsEvent = Boolean(
    calendarEventForMonth(backtestMonth) || monthContainsSemanaSanta(backtestMonth)
  );
  const calibration = validationIsEvent
    ? rawCalibration
    : 1 + (rawCalibration - 1) * CALIBRATION_SHRINK;
  const averages = scaleForecastAverages(selected.averages, calibration);

  return {
    averages,
    trend: calibration,
    recentMonths: selected.sourceMonths,
    method: resolvedMethod,
    backtestMonth,
    backtestActual,
    backtestForecast: previousPrediction,
    backtestError: Number.isFinite(backtestError) ? backtestError : 0,
  };
}

function calculateForecastModelSeasonal(records, selectedMonth, seasonalWeight, forecastOptions = {}) {
  const product = normalizeProduct(records[0]?.producto || "");
  const useConservativeBacktest = /\b(GALLETA|BOLLO|PAN)\b/.test(product);
  const legacy = calculateForecastModelLegacy(records, selectedMonth, !useConservativeBacktest, forecastOptions);
  if (useConservativeBacktest) return legacy;
  const monthlyData = buildMonthlyForecastData(records);
  const candidates = buildForecastCandidates(monthlyData, selectedMonth, forecastOptions);
  const seasonal = candidates.get("Estacional-reciente") || candidates.get("Mismo mes año anterior");
  const recent = candidates.get("Último mes por día") || candidates.get("Último total mensual");
  // Hueco reciente (GELATINA FRESA $150 may–jun = 0) no debe tirar el
  // mismo mes del año anterior si ESE mes sí vendió. Un proxy de vecinos
  // no cuenta: CAJITA FELIZ no tiene julio 2025 y debe seguir en 0.
  if (!recent || recent.total <= 0.5) {
    const priorMonth = sameMonthPreviousYear(selectedMonth);
    const priorYearTotal = monthTotalFromData(monthlyData, priorMonth);
    const targetYear = String(selectedMonth || "").slice(0, 4);
    const sameYearMonths = [...monthlyData.keys()].filter((month) => month.slice(0, 4) === targetYear && month < selectedMonth);
    // Hueco honesto: hay meses del año en curso en cero y el MISMO mes del
    // año anterior sí vendió. Nunca se usa el mes pronosticado.
    if (forecastOptions.allowYearOverYear === false && !sameYearMonths.length) {
      return legacy;
    }
    let sameMonth = candidates.get("Mismo mes año anterior");
    if ((!sameMonth || !(sameMonth.total > 10) || !(sameMonth.sourceMonths || []).includes(priorMonth)) && priorYearTotal > 40) {
      const priorYearCandidates = buildForecastCandidates(monthlyData, selectedMonth, {
        ...forecastOptions,
        allowYearOverYear: true,
      });
      sameMonth = priorYearCandidates.get("Mismo mes año anterior");
    }
    const usedSamePriorMonth = Boolean(
      priorMonth
      && priorMonth < selectedMonth
      && priorYearTotal > 40
      && sameMonth
      && sameMonth.total > 10
      && (sameMonth.sourceMonths || []).includes(priorMonth)
    );
    if (usedSamePriorMonth) {
      return {
        ...legacy,
        averages: sameMonth.averages,
        method: "Reactivación estacional",
        recentMonths: [...new Set([...legacy.recentMonths, ...sameMonth.sourceMonths])],
      };
    }
    return legacy;
  }
  if (!seasonal) return legacy;
  if (Math.abs(seasonal.total - recent.total) / recent.total < 0.08) {
    return legacy;
  }
  let weight = seasonalWeight;
  if (recent.total > seasonal.total) {
    const lift = (recent.total - seasonal.total) / recent.total;
    weight = clamp(seasonalWeight - lift, Math.min(0.25, seasonalWeight), seasonalWeight);
  }
  return {
    ...legacy,
    averages: blendForecastAverages(seasonal.averages, legacy.averages, weight),
    method: `Resguardo estacional ${Math.round(weight * 100)}%`,
    recentMonths: [...new Set([...legacy.recentMonths, ...seasonal.sourceMonths])],
  };
}

// El peso estacional fijo falla cuando la forma del año anterior contradice la tendencia
// reciente. Aqui se elige por producto probando varios pesos contra los ultimos meses,
// usando solo informacion anterior a cada mes de validacion.
function calculateForecastModelSeasonalAdaptive(records, selectedMonth, candidateWeights = [0, 0.25, 0.5, 0.75, 1], forecastOptions = {}) {
  const monthlyData = buildMonthlyForecastData(records);
  const targetYear = String(selectedMonth || "").slice(0, 4);
  const historicalMonths = [...monthlyData.keys()].filter((month) => {
    if (!(month < selectedMonth)) return false;
    return forecastOptions.allowYearOverYear !== false || month.slice(0, 4) === targetYear;
  }).sort();
  const validationMonths = historicalMonths.slice(-3);
  const fallbackWeight = ["Otros", "Mini medianos"].includes(productCategory(records[0]?.producto || "")) ? 0.5 : 0.75;
  if (validationMonths.length < 2) return calculateForecastModelSeasonal(records, selectedMonth, fallbackWeight, forecastOptions);

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
      const model = calculateForecastModelSeasonal(priorRecords, validationMonth, candidateWeight, forecastOptions);
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

  const fallbackScore = scores.get(fallbackWeight)?.scale
    ? scores.get(fallbackWeight).error / scores.get(fallbackWeight).scale
    : Infinity;
  if (!(bestScore < fallbackScore * 0.96)) bestWeight = fallbackWeight;

  const model = calculateForecastModelSeasonal(records, selectedMonth, bestWeight, forecastOptions);
  return { ...model, method: `Estacional adaptativo ${Math.round(bestWeight * 100)}%` };
}

function calculateForecastModelForVersion(records, selectedMonth, product, modelVersion) {
  const version = String(modelVersion || FORECAST_MODEL_VERSION);
  if (version === "seasonalAdaptive") return calculateForecastModelSeasonalAdaptive(records, selectedMonth);
  if (version === "categorySeasonal" || version.startsWith("csG")) {
    const options = {
      ...forecastTuningOptions(version),
      allowYearOverYear: !forecastHidesPriorYearMonths(records, selectedMonth),
    };
    const category = productCategory(product);
    const seasonalWeight = ["Otros", "Mini medianos"].includes(category) ? 0.5 : 0.75;
    if (category === "Pasteles grandes" || category === "Pasteles medianos" || category === "Pasteles chicos" || category === "Mini medianos") {
      const monthlyData = buildMonthlyForecastData(records);
      const latest = previousMonthKey(selectedMonth);
      const latestData = latest ? monthlyData.get(latest) : null;
      const hasRecentDaily = Boolean(
        latestData &&
        !latestData.filledFromMonthlyTotal &&
        (latestData.valuesByDate?.size || 0) >= 20 &&
        latest &&
        !String(latest).endsWith("-05")
      );
      const seasonalFade = options.allowYearOverYear !== false
        && priorYearLooksLikeSeasonalFade(monthlyData, selectedMonth);
      // Impulso por SKU si ESE producto aceleró en la 2ª quincena. No es un
      // empuje de categoría: FRUTAS MED sin diario no se mueve. Gelatinas
      // quedan fuera: en jun–ago el diario aceleró y julio quedó plano.
      const allowMomentum = (
        category === "Pasteles grandes"
        || category === "Pasteles medianos"
        || category === "Mini medianos"
      ) && hasRecentDaily && !seasonalFade;
      const cakeOptions = {
        ...options,
        dipRatio: 0.92,
        dampenDecline: category === "Pasteles grandes" && allowMomentum ? 0.4 : undefined,
        momentumStrength: 0.4,
        momentumCap: 1.12,
        momentumTrigger: 1.1,
      };
      const model = category === "Pasteles grandes"
        ? calculateForecastModelSeasonalAdaptive(
            records,
            selectedMonth,
            [0.25, 0.5, 0.75],
            cakeOptions
          )
        : calculateForecastModelSeasonal(records, selectedMonth, seasonalWeight, cakeOptions);
      const lifted = applyColdStartEventUplift(
        liftColdStartMadresCalibration(model, records, selectedMonth, product),
        records,
        selectedMonth,
        product
      );
      return allowMomentum
        ? applyRecentMomentum(lifted, records, selectedMonth, cakeOptions)
        : lifted;
    }
    return applyColdStartEventUplift(
      liftColdStartMadresCalibration(
        calculateForecastModelSeasonal(records, selectedMonth, seasonalWeight, options),
        records,
        selectedMonth,
        product
      ),
      records,
      selectedMonth,
      product
    );
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

function calculateDailyForecast({ monthlyRows, ventasReales, realProduction, selectedMonth, dailyBufferPct, activePromos = [], dailyBranchStock = [], dailyColdRoom = [] }) {
  const realDailyMap = aggregateDailyProductionRows(realProduction);
  const salesDailyMap = aggregateDailySalesRows(ventasReales);
  const stockByProductDate = sumDailyBranchStockByProductDate(dailyBranchStock);
  const coldByProductDate = mapDailyColdRoomByProductDate(dailyColdRoom);
  const monthDates = datesForMonth(selectedMonth);
  const monthKeySet = new Set(monthDates.map((date) => dateKey(date)));
  const productRows = monthlyRows.filter((row) => isValidProduct(row.producto) && !isSliceProduct(row.producto));

  return productRows.flatMap((productRow) => {
    const product = productRow.producto;
    const demandByDate = monthDates.map((date) => {
      const pronosticoVentaDia = getWeekdayAverage(productRow, weekdayLabel(date.getDay()));
      return {
        date,
        key: dateKey(date),
        weekday: date.getDay(),
        pronosticoVentaDia,
        colchonDiario: pronosticoVentaDia * (dailyBufferPct / 100),
        baseConColchonDia: pronosticoVentaDia * (1 + dailyBufferPct / 100),
      };
    });
    const productionWeights = demandByDate.map((row) => {
      if (row.weekday === 0) return 0;
      const sundayDemand = demandByDate
        .filter((candidate) => candidate.weekday === 0 && productionDateKeyForDemand(candidate.date, monthKeySet) === row.key)
        .reduce((sum, candidate) => sum + candidate.baseConColchonDia, 0);
      return row.baseConColchonDia + sundayDemand;
    });
    const allocated = allocateDailyProduction(product, productionWeights, productRow.produccionSugerida);

    return demandByDate.map((row, index) => {
      const promo = findActivePromoForProduct(activePromos, product, row.date);
      const produccionBrutaDia = promo && row.weekday !== 0
        ? applyPromoUpliftToQuantity(product, allocated[index], promo)
        : allocated[index];
      const stockKey = `${row.key}|${normalizeProduct(product)}`;
      const hasDailyBranchStock = stockByProductDate.has(stockKey);
      const inventarioSucursalesDia = hasDailyBranchStock ? stockByProductDate.get(stockKey) : 0;
      const hasDailyColdRoom = coldByProductDate.has(stockKey);
      const cuartoFrioDia = hasDailyColdRoom ? coldByProductDate.get(stockKey) : 0;
      const produccionSugeridaDia = hasDailyBranchStock
        ? applyDailyBranchStockToPlantSuggestion(produccionBrutaDia, inventarioSucursalesDia)
        : produccionBrutaDia;
      const aProducirDia = applyInventoryToProductionSuggestion(
        produccionBrutaDia,
        hasDailyBranchStock ? inventarioSucursalesDia : 0,
        hasDailyColdRoom ? cuartoFrioDia : 0
      );
      const realKey = `${product}|${row.key}`;
      const hasRealData = realDailyMap.has(realKey);
      const produccionRealDia = hasRealData ? realDailyMap.get(realKey) : null;
      const diferenciaPiezas = hasRealData ? produccionRealDia - produccionSugeridaDia : null;
      const hasVentaReal = salesDailyMap.has(realKey);
      const ventaRealDia = hasVentaReal ? salesDailyMap.get(realKey) : null;
      const diferenciaVenta = hasVentaReal ? ventaRealDia - row.pronosticoVentaDia : null;
      const precisionVenta = precisionScore(row.pronosticoVentaDia, ventaRealDia);
      const productionTarget = productionDateKeyForDemand(row.date, monthKeySet);
      const sundayMoved = row.weekday === 0 && productionTarget !== row.key;
      const receivedSunday = row.weekday === 6 && productionWeights[index] > row.baseConColchonDia + 0.01;

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

      let reglaOperativa = getReglaOperativaLabel(product, productionWeights[index]);
      if (row.weekday === 0) reglaOperativa = "Domingo: no producir; demanda al sábado";
      else if (receivedSunday) reglaOperativa = `${reglaOperativa} · incluye demanda del domingo`;
      if (hasDailyBranchStock && row.weekday !== 0) {
        reglaOperativa = `${reglaOperativa} · menos ${formatNumber(inventarioSucursalesDia, 0)} en sucursales`;
      }
      if (hasDailyColdRoom && row.weekday !== 0) {
        reglaOperativa = `${reglaOperativa} · menos ${formatNumber(cuartoFrioDia, 0)} en cuarto frío`;
      }

      return {
        fecha: row.key,
        fechaDisplay: displayDate(row.date),
        weekday: row.weekday,
        dia: weekdayLabel(row.weekday),
        producto: product,
        promedioUsado: row.pronosticoVentaDia,
        pronosticoVentaDia: row.pronosticoVentaDia,
        colchonDiario: row.colchonDiario,
        baseConColchonDia: row.baseConColchonDia,
        reglaOperativa,
        produccionBrutaDia,
        inventarioSucursalesDia,
        hasDailyBranchStock,
        cuartoFrioDia,
        hasDailyColdRoom,
        produccionSugeridaDia,
        aProducirDia,
        promoActiva: Boolean(promo && row.weekday !== 0),
        promoEtiqueta: promo && row.weekday !== 0 ? formatPromoUpliftLabel(promo) : "",
        produccionDestino: sundayMoved ? productionTarget : row.key,
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
  const aProducirMensual = rows.reduce((sum, row) => sum + (row.aProducirDia ?? row.produccionSugeridaDia), 0);
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
    aProducirMensual,
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

function buildValidationAlerts(productSummary, homologationRows, historicalVentas, selectedMonth, stockSheetNotice = null) {
  const alerts = [];

  if (stockSheetNotice?.missingTotal) {
    alerts.push({
      tipo: "Catálogo",
      producto: stockSheetNotice.chosenSheet,
      detalle: stockSheetNotice.message,
      severidad: "Alta",
    });
  }

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

function withForecastDisplayDefaults(row) {
  const pronosticoVenta = toNumber(row.pronosticoVenta ?? row.pronosticoBase);
  return {
    promedioHistorico: 0,
    promedioDiario: 0,
    registrosHistoricos: 0,
    promedioLunes: 0,
    promedioMartes: 0,
    promedioMiercoles: 0,
    promedioJueves: 0,
    promedioViernes: 0,
    promedioSabado: 0,
    promedioDomingo: 0,
    tasaBajas: 0,
    bajasEsperadas: 0,
    colchonOperativo: 0,
    baseConColchon: pronosticoVenta,
    reglaOperativa: "",
    produccionSugerida: 0,
    inventarioObjetivo: 0,
    produccionBalanceada: 0,
    totalSuc: 0,
    cf: 0,
    sumaSucCf: 0,
    produccionRecomendada: 0,
    tendenciaAplicada: row.tendenciaAplicada,
    mesesUsados: row.mesesUsados || "",
    metodoPronostico: row.metodoPronostico || "",
    catalogCleanup: "",
    promoActiva: false,
    promoEtiqueta: "",
    produccionReal: 0,
    hasRealData: false,
    diferenciaReal: 0,
    precision: null,
    confianza: 0,
    estatus: "Sin dato real",
    ...row,
    pronosticoVenta,
    demandaPronosticada: toNumber(row.demandaPronosticada ?? pronosticoVenta),
  };
}

function snapshotForecastRowsForFreeze(forecastRows) {
  return (forecastRows || []).map((row) => ({
    producto: row.producto,
    orden: row.orden,
    promedioHistorico: toNumber(row.promedioHistorico),
    promedioDiario: toNumber(row.promedioDiario),
    registrosHistoricos: toNumber(row.registrosHistoricos),
    promedioLunes: toNumber(row.promedioLunes),
    promedioMartes: toNumber(row.promedioMartes),
    promedioMiercoles: toNumber(row.promedioMiercoles),
    promedioJueves: toNumber(row.promedioJueves),
    promedioViernes: toNumber(row.promedioViernes),
    promedioSabado: toNumber(row.promedioSabado),
    promedioDomingo: toNumber(row.promedioDomingo),
    demandaPronosticada: toNumber(row.demandaPronosticada ?? row.pronosticoVenta),
    pronosticoVenta: toNumber(row.pronosticoVenta),
    tasaBajas: toNumber(row.tasaBajas),
    bajasEsperadas: toNumber(row.bajasEsperadas),
    colchonOperativo: toNumber(row.colchonOperativo),
    baseConColchon: toNumber(row.baseConColchon),
    reglaOperativa: row.reglaOperativa || "",
    produccionSugerida: toNumber(row.produccionSugerida),
    inventarioObjetivo: toNumber(row.inventarioObjetivo),
    tendenciaAplicada: row.tendenciaAplicada,
    mesesUsados: row.mesesUsados,
    metodoPronostico: row.metodoPronostico,
    catalogCleanup: row.catalogCleanup || "",
    promoActiva: Boolean(row.promoActiva),
    promoEtiqueta: row.promoEtiqueta || "",
  }));
}

function hydrateForecastFromOperationalRows(operationalRows, selectedMonth) {
  const days = datesForMonth(selectedMonth).length || 1;
  return (operationalRows || []).map((row) => {
    const pronosticoVenta = toNumber(row.pronosticoBase ?? row.pronosticoVenta);
    const daily = days > 0 ? pronosticoVenta / days : 0;
    return withForecastDisplayDefaults({
      producto: row.producto,
      orden: row.orden,
      promedioLunes: daily,
      promedioMartes: daily,
      promedioMiercoles: daily,
      promedioJueves: daily,
      promedioViernes: daily,
      promedioSabado: daily,
      promedioDomingo: daily,
      promedioDiario: daily,
      demandaPronosticada: pronosticoVenta,
      pronosticoVenta,
      baseConColchon: pronosticoVenta,
      produccionSugerida: toNumber(row.produccionSugerida) || getProduccionSugerida(row.producto, pronosticoVenta * (1 + 10 / 100)),
      metodoPronostico: row.metodoPronostico,
      tendenciaAplicada: row.tendenciaAplicada,
      mesesUsados: row.mesesUsados,
    });
  });
}

function resolveEffectiveForecast({ liveForecast = [], frozenSnapshot = null, selectedMonth = "" } = {}) {
  const period = frozenSnapshot?.periodo || frozenSnapshot?.contenido?.selectedMonth || "";
  const version = frozenSnapshot?.version || null;
  if (!frozenSnapshot || !selectedMonth || period !== selectedMonth) {
    return {
      rows: liveForecast,
      source: "live",
      frozenVersion: null,
      label: "pronóstico vigente",
    };
  }
  const content = frozenSnapshot.contenido || {};
  if (Array.isArray(content.forecastRows) && content.forecastRows.length) {
    return {
      rows: content.forecastRows.map((row) => withForecastDisplayDefaults(row)),
      source: "frozen",
      frozenVersion: version,
      label: version ? `pronóstico congelado v${version}` : "pronóstico congelado",
    };
  }
  if (Array.isArray(content.rows) && content.rows.length) {
    return {
      rows: hydrateForecastFromOperationalRows(content.rows, selectedMonth),
      source: "frozen-operational",
      frozenVersion: version,
      label: version ? `pronóstico congelado v${version}` : "pronóstico congelado",
    };
  }
  return {
    rows: liveForecast,
    source: "live",
    frozenVersion: version,
    label: "pronóstico vigente",
  };
}

const MONTHLY_REVIEW_STATUSES = ["ACTIVO", "BAJA", "BAJO PEDIDO", "ESTACIONAL"];

function countCapturedProductStatuses(inputs) {
  return Object.values(inputs || {}).filter((input) => MONTHLY_REVIEW_STATUSES.includes(String(input?.status || ""))).length;
}

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
  inventoryUsable = false,
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
    const capturedInventory = input.inventoryOverride === null || input.inventoryOverride === undefined || input.inventoryOverride === ""
      ? (hasLoadedInventory ? existenceByProduct.get(product) : 0)
      : Math.max(0, toNumber(input.inventoryOverride));
    const deductedInventory = inventoryUsable ? capturedInventory : 0;
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
      proposed = Math.max(0, getProduccionSugerida(product, baseForecast * (1 + marginPct / 100) - deductedInventory));
      if (!inventoryUsable && (hasLoadedInventory || capturedInventory > 0)) {
        reasons.push("Existencias fuera de ventana o sin fecha de corte: no se descontaron.");
        if (severity === "ok") severity = "warn";
      } else if (hasLoadedInventory || input.inventoryOverride !== null && input.inventoryOverride !== undefined && input.inventoryOverride !== "") {
        reasons.push(`Se descontaron ${formatNumber(deductedInventory, 0)} piezas de existencia.`);
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
      inventory: capturedInventory,
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
    { Indicador: "Fecha de corte existencias", Valor: review.inventoryCutoff || "Sin fecha" },
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
    { Regla: "Existencias", Detalle: "Solo se descuentan si la fecha de corte cae entre el mes anterior y el mes planificado." },
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
    { Indicador: "A producir mensual", Valor: summary.aProducirMensual ?? summary.produccionSugeridaMensual },
    { Indicador: "Regla domingo", Valor: "Sin produccion; la demanda se cubre el sabado. La suma diaria iguala el total mensual." },
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
    "Produccion bruta dia": row.produccionBrutaDia ?? row.produccionSugeridaDia,
    "Inventario sucursales": row.hasDailyBranchStock ? row.inventarioSucursalesDia : "",
    "Pedido planta": row.produccionSugeridaDia,
    "Cuarto frio": row.hasDailyColdRoom ? row.cuartoFrioDia : "",
    "A producir": row.aProducirDia ?? row.produccionSugeridaDia,
    "Produccion sugerida dia": row.produccionSugeridaDia,
    "Promo activa": row.promoActiva ? row.promoEtiqueta || "Sí" : "",
    "Produccion destino": row.produccionDestino || row.fecha,
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

function SectionDisclosure({
  className = "",
  summaryClassName = "compact-analytics-summary",
  contentClassName = "compact-analytics-content",
  eyebrow,
  title,
  description,
  badge,
  icon: Icon,
  children,
  open,
  onToggle,
  defaultOpen = false,
}) {
  const detailsProps = typeof onToggle === "function"
    ? {
      open,
      onToggle: (event) => onToggle(event.currentTarget.open),
    }
    : { defaultOpen };
  return (
    <details className={className} {...detailsProps}>
      <summary className={summaryClassName}>
        {Icon ? <span className="compact-analytics-icon"><Icon size={19} /></span> : null}
        <div>
          {eyebrow ? <span className="eyebrow">{eyebrow}</span> : null}
          <strong>{title}</strong>
          {description ? <small>{description}</small> : null}
        </div>
        {badge}
      </summary>
      <div className={contentClassName}>{children}</div>
    </details>
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
  resumeFromBatch = null,
}) {
  if (!pendingRows.length && !status) return null;
  const statusIsError = status.startsWith("No ") || status.includes("Falló");
  const buttonLabel = importing
    ? "Guardando..."
    : resumeFromBatch
      ? `Reanudar desde lote ${resumeFromBatch}`
      : "Guardar en la base";
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
            <Database size={18} /> {buttonLabel}
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
      <p className={`sales-import-status ${statusIsError ? "error" : ""}`}>{status}</p>
      {preview.monthlyTotals > 0 && pendingRows.length > 0 && (
        <p className="sales-import-note">Los totales mensuales no se convierten en registros diarios.</p>
      )}
    </section>
  );
}

function FreezeReadinessStrip({ readiness, selectedMonth }) {
  if (!selectedMonth) return null;
  const previousShort = shortMonthLabel(readiness.previousMonthKey);
  const targetShort = shortMonthLabel(selectedMonth);
  const syncOk = !readiness.blockers.some((item) => item.code === "sync-pending" || item.code === "sync-incomplete");
  const targetOk = !readiness.blockers.some((item) => item.code === "target-has-sales");
  const chips = [
    { key: "sync", label: "Sync", value: syncOk ? "OK" : "No", tone: syncOk ? "ok" : "blocked" },
    {
      key: "close",
      label: `Cierre ${previousShort}`,
      value: readiness.previousMonth.hasClose ? "OK" : "No",
      tone: readiness.previousMonth.hasClose ? "ok" : "blocked",
    },
    {
      key: "daily",
      label: `Diario ${previousShort}`,
      value: readiness.previousMonth.hasDaily
        ? "OK"
        : readiness.previousMonth.hasClose
          ? "Cierre"
          : `${readiness.previousMonth.dailyDays}/${readiness.previousMonth.daysInMonth || 0}`,
      tone: readiness.previousMonth.hasDaily
        ? "ok"
        : readiness.previousMonth.hasClose
          ? "warn"
          : "blocked",
    },
    {
      key: "target",
      label: `${targetShort} sin ventas`,
      value: targetOk ? "OK" : "Hay ventas",
      tone: targetOk ? "ok" : "blocked",
    },
  ];
  const alert = readiness.blockers[0] || (readiness.canFreeze ? readiness.warnings[0] : null);
  return (
    <section className={`freeze-strip ${readiness.canFreeze ? (readiness.warnings.length ? "warning" : "success") : "warning"}`}>
      <div className="freeze-strip-stats">
        {chips.map((chip) => (
          <div key={chip.key} className={`freeze-strip-stat ${chip.tone}`}>
            <strong>{chip.value}</strong>
            <span>{chip.label}</span>
          </div>
        ))}
      </div>
      {alert && (
        <p className={`freeze-strip-alert ${readiness.canFreeze ? "warn" : "blocked"}`}>{alert.message}</p>
      )}
    </section>
  );
}

function FrozenMonthBanner({ forecastLock, selectedMonth }) {
  if (!selectedMonth || !forecastLock || forecastLock.source === "live") return null;
  return (
    <p className="freeze-strip-alert ok month-frozen-banner">
      {displayMonthLabel(selectedMonth)} está congelado{forecastLock.frozenVersion ? ` (v${forecastLock.frozenVersion})` : ""}.
      El pronóstico y las sugerencias de planta de este mes no cambian si cargas más datos.
    </p>
  );
}

function ForecastAccuracyPanel({ health, selectedMonth }) {
  const backtests = health?.backtests || [];
  const [focusMonth, setFocusMonth] = useState("");
  const selected = backtests.find((row) => row.month === focusMonth) || health?.latestBacktest || null;
  const weighted = weightedWapeFromBacktests(backtests);
  const windowLabel = describeAccuracyWindow(backtests);
  const weightedTone = forecastAccuracyTone(weighted);
  if (!selectedMonth) return null;

  return (
    <section className="forecast-accuracy-panel">
      <div className="forecast-accuracy-heading">
        <div>
          <span className="eyebrow">Control del pronóstico</span>
          <strong>WAPE de meses cerrados</strong>
          <p>Misma cuenta que las pruebas de exactitud: oculta el mes, pronostica y compara contra la venta cargada.</p>
        </div>
        <div className={`forecast-accuracy-weighted ${weightedTone}`}>
          <strong>{weighted == null ? "Sin cierre" : `${weighted.toFixed(1)}%`}</strong>
          <span>{windowLabel ? `WAPE ponderado ${windowLabel}` : "WAPE ponderado"}</span>
        </div>
      </div>
      {backtests.length ? (
        <>
          <div className="forecast-accuracy-months" role="tablist" aria-label="Mes cerrado">
            {backtests.map((row) => {
              const active = selected?.month === row.month;
              return (
                <button
                  key={row.month}
                  type="button"
                  role="tab"
                  aria-selected={active}
                  className={`forecast-accuracy-month ${forecastAccuracyTone(row.wape)}${active ? " active" : ""}`}
                  onClick={() => setFocusMonth(row.month)}
                >
                  <strong>{row.wape == null ? "—" : `${row.wape.toFixed(1)}%`}</strong>
                  <span>{shortMonthLabel(row.month)}</span>
                </button>
              );
            })}
          </div>
          <div className="forecast-accuracy-table-wrap">
            <table className="forecast-accuracy-table">
              <thead>
                <tr>
                  <th>Mes</th>
                  <th>Venta real</th>
                  <th>Pronóstico</th>
                  <th>WAPE</th>
                  <th>MAE</th>
                  <th>±15</th>
                  <th>Peor producto</th>
                </tr>
              </thead>
              <tbody>
                {backtests.map((row) => {
                  const worst = row.topErrors?.[0];
                  const active = selected?.month === row.month;
                  return (
                    <tr key={row.month} className={active ? "is-selected" : undefined}>
                      <td>{displayMonthLabel(row.month)}</td>
                      <td>{formatNumber(row.actual, 0)}</td>
                      <td>{formatNumber(row.forecast, 0)}</td>
                      <td>{row.wape == null ? "—" : formatPercent(row.wape, 1)}</td>
                      <td>{row.mae == null ? "—" : formatNumber(row.mae, 1)}</td>
                      <td>{`${row.inside15 || 0}/${row.products || 0}`}</td>
                      <td>{worst ? `${worst.producto} (${worst.error > 0 ? "+" : ""}${Math.round(worst.error)})` : "—"}</td>
                    </tr>
                  );
                })}
              </tbody>
            </table>
          </div>
        </>
      ) : (
        <p className="forecast-accuracy-empty">Carga ventas de meses anteriores para ver el WAPE automático. No hace falta correr scripts.</p>
      )}
    </section>
  );
}

function ForecastHealthStrip({ health, selectedMonth }) {
  if (!selectedMonth) return null;
  const latest = health.latestBacktest;
  const chips = [
    {
      key: "stock",
      label: "Stock / catálogo",
      value: health.catalogCount ? String(health.catalogCount) : "No",
      tone: health.catalogCount ? "ok" : "blocked",
    },
    {
      key: "months",
      label: "Meses de venta",
      value: health.months.length ? String(health.months.length) : "No",
      tone: health.months.length ? "ok" : "blocked",
    },
    {
      key: "wape",
      label: latest ? `WAPE ${shortMonthLabel(latest.month)}` : "WAPE de control",
      value: latest?.wape == null ? "Sin cierre" : `${latest.wape.toFixed(1)}%`,
      tone: health.wapeTone === "muted" ? (health.months.length ? "warn" : "blocked") : health.wapeTone,
    },
    {
      key: "forecast",
      label: `Pronóstico ${shortMonthLabel(selectedMonth)}`,
      value: health.currentTotal > 0 ? formatNumber(health.currentTotal, 0) : "Cero",
      tone: health.currentTotal > 0 ? "ok" : "blocked",
    },
  ];
  const alert =
    health.checks.find((item) => item.level === "error") ||
    health.checks.find((item) => item.level === "warning") ||
    health.checks.find((item) => item.level === "ok");
  const topErrors = (latest?.topErrors || []).slice(0, 5);
  return (
    <section className={`freeze-strip forecast-health-strip ${health.healthy ? "success" : "warning"}`}>
      <div className="freeze-strip-stats">
        {chips.map((chip) => (
          <div key={chip.key} className={`freeze-strip-stat ${chip.tone}`}>
            <strong>{chip.value}</strong>
            <span>{chip.label}</span>
          </div>
        ))}
      </div>
      {alert && (
        <p className={`freeze-strip-alert ${alert.level === "error" ? "blocked" : alert.level === "ok" ? "ok" : "warn"}`}>
          {alert.message}
        </p>
      )}
      {topErrors.length > 0 && (
        <details className="forecast-health-more">
          <summary>Ver productos con más error</summary>
          <p className="freeze-strip-alert warn forecast-health-errors">
            Más error absoluto en {shortMonthLabel(latest.month)}:{" "}
            {topErrors.map((row) => `${row.producto} (${row.error > 0 ? "+" : ""}${Math.round(row.error)})`).join(" · ")}
          </p>
        </details>
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
              minLength={needsSetup ? 4 : undefined}
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
  const [stockSheetNotice, setStockSheetNotice] = useState(null);
  const [ventas, setVentas] = useState([]);
  const [ventasValidacion, setVentasValidacion] = useState([]);
  const [bajas, setBajas] = useState([]);
  const [existencias, setExistencias] = useState([]);
  const [existenciasCutoff, setExistenciasCutoff] = useState({ date: "", source: "", fileName: "" });
  const [dailyBranchStock, setDailyBranchStock] = useState([]);
  const [dailyColdRoom, setDailyColdRoom] = useState([]);
  const [stockCaptureDate, setStockCaptureDate] = useState(() => defaultInventoryDate(defaultMonthValue()));
  const [stockCaptureQuery, setStockCaptureQuery] = useState("");
  const [newSucursalName, setNewSucursalName] = useState("");
  const [manualSucursales, setManualSucursales] = useState([]);
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
  const [dailyDateFilter, setDailyDateFilter] = useState(() => defaultInventoryDate(defaultMonthValue()));
  const [dailyProductQuery, setDailyProductQuery] = useState("");
  const [dailyWeekdayFilter, setDailyWeekdayFilter] = useState("");
  const [selectedWeekKey, setSelectedWeekKey] = useState("");
  const [onlyDailyShortage, setOnlyDailyShortage] = useState(false);
  const [onlyDailyOverproduction, setOnlyDailyOverproduction] = useState(false);
  const [onlySalesMismatch, setOnlySalesMismatch] = useState(false);
  const [validationProduct, setValidationProduct] = useState("");
  const [productAliases, setProductAliases] = useState(loadStoredProductAliases);
  const [activePromos, setActivePromos] = useState(loadStoredActivePromos);
  const [promoForm, setPromoForm] = useState(() => emptyPromoForm());
  const [promoFormError, setPromoFormError] = useState("");
  const [promoOpen, setPromoOpen] = useState(() => loadStoredActivePromos().some((promo) => isPromoListedAsActive(promo)));
  const [showPromoForm, setShowPromoForm] = useState(false);
  const [cloudStatus, setCloudStatus] = useState("Buscando respaldo...");
  const [databaseSync, setDatabaseSync] = useState(null);
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
  const [salesImportRun, setSalesImportRun] = useState(null);
  const [salesExcelSources, setSalesExcelSources] = useState({});
  const [salesMonthOverrides, setSalesMonthOverrides] = useState({});
  const [salesSourceDecisions, setSalesSourceDecisions] = useState([]);
  const [pendingProductionImport, setPendingProductionImport] = useState([]);
  const [productionImportFile, setProductionImportFile] = useState("");
  const [productionImporting, setProductionImporting] = useState(false);
  const [productionImportStatus, setProductionImportStatus] = useState("");
  const [productionImportRun, setProductionImportRun] = useState(null);
  const [pendingWasteImport, setPendingWasteImport] = useState([]);
  const [wasteImportFile, setWasteImportFile] = useState("");
  const [wasteImporting, setWasteImporting] = useState(false);
  const [wasteImportStatus, setWasteImportStatus] = useState("");
  const [wasteImportRun, setWasteImportRun] = useState(null);
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
        fetchOperationalSync("/api/ventas/sync", session.token),
        fetchOperationalSync("/api/produccion-real/sync", session.token),
        fetchOperationalSync("/api/bajas/sync", session.token),
      ]);
      if ([salesSync, productionSync, wasteSync].some((result) => result.status === 401)) {
        onLogout();
        return;
      }
      if (!active) return;

      const remoteSales = (salesSync.ok ? salesSync.rows : []).map((row) => ({
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
      const remoteProduction = (productionSync.ok ? productionSync.rows : []).map((row) => ({
        fecha: row.fecha,
        fechaKey: row.fecha,
        producto: normalizeProduct(row.producto_codigo),
        productoOriginal: row.producto_nombre,
        cantidad: Number(row.cantidad),
        turno: row.turno || "",
        sourceFile: "Base de datos",
        databaseSynced: true,
      }));
      const remoteWaste = (wasteSync.ok ? wasteSync.rows : []).map((row) => ({
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
      setExistenciasCutoff(data.existenciasCutoff && typeof data.existenciasCutoff === "object"
        ? {
            date: dateKey(data.existenciasCutoff.date) || "",
            source: String(data.existenciasCutoff.source || ""),
            fileName: String(data.existenciasCutoff.fileName || ""),
          }
        : { date: "", source: "", fileName: "" });
      setDailyBranchStock(sanitizeDailyBranchStock(data.dailyBranchStock));
      setDailyColdRoom(sanitizeDailyColdRoom(data.dailyColdRoom));
      if (Array.isArray(data.manualSucursales)) {
        setManualSucursales(data.manualSucursales.map((name) => String(name || "").trim()).filter(Boolean));
      }
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
      if (Array.isArray(data.activePromos)) {
        setActivePromos(sanitizeActivePromos(data.activePromos));
      }
      if (data.selectedMonth) setSelectedMonth(data.selectedMonth);
      if (Number.isFinite(Number(data.dailyBufferPct))) setDailyBufferPct(Number(data.dailyBufferPct));
      setLastBackup(snapshot);
      setHasUnsavedChanges(false);
      const nextSync = {
        sales: { ok: salesSync.ok, count: remoteSales.length, error: salesSync.error },
        production: { ok: productionSync.ok, count: remoteProduction.length, error: productionSync.error },
        waste: { ok: wasteSync.ok, count: remoteWaste.length, error: wasteSync.error },
      };
      setDatabaseSync(nextSync);
      setCloudStatus(databaseSyncStatusText(snapshot, nextSync));
    }
    restoreData().catch((error) => {
      if (!active) return;
      if (error.status === 401) {
        onLogout();
        return;
      }
      setDatabaseSync({
        sales: { ok: false, count: 0, error: error.message },
        production: { ok: false, count: 0, error: error.message },
        waste: { ok: false, count: 0, error: error.message },
      });
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
        existenciasCutoff,
        dailyBranchStock: sanitizeDailyBranchStock(dailyBranchStock),
        dailyColdRoom: sanitizeDailyColdRoom(dailyColdRoom),
        manualSucursales,
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
        activePromos,
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
      setCloudStatus(
        isDatabaseSyncComplete(databaseSync)
          ? `Respaldo v${response.version} guardado correctamente`
          : `Respaldo v${response.version} guardado. Sincronización incompleta: no se puede congelar.`
      );
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

  async function handleStockFile(file) {
    if (!file) return;
    const workbook = await readWorkbook(file);
    setStockRows(parseStock(workbook));
    const notice = assessStockSheetSelection(workbook);
    setStockSheetNotice(notice.missingTotal ? notice : null);
    if (notice.missingTotal) setToast({ tone: "warning", message: notice.message });
    setFiles((f) => ({ ...f, stock: file.name }));
    setHasUnsavedChanges(true);
  }

  async function handleExistenciasFile(file) {
    if (!file) return;
    const workbook = await readWorkbook(file);
    const parsed = parseExistencias(workbook);
    const inferred = inferInventoryCutoffDate(workbook, file.name);
    setExistencias(parsed);
    setExistenciasCutoff({
      date: inferred.date || "",
      source: inferred.source || "",
      fileName: file.name,
    });
    setFiles((current) => ({ ...current, existencias: file.name }));
    setHasUnsavedChanges(true);
  }

  async function handleDailyBranchStockFile(file) {
    if (!file) return;
    const workbook = await readWorkbook(file);
    const parsed = parseDailyInventory(workbook, stockCaptureDate);
    if (!parsed.branchStock.length && !parsed.coldRoom.length) {
      setToast({ tone: "warning", message: "No se reconocieron filas de inventario. Usa sucursales y/o Cuarto frío (RAIZ: Total suc + Restante CF). TOTAL A TENER y sumas no entran como sucursal." });
      return;
    }
    if (parsed.branchStock.length) {
      setDailyBranchStock((current) => upsertDailyBranchStock(current, parsed.branchStock));
      const importedBranches = [...new Set(parsed.branchStock.map((row) => row.sucursal).filter(Boolean))];
      setManualSucursales((current) => collectSucursales({ extra: [...current, ...importedBranches] }));
    }
    if (parsed.coldRoom.length) {
      setDailyColdRoom((current) => upsertDailyColdRoom(current, parsed.coldRoom));
    }
    setFiles((current) => ({ ...current, dailyBranchStock: file.name }));
    setHasUnsavedChanges(true);
    const dates = [...new Set([
      ...parsed.branchStock.map((row) => row.fecha),
      ...parsed.coldRoom.map((row) => row.fecha),
    ])].sort();
    if (dates.length === 1) {
      setStockCaptureDate(dates[0]);
      setDailyDateFilter(dates[0]);
    }
    const dateLabel = dates[0] === dates.at(-1)
      ? displayDate(dates[0])
      : `${displayDate(dates[0])} a ${displayDate(dates.at(-1))}`;
    const parts = [];
    if (parsed.branchStock.length) parts.push(`${parsed.branchStock.length} sucursales`);
    if (parsed.coldRoom.length) parts.push(`${parsed.coldRoom.length} cuarto frío`);
    setToast({
      tone: "success",
      message: `Inventario diario: ${parts.join(" · ")} (${dateLabel}).`,
    });
  }

  function updateDailyBranchStockCell(fecha, sucursal, producto, rawValue) {
    const key = dailyBranchStockKey({ fecha, sucursal, producto });
    if (!fecha || !sucursal || !producto) return;
    if (rawValue === "" || rawValue === null || rawValue === undefined) {
      setDailyBranchStock((current) => removeDailyBranchStockKey(current, key));
      setHasUnsavedChanges(true);
      return;
    }
    setDailyBranchStock((current) => upsertDailyBranchStock(current, [{
      fecha,
      sucursal,
      producto,
      cantidad: rawValue,
    }]));
    setHasUnsavedChanges(true);
  }

  function updateDailyColdRoomCell(fecha, producto, rawValue) {
    const key = dailyColdRoomKey({ fecha, producto });
    if (!fecha || !producto) return;
    if (rawValue === "" || rawValue === null || rawValue === undefined) {
      setDailyColdRoom((current) => removeDailyColdRoomKey(current, key));
      setHasUnsavedChanges(true);
      return;
    }
    setDailyColdRoom((current) => upsertDailyColdRoom(current, [{
      fecha,
      producto,
      cantidad: rawValue,
    }]));
    setHasUnsavedChanges(true);
  }

  function addManualSucursal() {
    const name = newSucursalName.trim();
    if (!name) return;
    setManualSucursales((current) => collectSucursales({ extra: [...current, name] }));
    setNewSucursalName("");
    setHasUnsavedChanges(true);
  }

  function clearDailyInventoryForDate(fecha) {
    if (!fecha) return;
    setDailyBranchStock((current) => current.filter((row) => row.fecha !== fecha));
    setDailyColdRoom((current) => current.filter((row) => row.fecha !== fecha));
    setHasUnsavedChanges(true);
  }

  async function handleSalesFiles(selectedFiles) {
    const filesToRead = Array.isArray(selectedFiles) ? selectedFiles : [selectedFiles];
    const loadedNames = new Set(String(files.ventas || "").split(", ").filter(Boolean));
    const validFiles = filesToRead.filter(Boolean);
    if (!validFiles.length) return;

    const parsedFiles = await Promise.all(
      validFiles.map(async (file) => {
        const workbook = await readWorkbook(file);
        const rows = parseSalesOrReturns(workbook, "ventas", file.name);
        return rows.map((row) => ({ ...row, sourceFile: file.name }));
      })
    );
    const incomingByName = Object.fromEntries(validFiles.map((file, index) => [file.name, parsedFiles[index]]));
    const nextSources = { ...salesExcelSources };
    if (!Object.keys(nextSources).length) {
      for (const row of ventas) {
        if (isDatabaseSalesSource(row.sourceFile)) continue;
        if (!nextSources[row.sourceFile]) nextSources[row.sourceFile] = [];
        nextSources[row.sourceFile].push(row);
      }
    }
    Object.assign(nextSources, incomingByName);
    const resolved = resolveCanonicalMonthSources(
      Object.entries(nextSources).map(([name, rows]) => ({ name, rows })),
      salesMonthOverrides
    );
    const excelRows = resolved.entries.flatMap((entry) => entry.rows.map((row) => ({ ...row, sourceFile: entry.name })));
    const dbRows = ventas.filter((row) => isDatabaseSalesSource(row.sourceFile));
    const pendingNames = new Set([
      ...pendingSalesImport.map((row) => row.sourceFile).filter(Boolean),
      ...validFiles.map((file) => file.name),
    ]);
    setSalesExcelSources(nextSources);
    setSalesSourceDecisions(resolved.decisions);
    setVentas([...dbRows, ...excelRows]);
    setPendingSalesImport(excelRows.filter((row) => pendingNames.has(row.sourceFile)));
    setSalesImportRun(null);
    setSalesImportFiles((current) => [...new Set([...current, ...validFiles.map((file) => file.name)])]);
    const decisionText = resolved.decisions.map((decision) => describeSourceDecision(decision)).join(" ");
    const parsedCount = parsedFiles.flat().length;
    setSalesImportStatus(
      parsedCount
        ? `${decisionText}${decisionText ? " " : ""}Revisa el resumen antes de guardar en la base.`
        : "No se reconocieron ventas en los archivos."
    );
    setFiles((current) => ({
      ...current,
      ventas: [...new Set([...loadedNames, ...validFiles.map((file) => file.name)])].join(", "),
    }));
    setHasUnsavedChanges(true);
  }

  function chooseSalesMonthSource(month, winner) {
    const overrides = { ...salesMonthOverrides, [month]: winner };
    const resolved = resolveCanonicalMonthSources(
      Object.entries(salesExcelSources).map(([name, rows]) => ({ name, rows })),
      overrides
    );
    const excelRows = resolved.entries.flatMap((entry) => entry.rows.map((row) => ({ ...row, sourceFile: entry.name })));
    const dbRows = ventas.filter((row) => isDatabaseSalesSource(row.sourceFile));
    const pendingNames = new Set(pendingSalesImport.map((row) => row.sourceFile).filter(Boolean));
    setSalesMonthOverrides(overrides);
    setSalesSourceDecisions(resolved.decisions);
    setVentas([...dbRows, ...excelRows]);
    setPendingSalesImport(excelRows.filter((row) => pendingNames.has(row.sourceFile)));
    setHasUnsavedChanges(true);
  }

  async function importSalesToDatabase() {
    if (!pendingSalesImport.length) return;
    setSalesImporting(true);
    setSalesImportStatus("Validando y guardando ventas...");
    const previousRun = salesImportRun;
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
        importRunId: previousRun?.importRunId,
        startBatch: previousRun?.nextBatch || 1,
        onProgress: (batch, total) => setSalesImportStatus(`Guardando ventas: lote ${batch} de ${total}...`),
      });
      const inserted = Number(previousRun?.inserted || 0) + Number(response.inserted || 0);
      const updated = Number(previousRun?.updated || 0) + Number(response.updated || 0);
      const importedKeys = new Set(rows.map(salesRecordKey));
      setVentas((current) => current.map((row) => importedKeys.has(salesRecordKey(row)) ? { ...row, databaseSynced: true } : row));
      setPendingSalesImport([]);
      setSalesImportRun(null);
      setSalesImportStatus("");
      const backup = await saveWorkspace({ silent: true, persistedKeys: { sales: importedKeys } });
      setToast({
        tone: backup.ok ? "success" : "warning",
        message: backup.ok
          ? `Ventas guardadas: ${inserted} nuevas, ${updated} actualizadas. Respaldo v${backup.version} creado.`
          : `Ventas guardadas en la base de datos, pero el respaldo automático falló.`,
      });
    } catch (error) {
      if (error.status === 401) onLogout();
      else {
        setSalesImportRun(accumulateImportRun(previousRun, error.importProgress));
        setSalesImportStatus(error.message);
      }
    } finally {
      setSalesImporting(false);
    }
  }

  async function handleProductionReal(file) {
    if (!file) return;
    const parsed = (await parseProductionRealFile(file)).map((row) => ({ ...row, sourceFile: file.name }));
    setRealProduction(parsed);
    setPendingProductionImport(parsed);
    setProductionImportRun(null);
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
    setWasteImportRun(null);
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
    const setRun = isProduction ? setProductionImportRun : setWasteImportRun;
    const previousRun = isProduction ? productionImportRun : wasteImportRun;
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
        importRunId: previousRun?.importRunId,
        startBatch: previousRun?.nextBatch || 1,
        onProgress: (batch, total) => setStatus(`Guardando ${isProduction ? "producción" : "bajas"}: lote ${batch} de ${total}...`),
      });
      const inserted = Number(previousRun?.inserted || 0) + Number(response.inserted || 0);
      const updated = Number(previousRun?.updated || 0) + Number(response.updated || 0);
      const keyForRow = isProduction ? productionRecordKey : wasteRecordKey;
      const importedKeys = new Set(rows.map(keyForRow));
      if (isProduction) {
        setRealProduction((current) => current.map((row) => importedKeys.has(productionRecordKey(row)) ? { ...row, databaseSynced: true } : row));
      } else {
        setBajas((current) => current.map((row) => importedKeys.has(wasteRecordKey(row)) ? { ...row, databaseSynced: true } : row));
      }
      clearPending([]);
      setRun(null);
      setStatus("");
      const backup = await saveWorkspace({
        silent: true,
        persistedKeys: isProduction ? { production: importedKeys } : { waste: importedKeys },
      });
      setToast({
        tone: backup.ok ? "success" : "warning",
        message: backup.ok
          ? `${isProduction ? "Producción" : "Bajas"} guardadas: ${inserted} nuevas, ${updated} actualizadas. Respaldo v${backup.version} creado.`
          : `${isProduction ? "Producción" : "Bajas"} guardadas en la base de datos, pero el respaldo automático falló.`,
      });
    } catch (error) {
      if (error.status === 401) onLogout();
      else {
        setRun(accumulateImportRun(previousRun, error.importProgress));
        setStatus(error.message);
      }
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

  function resetPromoForm() {
    setPromoForm(emptyPromoForm());
    setPromoFormError("");
    setShowPromoForm(false);
  }

  function submitPromo(event) {
    event.preventDefault();
    const catalogName = findOfficialProduct(promoForm.producto, officialProducts) || normalizeProduct(promoForm.producto);
    if (!catalogName) {
      setPromoFormError("Elige un producto del catálogo o escribe un nombre reconocido.");
      return;
    }
    const multiplier = Number(promoForm.multiplier);
    const extraPiecesPerDay = Number(promoForm.extraPiecesPerDay);
    if (!Number.isFinite(multiplier) || multiplier <= 0) {
      setPromoFormError("El multiplicador debe ser mayor a 0. Usa 1 si solo quieres evitar que se apague el SKU.");
      return;
    }
    if (!Number.isFinite(extraPiecesPerDay) || extraPiecesPerDay < 0) {
      setPromoFormError("Las piezas extra por día no pueden ser negativas.");
      return;
    }
    const startDate = dateKey(promoForm.startDate) || todayKey();
    const nextPromo = normalizeActivePromo({
      id: promoForm.id || createPromoId(),
      producto: catalogName,
      startDate,
      durationPreset: promoForm.durationPreset,
      multiplier,
      extraPiecesPerDay,
      note: promoForm.note,
      active: true,
      createdAt: activePromos.find((promo) => promo.id === promoForm.id)?.createdAt,
      updatedAt: new Date().toISOString(),
    });
    if (!nextPromo) {
      setPromoFormError("No se pudo registrar la promo. Revisa el producto y las fechas.");
      return;
    }
    setActivePromos((current) => {
      const withoutSameProduct = current.map((promo) => {
        if (promo.id === nextPromo.id) return nextPromo;
        if (promo.active && productsMatch(promo.producto, nextPromo.producto)) {
          return { ...promo, active: false, deactivatedAt: new Date().toISOString() };
        }
        return promo;
      });
      if (withoutSameProduct.some((promo) => promo.id === nextPromo.id)) return withoutSameProduct;
      return [...withoutSameProduct, nextPromo];
    });
    setHasUnsavedChanges(true);
    setToast({
      tone: "success",
      message: promoForm.id
        ? `Promo de ${catalogName} actualizada.`
        : `Promo activa registrada para ${catalogName}.`,
    });
    setPromoOpen(true);
    resetPromoForm();
  }

  function editPromo(promo) {
    setPromoOpen(true);
    setShowPromoForm(true);
    setPromoForm({
      id: promo.id,
      producto: promo.producto,
      startDate: promo.startDate,
      durationPreset: PROMO_DURATION_PRESETS.some((item) => item.value === promo.durationPreset)
        ? promo.durationPreset
        : promo.endDate
          ? "3dias"
          : "hasta_desactivar",
      multiplier: promo.multiplier,
      extraPiecesPerDay: promo.extraPiecesPerDay,
      note: promo.note || "",
    });
    setPromoFormError("");
  }

  function deactivatePromo(promoId) {
    const target = activePromos.find((promo) => promo.id === promoId);
    setActivePromos((current) =>
      current.map((promo) =>
        promo.id === promoId
          ? { ...promo, active: false, deactivatedAt: new Date().toISOString() }
          : promo
      )
    );
    if (promoForm.id === promoId) resetPromoForm();
    setHasUnsavedChanges(true);
    setToast({
      tone: "success",
      message: target ? `Promo de ${target.producto} desactivada.` : "Promo desactivada.",
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

  useEffect(() => {
    try {
      localStorage.setItem(ACTIVE_PROMOS_STORAGE_KEY, JSON.stringify(activePromos));
    } catch {
      // Si el navegador bloquea localStorage, la app debe seguir funcionando.
    }
  }, [activePromos]);

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
  const effectiveDailyBranchStock = useMemo(
    () => applyProductAliases(sanitizeDailyBranchStock(dailyBranchStock), productAliases, officialProducts),
    [dailyBranchStock, productAliases, officialProducts]
  );
  const effectiveDailyColdRoom = useMemo(
    () => applyProductAliases(sanitizeDailyColdRoom(dailyColdRoom), productAliases, officialProducts),
    [dailyColdRoom, productAliases, officialProducts]
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

  const salesMonthCoverage = useMemo(() => {
    const extra = [];
    if (monthlyClosePeriod === selectedMonth) extra.push(...effectiveMonthlyCloseSales);
    extra.push(...effectiveVentasValidacion.filter((record) => {
      const monthKey = monthKeyFromRecord(record);
      return !monthKey || monthKey === selectedMonth;
    }));
    return buildSalesMonthCoverage(extra.length ? [...effectiveVentas, ...extra] : effectiveVentas);
  }, [effectiveVentas, effectiveMonthlyCloseSales, effectiveVentasValidacion, monthlyClosePeriod, selectedMonth]);

  const pendingIncompleteSalesMonths = useMemo(() => {
    const pending = buildSalesMonthCoverage(pendingSalesImport);
    const combined = new Map(salesMonthCoverage.map((row) => [row.monthKey, row]));
    return pending
      .map((row) => ({ ...row, combinedStatus: combined.get(row.monthKey)?.status || row.status }))
      .filter((row) => row.combinedStatus !== "complete");
  }, [pendingSalesImport, salesMonthCoverage]);

  const historicalMonthKeys = useMemo(() => {
    const keys = new Set(historicalVentas.map((record) => monthKeyFromRecord(record)).filter(Boolean));
    return [...keys].sort();
  }, [historicalVentas]);

  const liveForecast = useMemo(
    () =>
      calculateForecast({
        stockRows,
        historicalVentas,
        bajas: effectiveBajas,
        existencias: effectiveExistencias,
        realProduction: effectiveRealProduction,
        selectedMonth,
        dailyBufferPct,
        activePromos,
      }),
    [
      stockRows,
      historicalVentas,
      effectiveBajas,
      effectiveExistencias,
      effectiveRealProduction,
      selectedMonth,
      dailyBufferPct,
      activePromos,
    ]
  );
  const effectiveForecastState = useMemo(
    () =>
      resolveEffectiveForecast({
        liveForecast,
        frozenSnapshot: monthlyReviewSource,
        selectedMonth,
      }),
    [liveForecast, monthlyReviewSource, selectedMonth]
  );
  const forecast = effectiveForecastState.rows;
  const operationalScenario = useMemo(
    () => buildOperationalForecastScenario(forecast),
    [forecast]
  );
  const operationalScenarioTotal = useMemo(
    () => operationalScenario.reduce((sum, row) => sum + row.pronosticoOperativo, 0),
    [operationalScenario]
  );
  const inventoryCutoff = useMemo(
    () => inventoryCutoffStatus(existenciasCutoff.date, selectedMonth),
    [existenciasCutoff.date, selectedMonth]
  );
  const inventoryBlocksApproval = existencias.length > 0 && inventoryCutoff.status !== "fresh";
  const freezeReadiness = useMemo(
    () =>
      assessForecastFreezeReadiness({
        selectedMonth,
        coverageRows: salesMonthCoverage,
        databaseSync,
        alreadyFrozen: Boolean(monthlyReviewSource),
        isAdmin: session.user.rol === "admin",
        capturedStatuses: countCapturedProductStatuses(monthlyReview.inputs),
        catalogCount: forecast.length,
      }),
    [
      selectedMonth,
      salesMonthCoverage,
      databaseSync,
      monthlyReviewSource,
      session.user.rol,
      monthlyReview.inputs,
      forecast.length,
    ]
  );
  const forecastHealth = useMemo(
    () =>
      buildForecastHealth({
        stockRows,
        ventas: effectiveVentas,
        selectedMonth,
        dailyBufferPct,
        currentForecastRows: forecast,
      }),
    [stockRows, effectiveVentas, selectedMonth, dailyBufferPct, forecast]
  );
  const monthlyReviewSourceRows = useMemo(
    () => monthlyReviewSource?.contenido?.rows?.length ? monthlyReviewSource.contenido.rows : operationalScenario,
    [monthlyReviewSource, operationalScenario]
  );
  const monthlyReviewRows = useMemo(
    () => buildMonthlyReviewRows({
      sourceRows: monthlyReviewSourceRows,
      forecastRows: liveForecast,
      historicalVentas,
      loadedExistencias: effectiveExistencias,
      inputs: monthlyReview.inputs,
      inventoryUsable: inventoryCutoff.status === "fresh",
    }),
    [monthlyReviewSourceRows, liveForecast, historicalVentas, effectiveExistencias, monthlyReview.inputs, inventoryCutoff.status]
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
    if (targetState === "approved" && inventoryBlocksApproval) {
      setToast({
        tone: "warning",
        message: inventoryCutoff.status === "missing"
          ? "Captura la fecha de corte de existencias antes de aprobar."
          : `Las existencias del ${displayDate(inventoryCutoff.cutoff)} están fuera de la ventana ${displayDate(inventoryCutoff.windowStart)} a ${displayDate(inventoryCutoff.windowEnd)}.`,
      });
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
            inventoryCutoff: inventoryCutoff.cutoff,
            inventoryCutoffStatus: inventoryCutoff.status,
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
    if (!liveForecast.length || forecastFreezing) return;
    if (!freezeReadiness.canFreeze) {
      setToast({
        tone: "warning",
        message: freezeReadiness.blockers[0]?.message || "Falta información para congelar este mes.",
      });
      return;
    }
    setForecastFreezing(true);
    try {
      let workspaceVersion = lastBackup?.version || null;
      if (hasUnsavedChanges || !lastBackup) {
        const saved = await saveWorkspace({ silent: true });
        if (!saved.ok) throw saved.error;
        workspaceVersion = saved.version;
      }
      const frozenAt = new Date().toISOString();
      const operationalRows = buildOperationalForecastScenario(liveForecast);
      const frozenContent = {
        schemaVersion: 2,
        frozenAt,
        selectedMonth,
        modelVersion: FORECAST_MODEL_VERSION,
        operationalMarginPct: OPERATIONAL_MARGIN_PCT,
        source: { workspaceVersion, files },
        rows: operationalRows,
        forecastRows: snapshotForecastRowsForFreeze(liveForecast),
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
        rows: operationalRows,
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
        activePromos,
        dailyBranchStock: effectiveDailyBranchStock,
        dailyColdRoom: effectiveDailyColdRoom,
      }),
    [forecast, ventasRealesMes, effectiveRealProduction, selectedMonth, dailyBufferPct, activePromos, effectiveDailyBranchStock, effectiveDailyColdRoom]
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
  const inventoryDate = [stockCaptureDate, dailyDateFilter, defaultInventoryDate(selectedMonth)].find(isPlausibleIsoDate)
    || defaultInventoryDate(selectedMonth);
  const knownSucursales = useMemo(
    () => collectSucursales({
      ventas: effectiveVentas,
      bajas: effectiveBajas,
      dailyBranchStock: effectiveDailyBranchStock,
      extra: manualSucursales,
    }),
    [effectiveVentas, effectiveBajas, effectiveDailyBranchStock, manualSucursales]
  );
  const stockRowsForDate = useMemo(
    () => effectiveDailyBranchStock.filter((row) => row.fecha === inventoryDate),
    [effectiveDailyBranchStock, inventoryDate]
  );
  const stockQtyByProductBranch = useMemo(() => {
    const map = new Map();
    for (const row of stockRowsForDate) {
      map.set(`${row.producto}|${norm(row.sucursal)}`, row.cantidad);
    }
    return map;
  }, [stockRowsForDate]);
  const coldRowsForDate = useMemo(
    () => effectiveDailyColdRoom.filter((row) => row.fecha === inventoryDate),
    [effectiveDailyColdRoom, inventoryDate]
  );
  const coldQtyByProduct = useMemo(() => {
    const map = new Map();
    for (const row of coldRowsForDate) {
      map.set(row.producto, row.cantidad);
    }
    return map;
  }, [coldRowsForDate]);
  const stockGridProducts = useMemo(() => {
    const catalog = officialProducts.length
      ? officialProducts
      : [...new Set([
        ...stockRowsForDate.map((row) => row.producto),
        ...coldRowsForDate.map((row) => row.producto),
      ])];
    const query = norm(stockCaptureQuery);
    return catalog.filter((product) => !query || product.includes(query));
  }, [officialProducts, stockRowsForDate, coldRowsForDate, stockCaptureQuery]);
  const inventoryCapturedDates = useMemo(
    () => [...new Set([
      ...effectiveDailyBranchStock.map((row) => row.fecha),
      ...effectiveDailyColdRoom.map((row) => row.fecha),
    ])].sort(),
    [effectiveDailyBranchStock, effectiveDailyColdRoom]
  );
  const inventoryDayCount = stockRowsForDate.length;
  const inventoryDayColdCount = coldRowsForDate.length;
  const inventoryDayHasCapture = inventoryDayCount > 0 || inventoryDayColdCount > 0;
  const inventoryDayProducts = new Set([
    ...stockRowsForDate.map((row) => row.producto),
    ...coldRowsForDate.map((row) => row.producto),
  ]).size;
  const inventoryDayPieces = stockRowsForDate.reduce((sum, row) => sum + row.cantidad, 0);
  const inventoryDayColdPieces = coldRowsForDate.reduce((sum, row) => sum + row.cantidad, 0);

  useEffect(() => {
    const next = defaultInventoryDate(selectedMonth);
    setStockCaptureDate((current) => {
      if (isPlausibleIsoDate(current) && selectedMonth && current.startsWith(selectedMonth)) return current;
      return next;
    });
    setDailyDateFilter((current) => {
      if (!current) return next;
      if (isPlausibleIsoDate(current) && selectedMonth && current.startsWith(selectedMonth)) return current;
      return next;
    });
  }, [selectedMonth]);
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
        selectedMonth,
        stockSheetNotice
      ),
    [productValidationSummary, homologationRows, historicalVentas, selectedMonth, stockSheetNotice]
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
    { label: "Ventas históricas", loaded: Boolean(files.ventas || ventas.length), primary: true },
    { label: "Stock fijo", loaded: Boolean(files.stock || stockRows.length), primary: true },
    { label: "Producción real", loaded: Boolean(files.real || realProduction.length) },
    { label: "Bajas/devoluciones", loaded: Boolean(files.bajas || bajas.length) },
    { label: "Existencias", loaded: Boolean(files.existencias || existencias.length), detail: existencias.length ? (inventoryCutoff.cutoff ? displayDate(inventoryCutoff.cutoff) : "Sin fecha de corte") : "" },
  ];
  const primaryFileItems = loadedFileItems.filter((item) => item.primary);
  const listedActivePromos = activePromos
    .filter((promo) => isPromoListedAsActive(promo))
    .sort((a, b) => a.producto.localeCompare(b.producto, "es") || a.startDate.localeCompare(b.startDate));
  const canSave = session.user.rol === "admin" || session.user.rol === "operador";
  const databaseSyncBlocked = !isDatabaseSyncComplete(databaseSync);
  const freezeBlockedReason = freezeReadiness.canFreeze
    ? (freezeReadiness.warnings[0]?.message || "Guarda una versión inmutable por producto y descarga el mismo pronóstico en Excel")
    : freezeReadiness.blockers[0]?.message || "Falta información para congelar este mes.";
  const cloudStatusTone = cloudStatus.startsWith("No se pudo")
    ? "error"
    : databaseSync && databaseSyncBlocked
      ? "warning"
      : "";

  return (
    <div className="app">
      <main className="main">
        <header className="top">
          <div>
              <span className="eyebrow">Archivo Maestro</span>
            <h2>Decisión de planta</h2>
            <p>Carga datos, revisa la salud del pronóstico y congela la producción sugerida del mes.</p>
          </div>
          <div className="top-toolbar">
            <label className="month-control">
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
              {effectiveForecastState.source !== "live" && (
                <span className="pill ok month-frozen-chip">
                  Congelado{effectiveForecastState.frozenVersion ? ` v${effectiveForecastState.frozenVersion}` : ""}
                </span>
              )}
            </label>
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
              <button
                className="primary"
                type="button"
                onClick={freezeAndExportForecast}
                disabled={!canSave || cloudSaving || forecastFreezing || !forecast.length || !freezeReadiness.canFreeze}
                title={freezeBlockedReason}
              >
                <ShieldCheck size={18} /> {forecastFreezing ? "Congelando..." : "Congelar mes"}
              </button>
              <button className="secondary" type="button" onClick={() => exportToExcel(filtered, summary)} disabled={!forecast.length}>
                <Download size={18} /> Exportar
              </button>
              <details className="more-menu">
                <summary className="secondary">Más opciones</summary>
                <div className="more-menu-panel">
                  <p>Filtros del Excel mensual. No cambian la tabla diaria.</p>
                  <div className="search">
                    <Search size={18} />
                    <input placeholder="Buscar producto del Excel..." value={query} onChange={(e) => setQuery(e.target.value)} />
                  </div>
                  <label className="check-control">
                    <input
                      type="checkbox"
                      checked={showMissingReal}
                      onChange={(e) => setShowMissingReal(e.target.checked)}
                    />
                    Incluir productos sin dato real
                  </label>
                  {session.user.rol === "admin" && (
                    <button className="secondary" type="button" onClick={toggleUsers}>
                      <UserRound size={18} /> Usuarios
                    </button>
                  )}
                </div>
              </details>
              <button className="icon-button" type="button" onClick={onLogout} title="Cerrar sesión" aria-label="Cerrar sesión">
                <LogOut size={18} />
              </button>
            </div>
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
                  <input type="password" minLength="4" value={userForm.password} onChange={(event) => setUserForm((value) => ({ ...value, password: event.target.value }))} required />
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
                    review: { ...monthlyReview, inventoryCutoff: inventoryCutoff.cutoff },
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
                    disabled={!monthlyReviewSource || monthlyReviewSaving || monthlyReviewSummary.pending > 0 || inventoryBlocksApproval}
                    title={
                      monthlyReviewSummary.pending
                        ? "Decide todas las propuestas antes de aprobar"
                        : inventoryBlocksApproval
                          ? "La fecha de corte de existencias debe estar en la ventana del mes"
                          : "Aprobar plan operativo"
                    }
                  >
                    <CheckCircle2 size={17} /> Aprobar plan
                  </button>
                )}
              </div>
            </div>

            <p className={`monthly-review-note ${monthlyReviewSource && !inventoryBlocksApproval ? "success" : "warning"}`}>
              {monthlyReviewSource
                ? `Fuente inmutable: pronóstico congelado ${selectedMonth}, versión ${monthlyReviewSource.version}. ${monthlyReview.state === "approved" ? "Editar cualquier dato abrirá un nuevo borrador." : "El pronóstico original no será modificado."}`
                : `Aún no existe un pronóstico congelado para ${selectedMonth}. Puedes revisar propuestas provisionales, pero debes usar “Congelar mes” antes de guardarlas.`}
              {existencias.length
                ? inventoryCutoff.status === "fresh"
                  ? ` Existencias con corte ${displayDate(inventoryCutoff.cutoff)} (ventana ${displayDate(inventoryCutoff.windowStart)} a ${displayDate(inventoryCutoff.windowEnd)}).`
                  : inventoryCutoff.status === "missing"
                    ? " Falta la fecha de corte de existencias: no se descuentan y no se puede aprobar."
                    : ` Existencias del ${displayDate(inventoryCutoff.cutoff)} fuera de ventana: no se descuentan y no se puede aprobar.`
                : ""}
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
                ? `Cierre ${selectedMonth}: ${formatNumber(monthlyClose.summary.productos)} productos comparados contra el ${effectiveForecastState.label}.`
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
              <span className="eyebrow">Paso 1 · Datos</span>
              <h3>Cargar stock y ventas</h3>
              <p>Estos dos archivos encienden el pronóstico. Bajas, existencias y producción real están en Más archivos.</p>
            </div>
            <div className="loaded-context">
              <span>Mes: {selectedMonth || "Sin mes"}</span>
              <span>Margen: {dailyBufferPct}%</span>
            </div>
          </div>
          <div className="file-status-grid primary-file-status-grid">
            {primaryFileItems.map((item) => (
              <div className="file-status-item" key={item.label}>
                <span className={`status-dot ${item.loaded ? "loaded" : ""}`} />
                <strong>{item.label}</strong>
                <small>{item.detail || (item.loaded ? "Cargado" : "Pendiente")}</small>
              </div>
            ))}
          </div>
          <div className={`cloud-status ${cloudStatusTone}`}>
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

        <section className="uploads required-uploads">
          <UploadBox
            title="Stock fijo"
            description="Productos oficiales, stock objetivo y orden."
            required
            onFile={handleStockFile}
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
        </section>

        <SectionDisclosure
          className="complementary-files advanced-details"
          summaryClassName="advanced-summary"
          contentClassName="advanced-details-content complementary-files-content"
          eyebrow="Opcional"
          title="Más archivos"
          description="Bajas, existencias y producción real. No bloquean el pronóstico base."
          badge={(
            <span className="advanced-count">
              {[
                files.bajas || bajas.length,
                files.existencias || existencias.length,
                files.real || realProduction.length,
              ].filter(Boolean).length}
              /3 cargados
            </span>
          )}
        >
          <section className="uploads optional-uploads">
            <UploadBox
              title="Bajas"
              description="Merma, devoluciones o bajas por producto."
              onFile={handleWasteFile}
              fileName={files.bajas}
            />
            <UploadBox
              title="Existencias"
              description="Corte mensual (EXISTENCIA EN SUCURSALES) para el balance. El inventario diario por sucursal se captura en Planta."
              onFile={handleExistenciasFile}
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
        </SectionDisclosure>

        {salesSourceDecisions.length > 0 && (
          <section className="sales-conflict-bar">
            <p>
              {salesSourceDecisions.some((decision) => decision.strategy === "keep-daily-and-close")
                ? "Si un mes trae cierre dedicado y Excel diario, se usan juntos: el cierre fija el total y el diario conserva la forma (sábado vs martes e impulso GDE)."
                : "Hay más de un archivo para el mismo mes. Se usó el archivo específico; puedes cambiar la fuente canónica."}
            </p>
            {salesSourceDecisions.map((decision) => {
              const dailyKept = (decision.kept || []).filter((name) => name !== decision.winner);
              return (
                <div className="sales-conflict-row" key={decision.month}>
                  <strong>{displayMonthLabel(decision.month)}</strong>
                  <span>
                    {decision.strategy === "keep-daily-and-close"
                      ? `Cierre ${decision.winner}${dailyKept.length ? ` + diario ${dailyKept.join(", ")}` : ""}`
                      : `Usando ${decision.winner}`}
                  </span>
                  {(decision.omitted || []).map((name) => (
                    <button key={name} className="secondary" type="button" onClick={() => chooseSalesMonthSource(decision.month, name)}>
                      Usar {name}
                    </button>
                  ))}
                </div>
              );
            })}
          </section>
        )}

        {existencias.length > 0 && (
          <section className={`inventory-cutoff-bar ${inventoryCutoff.status === "fresh" ? "success" : "warning"}`}>
            <label>
              Fecha de corte de existencias
              <input
                type="date"
                value={existenciasCutoff.date || ""}
                onChange={(event) => {
                  setExistenciasCutoff((current) => ({
                    ...current,
                    date: event.target.value,
                    source: "captura manual",
                  }));
                  setHasUnsavedChanges(true);
                }}
              />
            </label>
            <p>
              {inventoryCutoff.status === "fresh"
                ? `Válidas para ${selectedMonth}: ${displayDate(inventoryCutoff.windowStart)} a ${displayDate(inventoryCutoff.windowEnd)}${existenciasCutoff.source ? ` · origen: ${existenciasCutoff.source}` : ""}.`
                : inventoryCutoff.status === "missing"
                  ? "Sin fecha de corte: la revisión no descuenta inventario y no se puede aprobar."
                  : `Fuera de ventana (${displayDate(inventoryCutoff.windowStart)} a ${displayDate(inventoryCutoff.windowEnd)}): no se descuentan y no se puede aprobar.`}
            </p>
          </section>
        )}

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
                  <Database size={18} /> {salesImporting
                    ? "Guardando..."
                    : salesImportRun?.nextBatch > 1
                      ? `Reanudar desde lote ${salesImportRun.nextBatch}`
                      : "Guardar ventas en la base"}
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

            <p className={`sales-import-status ${salesImportStatus.startsWith("No ") || salesImportStatus.includes("Falló") ? "error" : ""}`}>
              {salesImportStatus}
            </p>
            {salesImportPreview.monthlyTotals > 0 && pendingSalesImport.length > 0 && (
              <p className="sales-import-note">
                Los totales mensuales permanecen disponibles para el pronóstico, pero no se guardan como ventas diarias.
              </p>
            )}
            {pendingIncompleteSalesMonths.length > 0 && (
              <p className="sales-import-note">
                Meses incompletos: {pendingIncompleteSalesMonths.map((row) => `${displayMonthLabel(row.monthKey)} (${salesCoverageStatusLabel(row.combinedStatus)})`).join("; ")}.
                Carga el cierre mensual y el detalle diario del mismo mes.
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
          resumeFromBatch={productionImportRun?.nextBatch > 1 ? productionImportRun.nextBatch : null}
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
          resumeFromBatch={wasteImportRun?.nextBatch > 1 ? wasteImportRun.nextBatch : null}
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

        <section className="executive-summary-section decision-section">
          <div className="section-heading compact-heading">
            <div>
              <span className="eyebrow">Paso 2 · Salud</span>
              <h3>Pronóstico listo para planta</h3>
              <p>Revisa el WAPE de meses cerrados, el sync y el total sugerido antes de congelar. Un mes congelado ya no se recalcula en silencio.</p>
            </div>
            <strong className="decision-month-chip">{selectedMonth || "Sin mes"} · margen {dailyBufferPct}%</strong>
          </div>

          <ForecastHealthStrip health={forecastHealth} selectedMonth={selectedMonth} />
          <ForecastAccuracyPanel health={forecastHealth} selectedMonth={selectedMonth} />
          <FrozenMonthBanner forecastLock={effectiveForecastState} selectedMonth={selectedMonth} />
          <FreezeReadinessStrip readiness={freezeReadiness} selectedMonth={selectedMonth} />

          <section className="executive-summary-kpis decision-kpis">
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
              caption={effectiveForecastState.source === "live"
                ? "Suma de pronósticos diarios"
                : `Fijado: ${effectiveForecastState.label}`}
            />
            <KpiCard
              icon={ShieldCheck}
              label="Producción sugerida mensual"
              value={formatNumber(dailySummary.produccionSugeridaMensual, 0)}
              caption={effectiveDailyBranchStock.length ? "Con regla operativa e inventario del día" : "Con regla operativa"}
            />
            <KpiCard
              icon={Target}
              label="Escenario operativo"
              value={formatNumber(operationalScenarioTotal, 0)}
              caption={`Colchón +${OPERATIONAL_MARGIN_PCT}% separado del estadístico`}
            />
          </section>
        </section>

        <SectionDisclosure
          className={`promo-section compact-analytics-section ${listedActivePromos.length ? "has-promos" : "is-empty"}`}
          eyebrow="Cuando haga falta"
          title="Promo activa"
          description={listedActivePromos.length
            ? "Empuje puntual sobre la producción sugerida. No cambia el WAPE histórico."
            : "Vacía. Ábrela solo si hay que empujar un SKU para que planta no se quede corta."}
          icon={Megaphone}
          badge={(
            <span className={`pill ${listedActivePromos.length ? "ok" : "muted"}`}>
              {listedActivePromos.length ? `${listedActivePromos.length} activa${listedActivePromos.length === 1 ? "" : "s"}` : "Sin promo"}
            </span>
          )}
          open={promoOpen}
          onToggle={setPromoOpen}
        >
          <div className="promo-grid">
            {(showPromoForm || !listedActivePromos.length) ? (
            <form className="promo-form" onSubmit={submitPromo}>
              <label>
                Producto
                <input
                  list="promo-product-options"
                  value={promoForm.producto}
                  onChange={(event) => setPromoForm((current) => ({ ...current, producto: event.target.value }))}
                  placeholder="Nombre del catálogo"
                  required
                  disabled={!canSave}
                />
                <datalist id="promo-product-options">
                  {officialProducts.map((product) => (
                    <option value={product} key={product} />
                  ))}
                </datalist>
              </label>
              <label>
                Inicio
                <input
                  type="date"
                  value={promoForm.startDate}
                  onChange={(event) => setPromoForm((current) => ({ ...current, startDate: event.target.value }))}
                  disabled={!canSave}
                  required
                />
              </label>
              <label>
                Duración
                <select
                  value={promoForm.durationPreset}
                  onChange={(event) => setPromoForm((current) => ({ ...current, durationPreset: event.target.value }))}
                  disabled={!canSave}
                >
                  {PROMO_DURATION_PRESETS.map((preset) => (
                    <option value={preset.value} key={preset.value}>{preset.label}</option>
                  ))}
                </select>
              </label>
              <label>
                Multiplicar pronóstico
                <input
                  type="number"
                  min="0.1"
                  step="0.1"
                  value={promoForm.multiplier}
                  onChange={(event) => setPromoForm((current) => ({ ...current, multiplier: event.target.value }))}
                  disabled={!canSave}
                />
              </label>
              <label>
                Piezas extra por día
                <input
                  type="number"
                  min="0"
                  step="1"
                  value={promoForm.extraPiecesPerDay}
                  onChange={(event) => setPromoForm((current) => ({ ...current, extraPiecesPerDay: event.target.value }))}
                  disabled={!canSave}
                />
              </label>
              <label className="promo-note-field">
                Nota
                <input
                  value={promoForm.note}
                  onChange={(event) => setPromoForm((current) => ({ ...current, note: event.target.value }))}
                  placeholder="Ej. sobra masa, reunión de lunes"
                  disabled={!canSave}
                />
              </label>
              <div className="promo-form-actions">
                <button className="primary" type="submit" disabled={!canSave}>
                  <Megaphone size={17} /> {promoForm.id ? "Guardar cambios" : "Activar promo"}
                </button>
                {promoForm.id && (
                  <button className="secondary" type="button" onClick={resetPromoForm}>Cancelar</button>
                )}
              </div>
              {promoFormError && <small className="promo-form-error">{promoFormError}</small>}
              <small className="promo-form-hint">
                ×1.3 sube la sugerencia de planta 30%. Las piezas extra se suman cada día de planta (lunes a sábado). ×1 sin extras solo evita que la limpieza de catálogo apague el SKU.
              </small>
            </form>
            ) : (
              <div className="promo-list-toolbar">
                <p>Las promos de abajo ya empujan la producción sugerida. El formulario queda oculto hasta que agregues otra.</p>
                <button className="secondary" type="button" onClick={() => setShowPromoForm(true)} disabled={!canSave}>
                  <Megaphone size={17} /> Nueva promo
                </button>
              </div>
            )}
            <div className="promo-list">
              {listedActivePromos.map((promo) => (
                <div className="promo-list-row" key={promo.id}>
                  <div>
                    <strong>{promo.producto}</strong>
                    <span>{formatPromoWindowLabel(promo)}</span>
                    {promo.note && <small>{promo.note}</small>}
                  </div>
                  <span className="pill ok">{formatPromoUpliftLabel(promo)}</span>
                  <div className="promo-list-actions">
                    <button className="secondary" type="button" onClick={() => editPromo(promo)} disabled={!canSave}>
                      Editar
                    </button>
                    <button className="secondary" type="button" onClick={() => deactivatePromo(promo.id)} disabled={!canSave}>
                      Desactivar
                    </button>
                  </div>
                </div>
              ))}
              {!listedActivePromos.length && (
                <div className="empty">No hay promos activas. Si deciden empujar un producto hoy, regístrenlo aquí para que planta no se quede corta.</div>
              )}
            </div>
          </div>
        </SectionDisclosure>

        <section className="daily-section">
          <div className="section-heading compact-heading">
            <div>
              <span className="eyebrow">Paso 3 · Planta</span>
              <h3>Producción diaria sugerida</h3>
              <p>Captura sucursales y cuarto frío. Pedido planta es el envío; A producir es lo que hay que fabricar. El mes se cambia arriba, junto a Congelar.</p>
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

          <section className="daily-stock-panel" aria-label="Inventario diario por sucursal y cuarto frío">
            <div className="daily-stock-heading">
              <span className="daily-stock-icon"><Warehouse size={20} /></span>
              <div>
                <span className="eyebrow">Inventario del día</span>
                <h4>Stock por sucursal, cuarto frío y SKU</h4>
                <p>Sucursales restan del pedido de planta. Cuarto frío resta de A producir. Vacío = no capturado. 0 = contado vacío.</p>
              </div>
              <span className={`pill ${inventoryDayHasCapture ? "ok" : "muted"}`}>
                {inventoryDayHasCapture
                  ? `${formatNumber(inventoryDayProducts)} SKU · ${formatNumber(inventoryDayPieces, 0)} suc. · ${formatNumber(inventoryDayColdPieces, 0)} CF · ${displayDate(inventoryDate)}`
                  : `Sin captura · ${displayDate(inventoryDate)}`}
              </span>
            </div>
            <div className="daily-stock-toolbar">
              <label>
                Fecha de inventario
                <input
                  type="date"
                  min="2020-01-01"
                  max="2100-12-31"
                  value={isPlausibleIsoDate(inventoryDate) ? inventoryDate : ""}
                  onChange={(event) => {
                    const next = event.target.value;
                    if (next && !isPlausibleIsoDate(next)) return;
                    setStockCaptureDate(next);
                    setDailyDateFilter(next);
                  }}
                />
              </label>
              <div className="search">
                <Search size={18} />
                <input
                  placeholder="Buscar SKU del inventario..."
                  value={stockCaptureQuery}
                  onChange={(event) => setStockCaptureQuery(event.target.value)}
                />
              </div>
              <label className="daily-stock-add-branch">
                Sucursal
                <span>
                  <input
                    value={newSucursalName}
                    onChange={(event) => setNewSucursalName(event.target.value)}
                    placeholder="Nombre de sucursal"
                    disabled={!canSave}
                    onKeyDown={(event) => {
                      if (event.key === "Enter") {
                        event.preventDefault();
                        addManualSucursal();
                      }
                    }}
                  />
                  <button className="secondary" type="button" onClick={addManualSucursal} disabled={!canSave || !newSucursalName.trim()}>
                    Agregar
                  </button>
                </span>
              </label>
              <label className="upload-button daily-stock-upload">
                <FileSpreadsheet size={17} />
                Importar Excel
                <input
                  type="file"
                  accept=".xlsx,.xls"
                  disabled={!canSave}
                  onChange={(event) => {
                    handleDailyBranchStockFile(event.target.files?.[0]);
                    event.target.value = "";
                  }}
                />
              </label>
              <button
                className="secondary"
                type="button"
                onClick={() => exportDailyBranchStockTemplate(inventoryDate, knownSucursales, officialProducts)}
              >
                <Download size={17} /> Plantilla
              </button>
              <button
                className="secondary"
                type="button"
                onClick={() => clearDailyInventoryForDate(inventoryDate)}
                disabled={!canSave || !inventoryDayHasCapture}
              >
                Vaciar este día
              </button>
            </div>
            {files.dailyBranchStock && <span className="file-name">Último Excel: {files.dailyBranchStock}</span>}
            {knownSucursales.length === 0 && !officialProducts.length && !coldRowsForDate.length ? (
              <div className="empty">Agrega una sucursal o importa un Excel RAIZ (sucursales y/o Cuarto frío / Restante CF). Las metas TOTAL A TENER no se importan como sucursal.</div>
            ) : !stockGridProducts.length ? (
              <div className="empty">
                {officialProducts.length
                  ? "Ningún SKU coincide con la búsqueda."
                  : "Carga el stock fijo para ver el catálogo, o importa un Excel con productos."}
              </div>
            ) : (
              <div className="daily-stock-table-wrap">
                <table className="daily-stock-table">
                  <thead>
                    <tr>
                      <th>Producto</th>
                      {knownSucursales.map((sucursal) => (
                        <th key={sucursal}>{sucursal}</th>
                      ))}
                      <th>Total sucursales</th>
                      <th className="daily-stock-cf">Cuarto frío</th>
                    </tr>
                  </thead>
                  <tbody>
                    {stockGridProducts.map((product) => {
                      const total = knownSucursales.reduce((sum, sucursal) => {
                        const value = stockQtyByProductBranch.get(`${product}|${norm(sucursal)}`);
                        return sum + (Number.isFinite(value) ? value : 0);
                      }, 0);
                      const captured = knownSucursales.some((sucursal) => stockQtyByProductBranch.has(`${product}|${norm(sucursal)}`));
                      const coldValue = coldQtyByProduct.has(product) ? coldQtyByProduct.get(product) : "";
                      return (
                        <tr key={product}>
                          <td>{product}</td>
                          {knownSucursales.map((sucursal) => {
                            const key = `${product}|${norm(sucursal)}`;
                            const value = stockQtyByProductBranch.has(key) ? stockQtyByProductBranch.get(key) : "";
                            return (
                              <td key={sucursal}>
                                <input
                                  className="daily-stock-qty"
                                  type="number"
                                  min="0"
                                  step="1"
                                  inputMode="numeric"
                                  value={value}
                                  disabled={!canSave}
                                  aria-label={`${product} en ${sucursal}`}
                                  onChange={(event) => updateDailyBranchStockCell(inventoryDate, sucursal, product, event.target.value)}
                                />
                              </td>
                            );
                          })}
                          <td className="strong">{captured ? formatNumber(total, 0) : "—"}</td>
                          <td className="daily-stock-cf">
                            <input
                              className="daily-stock-qty"
                              type="number"
                              min="0"
                              step="1"
                              inputMode="numeric"
                              value={coldValue}
                              disabled={!canSave}
                              aria-label={`${product} en cuarto frío`}
                              onChange={(event) => updateDailyColdRoomCell(inventoryDate, product, event.target.value)}
                            />
                          </td>
                        </tr>
                      );
                    })}
                  </tbody>
                </table>
              </div>
            )}
            {inventoryCapturedDates.length > 1 && (
              <small className="daily-stock-hint">
                Días con inventario: {inventoryCapturedDates.map((date) => displayDate(date)).join(" · ")}
              </small>
            )}
          </section>

          <section className="controls daily-controls">
            <label>
              Fecha
              <input
                type="date"
                min="2020-01-01"
                max="2100-12-31"
                value={isPlausibleIsoDate(dailyDateFilter) ? dailyDateFilter : ""}
                onChange={(e) => {
                  const next = e.target.value;
                  if (next && !isPlausibleIsoDate(next)) return;
                  setDailyDateFilter(next);
                  if (next) setStockCaptureDate(next);
                }}
              />
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
          </section>

          <details className="advanced-details daily-more-filters">
            <summary className="advanced-summary">
              <div>
                <span className="eyebrow">Filtros extra</span>
                <strong>Más filtros de la tabla</strong>
                <small>Faltante, sobreproducción y recortes de revisión</small>
              </div>
              <span className="advanced-count">
                {[onlyDailyShortage, onlyDailyOverproduction].filter(Boolean).length || "Ocultos"}
              </span>
            </summary>
            <div className="advanced-details-content daily-more-filters-content">
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
            </div>
          </details>

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
                  <th>Bruto planta</th>
                  <th>Inventario sucursales</th>
                  <th>Cuarto frío</th>
                  <th>Pedido planta</th>
                  <th>A producir</th>
                </tr>
              </thead>
              <tbody>
                {filteredDailyRows.map((row) => (
                  <tr key={`${row.fecha}-${row.producto}`}>
                    <td>{row.fecha}</td>
                    <td>{row.dia}</td>
                    <td>
                      {row.producto}
                      {row.promoActiva && (
                        <span className="pill ok promo-day-pill" title={row.promoEtiqueta}>Promo</span>
                      )}
                    </td>
                    <td>{row.promedioUsado.toFixed(2)}</td>
                    <td>{row.pronosticoVentaDia.toFixed(2)}</td>
                    <td>{row.colchonDiario.toFixed(2)}</td>
                    <td>{row.baseConColchonDia.toFixed(2)}</td>
                    <td>{row.produccionBrutaDia ?? row.produccionSugeridaDia}</td>
                    <td>{row.hasDailyBranchStock ? formatNumber(row.inventarioSucursalesDia, 0) : "—"}</td>
                    <td>{row.hasDailyColdRoom ? formatNumber(row.cuartoFrioDia, 0) : "—"}</td>
                    <td>{row.produccionSugeridaDia}</td>
                    <td className="strong">{row.aProducirDia ?? row.produccionSugeridaDia}</td>
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
            {dailyRows.length > 0 && (
              <small className="daily-stock-hint">
                Pedido planta = max(0, bruto − sucursales). A producir = max(0, bruto − sucursales − cuarto frío). Sin captura no se descuenta.
              </small>
            )}
          </section>
        </section>

        <section className="secondary-tools-heading">
          <div>
            <span className="eyebrow">Cuando haga falta</span>
            <h3>Seguimiento y auditoría</h3>
            <p>Avance semanal, cierre, homologación y validación. Cerrados para no competir con la decisión de planta.</p>
          </div>
        </section>

        <SectionDisclosure
          className="notes-section compact-analytics-section"
          eyebrow="Lectura rápida"
          title="Notas de interpretación"
          description="Reglas de domingo, lotes de pastel y cómo se lee una promo."
          icon={FileSpreadsheet}
          badge={<span className="pill muted">Guía</span>}
        >
          <div className="notes-list">
            <p>El pronóstico elige el método con menor error en el mes anterior y aplica una calibración limitada.</p>
            <p>El pronóstico de venta se reparte por día de semana. El domingo no se produce y su demanda pasa al sábado.</p>
            <p>Las existencias de corte mensual solo se descuentan en el balance si la fecha cae entre el mes anterior y el mes planificado.</p>
            <p>El inventario diario se captura en Planta. Pedido planta (envío) = max(0, bruto con promo − sucursales). A producir (fabricar) = max(0, bruto con promo − sucursales − cuarto frío). Vacío no descuenta; 0 sí. No se vuelve a armar lote de pastel después de restar inventario.</p>
            <p>Para pasteles GDE, MED y CH, cada día de planta (lunes a sábado) se produce 0 o un lote de 10, 15, 20… Un 13 se hace 15; menos de 8 no se produce. El domingo queda en cero y su demanda pasa al sábado, que también sale en lote.</p>
            <p>Una promo activa es un overlay de planta: no reescribe el WAPE histórico. Mientras dura, no se apaga el SKU por la limpieza de catálogo y la producción sugerida aplica el multiplicador y/o las piezas extra.</p>
            <p>La vista Validación de cálculos permite auditar cada producto.</p>
          </div>
        </SectionDisclosure>

        <SectionDisclosure
          className="validation-section compact-analytics-section"
          eyebrow="Auditoría paso a paso"
          title="Validación de cálculos"
          description="Ventas leídas, promedios y el cálculo diario de un producto."
          icon={BarChart3}
          badge={<span className="pill muted">{validationProduct || "Sin producto"}</span>}
        >
          <div className="section-heading compact-heading">
            <div>
              <span className="eyebrow">Producto a auditar</span>
              <h3>Desglose del cálculo</h3>
              <p>Revisa las ventas leídas, los promedios aplicados y el cálculo diario completo.</p>
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
                        <th>Bruto planta</th>
                        <th>Inventario sucursales</th>
                        <th>Cuarto frío</th>
                        <th>Pedido planta</th>
                        <th>A producir</th>
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
                          <td>{formatNumber(row.produccionBrutaDia ?? row.produccionSugeridaDia)}</td>
                          <td>{row.hasDailyBranchStock ? formatNumber(row.inventarioSucursalesDia, 0) : "—"}</td>
                          <td>{row.hasDailyColdRoom ? formatNumber(row.cuartoFrioDia, 0) : "—"}</td>
                          <td>
                            {formatNumber(row.produccionSugeridaDia)}
                            {row.promoActiva ? ` · promo ${row.promoEtiqueta}` : ""}
                          </td>
                          <td className="strong">{formatNumber(row.aProducirDia ?? row.produccionSugeridaDia)}</td>
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
        </SectionDisclosure>

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
    async function restoreSession() {
      try {
        const me = await apiRequest("/api/auth/me", { token: session?.token });
        if (!active) return;
        if (!me?.user) {
          const invalid = new Error("Sesión inválida");
          invalid.status = 401;
          throw invalid;
        }
        const nextSession = { token: session?.token || "", user: me.user };
        setSession(nextSession);
        writeStoredSession(nextSession);
        setNeedsSetup(false);
        setAuthError("");
      } catch (error) {
        if (!active) return;
        if (error.status === 401) {
          clearStoredSession();
          setSession(null);
          try {
            const status = await apiRequest("/api/auth/status");
            if (!active) return;
            setNeedsSetup(Boolean(status.needsSetup));
            setAuthError("");
          } catch (statusError) {
            setAuthError(`No se pudo conectar con el servidor: ${statusError.message}`);
          }
        } else {
          setAuthError(`No se pudo conectar con el servidor: ${error.message}`);
        }
      } finally {
        if (active) setAuthChecking(false);
      }
    }
    restoreSession();
    return () => {
      active = false;
    };
  }, []);

  async function authenticate(credentials) {
    setAuthSubmitting(true);
    setAuthError("");
    try {
      const response = await apiRequest(needsSetup ? "/api/auth/setup" : "/api/auth/login", {
        method: "POST",
        body: credentials,
      });
      const nextSession = { token: response.token, user: response.user };
      writeStoredSession(nextSession);
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
    clearStoredSession();
    setSession(null);
    setNeedsSetup(false);
    setAuthError("");
    apiRequest("/api/auth/logout", { token: current?.token, method: "POST" }).catch(() => {});
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

export {
  assessForecastFreezeReadiness,
  assessStockSheetSelection,
  analyzeForecastProductErrors,
  buildForecastHealth,
  buildMonthlyCloseSummary,
  buildOperationalForecastScenario,
  describeAccuracyWindow,
  hydrateForecastFromOperationalRows,
  resolveEffectiveForecast,
  snapshotForecastRowsForFreeze,
  summarizeForecastAccuracy,
  weightedWapeFromBacktests,
  buildSalesMonthCoverage,
  buildWeeklyProgress,
  calculateForecast,
  calculateDailyForecast,
  buildMonthlyForecastData,
  normalizeActivePromo,
  sanitizeActivePromos,
  isPromoActiveOnDate,
  findActivePromoForProduct,
  applyPromoUpliftToQuantity,
  getProduccionSugerida,
  consolidateOperationalRowsForUpload,
  consolidateSalesRowsForUpload,
  countCapturedProductStatuses,
  filterVentasBeforeMonth,
  inferMonthHintFromFileName,
  monthsNamedInFileName,
  parseBajasReport,
  parseBajasSummaryWorkbook,
  parseExistencias,
  parseDailyBranchStock,
  parseDailyInventory,
  parseDailyColdRoom,
  sanitizeDailyBranchStock,
  sanitizeDailyColdRoom,
  upsertDailyBranchStock,
  upsertDailyColdRoom,
  sumDailyBranchStockByProductDate,
  mapDailyColdRoomByProductDate,
  applyDailyBranchStockToPlantSuggestion,
  applyInventoryToProductionSuggestion,
  collectSucursales,
  parseMonthlySummaryWorkbook,
  parseProductionReal,
  parseSalesOrReturns,
  parseStock,
  resolveCanonicalMonthSources,
  sourceKindForMonth,
  describeSourceDecision,
  computeAnnualGrowthFactor,
  resolvePriorYearSeasonal,
  computeRecentMomentumFactor,
  computeEventCarryoverScale,
  computeImpulseCarryoverScale,
  calendarEventForMonth,
  normalizeProduct,
  productMatchKey,
  findOfficialProduct,
  resolveOfficialProduct,
  lookupBuiltinProductAlias,
  countSameYearConsecutiveRecentMonths,
  forecastHidesPriorYearMonths,
  prepareProductForecastHistory,
  buildColdStartPriorYearModel,
  applyColdStartDormantPriorYearGuard,
  applyEventPriorYearIndex,
  applyPriorYearPulseFade,
  lentDaysInMonth,
  isPriceTaggedProduct,
  isPromotionalProduct,
  isOperationalCakeProduct,
  applyCatalogOutlierCleanup,
};

if (typeof document !== "undefined") {
  createRoot(document.getElementById("root")).render(<App />);
}
