const STORAGE_KEY = "archivo-maestro-session-v1";

export function dateFromKey(value) {
  if (value instanceof Date && !Number.isNaN(value.getTime())) return value;
  const raw = String(value ?? "").trim();
  const match = raw.match(/^(\d{4})-(\d{2})-(\d{2})/);
  if (!match) return value ? new Date(value) : null;
  return new Date(Number(match[1]), Number(match[2]) - 1, Number(match[3]));
}

function serializeRecord(record) {
  if (!record || typeof record !== "object") return record;
  const fechaKey =
    record.fechaKey ||
    (record.fecha instanceof Date && !Number.isNaN(record.fecha.getTime())
      ? `${record.fecha.getFullYear()}-${String(record.fecha.getMonth() + 1).padStart(2, "0")}-${String(record.fecha.getDate()).padStart(2, "0")}`
      : record.fecha
        ? String(record.fecha).slice(0, 10)
        : "");
  return {
    ...record,
    fecha: fechaKey || record.fecha,
    fechaKey: fechaKey || record.fechaKey || "",
  };
}

function reviveRecord(record) {
  if (!record || typeof record !== "object") return record;
  const fecha = dateFromKey(record.fecha || record.fechaKey);
  return {
    ...record,
    fecha,
    fechaKey: record.fechaKey || (fecha instanceof Date && !Number.isNaN(fecha.getTime())
      ? `${fecha.getFullYear()}-${String(fecha.getMonth() + 1).padStart(2, "0")}-${String(fecha.getDate()).padStart(2, "0")}`
      : ""),
  };
}

export function loadSession() {
  try {
    const raw = localStorage.getItem(STORAGE_KEY);
    if (!raw) return null;
    const parsed = JSON.parse(raw);
    return {
      ...parsed,
      stockRows: Array.isArray(parsed.stockRows) ? parsed.stockRows : [],
      ventas: Array.isArray(parsed.ventas) ? parsed.ventas.map(reviveRecord) : [],
      bajas: Array.isArray(parsed.bajas) ? parsed.bajas.map(reviveRecord) : [],
      existencias: Array.isArray(parsed.existencias) ? parsed.existencias : [],
      realProduction: Array.isArray(parsed.realProduction) ? parsed.realProduction.map(reviveRecord) : [],
      files: parsed.files && typeof parsed.files === "object" ? parsed.files : {},
    };
  } catch {
    return null;
  }
}

export function saveSession(session) {
  try {
    const payload = {
      version: 1,
      savedAt: new Date().toISOString(),
      selectedMonth: session.selectedMonth || "",
      selectedMonthTouched: Boolean(session.selectedMonthTouched),
      dailyBufferPct: Number(session.dailyBufferPct) || 0,
      weekendBoost: Number(session.weekendBoost) || 1,
      showMissingReal: Boolean(session.showMissingReal),
      files: session.files || {},
      stockRows: (session.stockRows || []).map(serializeRecord),
      ventas: (session.ventas || []).map(serializeRecord),
      bajas: (session.bajas || []).map(serializeRecord),
      existencias: session.existencias || [],
      realProduction: (session.realProduction || []).map(serializeRecord),
    };
    localStorage.setItem(STORAGE_KEY, JSON.stringify(payload));
    return true;
  } catch {
    return false;
  }
}

export function clearSession() {
  try {
    localStorage.removeItem(STORAGE_KEY);
  } catch {
    // ignore quota / private mode
  }
}
