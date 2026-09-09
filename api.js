const API_BASE = String(import.meta.env.VITE_API_URL || "").replace(/\/$/, "");

function apiUrl(path) {
  const suffix = path.startsWith("/") ? path : `/${path}`;
  return API_BASE ? `${API_BASE}${suffix}` : suffix;
}

async function readError(response) {
  const text = await response.text();
  try {
    const json = JSON.parse(text);
    return json.error || text || response.statusText;
  } catch {
    return text || response.statusText;
  }
}

export async function apiRequest(path, options = {}) {
  const response = await fetch(apiUrl(path), {
    headers: {
      "Content-Type": "application/json",
      ...(options.headers || {}),
    },
    ...options,
  });
  if (!response.ok) {
    throw new Error(await readError(response));
  }
  if (response.status === 204) return null;
  return response.json();
}

export async function checkApiHealth() {
  try {
    const data = await apiRequest("/api/health");
    return Boolean(data?.ok);
  } catch {
    return false;
  }
}

function chunk(rows, size = 400) {
  const out = [];
  for (let i = 0; i < rows.length; i += size) out.push(rows.slice(i, i + size));
  return out;
}

async function postChunks(path, key, rows) {
  if (!rows.length) return { ok: true, insertedOrUpdated: 0, received: 0 };
  let insertedOrUpdated = 0;
  let received = 0;
  for (const part of chunk(rows)) {
    const result = await apiRequest(path, {
      method: "POST",
      body: JSON.stringify({ [key]: part }),
    });
    insertedOrUpdated += Number(result?.insertedOrUpdated || 0);
    received += Number(result?.received || part.length);
  }
  return { ok: true, insertedOrUpdated, received };
}

export function toApiDate(value) {
  if (value instanceof Date && !Number.isNaN(value.getTime())) {
    return `${value.getFullYear()}-${String(value.getMonth() + 1).padStart(2, "0")}-${String(value.getDate()).padStart(2, "0")}`;
  }
  const raw = String(value ?? "").trim();
  return /^\d{4}-\d{2}-\d{2}/.test(raw) ? raw.slice(0, 10) : "";
}

export async function loadWorkspaceFromApi() {
  const data = await apiRequest("/api/workspace");
  return data;
}

export async function saveWorkspaceToApi({ stockRows, ventas, realProduction, dailyRows, selectedMonth }) {
  const ventasPayload = (ventas || [])
    .map((row) => {
      const fecha = toApiDate(row.fecha || row.fechaKey);
      if (!fecha || !row.producto) return null;
      return {
        fecha,
        producto: row.producto,
        producto_codigo: row.producto,
        producto_nombre: row.productoOriginal || row.producto,
        cantidad: Number(row.cantidad) || 0,
        importe: Number(row.importe) || 0,
        nombre_origen: row.productoOriginal || "",
      };
    })
    .filter(Boolean);

  const stockPayload = (stockRows || []).map((row, index) => ({
    mes: selectedMonth,
    producto: row.producto,
    producto_codigo: row.producto,
    producto_nombre: row.productoOriginal || row.producto,
    cantidad: Number(row.stock) || 0,
    orden: row.orden || index + 1,
  }));

  const produccionPayload = (realProduction || [])
    .map((row) => {
      const fecha = toApiDate(row.fecha || row.fechaKey);
      if (!fecha || !row.producto) return null;
      return {
        fecha,
        producto: row.producto,
        producto_codigo: row.producto,
        producto_nombre: row.productoOriginal || row.producto,
        cantidad: Number(row.cantidad) || 0,
      };
    })
    .filter(Boolean);

  const pronosticoPayload = (dailyRows || [])
    .map((row) => {
      const fecha = toApiDate(row.fecha);
      if (!fecha || !row.producto) return null;
      return {
        fecha,
        producto: row.producto,
        producto_codigo: row.producto,
        producto_nombre: row.producto,
        cantidad_pronosticada: Number(row.produccionSugeridaDia) || 0,
        metodo: "weekday_colchon_regla",
      };
    })
    .filter(Boolean);

  const results = {};
  if (ventasPayload.length) results.ventas = await postChunks("/api/ventas/bulk", "ventas", ventasPayload);
  if (stockPayload.length) results.stock = await postChunks("/api/stock/bulk", "stock", stockPayload);
  if (produccionPayload.length) {
    results.produccion = await postChunks("/api/produccion-real/bulk", "produccion", produccionPayload);
  }
  if (pronosticoPayload.length) {
    results.pronostico = await postChunks("/api/pronostico/bulk", "pronostico", pronosticoPayload);
  }
  return results;
}

export function mapApiWorkspace(data) {
  const ventas = (data?.ventas || []).map((row) => ({
    fecha: dateSafe(row.fecha),
    producto: row.producto_codigo || row.producto,
    productoOriginal: row.producto_nombre || row.producto_codigo || row.producto,
    cantidad: Number(row.cantidad) || 0,
    importe: Number(row.importe) || 0,
  }));
  const stockRows = (data?.stock || []).map((row, index) => ({
    producto: row.producto_codigo || row.producto,
    productoOriginal: row.producto_nombre || row.producto_codigo || row.producto,
    stock: Number(row.cantidad) || 0,
    orden: index + 1,
  }));
  const realProduction = (data?.produccion || []).map((row) => ({
    fecha: dateSafe(row.fecha),
    fechaKey: String(row.fecha || "").slice(0, 10),
    producto: row.producto_codigo || row.producto,
    productoOriginal: row.producto_nombre || row.producto_codigo || row.producto,
    cantidad: Number(row.cantidad) || 0,
  }));
  return { ventas, stockRows, realProduction };
}

function dateSafe(value) {
  const raw = String(value ?? "").slice(0, 10);
  const match = raw.match(/^(\d{4})-(\d{2})-(\d{2})$/);
  if (!match) return value;
  return new Date(Number(match[1]), Number(match[2]) - 1, Number(match[3]));
}
