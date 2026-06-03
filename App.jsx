import React, { useMemo, useState } from "react";
import { createRoot } from "react-dom/client";
import * as XLSX from "xlsx";
import {
  AlertTriangle,
  BarChart3,
  CheckCircle2,
  Database,
  Download,
  FileSpreadsheet,
  PackageCheck,
  RefreshCw,
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
  "TOTAL AREA",
  "TOTAL ÁREA",
  "ESPECIALIDAD",
]);

const STATUS_META = {
  "Sin dato real": { className: "muted", label: "Sin dato real" },
  "No producir": { className: "muted", label: "No producir" },
  "Dentro de rango": { className: "ok", label: "Dentro de rango" },
  "Riesgo faltante": { className: "danger", label: "Riesgo faltante" },
  Sobreproduccion: { className: "warn", label: "Sobreproducción" },
  Revisar: { className: "warn", label: "Revisar" },
};

function norm(value) {
  return String(value ?? "")
    .trim()
    .toUpperCase()
    .normalize("NFD")
    .replace(/[\u0300-\u036f]/g, "");
}

function isValidProduct(value) {
  const p = norm(value);
  if (!p) return false;
  if (INVALID_PRODUCTS.has(p)) return false;
  if (p.startsWith("TOTAL")) return false;
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

function addDays(date, days) {
  const next = new Date(date);
  next.setDate(next.getDate() + days);
  return next;
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
    const product = norm(rows[i][0]);
    const stock = toNumber(rows[i][1]);
    if (isValidProduct(product)) {
      parsed.push({ producto: product, stock, orden: parsed.length + 1 });
    }
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
    const product = norm(rows[i][productCol]);
    if (!isValidProduct(product)) continue;
    parsed.push({
      producto: product,
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
      const product = norm(rows[i][productCol]);
      if (!isValidProduct(product)) continue;
      const cantidad = toNumber(rows[i][qtyCol]);
      const importe = amountCol >= 0 ? toNumber(rows[i][amountCol]) : 0;
      if (cantidad === 0 && importe === 0) continue;
      out.push({ fecha, producto: product, cantidad, importe, tipo: type });
    }
  }
  return out;
}

function parseWideSales(workbook, type = "ventas") {
  const rows = rowsFromFirstSheet(workbook);
  if (rows.length < 4) return [];
  const weekdays = rows[0].map(norm);
  const parsed = [];
  for (let r = 3; r < rows.length; r++) {
    const product = norm(rows[r][0]);
    if (!isValidProduct(product)) continue;
    for (let c = 1; c < rows[r].length; c++) {
      if (!weekdays[c]) continue;
      const cantidad = toNumber(rows[r][c]);
      if (cantidad === 0) continue;
      parsed.push({
        fecha: new Date(2026, 0, c),
        producto: product,
        cantidad,
        importe: 0,
        weekday: weekdays[c],
        tipo: type,
      });
    }
  }
  return parsed;
}

function parseSalesOrReturns(workbook, type) {
  const bySheets = parseMonthlyDailySheets(workbook, type);
  if (bySheets.length > 0) return bySheets;
  return parseWideSales(workbook, type);
}

function parseProductionReal(workbook) {
  const rows = rowsFromFirstSheet(workbook);
  let headerIndex = -1;
  let productCol = 0;
  let qtyCol = 1;

  for (let i = 0; i < Math.min(rows.length, 15); i++) {
    const row = rows[i].map(norm);
    const p = row.findIndex((x) => x.includes("PRODUCTO") || x.includes("DESCRIPCION") || x.includes("ARTICULO"));
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
      break;
    }
  }

  const start = headerIndex >= 0 ? headerIndex + 1 : 0;
  const map = new Map();
  for (let i = start; i < rows.length; i++) {
    const product = norm(rows[i][productCol]);
    if (!isValidProduct(product)) continue;
    const cantidad = toNumber(rows[i][qtyCol]);
    if (cantidad === 0) continue;
    map.set(product, (map.get(product) || 0) + cantidad);
  }
  return [...map.entries()].map(([producto, cantidad]) => ({ producto, cantidad }));
}

function groupByProduct(records) {
  const map = new Map();
  for (const r of records) {
    const current = map.get(r.producto) || [];
    current.push(r);
    map.set(r.producto, current);
  }
  return map;
}

function calculateForecast({ stockRows, ventas, bajas, existencias, realProduction, days, weekendBoost, bufferPct }) {
  const ventasByProduct = groupByProduct(ventas);
  const bajasByProduct = groupByProduct(bajas);
  const existMap = new Map(existencias.map((e) => [e.producto, e]));
  const realMap = new Map(realProduction.map((e) => [e.producto, e.cantidad]));
  const horizonFactor = horizonWeekendFactor(days, weekendBoost);

  return stockRows.map((s) => {
    const v = ventasByProduct.get(s.producto) || [];
    const b = bajasByProduct.get(s.producto) || [];

    const values = v.map((x) => x.cantidad);
    const recentValues = values.slice(-28);
    const promedioReciente = recentValues.length ? recentValues.reduce((a, n) => a + n, 0) / recentValues.length : 0;
    const promedioHistorico = values.length ? values.reduce((a, n) => a + n, 0) / values.length : 0;
    const promedioDiario = promedioReciente * 0.65 + promedioHistorico * 0.35;
    const demandaPronosticada = promedioDiario * days * horizonFactor;

    const bajasTotal = b.reduce((a, n) => a + n.cantidad, 0);
    const ventasTotal = v.reduce((a, n) => a + n.cantidad, 0);
    const tasaBajas = ventasTotal > 0 ? bajasTotal / ventasTotal : 0;
    const bajasEsperadas = demandaPronosticada * tasaBajas;
    const colchonOperativo = Math.ceil((demandaPronosticada + bajasEsperadas) * bufferPct);
    const produccionPronosticada = Math.ceil(demandaPronosticada + bajasEsperadas + colchonOperativo);

    const ex = existMap.get(s.producto) || { totalSuc: 0, cf: 0, sumaSucCf: 0 };
    const sumaSucCf = ex.sumaSucCf || ex.totalSuc + ex.cf;
    const inventarioObjetivo = s.stock;
    const produccionRecomendada = Math.max(0, Math.ceil(inventarioObjetivo + produccionPronosticada - sumaSucCf));
    const produccionReal = realMap.get(s.producto) || 0;
    const diferenciaReal = produccionReal - produccionRecomendada;
    const cumplimiento =
      produccionRecomendada > 0 ? (produccionReal / produccionRecomendada) * 100 : produccionReal > 0 ? 100 : null;

    const confianza =
      promedioHistorico > 0
        ? Math.max(0, Math.min(100, 100 - (Math.abs(promedioReciente - promedioHistorico) / promedioHistorico) * 100))
        : values.length > 0
          ? 50
          : 0;

    let estatus = "Sin dato real";
    if (produccionRecomendada === 0) estatus = "No producir";
    else if (produccionReal === 0) estatus = "Sin dato real";
    else if (cumplimiento < 90) estatus = "Riesgo faltante";
    else if (cumplimiento > 115) estatus = "Sobreproduccion";
    else if (confianza < 65) estatus = "Revisar";
    else estatus = "Dentro de rango";

    return {
      producto: s.producto,
      orden: s.orden,
      promedioReciente,
      promedioHistorico,
      promedioDiario,
      demandaPronosticada,
      tasaBajas,
      bajasEsperadas,
      colchonOperativo,
      produccionPronosticada,
      inventarioObjetivo,
      totalSuc: ex.totalSuc || 0,
      cf: ex.cf || 0,
      sumaSucCf,
      produccionRecomendada,
      produccionReal,
      diferenciaReal,
      cumplimiento,
      confianza,
      estatus,
    };
  });
}

function exportToExcel(rows, summary) {
  const detalle = rows.map((r) => ({
    Producto: r.producto,
    "Promedio reciente": Number(r.promedioReciente.toFixed(2)),
    "Promedio historico": Number(r.promedioHistorico.toFixed(2)),
    "Demanda pronosticada": Number(r.demandaPronosticada.toFixed(2)),
    "Bajas esperadas": Number(r.bajasEsperadas.toFixed(2)),
    "Colchon operativo": r.colchonOperativo,
    "Produccion pronosticada": r.produccionPronosticada,
    "Inventario objetivo": r.inventarioObjetivo,
    "Existencia sucursales + CF": r.sumaSucCf,
    "Produccion recomendada": r.produccionRecomendada,
    "Produccion real": r.produccionReal,
    "Diferencia real vs recomendada": r.diferenciaReal,
    "Cumplimiento %": r.cumplimiento === null ? "" : Number(r.cumplimiento.toFixed(1)),
    Confianza: Number(r.confianza.toFixed(1)),
    Estatus: STATUS_META[r.estatus]?.label || r.estatus,
  }));

  const resumen = [
    { Indicador: "Produccion pronosticada con colchon", Valor: summary.totalPronosticada },
    { Indicador: "Produccion recomendada", Valor: summary.totalRecomendada },
    { Indicador: "Produccion real", Valor: summary.totalReal },
    { Indicador: "Brecha real vs recomendada", Valor: summary.brechaTotal },
    { Indicador: "Cumplimiento ejecutivo", Valor: `${summary.cumplimientoEjecutivo.toFixed(1)}%` },
    { Indicador: "Productos en riesgo", Valor: summary.riesgoFaltante },
    { Indicador: "Productos con sobreproduccion", Valor: summary.sobreproduccion },
  ];

  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(resumen), "Dashboard");
  XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(detalle), "Detalle");
  XLSX.writeFile(wb, "dashboard_produccion_archivo_maestro.xlsx");
}

function UploadBox({ title, description, onFile, fileName, required }) {
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
        <input type="file" accept=".xlsx,.xls" onChange={(e) => onFile(e.target.files?.[0])} />
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
  const [days, setDays] = useState(7);
  const [bufferPct, setBufferPct] = useState(0.12);

  async function handleFile(file, parser, key, setter) {
    if (!file) return;
    const wb = await readWorkbook(file);
    setter(parser(wb));
    setFiles((f) => ({ ...f, [key]: file.name }));
  }

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
        bufferPct,
      }),
    [stockRows, ventas, bajas, existencias, realProduction, days, weekendBoost, bufferPct]
  );

  const filtered = forecast.filter((r) => r.producto.includes(norm(query)));

  const summary = useMemo(() => {
    const totalPronosticada = forecast.reduce((a, r) => a + r.produccionPronosticada, 0);
    const totalRecomendada = forecast.reduce((a, r) => a + r.produccionRecomendada, 0);
    const totalReal = forecast.reduce((a, r) => a + r.produccionReal, 0);
    const totalColchon = forecast.reduce((a, r) => a + r.colchonOperativo, 0);
    const brechaTotal = totalReal - totalRecomendada;
    const cumplimientoEjecutivo = totalRecomendada > 0 ? (totalReal / totalRecomendada) * 100 : 0;
    const confianza = forecast.length ? forecast.reduce((a, r) => a + r.confianza, 0) / forecast.length : 0;
    const riesgoFaltante = forecast.filter((r) => r.estatus === "Riesgo faltante").length;
    const sobreproduccion = forecast.filter((r) => r.estatus === "Sobreproduccion").length;
    const sinDatoReal = forecast.filter((r) => r.estatus === "Sin dato real").length;
    return {
      totalPronosticada,
      totalRecomendada,
      totalReal,
      totalColchon,
      brechaTotal,
      cumplimientoEjecutivo,
      confianza,
      riesgoFaltante,
      sobreproduccion,
      sinDatoReal,
    };
  }, [forecast]);

  const topFaltantes = [...forecast]
    .filter((r) => r.diferenciaReal < 0)
    .sort((a, b) => a.diferenciaReal - b.diferenciaReal)
    .slice(0, 5);

  const topExcedentes = [...forecast]
    .filter((r) => r.diferenciaReal > 0)
    .sort((a, b) => b.diferenciaReal - a.diferenciaReal)
    .slice(0, 5);

  const chartRows = [
    { label: "Pronosticada", value: summary.totalPronosticada },
    { label: "Recomendada", value: summary.totalRecomendada },
    { label: "Real", value: summary.totalReal },
  ];
  const chartMax = Math.max(1, ...chartRows.map((r) => r.value));

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
          <span>Demanda por horizonte + bajas esperadas + colchón operativo.</span>
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
            <h2>Producción pronosticada vs real</h2>
            <p>Calcula colchón operativo, detecta brechas y prioriza productos con riesgo de faltante.</p>
          </div>
          <button className="primary" onClick={() => exportToExcel(filtered, summary)} disabled={!forecast.length}>
            <Download size={18} /> Exportar dashboard
          </button>
        </header>

        <section className="kpis">
          <KpiCard
            icon={ShieldCheck}
            label="Pronóstico con colchón"
            value={formatNumber(summary.totalPronosticada)}
            caption={`Colchón: ${formatNumber(summary.totalColchon)}`}
          />
          <KpiCard
            icon={Target}
            label="Producción recomendada"
            value={formatNumber(summary.totalRecomendada)}
            caption={`${days} días de horizonte`}
          />
          <KpiCard
            icon={Database}
            label="Producción real"
            value={formatNumber(summary.totalReal)}
            caption={`Brecha: ${formatNumber(summary.brechaTotal)}`}
            tone={summary.brechaTotal < 0 ? "danger" : summary.brechaTotal > 0 ? "warn" : "ok"}
          />
          <KpiCard
            icon={CheckCircle2}
            label="Cumplimiento"
            value={formatPercent(summary.cumplimientoEjecutivo, 1)}
            caption={`${formatNumber(summary.confianza)}% confianza promedio`}
            tone={summary.cumplimientoEjecutivo < 90 ? "danger" : summary.cumplimientoEjecutivo > 115 ? "warn" : "ok"}
          />
        </section>

        <section className="executive-grid">
          <div className="panel">
            <div className="panel-title">
              <div>
                <h3>Lectura ejecutiva</h3>
                <p>{summary.sinDatoReal ? "Carga producción real para completar la comparación." : "Comparativo consolidado contra producción real."}</p>
              </div>
              <BarChart3 size={22} />
            </div>
            <div className="bars">
              {chartRows.map((row) => (
                <div className="bar-row" key={row.label}>
                  <span>{row.label}</span>
                  <div className="bar-track">
                    <div className={`bar-fill ${row.label.toLowerCase()}`} style={{ width: `${(row.value / chartMax) * 100}%` }} />
                  </div>
                  <strong>{formatNumber(row.value)}</strong>
                </div>
              ))}
            </div>
          </div>

          <div className="panel status-panel">
            <div className="metric-line danger">
              <AlertTriangle size={20} />
              <div>
                <strong>{summary.riesgoFaltante}</strong>
                <span>productos con riesgo de faltante</span>
              </div>
            </div>
            <div className="metric-line warn">
              <RefreshCw size={20} />
              <div>
                <strong>{summary.sobreproduccion}</strong>
                <span>productos con sobreproducción</span>
              </div>
            </div>
            <div className="metric-line muted">
              <FileSpreadsheet size={20} />
              <div>
                <strong>{summary.sinDatoReal}</strong>
                <span>productos sin dato real</span>
              </div>
            </div>
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
            description="Archivo real producido por producto para comparar."
            onFile={(file) => handleFile(file, parseProductionReal, "real", setRealProduction)}
            fileName={files.real}
          />
        </section>

        <section className="controls">
          <div className="search">
            <Search size={18} />
            <input placeholder="Buscar producto..." value={query} onChange={(e) => setQuery(e.target.value)} />
          </div>
          <label>
            Horizonte
            <input min="1" type="number" value={days} onChange={(e) => setDays(Math.max(1, Number(e.target.value)))} />
          </label>
          <label>
            Colchón %
            <input
              min="0"
              max="100"
              type="number"
              value={Math.round(bufferPct * 100)}
              onChange={(e) => setBufferPct(Math.max(0, Number(e.target.value)) / 100)}
            />
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
        </section>

        <section className="priority-grid">
          <div className="panel">
            <h3>Mayores faltantes</h3>
            {topFaltantes.length ? (
              topFaltantes.map((item) => (
                <div className="priority-row" key={item.producto}>
                  <span>{item.producto}</span>
                  <strong>{formatNumber(item.diferenciaReal)}</strong>
                </div>
              ))
            ) : (
              <p className="soft">Sin faltantes detectados con los datos actuales.</p>
            )}
          </div>
          <div className="panel">
            <h3>Mayores excedentes</h3>
            {topExcedentes.length ? (
              topExcedentes.map((item) => (
                <div className="priority-row" key={item.producto}>
                  <span>{item.producto}</span>
                  <strong>+{formatNumber(item.diferenciaReal)}</strong>
                </div>
              ))
            ) : (
              <p className="soft">Sin excedentes detectados con los datos actuales.</p>
            )}
          </div>
        </section>

        <section className="table-card">
          <table>
            <thead>
              <tr>
                <th>Producto</th>
                <th>Prom. reciente</th>
                <th>Demanda pron.</th>
                <th>Bajas esp.</th>
                <th>Colchón</th>
                <th>Pron. con colchón</th>
                <th>Stock objetivo</th>
                <th>Existencias</th>
                <th>Prod. recomendada</th>
                <th>Prod. real</th>
                <th>Brecha</th>
                <th>Cumpl.</th>
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
                    <td>{r.demandaPronosticada.toFixed(1)}</td>
                    <td>{r.bajasEsperadas.toFixed(1)}</td>
                    <td>{r.colchonOperativo}</td>
                    <td>{r.produccionPronosticada}</td>
                    <td>{r.inventarioObjetivo}</td>
                    <td>{r.sumaSucCf}</td>
                    <td className="strong">{r.produccionRecomendada}</td>
                    <td>{r.produccionReal}</td>
                    <td className={r.diferenciaReal < 0 ? "negative" : r.diferenciaReal > 0 ? "positive" : ""}>
                      {r.diferenciaReal > 0 ? "+" : ""}
                      {formatNumber(r.diferenciaReal)}
                    </td>
                    <td>{r.cumplimiento === null ? "-" : formatPercent(r.cumplimiento, 0)}</td>
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
        </section>
      </main>
    </div>
  );
}

createRoot(document.getElementById("root")).render(<App />);
