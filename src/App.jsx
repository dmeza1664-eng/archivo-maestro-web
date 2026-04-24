
import React, { useMemo, useState } from "react";
import { createRoot } from "react-dom/client";
import * as XLSX from "xlsx";
import {
  Upload,
  FileSpreadsheet,
  RefreshCw,
  Download,
  Search,
  PackageCheck,
  AlertTriangle,
  CheckCircle2,
  BarChart3,
  Database,
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
  const n = Number(value);
  return Number.isFinite(n) ? n : 0;
}

function getWeekday(date) {
  const d = new Date(date);
  const map = ["DOMINGO", "LUNES", "MARTES", "MIERCOLES", "JUEVES", "VIERNES", "SABADO"];
  return map[d.getDay()];
}

function excelDateToJSDate(serial) {
  if (serial instanceof Date) return serial;
  if (typeof serial === "number") {
    const utcDays = Math.floor(serial - 25569);
    const utcValue = utcDays * 86400;
    return new Date(utcValue * 1000);
  }
  const d = new Date(serial);
  return isNaN(d.getTime()) ? null : d;
}

async function readWorkbook(file) {
  const data = await file.arrayBuffer();
  return XLSX.read(data, { type: "array", cellDates: true });
}

function parseStock(workbook) {
  const sheet = workbook.Sheets[workbook.SheetNames[0]];
  const rows = XLSX.utils.sheet_to_json(sheet, { header: 1, defval: "" });
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
  const sheet = workbook.Sheets[workbook.SheetNames[0]];
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

  if (headerIndex < 0) {
    return [];
  }

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

    let fecha = null;
    const parsedDate = new Date(sheetName);
    if (!isNaN(parsedDate.getTime())) fecha = parsedDate;
    const dayOnly = String(sheetName).match(/\d{1,2}/);
    if (!fecha && dayOnly) {
      fecha = new Date(2026, 0, Number(dayOnly[0]));
    }
    if (!fecha) fecha = new Date();

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
  const sheet = workbook.Sheets[workbook.SheetNames[0]];
  const rows = XLSX.utils.sheet_to_json(sheet, { header: 1, defval: "" });
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

function groupByProduct(records) {
  const map = new Map();
  for (const r of records) {
    const current = map.get(r.producto) || [];
    current.push(r);
    map.set(r.producto, current);
  }
  return map;
}

function calculateForecast({ stockRows, ventas, bajas, existencias, days, weekendBoost }) {
  const ventasByProduct = groupByProduct(ventas);
  const bajasByProduct = groupByProduct(bajas);
  const existMap = new Map(existencias.map((e) => [e.producto, e]));

  return stockRows.map((s) => {
    const v = ventasByProduct.get(s.producto) || [];
    const b = bajasByProduct.get(s.producto) || [];

    const values = v.map((x) => x.cantidad);
    const recentValues = values.slice(-28);
    const promedioReciente = recentValues.length ? recentValues.reduce((a, n) => a + n, 0) / recentValues.length : 0;
    const promedioHistorico = values.length ? values.reduce((a, n) => a + n, 0) / values.length : 0;

    let pronosticoBruto = promedioReciente * 0.6 + promedioHistorico * 0.4;
    const today = new Date();
    const isWeekend = [0, 6].includes(today.getDay());
    if (isWeekend) pronosticoBruto *= weekendBoost;

    const bajasTotal = b.reduce((a, n) => a + n.cantidad, 0);
    const ventasTotal = v.reduce((a, n) => a + n.cantidad, 0);
    const tasaBajas = ventasTotal > 0 ? bajasTotal / ventasTotal : 0;

    const produccionSugerida = Math.ceil(pronosticoBruto);
    const ex = existMap.get(s.producto) || { totalSuc: 0, cf: 0, sumaSucCf: 0 };
    const sumaSucCf = ex.sumaSucCf || ex.totalSuc + ex.cf;
    const produccionBalanceada = Math.ceil((s.stock - sumaSucCf + produccionSugerida) / 2);

    const confianza =
      promedioHistorico > 0
        ? Math.max(0, Math.min(100, 100 - (Math.abs(promedioReciente - promedioHistorico) / promedioHistorico) * 100))
        : 0;

    return {
      producto: s.producto,
      orden: s.orden,
      promedioReciente,
      promedioHistorico,
      pronosticoBruto,
      bajasEsperadas: pronosticoBruto * tasaBajas,
      produccionSugerida,
      stock: s.stock,
      totalSuc: ex.totalSuc || 0,
      cf: ex.cf || 0,
      sumaSucCf,
      produccionBalanceada,
      confianza,
      estatus: produccionBalanceada <= 0 ? "No producir" : confianza < 75 ? "Revisar" : "Normal",
    };
  });
}

function exportToExcel(rows) {
  const ws = XLSX.utils.json_to_sheet(rows);
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, "Pronostico");
  XLSX.writeFile(wb, "pronostico_archivo_maestro.xlsx");
}

function UploadBox({ title, description, onFile, fileName }) {
  return (
    <div className="upload-card">
      <div className="upload-icon"><Upload size={22} /></div>
      <h3>{title}</h3>
      <p>{description}</p>
      <label className="upload-button">
        Seleccionar Excel
        <input type="file" accept=".xlsx,.xls" onChange={(e) => onFile(e.target.files?.[0])} />
      </label>
      {fileName && <span className="file-name">{fileName}</span>}
    </div>
  );
}

function App() {
  const [stockRows, setStockRows] = useState([]);
  const [ventas, setVentas] = useState([]);
  const [bajas, setBajas] = useState([]);
  const [existencias, setExistencias] = useState([]);
  const [files, setFiles] = useState({});
  const [query, setQuery] = useState("");
  const [weekendBoost, setWeekendBoost] = useState(1.15);
  const [days, setDays] = useState(7);

  async function handleStock(file) {
    if (!file) return;
    const wb = await readWorkbook(file);
    setStockRows(parseStock(wb));
    setFiles((f) => ({ ...f, stock: file.name }));
  }

  async function handleVentas(file) {
    if (!file) return;
    const wb = await readWorkbook(file);
    setVentas(parseSalesOrReturns(wb, "ventas"));
    setFiles((f) => ({ ...f, ventas: file.name }));
  }

  async function handleBajas(file) {
    if (!file) return;
    const wb = await readWorkbook(file);
    setBajas(parseSalesOrReturns(wb, "bajas"));
    setFiles((f) => ({ ...f, bajas: file.name }));
  }

  async function handleExistencias(file) {
    if (!file) return;
    const wb = await readWorkbook(file);
    setExistencias(parseExistencias(wb));
    setFiles((f) => ({ ...f, existencias: file.name }));
  }

  const forecast = useMemo(
    () => calculateForecast({ stockRows, ventas, bajas, existencias, days, weekendBoost }),
    [stockRows, ventas, bajas, existencias, days, weekendBoost]
  );

  const filtered = forecast.filter((r) => r.producto.includes(norm(query)));

  const totalSugerida = forecast.reduce((a, r) => a + r.produccionSugerida, 0);
  const totalBalanceada = forecast.reduce((a, r) => a + Math.max(0, r.produccionBalanceada), 0);
  const confianza = forecast.length ? forecast.reduce((a, r) => a + r.confianza, 0) / forecast.length : 0;

  return (
    <div className="app">
      <aside className="sidebar">
        <div className="brand">
          <div className="brand-icon"><PackageCheck /></div>
          <div>
            <h1>Archivo Maestro</h1>
            <p>Web real con Excel</p>
          </div>
        </div>
        <div className="formula">
          <strong>Fórmula clave</strong>
          <span>Producción balanceada = (Stock - Suma Suc. y C.F. + Producción sugerida) / 2</span>
        </div>
      </aside>

      <main className="main">
        <header className="top">
          <div>
            <h2>Sistema PRO conectado a Excel</h2>
            <p>Carga tus archivos y calcula pronóstico + producción balanceada.</p>
          </div>
          <button className="primary" onClick={() => exportToExcel(filtered)}>
            <Download size={18} /> Exportar Excel
          </button>
        </header>

        <section className="grid kpis">
          <div className="card"><BarChart3 /><span>Producción sugerida</span><strong>{totalSugerida}</strong></div>
          <div className="card"><Database /><span>Producción balanceada</span><strong>{totalBalanceada}</strong></div>
          <div className="card"><CheckCircle2 /><span>Confianza promedio</span><strong>{confianza.toFixed(0)}%</strong></div>
          <div className="card"><AlertTriangle /><span>Productos</span><strong>{forecast.length}</strong></div>
        </section>

        <section className="uploads">
          <UploadBox title="Stock fijo" description="Productos oficiales, stock y orden." onFile={handleStock} fileName={files.stock} />
          <UploadBox title="Ventas" description="Ventas históricas por día y producto." onFile={handleVentas} fileName={files.ventas} />
          <UploadBox title="Bajas" description="Devoluciones o bajas por producto." onFile={handleBajas} fileName={files.bajas} />
          <UploadBox title="Existencias" description="TOTAL SUC., C.F. y SUMA SUC. Y C.F." onFile={handleExistencias} fileName={files.existencias} />
        </section>

        <section className="controls">
          <div className="search">
            <Search size={18} />
            <input placeholder="Buscar producto..." value={query} onChange={(e) => setQuery(e.target.value)} />
          </div>
          <label>
            Factor fin semana
            <input type="number" step="0.01" value={weekendBoost} onChange={(e) => setWeekendBoost(Number(e.target.value))} />
          </label>
          <label>
            Días
            <input type="number" value={days} onChange={(e) => setDays(Number(e.target.value))} />
          </label>
          <button className="secondary"><RefreshCw size={18} /> Recalcular</button>
        </section>

        <section className="table-card">
          <table>
            <thead>
              <tr>
                <th>Producto</th>
                <th>Prom. reciente</th>
                <th>Prom. histórico</th>
                <th>Pronóstico bruto</th>
                <th>Prod. sugerida</th>
                <th>Stock</th>
                <th>Total suc.</th>
                <th>C.F.</th>
                <th>Suma suc. y C.F.</th>
                <th>Prod. balanceada</th>
                <th>Confianza</th>
                <th>Estatus</th>
              </tr>
            </thead>
            <tbody>
              {filtered.map((r) => (
                <tr key={r.producto}>
                  <td>{r.producto}</td>
                  <td>{r.promedioReciente.toFixed(2)}</td>
                  <td>{r.promedioHistorico.toFixed(2)}</td>
                  <td>{r.pronosticoBruto.toFixed(2)}</td>
                  <td>{r.produccionSugerida}</td>
                  <td>{r.stock}</td>
                  <td>{r.totalSuc}</td>
                  <td>{r.cf}</td>
                  <td>{r.sumaSucCf}</td>
                  <td className="strong">{Math.max(0, r.produccionBalanceada)}</td>
                  <td>{r.confianza.toFixed(0)}%</td>
                  <td><span className={`pill ${r.estatus === "Revisar" ? "warn" : r.estatus === "No producir" ? "muted" : "ok"}`}>{r.estatus}</span></td>
                </tr>
              ))}
            </tbody>
          </table>
          {!forecast.length && (
            <div className="empty">
              Carga primero el archivo de <strong>stock fijo</strong> y después ventas/existencias para ver resultados.
            </div>
          )}
        </section>
      </main>
    </div>
  );
}

createRoot(document.getElementById("root")).render(<App />);
