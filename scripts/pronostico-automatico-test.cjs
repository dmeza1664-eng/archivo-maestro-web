/**
 * Pronóstico automático (scripts/pronostico-automatico.cjs).
 * DATOS SINTÉTICOS: el fixture sintético del repo repartido en sucursales
 * inventadas, más un canal "Planta León" y filas de Suc. Amado Nervo antes de
 * ago-2024 inventadas. No son ventas reales de Pastelería Pepes. La "base" es
 * una simulación en memoria de las tablas dbo (no hay conexión real).
 *
 * Protege:
 *  - modo archivos: la serie oficial = venta de sucursales (sin Planta, sin
 *    Amado Nervo antes de ago-2024, sin productos fuera del catálogo); el
 *    pronóstico y el error walk-forward son idénticos a llamar al modelo
 *    directo (la automatización no cambia resultados); fecha de corte y mes
 *    incompleto; archivo anual gana sobre parcial; salidas CSV/MD y log;
 *  - modo base: mismas salidas que modo archivos con los mismos datos; solo
 *    consultas SELECT a las 7 tablas dbo; ventas canceladas fuera; Produccion
 *    se relee completa (una orden que pasa de Peticion a Finalizado cambia el
 *    cruce); variables faltantes; la contraseña nunca aparece en log ni salidas;
 *    la copia "extracto/" de modo base vuelve a correr igual en modo archivos.
 */
const fs = require("fs");
const os = require("os");
const path = require("path");
const auto = require("./pronostico-automatico.cjs");
const datos = require("./lib/pepes-datos.cjs");
const base = require("./lib/pepes-fuente-base.cjs");

const FIXTURE = JSON.parse(fs.readFileSync(path.join(__dirname, "fixtures", "sintetico-pronostico-2024-2025.json"), "utf8"));
const PASSWORD = "Clave-Sintetica-123!";

function assert(condition, message) {
  if (!condition) throw new Error(`[pronostico-automatico] ${message}`);
}

const tmp = fs.mkdtempSync(path.join(os.tmpdir(), "pronostico-auto-"));
const productos = FIXTURE.stock.map((s) => s.producto);
const pkDe = (p) => String(1000 + productos.indexOf(p));

// --- Datos sintéticos por mes × producto × sucursal -------------------------
function filasSinteticas() {
  const ventas = [];
  for (const [producto, meses] of Object.entries(FIXTURE.cierresMensuales)) {
    for (const [mes, qty] of Object.entries(meses)) {
      const centro = Math.round(qty * 0.6);
      ventas.push({ mes, pk: pkDe(producto), producto, suc: "Suc. Centro", sucPK: "1", bod: "11", qty: centro });
      ventas.push({ mes, pk: pkDe(producto), producto, suc: "Suc. Norte", sucPK: "2", bod: "21", qty: qty - centro });
      ventas.push({ mes, pk: pkDe(producto), producto, suc: "Planta León", sucPK: "9", bod: "91", qty: Math.round(qty * 1.3) });
      if (mes < "2024-08") ventas.push({ mes, pk: pkDe(producto), producto, suc: "Suc. Amado Nervo", sucPK: "3", bod: "31", qty: 777 });
    }
  }
  for (const mes of ["2024-06", "2025-06"]) ventas.push({ mes, pk: "5000", producto: "VELA MAGICA", suc: "Suc. Centro", sucPK: "1", bod: "11", qty: 900 }); // fuera del catálogo
  // Enero 2026 parcial (hasta el día 15) + un PK nuevo que se asigna por código exacto.
  ventas.push({ mes: "2026-01", pk: pkDe("MOKA GDE"), producto: "MOKA GDE", suc: "Suc. Centro", sucPK: "1", bod: "11", qty: 300 });
  ventas.push({ mes: "2026-01", pk: "7001", producto: "MOKA GDE NUEVO PK", codigo: "C-MOKA", suc: "Suc. Norte", sucPK: "2", bod: "21", qty: 25 });
  ventas.push({ mes: "2026-01", pk: pkDe("MOKA GDE"), producto: "MOKA GDE", suc: "Planta León", sucPK: "9", bod: "91", qty: 999 });
  return ventas;
}

function produccionSintetica() {
  const out = [];
  for (const [producto, meses] of Object.entries(FIXTURE.cierresMensuales)) {
    for (const [mes, qty] of Object.entries(meses)) out.push({ mes, pk: pkDe(producto), producto, producida: Math.round(qty * 1.05), pedida: Math.round(qty * 1.1) });
  }
  out.push({ mes: "2026-01", pk: pkDe("MOKA GDE"), producto: "MOKA GDE", producida: 310, pedida: 320 });
  return out;
}

const esc = (v) => datos.csvCell(v);
function escribirCsv(file, cols, filas) {
  fs.writeFileSync(file, [cols.join(","), ...filas.map((r) => cols.map((c) => esc(r[c])).join(","))].join("\n") + "\n");
}

function prepararCatalogo(dir) {
  fs.mkdirSync(dir, { recursive: true });
  const filas = productos.map((p) => ({ nProductoPK: pkDe(p), cCodigo: `C${pkDe(p)}`, cDescripcion: p, qtyTotal2025: 1, producto: p, method: "exact_code", confidence: 1, inCatalog: "true", promotional: "false", slice: "false", inWapeUniverse: "true", notes: "" }));
  filas.push({ nProductoPK: "5000", cCodigo: "152", cDescripcion: "VELA MAGICA", qtyTotal2025: 1, producto: "", method: "no_match", confidence: 0, inCatalog: "false", promotional: "false", slice: "false", inWapeUniverse: "false", notes: "" });
  escribirCsv(path.join(dir, "mapeo.csv"), ["nProductoPK", "cCodigo", "cDescripcion", "qtyTotal2025", "producto", "method", "confidence", "inCatalog", "promotional", "slice", "inWapeUniverse", "notes"], filas);
  fs.writeFileSync(path.join(dir, "codigos.json"), JSON.stringify({ "C-MOKA": { codigo: "C-MOKA", producto: "MOKA GDE", inCatalog: true } }));
  fs.writeFileSync(path.join(dir, "stock.json"), JSON.stringify(FIXTURE.stock));
  fs.writeFileSync(path.join(dir, "referencia.json"), JSON.stringify({ modelo: "sintetico", cortes: {} }));
}

function prepararArchivos(dir) {
  fs.mkdirSync(dir, { recursive: true });
  const v = filasSinteticas().map((r) => ({ mes: r.mes, nProductoPK: r.pk, cCodigo: r.codigo || `C${r.pk}`, cDescripcion: r.producto, nBodegaPK: r.bod, bodegaNombre: "Piso de venta", nSucursalPK: r.sucPK, sucursalNombre: r.suc, qty: r.qty, subtotal: 0, total: 0, lineas: 1 }));
  const p = produccionSintetica().map((r) => ({ mes: r.mes, nProductoPK: r.pk, cCodigo: `C${r.pk}`, cDescripcion: r.producto, qtyPedida: r.pedida, qtyProducida: r.producida, qtyCancelada: 0, qtySalida: r.producida, lineas: 1 }));
  for (const y of ["2024", "2025"]) {
    escribirCsv(path.join(dir, `ventas-${y}-mensual-producto-bodega.csv`), datos.VENTAS_COLUMNAS, v.filter((r) => r.mes.startsWith(y)));
    escribirCsv(path.join(dir, `produccion-${y}-mensual-producto.csv`), datos.PRODUCCION_COLUMNAS, p.filter((r) => r.mes.startsWith(y)));
  }
  escribirCsv(path.join(dir, "part-ventas-bodega-2026-01.csv"), datos.VENTAS_COLUMNAS, v.filter((r) => r.mes === "2026-01"));
  escribirCsv(path.join(dir, "part-produccion-2026-01.csv"), datos.PRODUCCION_COLUMNAS, p.filter((r) => r.mes === "2026-01"));
  // Parcial de un mes que ya está en el archivo anual: debe ignorarse (si se sumara, duplicaría marzo).
  escribirCsv(path.join(dir, "part-ventas-bodega-2025-03.csv"), datos.VENTAS_COLUMNAS, v.filter((r) => r.mes === "2025-03").map((r) => ({ ...r, qty: 99999 })));
  fs.writeFileSync(path.join(dir, "extract-meta.json"), JSON.stringify({ database: "sintetica", maxVentaDCreada: "2026-01-15T19:48:39.593Z" }));
}

// --- Base sintética en memoria (tablas dbo) y ejecutor simulado ---------------
function crearBaseSintetica() {
  const Sucursales = [{ nSucursalPK: 1, cNombre: "Suc. Centro" }, { nSucursalPK: 2, cNombre: "Suc. Norte" }, { nSucursalPK: 3, cNombre: "Suc. Amado Nervo" }, { nSucursalPK: 9, cNombre: "Planta León" }];
  const Bodegas = [{ nBodegaPK: 11, nSucursalPK: 1, cNombre: "Piso de venta" }, { nBodegaPK: 21, nSucursalPK: 2, cNombre: "Piso de venta" }, { nBodegaPK: 31, nSucursalPK: 3, cNombre: "Piso de venta" }, { nBodegaPK: 91, nSucursalPK: 9, cNombre: "Piso de venta" }];
  const Productos = new Map();
  const Ventas = [];
  const VentaDet = [];
  let vpk = 1;
  for (const r of filasSinteticas()) {
    Productos.set(Number(r.pk), { nProductoPK: Number(r.pk), cCodigo: r.codigo || `C${r.pk}`, cDescripcion: r.producto });
    // Dos tickets por fila (día 5 y día 12) + un ticket cancelado y una línea cancelada que no deben contar.
    const a = Math.floor(r.qty / 2);
    for (const [dia, q] of [["05", a], ["12", r.qty - a]]) {
      Ventas.push({ nVentaPK: vpk, dCreada: `${r.mes}-${dia}T13:30:00`, nBodegaPK: Number(r.bod), dCancelada: null });
      VentaDet.push({ nVentaPK: vpk, nProductoPK: Number(r.pk), nCantidad: q, dCancelado: null, mSubTotal: q * 10, mTotal: q * 11.6, nPromocionPK: null });
      vpk += 1;
    }
    Ventas.push({ nVentaPK: vpk, dCreada: `${r.mes}-07T10:00:00`, nBodegaPK: Number(r.bod), dCancelada: `${r.mes}-07T10:05:00` });
    VentaDet.push({ nVentaPK: vpk, nProductoPK: Number(r.pk), nCantidad: 5000, dCancelado: null, mSubTotal: 0, mTotal: 0, nPromocionPK: null });
    VentaDet.push({ nVentaPK: vpk - 1, nProductoPK: Number(r.pk), nCantidad: 4000, dCancelado: `${r.mes}-12T14:00:00`, mSubTotal: 0, mTotal: 0, nPromocionPK: null });
    vpk += 1;
  }
  // La última venta: 2026-01-15 19:48.
  Ventas.push({ nVentaPK: vpk, dCreada: "2026-01-15T19:48:39", nBodegaPK: 11, dCancelada: null });
  VentaDet.push({ nVentaPK: vpk, nProductoPK: 5000, nCantidad: 1, dCancelado: null, mSubTotal: 0, mTotal: 0, nPromocionPK: null });
  Productos.set(5000, { nProductoPK: 5000, cCodigo: "152", cDescripcion: "VELA MAGICA" });
  const Produccion = [];
  const ProduccionDet = [];
  let ppk = 1;
  for (const r of produccionSintetica()) {
    // Pedida a fin del mes anterior y entregada en el mes: el mes de la orden es el de entrega.
    const peticion = r.mes < "2026-01" ? `${datos.siguienteMes(r.mes, -1)}-28T08:00:00` : `${r.mes}-10T08:00:00`;
    Produccion.push({ nProduccionPK: ppk, dFechaPeticion: peticion, dFechaEntrega: `${r.mes}-03T08:00:00`, cEstado: "Finalizado" });
    ProduccionDet.push({ nProduccionPK: ppk, nProductoPK: Number(r.pk), nCantidad: r.pedida, nProducidos: r.producida, nCancelados: 0, nSalida: r.producida, cEstado: "Finalizado" });
    ppk += 1;
  }
  return { Sucursales, Bodegas, Productos, Ventas, VentaDet, Produccion, ProduccionDet, db: "pepes_devBI_espejo" };
}

function crearEjecutorSimulado(bd, registro) {
  const aDate = (s) => new Date(`${s}Z`);
  const tag = (sql) => (String(sql).match(/\/\* pronostico:([a-z_]+) \*\//) || [])[1];
  return async () => ({
    async query(sql, params = {}) {
      base.asegurarSoloLectura(sql); // el simulador también rechaza escrituras
      registro.push({ tag: tag(sql), params });
      const ventaDe = new Map(bd.Ventas.map((v) => [v.nVentaPK, v]));
      const bodega = new Map(bd.Bodegas.map((b) => [b.nBodegaPK, b]));
      const suc = new Map(bd.Sucursales.map((s) => [s.nSucursalPK, s]));
      switch (tag(sql)) {
        case "base_actual": return [{ db: bd.db }];
        case "max_venta": {
          const vivas = bd.Ventas.filter((v) => !v.dCancelada && v.dCreada < params.limite).map((v) => v.dCreada).sort();
          return [{ maxVenta: vivas.length ? aDate(vivas[vivas.length - 1]) : null }];
        }
        case "ventas_mes_bodega": {
          // El simulador aplica los filtros de cancelación solo si la consulta los trae,
          // así que quitar un filtro del SQL rompe la igualdad con los extractos.
          const sinCabCancelada = /\bv\.dCancelada IS NULL\b/.test(sql);
          const sinLineaCancelada = /\bvd\.dCancelado IS NULL\b/.test(sql);
          const g = new Map();
          for (const d of bd.VentaDet) {
            const v = ventaDe.get(d.nVentaPK);
            if (!v || (sinCabCancelada && v.dCancelada) || (sinLineaCancelada && d.dCancelado) || !(v.dCreada >= params.desde && v.dCreada < params.hasta)) continue;
            const b = bodega.get(v.nBodegaPK); const s = suc.get(b.nSucursalPK); const p = bd.Productos.get(d.nProductoPK);
            const k = [v.dCreada.slice(0, 7), d.nProductoPK, b.nBodegaPK].join("|");
            const cur = g.get(k) || { mes: v.dCreada.slice(0, 7), nProductoPK: d.nProductoPK, cCodigo: p.cCodigo, cDescripcion: p.cDescripcion, nBodegaPK: b.nBodegaPK, bodegaNombre: b.cNombre, nSucursalPK: s.nSucursalPK, sucursalNombre: s.cNombre, qty: 0, subtotal: 0, total: 0, lineas: 0, qtyPromocion: 0 };
            cur.qty += d.nCantidad; cur.subtotal += d.mSubTotal; cur.total += d.mTotal; cur.lineas += 1; if (d.nPromocionPK) cur.qtyPromocion += d.nCantidad;
            g.set(k, cur);
          }
          return [...g.values()];
        }
        case "produccion_mes_producto": {
          const cab = new Map(bd.Produccion.map((p) => [p.nProduccionPK, p]));
          const porEntrega = /COALESCE\(p\.dFechaEntrega,\s*p\.dFechaPeticion\)/.test(sql);
          const g = new Map();
          for (const d of bd.ProduccionDet) {
            const c = cab.get(d.nProduccionPK); const fecha = porEntrega ? (c.dFechaEntrega || c.dFechaPeticion) : c.dFechaPeticion;
            if (!(fecha >= params.desde)) continue;
            const p = bd.Productos.get(d.nProductoPK);
            const k = `${fecha.slice(0, 7)}|${d.nProductoPK}`;
            const cur = g.get(k) || { mes: fecha.slice(0, 7), nProductoPK: d.nProductoPK, cCodigo: p.cCodigo, cDescripcion: p.cDescripcion, qtyPedida: 0, qtyProducida: 0, qtyCancelada: 0, qtySalida: 0, lineas: 0, lineasAbiertas: 0 };
            cur.qtyPedida += d.nCantidad; cur.qtyProducida += d.nProducidos; cur.qtyCancelada += d.nCancelados || 0; cur.qtySalida += d.nSalida || 0; cur.lineas += 1;
            if (!["Finalizado", "Cancelado"].includes(c.cEstado)) cur.lineasAbiertas += 1;
            g.set(k, cur);
          }
          return [...g.values()];
        }
        case "produccion_estado": {
          const sel = bd.Produccion.filter((p) => (p.dFechaEntrega || p.dFechaPeticion) >= params.desde);
          const maxP = sel.map((p) => p.dFechaPeticion).sort().pop();
          return [{ ordenes: sel.length, ordenesAbiertas: sel.filter((p) => !["Finalizado", "Cancelado"].includes(p.cEstado)).length, maxPeticion: maxP ? aDate(maxP) : null }];
        }
        default: throw new Error(`consulta desconocida para la base sintética: ${sql.slice(0, 60)}`);
      }
    },
    async close() { registro.push({ tag: "close" }); },
  });
}

function leerSalida(dir, nombre) {
  return fs.readFileSync(path.join(dir, nombre), "utf8");
}
function csvDe(dir, nombre) {
  return datos.parseCsv(leerSalida(dir, nombre));
}
function sinVolatiles(resumen) {
  const { generado, fuente, modo, produccionEstado, advertencias, codigo, ...resto } = resumen;
  return resto;
}

(async () => {
  const app = await auto.cargarApp();
  const cat = path.join(tmp, "catalogo");
  prepararCatalogo(cat);
  const envBase = {
    PEPES_SILENCIOSO: "1",
    PEPES_ENV_FILE: path.join(tmp, "no-existe.env"),
    PEPES_STOCK_IDEAL: path.join(cat, "stock.json"),
    PEPES_MAPEO_PRODUCTOS: path.join(cat, "mapeo.csv"),
    PEPES_MAPEO_CODIGOS: path.join(cat, "codigos.json"),
    PEPES_REFERENCIA_METRICAS: path.join(cat, "referencia.json"),
  };

  // --- Lector de CSV ---
  const parsed = datos.parseCsv('a,b,c\n1,"x, ""y""",3\r\n2,"línea\nnueva",4\n');
  assert(parsed.length === 2 && parsed[0].b === 'x, "y"' && parsed[1].b === "línea\nnueva" && parsed[1].c === "4", "CSV con comas, comillas y saltos de línea");
  assert(datos.ultimoMesCompleto("2026-09-20") === "2026-08" && datos.ultimoMesCompleto("2026-09-30") === "2026-09" && datos.ultimoMesCompleto("2024-02-29") === "2024-02", "último mes completo según la fecha de corte");

  // --- Modo archivos ---
  const dirArchivos = path.join(tmp, "extractos");
  prepararArchivos(dirArchivos);
  const r1 = await auto.ejecutar({ app, env: { ...envBase, PEPES_FUENTE: "archivos", PEPES_DATOS_DIRS: dirArchivos, PEPES_SALIDA_DIR: path.join(tmp, "salida-archivos") }, ahora: new Date(2026, 0, 16, 7, 0, 0) });
  assert(r1.resumen.fechaCorte === "2026-01-15" && r1.resumen.ultimoMesCompleto === "2025-12", "fecha de corte desde extract-meta y último mes completo");
  assert(r1.resumen.mesesObjetivo.join() === "2026-01", "por defecto pronostica el mes siguiente al último mes completo");
  const serie = JSON.parse(leerSalida(r1.dir, "serie-sucursal.json")).rows;
  const esperado = [];
  for (const [producto, meses] of Object.entries(FIXTURE.cierresMensuales)) for (const [mes, qty] of Object.entries(meses)) esperado.push(`${mes}|${producto}|${qty}`);
  assert(JSON.stringify(serie.map((r) => `${r.mes}|${r.producto}|${r.cantidad}`).sort()) === JSON.stringify(esperado.sort()),
    "la serie oficial es la venta de sucursales: sin Planta León, sin Amado Nervo antes de ago-2024, sin productos fuera del catálogo y sin duplicar el parcial de marzo");
  // El modelo no cambia: pronóstico y walk-forward idénticos a llamar calculateForecast directo.
  const limpia = datos.hidratarSerie(serie);
  const directo = app.calculateForecast({ stockRows: FIXTURE.stock, historicalVentas: app.filterVentasBeforeMonth(limpia, "2026-01"), bajas: [], existencias: [], realProduction: [], selectedMonth: "2026-01", dailyBufferPct: 10 });
  const pron = csvDe(r1.dir, "pronostico-2026-01.csv");
  assert(pron.length === directo.length && pron.length === 6, "un renglón por producto del catálogo");
  for (const d of directo) {
    const f = pron.find((p) => p.producto === d.producto);
    assert(f && Number(f.pronostico) === +d.pronosticoVenta.toFixed(2) && Number(f.plan_modelo) === d.produccionSugerida, `${d.producto}: pronóstico idéntico al modelo`);
  }
  const moka = pron.find((p) => p.producto === "MOKA GDE");
  assert(Number(moka.venta_al_corte) === 325 && Number(moka.produccion_al_corte) === 310, "avance del mes en curso: venta de sucursales (incluye el PK nuevo asignado por código, sin Planta) y producción");
  let abs = 0; let act = 0;
  for (let m = 1; m <= 12; m += 1) {
    const mes = `2025-${String(m).padStart(2, "0")}`;
    const fr = app.calculateForecast({ stockRows: FIXTURE.stock, historicalVentas: app.filterVentasBeforeMonth(limpia, mes), bajas: [], existencias: [], realProduction: [], selectedMonth: mes, dailyBufferPct: 10 });
    const am = new Map(limpia.filter((r) => r.mes === mes).map((r) => [r.producto, r.cantidad]));
    const an = app.analyzeForecastProductErrors(fr, am, { topN: 15 });
    abs += +an.rows.reduce((s, r) => s + r.absoluteError, 0).toFixed(2); act += +Number(an.actual).toFixed(2);
    assert(r1.resumen.metricas["2025"].porMes[mes.slice(5)] === +Number(an.wape).toFixed(2), `${mes}: WAPE idéntico al modelo directo`);
  }
  assert(r1.resumen.metricas["2025"].wape === +((abs / act) * 100).toFixed(2), "WAPE 2025 idéntico al cálculo directo");
  assert(!r1.resumen.metricas["2024"], "sin historia previa a 2024 no se evalúa 2024 (arranque en frío)");
  const mensual = csvDe(r1.dir, "cruce-mensual.csv");
  const prod2025 = produccionSintetica().filter((r) => r.mes.startsWith("2025")).reduce((s, r) => s + r.producida, 0);
  assert(mensual.length === 12 && mensual.reduce((s, r) => s + Number(r.produccion_real), 0) === prod2025, "cruce mensual con la producción real de cada mes");
  const cpm = csvDe(r1.dir, "cruce-producto-mes.csv");
  const f1 = cpm.find((r) => r.mes === "2025-05" && r.producto === "MOKA GDE");
  assert(f1 && Number(f1.venta_sucursales) === 875 && Number(f1.produccion_real) === Math.round(875 * 1.05)
    && Number(f1.sobrante_real) === Number(f1.produccion_real) - 875 && Number(f1.faltante_real) === 0, "cruce producto × mes: venta, producción, sobrante y faltante");
  const md = leerSalida(r1.dir, "resumen.md");
  assert(/Fecha de corte \(última venta\):\*\* 2026-01-15/.test(md) && /## Cruce producción contra pronóstico/.test(md) && /## Pronóstico de 2026-01/.test(md), "resumen en Markdown");
  const logTxt = leerSalida(r1.dir, "corrida.log");
  assert(/FECHA DE CORTE: 2026-01-15/.test(logTxt) && /Modo: archivos/.test(logTxt) && /Fin OK/.test(logTxt) && /2026-01 está incompleto/.test(logTxt), "log con modo, fecha de corte y fin");
  assert(/\tOK\tarchivos\tcorte 2026-01-15/.test(fs.readFileSync(path.join(tmp, "salida-archivos", "historial-corridas.log"), "utf8")), "historial de corridas");

  // --- Guardia de solo lectura ---
  for (const sql of Object.values(base.SQL)) base.asegurarSoloLectura(sql);
  const malas = [
    "INSERT INTO dbo.Ventas VALUES (1)", "UPDATE dbo.Produccion SET cEstado='x'", "DELETE FROM dbo.Ventas",
    "SELECT * INTO dbo.Copia FROM dbo.Ventas", "SELECT 1 FROM dbo.Ventas; DROP TABLE dbo.Ventas", "EXEC sp_who",
    "SELECT cUsuario FROM dbo.Usuarios", "SELECT * FROM dbo.ChoferTraspasos", "WITH x AS (SELECT 1 AS a) DELETE FROM dbo.Ventas",
    "select 1 from dbo.Ventas -- ok\n; truncate table dbo.Ventas", "",
  ];
  for (const sql of malas) {
    let rechazada = false;
    try { base.asegurarSoloLectura(sql); } catch (e) { rechazada = e instanceof base.ErrorSoloLectura; }
    assert(rechazada, `rechaza consulta que no es de solo lectura: ${sql.slice(0, 50)}`);
  }
  assert(base.asegurarSoloLectura("SELECT cEstado FROM dbo.Produccion WHERE cEstado = 'DELETE; INSERT'"), "las palabras dentro de textos entre comillas no cuentan");

  // --- Configuración de modo base ---
  let msg = "";
  try { base.configDesdeEntorno({ PEPES_DB_SERVER: "srv", PEPES_DB_PASSWORD: PASSWORD }); } catch (e) { msg = e.message; }
  assert(/PEPES_DB_NAME/.test(msg) && /PEPES_DB_USER/.test(msg) && !msg.includes(PASSWORD) && !/PEPES_DB_SERVER,/.test(msg), "lista las variables que faltan, sin valores");
  const cfgOk = base.configDesdeEntorno({ PEPES_DB_SERVER: "espejo.local", PEPES_DB_NAME: "pepes_devBI_espejo", PEPES_DB_USER: "lector", PEPES_DB_PASSWORD: PASSWORD });
  assert(cfgOk.port === 1433 && cfgOk.encrypt === true && cfgOk.trustServerCertificate === false && !base.describirConexion(cfgOk).includes(PASSWORD), "valores por defecto y descripción sin contraseña");

  // Ejecutor real con un módulo mssql falso: pide ApplicationIntent=ReadOnly y nunca manda escrituras.
  const enviadas = [];
  let opciones = null;
  const falso = {
    VarChar: (n) => `VarChar(${n})`,
    ConnectionPool: class { constructor(o) { opciones = o; } async connect() { return { request: () => { const inputs = {}; return { input: (k, t, v) => { inputs[k] = { t, v }; }, query: async (s) => { enviadas.push({ s, inputs }); return { recordset: [{ db: "pepes_devBI_espejo" }] }; } }; }, close: async () => {} }; } },
  };
  const ej = await base.crearEjecutorMssql(cfgOk, { cargar: () => falso });
  assert(opciones.options.readOnlyIntent === true && opciones.password === PASSWORD && opciones.database === "pepes_devBI_espejo", "conexión con readOnlyIntent y credenciales solo del entorno");
  await ej.query(base.SQL.baseActual);
  let bloqueada = false;
  try { await ej.query("DELETE FROM dbo.Ventas"); } catch (e) { bloqueada = e instanceof base.ErrorSoloLectura; }
  assert(bloqueada && enviadas.length === 1, "una escritura nunca llega al servidor");
  let conexionMsg = "";
  const falsoError = { ConnectionPool: class { async connect() { throw new Error(`Login failed for user 'lector' password '${PASSWORD}'`); } } };
  try { await base.crearEjecutorMssql(cfgOk, { cargar: () => falsoError }); } catch (e) { conexionMsg = e.message; }
  assert(/no se pudo conectar/.test(conexionMsg) && !conexionMsg.includes(PASSWORD) && conexionMsg.includes("***"), "los errores de conexión ocultan la contraseña");

  // --- Modo base con la base sintética ---
  const bd = crearBaseSintetica();
  const registro = [];
  const envDb = { ...envBase, PEPES_FUENTE: "base", PEPES_DB_SERVER: "espejo.local", PEPES_DB_NAME: "pepes_devBI_espejo", PEPES_DB_USER: "lector", PEPES_DB_PASSWORD: PASSWORD, PEPES_SALIDA_DIR: path.join(tmp, "salida-base") };
  const r2 = await auto.ejecutar({ app, env: envDb, crearEjecutor: crearEjecutorSimulado(bd, registro), ahora: new Date(2026, 0, 16, 7, 0, 0) });
  assert(JSON.stringify(sinVolatiles(r2.resumen)) === JSON.stringify(sinVolatiles(r1.resumen)), "modo base da las mismas cifras que modo archivos con los mismos datos (canceladas fuera)");
  for (const nombre of ["cruce-producto-mes.csv", "cruce-mensual.csv", "cruce-producto-anual.csv", "pronostico-2026-01.csv", "serie-sucursal.json"]) {
    // Única diferencia esperada: modo base sabe cuántas líneas de producción siguen abiertas (los extractos no).
    const normal = (dir) => (nombre.endsWith(".csv")
      ? JSON.stringify(csvDe(dir, nombre).map(({ lineas_produccion_abiertas: _l, ...r }) => r))
      : leerSalida(dir, nombre).replace(/"source":"[^"]*"/, ""));
    assert(normal(r1.dir) === normal(r2.dir), `${nombre}: idéntico en modo base y modo archivos`);
  }
  const ventasQ = registro.filter((q) => q.tag === "ventas_mes_bodega");
  assert(ventasQ.length === 25 && ventasQ[0].params.desde === "2024-01-01T00:00:00" && ventasQ[24].params.hasta === "2026-02-01T00:00:00", "ventas mes por mes desde PEPES_DESDE hasta el mes de corte");
  const prodQ = registro.filter((q) => q.tag === "produccion_mes_producto");
  assert(prodQ.length === 1 && Object.keys(prodQ[0].params).join() === "desde", "Produccion se lee completa desde PEPES_DESDE, sin tope de fecha (incluye entregas a futuro)");
  assert(registro[registro.length - 1].tag === "close", "la conexión se cierra");
  for (const nombre of fs.readdirSync(r2.dir)) {
    const p = path.join(r2.dir, nombre);
    const contenido = fs.statSync(p).isDirectory() ? fs.readdirSync(p).map((n) => fs.readFileSync(path.join(p, n), "utf8")).join("") : fs.readFileSync(p, "utf8");
    assert(!contenido.includes(PASSWORD), `la contraseña no aparece en ${nombre}`);
  }
  assert(/Conectando a SQL Server espejo\.local:1433, base pepes_devBI_espejo, usuario lector \(solo lectura\)/.test(leerSalida(r2.dir, "corrida.log")), "el log dice a qué base se conectó");

  // La copia extracto/ de modo base vuelve a correr igual en modo archivos.
  const r3 = await auto.ejecutar({ app, env: { ...envBase, PEPES_FUENTE: "archivos", PEPES_DATOS_DIRS: path.join(r2.dir, "extracto"), PEPES_SALIDA_DIR: path.join(tmp, "salida-reproceso") }, ahora: new Date(2026, 0, 16, 8, 0, 0) });
  assert(leerSalida(r3.dir, "cruce-producto-mes.csv") === leerSalida(r2.dir, "cruce-producto-mes.csv"), "extracto/ de modo base reproduce la corrida en modo archivos");

  // Orden abierta que cambia de estado: la corrida siguiente la toma completa.
  const ordenDic = bd.Produccion.find((p) => p.dFechaEntrega.startsWith("2025-12"));
  const detDic = bd.ProduccionDet.find((d) => d.nProduccionPK === ordenDic.nProduccionPK);
  const final = detDic.nProducidos;
  ordenDic.cEstado = "Peticion"; detDic.nProducidos = 0; detDic.cEstado = "Pedido";
  const r4 = await auto.ejecutar({ app, env: { ...envDb, PEPES_SALIDA_DIR: path.join(tmp, "salida-base-2") }, crearEjecutor: crearEjecutorSimulado(bd, []), ahora: new Date(2026, 0, 16, 9, 0, 0) });
  assert(r4.resumen.advertencias.some((a) => /1 órdenes de producción siguen abiertas/.test(a)), "avisa de órdenes abiertas");
  const dic4 = csvDe(r4.dir, "cruce-mensual.csv").find((r) => r.mes === "2025-12");
  ordenDic.cEstado = "Finalizado"; detDic.nProducidos = final; detDic.cEstado = "Finalizado";
  const r5 = await auto.ejecutar({ app, env: { ...envDb, PEPES_SALIDA_DIR: path.join(tmp, "salida-base-3") }, crearEjecutor: crearEjecutorSimulado(bd, []), ahora: new Date(2026, 0, 16, 10, 0, 0) });
  const dic5 = csvDe(r5.dir, "cruce-mensual.csv").find((r) => r.mes === "2025-12");
  assert(Number(dic5.produccion_real) - Number(dic4.produccion_real) === final && !r5.resumen.advertencias.some((a) => /abiertas/.test(a)), "al finalizar la orden, la producción de diciembre se actualiza en la siguiente corrida");

  // Base equivocada: se detiene.
  let errDb = "";
  try { await auto.ejecutar({ app, env: { ...envDb, PEPES_DB_NAME: "otra_base", PEPES_SALIDA_DIR: path.join(tmp, "salida-error") }, crearEjecutor: crearEjecutorSimulado(bd, []), ahora: new Date() }); } catch (e) { errDb = e.message; }
  assert(/se esperaba "otra_base"/.test(errDb), "se detiene si la conexión abre otra base");
  assert(/\tERROR\tbase\t/.test(fs.readFileSync(path.join(tmp, "salida-error", "historial-corridas.log"), "utf8")), "el error queda en el historial");

  if (!process.env.CONSERVAR_TMP) fs.rmSync(tmp, { recursive: true, force: true });
  console.log("pronostico-automatico-test ok");
})().catch((error) => {
  console.error(error);
  process.exit(1);
});
