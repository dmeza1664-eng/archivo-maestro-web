#!/usr/bin/env node
/**
 * Pronóstico automático (Pastelería Pepes) — un solo comando:
 *
 *   npm run pronostico:auto                      (modo archivos: extractos CSV)
 *   PEPES_FUENTE=base npm run pronostico:auto    (modo base: copia espejo de pepes_devBI)
 *
 * 1. Lee ventas y producción de la fuente configurada (archivos o base SQL Server
 *    de solo lectura).
 * 2. Arma la serie oficial de demanda: venta de sucursales al público por mes y
 *    producto del catálogo (sin Planta León ni Amado Nervo antes de ago-2024).
 * 3. Pronostica el mes siguiente al último mes completo con el modelo del repo
 *    (App.jsx → calculateForecast), sin cambiar nada del modelo.
 * 4. Recalcula el pronóstico walk-forward de cada mes cerrado (solo con meses
 *    anteriores) y lo cruza con la venta y la producción real, por producto y mes.
 * 5. Escribe CSV, Markdown y un log con la fecha de corte en
 *    salidas/pronostico-automatico/corrida-AAAAMMDD-HHMMSS/.
 *
 * Configuración: variables de entorno (ver README, sección "Pronóstico automático")
 * o archivo .env.pronostico (ignorado por git). Nunca poner credenciales en el repo.
 */
const fs = require("fs");
const path = require("path");
const Module = require("module");
const { execFileSync } = require("child_process");

const datos = require("./lib/pepes-datos.cjs");
const { leerArchivos, separarDirectorios } = require("./lib/pepes-fuente-archivos.cjs");
const base = require("./lib/pepes-fuente-base.cjs");
const cruce = require("./lib/pepes-cruce.cjs");

const ROOT = path.join(__dirname, "..");

async function cargarApp(root = ROOT) {
  const esbuild = require("esbuild");
  const built = await esbuild.build({
    entryPoints: [path.join(root, "App.jsx")], bundle: true, platform: "node", format: "cjs", write: false,
    loader: { ".css": "text" },
    define: { "import.meta.env.VITE_API_URL": JSON.stringify(""), "import.meta.env.DEV": "false", "import.meta.env.PROD": "true" },
    logLevel: "silent",
  });
  const m = new Module("pronostico-automatico");
  m.filename = path.join(root, "scripts", "pronostico-automatico.bundle.cjs");
  m.paths = Module._nodeModulePaths(root);
  m._compile(built.outputFiles[0].text, m.filename);
  return m.exports;
}

function parseArgs(argv) {
  const out = {};
  for (const a of argv) {
    const m = String(a).match(/^--([a-z-]+)(?:=(.*))?$/);
    if (m) out[m[1]] = m[2] ?? "true";
  }
  return out;
}

function mesDe(valor) {
  const s = String(valor || "").trim();
  if (!/^\d{4}-\d{2}(-\d{2})?$/.test(s)) throw new Error(`Mes inválido: "${valor}" (usa AAAA-MM).`);
  return s.slice(0, 7);
}

function cargarArchivoEntorno(env, root) {
  const archivo = env.PEPES_ENV_FILE || path.join(root, ".env.pronostico");
  if (!fs.existsSync(archivo)) return null;
  const parsed = require("dotenv").parse(fs.readFileSync(archivo));
  for (const [k, v] of Object.entries(parsed)) if (env[k] == null || env[k] === "") env[k] = v;
  return archivo;
}

function leerConfig(argv, env, root = ROOT) {
  const args = parseArgs(argv);
  const modo = String(args.modo || env.PEPES_FUENTE || "archivos").toLowerCase();
  if (!["archivos", "base"].includes(modo)) throw new Error(`PEPES_FUENTE debe ser "archivos" o "base" (llegó "${modo}").`);
  const corte = args.corte || env.PEPES_FECHA_CORTE || null;
  if (corte && !/^\d{4}-\d{2}-\d{2}$/.test(corte)) throw new Error(`PEPES_FECHA_CORTE inválida: "${corte}" (usa AAAA-MM-DD).`);
  const meses = args.mes || env.PEPES_MESES_OBJETIVO || "";
  return {
    modo,
    datosDirs: separarDirectorios(args.datos || env.PEPES_DATOS_DIRS || ""),
    stockIdeal: args.stock || env.PEPES_STOCK_IDEAL || path.join(root, "wape-excel", "stock_ideal.xlsx"),
    mapeoProductos: env.PEPES_MAPEO_PRODUCTOS || path.join(root, "datos", "catalogo", "mapeo-productos.csv"),
    mapeoCodigos: env.PEPES_MAPEO_CODIGOS || path.join(root, "datos", "catalogo", "stock-codigo-map.json"),
    referencia: env.PEPES_REFERENCIA_METRICAS || path.join(root, "datos", "referencia", "metricas-modelo.json"),
    salidaBase: args.salida || env.PEPES_SALIDA_DIR || path.join(root, "salidas", "pronostico-automatico"),
    desde: mesDe(env.PEPES_DESDE || "2024-01"),
    fechaCorte: corte,
    mesesObjetivo: meses ? meses.split(",").map((m) => mesDe(m)) : null,
    colchonPct: Number(env.PEPES_COLCHON_PCT || 10),
  };
}

function horaLocal(d = new Date()) {
  const pad = (n) => String(n).padStart(2, "0");
  const off = -d.getTimezoneOffset();
  const sign = off >= 0 ? "+" : "-";
  const hh = Math.floor(Math.abs(off) / 60);
  const mm = Math.abs(off) % 60;
  return `${d.getFullYear()}-${pad(d.getMonth() + 1)}-${pad(d.getDate())} ${pad(d.getHours())}:${pad(d.getMinutes())}:${pad(d.getSeconds())} (UTC${sign}${hh}${mm ? `:${pad(mm)}` : ""})`;
}

function sello(d) {
  const pad = (n) => String(n).padStart(2, "0");
  return `${d.getFullYear()}${pad(d.getMonth() + 1)}${pad(d.getDate())}-${pad(d.getHours())}${pad(d.getMinutes())}${pad(d.getSeconds())}`;
}

function versionCodigo(root) {
  try {
    const commit = execFileSync("git", ["rev-parse", "--short", "HEAD"], { cwd: root, stdio: ["ignore", "pipe", "ignore"] }).toString().trim();
    const sucio = execFileSync("git", ["status", "--porcelain", "--untracked-files=no", "--", "App.jsx", "scripts"], { cwd: root, stdio: ["ignore", "pipe", "ignore"] }).toString().trim();
    return sucio ? `${commit} (con cambios locales en App.jsx/scripts)` : commit;
  } catch {
    return "desconocido (sin git)";
  }
}

function cargarStock(app, archivo) {
  if (!fs.existsSync(archivo)) throw new Error(`No existe el catálogo stock_ideal: ${archivo} (PEPES_STOCK_IDEAL).`);
  if (/\.json$/i.test(archivo)) {
    const j = JSON.parse(fs.readFileSync(archivo, "utf8"));
    return Array.isArray(j) ? j : j.stock;
  }
  const XLSX = require("xlsx");
  return app.parseStock(XLSX.readFile(archivo, { cellDates: true }));
}

const fmt = (n, d = 0) => (n == null || Number.isNaN(Number(n)) ? "—" : Number(n).toLocaleString("en-US", { minimumFractionDigits: d, maximumFractionDigits: d }));
const fmtPct = (n) => (n == null ? "—" : `${Number(n).toFixed(2)}%`);

function tablaMd(filas, columnas) {
  const head = `| ${columnas.map((c) => c.titulo).join(" | ")} |`;
  const sep = `|${columnas.map((c) => (c.izq ? "---" : "---:")).join("|")}|`;
  return [head, sep, ...filas.map((f) => `| ${columnas.map((c) => c.valor(f)).join(" | ")} |`)].join("\n");
}

function escribirExtractoDesdeBase(dir, fuente) {
  fs.mkdirSync(dir, { recursive: true });
  const porAnio = (filas) => {
    const g = new Map();
    for (const r of filas) { const y = String(r.mes).slice(0, 4); if (!g.has(y)) g.set(y, []); g.get(y).push(r); }
    return g;
  };
  for (const [y, filas] of porAnio(fuente.ventas)) fs.writeFileSync(path.join(dir, `ventas-${y}-mensual-producto-bodega.csv`), datos.toCsv(filas, [...datos.VENTAS_COLUMNAS, "qtyPromocion"]));
  for (const [y, filas] of porAnio(fuente.produccion)) fs.writeFileSync(path.join(dir, `produccion-${y}-mensual-producto.csv`), datos.toCsv(filas, [...datos.PRODUCCION_COLUMNAS, "lineasAbiertas"]));
  fs.writeFileSync(path.join(dir, "extract-meta.json"), JSON.stringify({
    database: fuente.fuente, extractedAt: new Date().toISOString(), readOnly: "SELECT only (pronostico-automatico)",
    maxVentaDCreada: fuente.maxVentaDCreada, produccionEstado: fuente.produccionEstado,
  }, null, 2));
}

function compararReferencia(referencia, evaluados) {
  const out = [];
  for (const [anio, ref] of Object.entries(referencia.cortes || {})) {
    const meses = ref.meses || Array.from({ length: 12 }, (_, i) => `${anio}-${String(i + 1).padStart(2, "0")}`);
    const sub = evaluados.filter((e) => meses.includes(e.mes));
    if (sub.length !== meses.length) { out.push({ anio, estado: "no evaluable", detalle: `faltan meses (${sub.length} de ${meses.length})` }); continue; }
    const obtenido = {
      wape: cruce.metricasDeMeses(sub).wape,
      wapeSinEnero: cruce.metricasDeMeses(sub, (m) => !m.endsWith("-01"))?.wape ?? null,
      errorTotalMensual: cruce.metricasDeMeses(sub).errorTotalMensual,
    };
    const difs = Object.keys(obtenido).filter((k) => ref[k] != null && obtenido[k] !== ref[k]);
    out.push({ anio, estado: difs.length ? "difiere" : "coincide", esperado: ref, obtenido, detalle: difs.map((k) => `${k} ${obtenido[k]} vs ${ref[k]}`).join("; ") });
  }
  return out;
}

/**
 * Corre todo el proceso. Opciones para pruebas: env (objeto), argv, root,
 * crearEjecutor (base simulada), ahora (Date), app (módulo ya cargado).
 */
async function ejecutar({ argv = [], env = process.env, root = ROOT, crearEjecutor = null, ahora = new Date(), app = null } = {}) {
  const inicio = ahora;
  const archivoEnv = cargarArchivoEntorno(env, root);
  const cfg = leerConfig(argv, env, root);
  const dir = path.join(cfg.salidaBase, `corrida-${sello(inicio)}`);
  fs.mkdirSync(dir, { recursive: true });
  const lineasLog = [];
  const log = (msg) => {
    const linea = `[${horaLocal(new Date())}] ${msg}`;
    lineasLog.push(linea);
    fs.appendFileSync(path.join(dir, "corrida.log"), `${linea}\n`);
    if (!env.PEPES_SILENCIOSO) console.log(linea);
  };
  let dbCfg = null;
  try {
    log(`Inicio del pronóstico automático. Código: ${versionCodigo(root)}. Modo: ${cfg.modo}.`);
    if (archivoEnv) log(`Variables leídas de ${archivoEnv} (sin mostrar valores).`);
    app = app || await cargarApp(root);

    // 1. Fuente
    let fuente;
    if (cfg.modo === "archivos") {
      fuente = leerArchivos({ dirs: cfg.datosDirs });
      for (const a of fuente.archivosLeidos) log(`  leído ${a.archivo} (${a.filas} filas)`);
    } else {
      dbCfg = base.configDesdeEntorno(env);
      log(`Conectando a ${base.describirConexion(dbCfg)}.`);
      const ejecutor = crearEjecutor ? await crearEjecutor(dbCfg) : await base.crearEjecutorMssql(dbCfg);
      try {
        fuente = await base.leerBase({ cfg: dbCfg, ejecutor, desde: `${cfg.desde}-01`, hoy: ahora, log });
      } finally {
        await ejecutor.close();
      }
      escribirExtractoDesdeBase(path.join(dir, "extracto"), fuente);
      log(`Copia de lo leído (mismo formato que los extractos): ${path.join(dir, "extracto")}`);
    }
    const fechaCorte = cfg.fechaCorte || fuente.fechaCorte;
    const ultimoCompleto = datos.ultimoMesCompleto(fechaCorte);
    const mesCorte = fechaCorte.slice(0, 7);
    log(`Fuente: ${fuente.fuente}. Ventas: ${fuente.ventas.length} filas; producción: ${fuente.produccion.length} filas.`);
    log(`FECHA DE CORTE: ${fechaCorte} (${cfg.fechaCorte ? "PEPES_FECHA_CORTE" : fuente.origenCorte}). Último mes completo: ${ultimoCompleto}.`);
    const advertencias = [...fuente.advertencias];
    if (mesCorte > ultimoCompleto) advertencias.push(`${mesCorte} está incompleto (venta hasta ${fechaCorte}): no entra a la historia ni al error; solo se muestra como avance.`);

    // 2. Serie oficial y producción
    const mapeo = datos.loadMapeo(cfg.mapeoProductos, cfg.mapeoCodigos);
    const esDemanda = datos.crearClasificadorDemanda(app);
    const ventasHastaCorte = fuente.ventas.filter((r) => String(r.mes) <= mesCorte);
    const oficial = datos.construirSerieSucursal(ventasHastaCorte, mapeo, esDemanda, { desde: cfg.desde, hasta: ultimoCompleto });
    const serieOficial = datos.hidratarSerie(oficial.rows);
    const hayPrevia = ventasHastaCorte.some((r) => String(r.mes) < cfg.desde);
    const anioDesde = cfg.desde.slice(0, 4);
    const finAnioDesde = `${anioDesde}-12` < ultimoCompleto ? `${anioDesde}-12` : ultimoCompleto;
    const serieExtendida = hayPrevia ? datos.hidratarSerie(datos.construirSerieSucursal(ventasHastaCorte, mapeo, esDemanda, { hasta: finAnioDesde }).rows) : null;
    const produccion = datos.agregarProduccion(fuente.produccion, mapeo);
    const ventaSucursalTodo = datos.agregarVentaSucursal(ventasHastaCorte, mapeo, esDemanda);
    const piezasSerie = oficial.rows.reduce((s, r) => s + r.cantidad, 0);
    log(`Serie oficial (venta de sucursales, ${cfg.desde} a ${ultimoCompleto}): ${oficial.rows.length} filas mes × producto, ${Math.round(piezasSerie).toLocaleString("en-US")} piezas.`);
    if (serieExtendida) log(`Historia previa a ${cfg.desde} disponible: se usa solo para evaluar ${anioDesde} (misma convención que el backtest oficial 2024 con historia 2023).`);
    const sinMapeo2 = oficial.sinMapeo.filter((x) => x.piezas > 0);
    if (sinMapeo2.length) log(`  ${sinMapeo2.length} nProductoPK sin fila en el mapeo (fuera del catálogo del pronóstico), ${Math.round(sinMapeo2.reduce((s, x) => s + x.piezas, 0)).toLocaleString("en-US")} piezas (todos los canales). Ver pk-sin-mapeo.csv.`);
    const stockRows = cargarStock(app, cfg.stockIdeal);
    log(`Catálogo stock_ideal: ${stockRows.length} productos (${cfg.stockIdeal}).`);

    // 3. Pronóstico del mes objetivo
    const mesesObjetivo = cfg.mesesObjetivo || [datos.siguienteMes(ultimoCompleto)];
    const pronosticos = {};
    for (const mes of mesesObjetivo) {
      if (mes > datos.siguienteMes(ultimoCompleto)) advertencias.push(`${mes} está a más de un mes del último mes completo (${ultimoCompleto}); el modelo está pensado para el mes siguiente.`);
      const filas = cruce.pronosticarMes(app, stockRows, serieOficial, mes, cfg.colchonPct);
      pronosticos[mes] = filas;
      const total = filas.reduce((s, r) => s + r.pronosticoVenta, 0);
      const plan = filas.reduce((s, r) => s + r.produccionSugerida, 0);
      const hastaHist = datos.siguienteMes(mes, -1) < ultimoCompleto ? datos.siguienteMes(mes, -1) : ultimoCompleto;
      log(`Pronóstico ${mes} (historia ${cfg.desde} a ${hastaHist}): ${Math.round(total).toLocaleString("en-US")} piezas; plan (pron. + ${cfg.colchonPct}%): ${Math.round(plan).toLocaleString("en-US")}.`);
    }

    // 4. Walk-forward y cruce
    const evaluados = [];
    const mesesEval = base.mesesEntre(cfg.desde, ultimoCompleto);
    for (const mes of mesesEval) {
      const serie = mes.startsWith(anioDesde) ? serieExtendida : serieOficial;
      if (!serie) continue;
      const e = cruce.evaluarMes(app, stockRows, serie, mes, cfg.colchonPct);
      if (e) evaluados.push(e);
    }
    if (!serieExtendida) log(`Sin historia previa a ${cfg.desde}: ${anioDesde} no se evalúa (arranque en frío).`);
    const metricas = cruce.metricasPorAnio(evaluados);
    for (const [anio, m] of Object.entries(metricas)) {
      log(`Error walk-forward ${anio} (${m.primerMes} a ${m.ultimoMes}): WAPE producto-mes ${fmtPct(m.wape)}; sin enero ${fmtPct(m.wapeSinEnero)}; error del total mensual ${fmtPct(m.errorTotalMensual)}.`);
    }
    if (evaluados.some((e) => e.mes < "2024-08")) advertencias.push("Antes de ago-2024 la venta de sucursales excluye a Suc. Amado Nervo (entonces también surtía) y la producción no: en esos meses la producción contra la venta no es comparable.");
    const filasCruce = cruce.cruceProductoMes(evaluados, produccion);
    const mensual = cruce.cruceMensual(evaluados, filasCruce);
    const anual = cruce.cruceProductoAnual(filasCruce);

    let comparacion = [];
    if (fs.existsSync(cfg.referencia)) {
      const referencia = JSON.parse(fs.readFileSync(cfg.referencia, "utf8"));
      comparacion = compararReferencia(referencia, evaluados);
      for (const c of comparacion) log(`Referencia ${referencia.modelo || ""} ${c.anio}: ${c.estado}${c.detalle ? ` (${c.detalle})` : ""}.`);
      if (comparacion.some((c) => c.estado === "difiere")) advertencias.push("Alguna métrica difiere de la referencia del modelo: revisa si cambiaron datos históricos (cancelaciones, recopia) o el código.");
    }

    // Avance del mes en curso (venta y producción parciales hasta la fecha de corte)
    const avance = [];
    for (const mes of mesesObjetivo) {
      for (const r of pronosticos[mes]) {
        const k = `${mes}\u0000${r.producto}`;
        avance.push({
          mes, orden: r.orden, producto: r.producto,
          pronostico: cruce.r2(r.pronosticoVenta), plan_modelo: r.produccionSugerida,
          venta_al_corte: mes === mesCorte ? Math.round(ventaSucursalTodo.get(k) || 0) : "",
          produccion_al_corte: produccion.has(k) ? Math.round(produccion.get(k).producida) : 0,
          lineas_produccion_abiertas: produccion.get(k)?.lineasAbiertas ?? "",
          metodo: r.metodoPronostico, meses_usados: r.mesesUsados,
        });
      }
    }

    // 5. Salidas
    const escribir = (nombre, contenido) => { fs.writeFileSync(path.join(dir, nombre), contenido); return nombre; };
    const archivos = [];
    for (const mes of mesesObjetivo) {
      archivos.push(escribir(`pronostico-${mes}.csv`, "\uFEFF" + datos.toCsv(avance.filter((a) => a.mes === mes),
        ["mes", "orden", "producto", "pronostico", "plan_modelo", "venta_al_corte", "produccion_al_corte", "lineas_produccion_abiertas", "metodo", "meses_usados"])));
    }
    archivos.push(escribir("cruce-producto-mes.csv", "\uFEFF" + datos.toCsv(filasCruce, Object.keys(filasCruce[0] || { mes: 1 }))));
    archivos.push(escribir("cruce-mensual.csv", "\uFEFF" + datos.toCsv(mensual, Object.keys(mensual[0] || { mes: 1 }))));
    archivos.push(escribir("cruce-producto-anual.csv", "\uFEFF" + datos.toCsv(anual, Object.keys(anual[0] || { anio: 1 }))));
    archivos.push(escribir("pk-sin-mapeo.csv", "\uFEFF" + datos.toCsv(sinMapeo2.map((x) => ({ ...x, piezas: Math.round(x.piezas) })), ["nProductoPK", "cCodigo", "cDescripcion", "piezas"])));
    archivos.push(escribir("serie-sucursal.json", JSON.stringify({ source: `pronostico-automatico (${fuente.modo})`, desde: cfg.desde, hasta: ultimoCompleto, rows: oficial.rows, rowCount: oficial.rows.length })));
    const resumen = {
      generado: horaLocal(inicio), codigo: versionCodigo(root), modo: fuente.modo, fuente: fuente.fuente,
      fechaCorte, ultimoMesCompleto: ultimoCompleto, mesesObjetivo, colchonPct: cfg.colchonPct,
      serieOficial: { desde: cfg.desde, hasta: ultimoCompleto, filas: oficial.rows.length, piezas: Math.round(piezasSerie) },
      metricas, referencia: comparacion,
      pronostico: Object.fromEntries(mesesObjetivo.map((m) => [m, {
        piezas: Math.round(pronosticos[m].reduce((s, r) => s + r.pronosticoVenta, 0)),
        plan: Math.round(pronosticos[m].reduce((s, r) => s + r.produccionSugerida, 0)),
        productos: pronosticos[m].length,
      }])),
      totalesPorAnio: Object.fromEntries(Object.keys(metricas).map((y) => [y, cruce.sumar(filasCruce.filter((f) => f.mes.startsWith(y)))]).map(([y, s]) => [y, Object.fromEntries(Object.entries(s).map(([k, v]) => [k, Math.round(v)]))])),
      produccionEstado: fuente.produccionEstado || null,
      advertencias,
    };
    archivos.push(escribir("metricas.json", JSON.stringify(resumen, null, 1)));
    archivos.push(escribir("resumen.md", resumenMarkdown({ resumen, mensual, anual, avance, mesesObjetivo, mesCorte })));
    for (const a of advertencias) log(`AVISO: ${a}`);
    log(`Archivos: ${archivos.join(", ")}, corrida.log. Carpeta: ${dir}`);
    fs.mkdirSync(cfg.salidaBase, { recursive: true });
    fs.appendFileSync(path.join(cfg.salidaBase, "historial-corridas.log"),
      `${horaLocal(inicio)}\tOK\t${fuente.modo}\tcorte ${fechaCorte}\tobjetivo ${mesesObjetivo.join(",")}\t${Object.entries(metricas).map(([y, m]) => `${y} ${m.wape}`).join(" · ")}\t${dir}\n`);
    fs.writeFileSync(path.join(cfg.salidaBase, "ultima-corrida.txt"), `${dir}\n`);
    log(`Fin OK (${((Date.now() - inicio.getTime()) / 1000).toFixed(1)} s desde el inicio).`);
    return { dir, resumen, archivos, log: lineasLog };
  } catch (error) {
    const msg = base.ocultarSecretos(error?.message || String(error), dbCfg);
    log(`ERROR: ${msg}`);
    fs.mkdirSync(cfg.salidaBase, { recursive: true });
    fs.appendFileSync(path.join(cfg.salidaBase, "historial-corridas.log"), `${horaLocal(inicio)}\tERROR\t${cfg.modo}\t${msg}\t${dir}\n`);
    const e = new Error(msg);
    e.dir = dir;
    throw e;
  }
}

function resumenMarkdown({ resumen, mensual, anual, avance, mesesObjetivo, mesCorte }) {
  const L = [];
  L.push("# Pronóstico automático — Pastelería Pepes", "");
  L.push(`- **Generado:** ${resumen.generado}`);
  L.push(`- **Código (modelo):** \`${resumen.codigo}\``);
  L.push(`- **Fuente:** ${resumen.modo} — ${resumen.fuente}`);
  L.push(`- **Fecha de corte (última venta):** ${resumen.fechaCorte} · **último mes completo:** ${resumen.ultimoMesCompleto}`);
  L.push(`- **Serie oficial:** venta de sucursales al público (sin Planta León; sin Suc. Amado Nervo antes de ago-2024), ${resumen.serieOficial.desde} a ${resumen.serieOficial.hasta}, ${fmt(resumen.serieOficial.piezas)} piezas.`);
  L.push(`- **Plan** = producción sugerida del modelo (pronóstico + ${resumen.colchonPct}% con la regla operativa de pasteles).`, "");
  for (const mes of mesesObjetivo) {
    const p = resumen.pronostico[mes];
    L.push(`## Pronóstico de ${mes}`, "");
    L.push(`Total: **${fmt(p.piezas)} piezas** de venta de sucursales; plan de producción **${fmt(p.plan)} piezas** (${p.productos} productos). Detalle: \`pronostico-${mes}.csv\`.`, "");
    const filas = avance.filter((a) => a.mes === mes).sort((a, b) => b.pronostico - a.pronostico).slice(0, 15);
    const cols = [
      { titulo: "Producto", izq: true, valor: (f) => f.producto },
      { titulo: "Pronóstico", valor: (f) => fmt(f.pronostico) },
      { titulo: "Plan", valor: (f) => fmt(f.plan_modelo) },
    ];
    if (mes === mesCorte) cols.push({ titulo: `Venta al ${resumen.fechaCorte}`, valor: (f) => fmt(f.venta_al_corte) });
    cols.push({ titulo: "Producción registrada", valor: (f) => fmt(f.produccion_al_corte) });
    L.push("Los 15 productos con mayor pronóstico:", "", tablaMd(filas, cols), "");
    if (mes === mesCorte) L.push(`_${mes} está en curso: la venta y la producción registradas son parciales (hasta ${resumen.fechaCorte})._`, "");
  }
  L.push("## Error del modelo (walk-forward: cada mes con solo meses anteriores)", "");
  L.push(tablaMd(Object.entries(resumen.metricas).map(([anio, m]) => ({ anio, ...m })), [
    { titulo: "Año", izq: true, valor: (f) => `${f.anio} (${f.primerMes.slice(5)}–${f.ultimoMes.slice(5)})` },
    { titulo: "WAPE producto-mes", valor: (f) => fmtPct(f.wape) },
    { titulo: "Sin enero", valor: (f) => fmtPct(f.wapeSinEnero) },
    { titulo: "Error del total mensual", valor: (f) => fmtPct(f.errorTotalMensual) },
    { titulo: "Sesgo neto", valor: (f) => fmtPct(f.netoPct) },
  ]), "");
  if (resumen.referencia.length) {
    L.push("Comparación con las cifras de referencia del modelo (`datos/referencia/metricas-modelo.json`):", "");
    for (const c of resumen.referencia) L.push(`- ${c.anio}: **${c.estado}**${c.detalle ? ` (${c.detalle})` : ""}`);
    L.push("");
  }
  L.push("## Cruce producción contra pronóstico, por mes", "");
  L.push("Sobrante/faltante se calculan producto por producto contra la venta de sucursales (no se compensan entre productos).", "");
  L.push(tablaMd(mensual, [
    { titulo: "Mes", izq: true, valor: (f) => f.mes },
    { titulo: "Venta suc.", valor: (f) => fmt(f.venta_sucursales) },
    { titulo: "Pronóstico", valor: (f) => fmt(f.pronostico) },
    { titulo: "Plan", valor: (f) => fmt(f.plan_modelo) },
    { titulo: "Producción", valor: (f) => fmt(f.produccion_real) },
    { titulo: "WAPE", valor: (f) => fmtPct(f.wape_producto_mes) },
    { titulo: "Error total mes", valor: (f) => fmtPct(f.error_total_mes_pct) },
    { titulo: "Prod. vs venta", valor: (f) => fmtPct(f.produccion_vs_venta_pct) },
    { titulo: "Sobrante / faltante real", valor: (f) => `${fmt(f.sobrante_real)} / ${fmt(f.faltante_real)}` },
    { titulo: "Sobrante / faltante con plan", valor: (f) => `${fmt(f.sobrante_plan)} / ${fmt(f.faltante_plan)}` },
  ]), "");
  L.push("### Totales por año", "");
  L.push(tablaMd(Object.entries(resumen.totalesPorAnio).map(([anio, s]) => ({ anio, ...s })), [
    { titulo: "Año", izq: true, valor: (f) => f.anio },
    { titulo: "Venta suc.", valor: (f) => fmt(f.venta_sucursales) },
    { titulo: "Pronóstico", valor: (f) => fmt(f.pronostico) },
    { titulo: "Plan", valor: (f) => fmt(f.plan_modelo) },
    { titulo: "Producción", valor: (f) => fmt(f.produccion_real) },
    { titulo: "Sobrante / faltante real", valor: (f) => `${fmt(f.sobrante_real)} / ${fmt(f.faltante_real)}` },
    { titulo: "Sobrante / faltante con plan", valor: (f) => `${fmt(f.sobrante_plan)} / ${fmt(f.faltante_plan)}` },
  ]), "");
  const ultimoAnio = Object.keys(resumen.metricas).pop();
  if (ultimoAnio) {
    const top = anual.filter((a) => a.anio === ultimoAnio).sort((a, b) => Math.abs(b.produccion_real - b.venta_sucursales) - Math.abs(a.produccion_real - a.venta_sucursales)).slice(0, 15);
    L.push(`## Productos con mayor diferencia entre producción y venta (${ultimoAnio})`, "");
    L.push(tablaMd(top, [
      { titulo: "Producto", izq: true, valor: (f) => f.producto },
      { titulo: "Venta suc.", valor: (f) => fmt(f.venta_sucursales) },
      { titulo: "Producción", valor: (f) => fmt(f.produccion_real) },
      { titulo: "Pronóstico", valor: (f) => fmt(f.pronostico) },
      { titulo: "Plan", valor: (f) => fmt(f.plan_modelo) },
      { titulo: "WAPE", valor: (f) => fmtPct(f.wape_pct) },
      { titulo: "Prod. vs venta", valor: (f) => fmtPct(f.produccion_vs_venta_pct) },
    ]), "");
  }
  if (resumen.advertencias.length) {
    L.push("## Avisos", "");
    for (const a of resumen.advertencias) L.push(`- ${a}`);
    L.push("");
  }
  L.push("Archivos: `pronostico-AAAA-MM.csv`, `cruce-producto-mes.csv`, `cruce-mensual.csv`, `cruce-producto-anual.csv`, `pk-sin-mapeo.csv`, `serie-sucursal.json`, `metricas.json`, `corrida.log`.", "");
  return L.join("\n");
}

module.exports = { ejecutar, leerConfig, cargarApp, compararReferencia, horaLocal };

if (require.main === module) {
  ejecutar({ argv: process.argv.slice(2) }).then(() => process.exit(0)).catch((error) => {
    console.error(`Pronóstico automático: ERROR — ${error.message}${error.dir ? ` (log en ${error.dir})` : ""}`);
    process.exit(1);
  });
}
