/**
 * Fuente "base": lee directamente la copia espejo de pepes_devBI (SQL Server)
 * con un usuario de SOLO LECTURA. Las consultas son las mismas que generaron
 * los extractos CSV (extract2026.js): mismas tablas dbo, mismos filtros.
 *
 *  - Conexión solo por variables de entorno (PEPES_DB_SERVER, PEPES_DB_NAME,
 *    PEPES_DB_USER, PEPES_DB_PASSWORD; opcionales PEPES_DB_PORT,
 *    PEPES_DB_ENCRYPT, PEPES_DB_TRUST_SERVER_CERT, PEPES_DB_TIMEOUT_MS).
 *    Nunca hay credenciales en el repo ni en el log.
 *  - Cada consulta pasa por asegurarSoloLectura: una sola sentencia SELECT,
 *    sin palabras de escritura y solo contra las 7 tablas dbo del Pronóstico.
 *    Además la conexión pide ApplicationIntent=ReadOnly.
 *  - Ventas se lee mes por mes (la consulta anual se queda sin tiempo).
 *  - Produccion se vuelve a leer COMPLETA en cada corrida (desde PEPES_DESDE,
 *    incluidas entregas a futuro): las órdenes cambian de estado
 *    (Peticion → Finalizado) y nProducidos se actualiza después.
 */
const TABLAS_PERMITIDAS = new Set(["dbo.ventas", "dbo.ventadet", "dbo.productos", "dbo.bodegas", "dbo.sucursales", "dbo.produccion", "dbo.producciondet"]);
const VARIABLES_REQUERIDAS = ["PEPES_DB_SERVER", "PEPES_DB_NAME", "PEPES_DB_USER", "PEPES_DB_PASSWORD"];

const SQL = {
  baseActual: "/* pronostico:base_actual */ SELECT DB_NAME() AS db",
  maxVenta: `/* pronostico:max_venta */
SELECT MAX(v.dCreada) AS maxVenta FROM dbo.Ventas v
WHERE v.dCancelada IS NULL AND v.dCreada < CONVERT(datetime, @limite, 126)`,
  ventasMes: `/* pronostico:ventas_mes_bodega */
SELECT FORMAT(v.dCreada,'yyyy-MM') AS mes, vd.nProductoPK, p.cCodigo, p.cDescripcion, b.nBodegaPK, b.cNombre AS bodegaNombre,
       s.nSucursalPK, s.cNombre AS sucursalNombre, SUM(vd.nCantidad) AS qty, SUM(vd.mSubTotal) AS subtotal, SUM(vd.mTotal) AS total,
       COUNT(*) AS lineas, SUM(CASE WHEN ISNULL(vd.nPromocionPK,0) <> 0 THEN vd.nCantidad ELSE 0 END) AS qtyPromocion
FROM dbo.VentaDet vd
INNER JOIN dbo.Ventas v ON v.nVentaPK = vd.nVentaPK
LEFT JOIN dbo.Productos p ON p.nProductoPK = vd.nProductoPK
LEFT JOIN dbo.Bodegas b ON b.nBodegaPK = v.nBodegaPK
LEFT JOIN dbo.Sucursales s ON s.nSucursalPK = b.nSucursalPK
WHERE v.dCreada >= CONVERT(datetime, @desde, 126) AND v.dCreada < CONVERT(datetime, @hasta, 126)
  AND v.dCancelada IS NULL AND vd.dCancelado IS NULL
GROUP BY FORMAT(v.dCreada,'yyyy-MM'), vd.nProductoPK, p.cCodigo, p.cDescripcion, b.nBodegaPK, b.cNombre, s.nSucursalPK, s.cNombre`,
  produccion: `/* pronostico:produccion_mes_producto */
SELECT FORMAT(COALESCE(p.dFechaEntrega,p.dFechaPeticion),'yyyy-MM') AS mes, pd.nProductoPK, pr.cCodigo, pr.cDescripcion,
       SUM(pd.nCantidad) AS qtyPedida, SUM(pd.nProducidos) AS qtyProducida, SUM(ISNULL(pd.nCancelados,0)) AS qtyCancelada,
       SUM(ISNULL(pd.nSalida,0)) AS qtySalida, COUNT(*) AS lineas,
       SUM(CASE WHEN p.cEstado IN ('Finalizado','Cancelado') THEN 0 ELSE 1 END) AS lineasAbiertas
FROM dbo.ProduccionDet pd
INNER JOIN dbo.Produccion p ON p.nProduccionPK = pd.nProduccionPK
LEFT JOIN dbo.Productos pr ON pr.nProductoPK = pd.nProductoPK
WHERE COALESCE(p.dFechaEntrega,p.dFechaPeticion) >= CONVERT(datetime, @desde, 126)
GROUP BY FORMAT(COALESCE(p.dFechaEntrega,p.dFechaPeticion),'yyyy-MM'), pd.nProductoPK, pr.cCodigo, pr.cDescripcion`,
  produccionEstado: `/* pronostico:produccion_estado */
SELECT COUNT(*) AS ordenes, SUM(CASE WHEN p.cEstado IN ('Finalizado','Cancelado') THEN 0 ELSE 1 END) AS ordenesAbiertas,
       MAX(p.dFechaPeticion) AS maxPeticion
FROM dbo.Produccion p
WHERE COALESCE(p.dFechaEntrega,p.dFechaPeticion) >= CONVERT(datetime, @desde, 126)`,
};

class ErrorSoloLectura extends Error {}

/** Rechaza cualquier consulta que no sea un SELECT de solo lectura a las tablas permitidas. */
function asegurarSoloLectura(sql) {
  const limpio = String(sql || "")
    .replace(/\/\*[\s\S]*?\*\//g, " ")
    .replace(/--[^\n]*/g, " ")
    .replace(/'(?:[^']|'')*'/g, "''")
    .trim();
  if (!limpio) throw new ErrorSoloLectura("consulta vacía");
  if (limpio.replace(/;\s*$/, "").includes(";")) throw new ErrorSoloLectura("solo se permite una sentencia por consulta");
  if (!/^(SELECT|WITH)\b/i.test(limpio)) throw new ErrorSoloLectura("solo se permiten consultas SELECT");
  const prohibida = limpio.match(/\b(INSERT|UPDATE|DELETE|MERGE|DROP|ALTER|CREATE|TRUNCATE|EXEC|EXECUTE|GRANT|REVOKE|DENY|BACKUP|RESTORE|INTO|DBCC|OPENROWSET|OPENQUERY|OPENDATASOURCE|BULK|SHUTDOWN|KILL|SET|DECLARE|WAITFOR|RECONFIGURE|USE)\b/i);
  if (prohibida) throw new ErrorSoloLectura(`palabra no permitida en consulta de solo lectura: ${prohibida[1].toUpperCase()}`);
  const tablas = [...limpio.matchAll(/\b(?:FROM|JOIN)\s+([\[\]\w.]+)/gi)].map((m) => m[1].replace(/[\[\]]/g, "").toLowerCase());
  for (const t of tablas) {
    if (!TABLAS_PERMITIDAS.has(t)) throw new ErrorSoloLectura(`tabla no permitida: ${t} (solo ${[...TABLAS_PERMITIDAS].join(", ")})`);
  }
  return true;
}

function booleano(valor, porDefecto) {
  if (valor == null || valor === "") return porDefecto;
  return /^(1|true|si|sí|yes)$/i.test(String(valor).trim());
}

/** Configuración de conexión desde variables de entorno. Nunca devuelve ni imprime la contraseña en mensajes. */
function configDesdeEntorno(env = process.env) {
  const faltan = VARIABLES_REQUERIDAS.filter((k) => !String(env[k] || "").trim());
  if (faltan.length) throw new Error(`Modo base: faltan variables de entorno: ${faltan.join(", ")}. Ver README (sección "Modo base").`);
  return {
    server: String(env.PEPES_DB_SERVER).trim(),
    database: String(env.PEPES_DB_NAME).trim(),
    user: String(env.PEPES_DB_USER).trim(),
    password: String(env.PEPES_DB_PASSWORD),
    port: Number(env.PEPES_DB_PORT || 1433),
    encrypt: booleano(env.PEPES_DB_ENCRYPT, true),
    trustServerCertificate: booleano(env.PEPES_DB_TRUST_SERVER_CERT, false),
    requestTimeout: Number(env.PEPES_DB_TIMEOUT_MS || 300000),
  };
}

function describirConexion(cfg) {
  return `SQL Server ${cfg.server}:${cfg.port}, base ${cfg.database}, usuario ${cfg.user} (solo lectura)`;
}

function ocultarSecretos(texto, cfg) {
  let s = String(texto ?? "");
  if (cfg?.password) s = s.split(cfg.password).join("***");
  return s;
}

/** Ejecutor real con el paquete mssql (ApplicationIntent=ReadOnly). */
async function crearEjecutorMssql(cfg, { cargar = () => require("mssql") } = {}) {
  let sql;
  try {
    sql = cargar();
  } catch {
    throw new Error("Modo base: falta el paquete mssql. Corre `npm ci` en la carpeta del repo.");
  }
  let pool;
  try {
    pool = await new sql.ConnectionPool({
      server: cfg.server, database: cfg.database, user: cfg.user, password: cfg.password, port: cfg.port,
      options: { encrypt: cfg.encrypt, trustServerCertificate: cfg.trustServerCertificate, readOnlyIntent: true, appName: "pronostico-automatico" },
      requestTimeout: cfg.requestTimeout, connectionTimeout: 30000, pool: { max: 2, min: 0 },
    }).connect();
  } catch (error) {
    throw new Error(`Modo base: no se pudo conectar (${describirConexion(cfg)}): ${ocultarSecretos(error?.message, cfg)}`);
  }
  return {
    async query(texto, params = {}) {
      asegurarSoloLectura(texto);
      const req = pool.request();
      for (const [k, v] of Object.entries(params)) req.input(k, sql.VarChar(40), String(v));
      try {
        return (await req.query(texto)).recordset;
      } catch (error) {
        throw new Error(`Modo base: falló la consulta: ${ocultarSecretos(error?.message, cfg)}`);
      }
    },
    async close() { await pool.close(); },
  };
}

function fechaIso(valor) {
  if (valor == null) return null;
  if (valor instanceof Date) return valor.toISOString();
  return String(valor);
}

function mesesEntre(desdeMes, hastaMes) {
  const out = [];
  let [y, m] = desdeMes.split("-").map(Number);
  const [yh, mh] = hastaMes.split("-").map(Number);
  while (y < yh || (y === yh && m <= mh)) {
    out.push(`${y}-${String(m).padStart(2, "0")}`);
    m += 1; if (m > 12) { m = 1; y += 1; }
  }
  return out;
}

function mesSiguiente(mes) {
  let [y, m] = mes.split("-").map(Number);
  m += 1; if (m > 12) { m = 1; y += 1; }
  return `${y}-${String(m).padStart(2, "0")}`;
}

const texto = (v) => (v == null ? "" : String(v));
function normalizarVenta(r) {
  return {
    mes: texto(r.mes), nProductoPK: texto(r.nProductoPK), cCodigo: texto(r.cCodigo), cDescripcion: texto(r.cDescripcion),
    nBodegaPK: texto(r.nBodegaPK), bodegaNombre: texto(r.bodegaNombre), nSucursalPK: texto(r.nSucursalPK), sucursalNombre: texto(r.sucursalNombre),
    qty: Number(r.qty) || 0, subtotal: Number(r.subtotal) || 0, total: Number(r.total) || 0, lineas: Number(r.lineas) || 0,
    qtyPromocion: r.qtyPromocion == null ? "" : Number(r.qtyPromocion) || 0,
  };
}
function normalizarProduccion(r) {
  return {
    mes: texto(r.mes), nProductoPK: texto(r.nProductoPK), cCodigo: texto(r.cCodigo), cDescripcion: texto(r.cDescripcion),
    qtyPedida: Number(r.qtyPedida) || 0, qtyProducida: Number(r.qtyProducida) || 0, qtyCancelada: Number(r.qtyCancelada) || 0,
    qtySalida: Number(r.qtySalida) || 0, lineas: Number(r.lineas) || 0, lineasAbiertas: r.lineasAbiertas == null ? "" : Number(r.lineasAbiertas) || 0,
  };
}

/**
 * Lee ventas y producción desde la base. `ejecutor` = { query(sql, params), close() }
 * (real con crearEjecutorMssql o simulado en pruebas).
 * desde: "YYYY-MM-01" (inicio de la serie); hoy: Date (para acotar ventas con fecha futura).
 */
async function leerBase({ cfg, ejecutor, desde = "2024-01-01", hoy = new Date(), log = () => {} }) {
  const q = async (sql, params) => { asegurarSoloLectura(sql); return ejecutor.query(sql, params); };
  const advertencias = [];
  const db = (await q(SQL.baseActual))[0]?.db;
  if (cfg?.database && db !== cfg.database) throw new Error(`Modo base: la conexión abrió la base "${db}" y se esperaba "${cfg.database}".`);
  const manana = new Date(hoy.getFullYear(), hoy.getMonth(), hoy.getDate() + 1);
  const limite = `${manana.getFullYear()}-${String(manana.getMonth() + 1).padStart(2, "0")}-${String(manana.getDate()).padStart(2, "0")}T00:00:00`;
  const maxVenta = fechaIso((await q(SQL.maxVenta, { limite }))[0]?.maxVenta);
  if (!maxVenta) throw new Error("Modo base: dbo.Ventas no tiene ventas (MAX(dCreada) vacío).");
  const fechaCorte = maxVenta.slice(0, 10);
  const desdeMes = desde.slice(0, 7);
  const meses = mesesEntre(desdeMes, fechaCorte.slice(0, 7));
  const ventas = [];
  for (const mes of meses) {
    const hastaMes = mesSiguiente(mes);
    const t0 = Date.now();
    const rows = await q(SQL.ventasMes, { desde: `${mes}-01T00:00:00`, hasta: `${hastaMes}-01T00:00:00` });
    ventas.push(...rows.map(normalizarVenta));
    log(`  ventas ${mes}: ${rows.length} filas (${((Date.now() - t0) / 1000).toFixed(1)} s)`);
  }
  const produccion = (await q(SQL.produccion, { desde: `${desdeMes}-01T00:00:00` })).map(normalizarProduccion);
  const estado = (await q(SQL.produccionEstado, { desde: `${desdeMes}-01T00:00:00` }))[0] || {};
  const maxPeticion = fechaIso(estado.maxPeticion);
  if (maxPeticion && maxPeticion.slice(0, 10) < addDays(fechaCorte, -2)) {
    advertencias.push(`La última orden de producción (${maxPeticion.slice(0, 10)}) es más vieja que la última venta (${fechaCorte}): revisa que Produccion se haya recopiado.`);
  }
  if (Number(estado.ordenesAbiertas) > 0) {
    advertencias.push(`${estado.ordenesAbiertas} órdenes de producción siguen abiertas (Peticion / En Proceso): sus piezas producidas pueden cambiar en la próxima copia.`);
  }
  return {
    modo: "base",
    fuente: describirConexion(cfg || { server: "?", port: "?", database: db, user: "?" }),
    ventas,
    produccion,
    fechaCorte,
    origenCorte: "MAX(dbo.Ventas.dCreada)",
    maxVentaDCreada: maxVenta,
    produccionEstado: { ordenes: Number(estado.ordenes) || 0, ordenesAbiertas: Number(estado.ordenesAbiertas) || 0, maxPeticion },
    archivosLeidos: [],
    advertencias,
  };
}

function addDays(fecha, dias) {
  const [y, m, d] = fecha.split("-").map(Number);
  const t = new Date(Date.UTC(y, m - 1, d + dias));
  return t.toISOString().slice(0, 10);
}

module.exports = {
  SQL, TABLAS_PERMITIDAS, VARIABLES_REQUERIDAS, ErrorSoloLectura,
  asegurarSoloLectura, configDesdeEntorno, describirConexion, ocultarSecretos, crearEjecutorMssql, leerBase, mesesEntre,
};
