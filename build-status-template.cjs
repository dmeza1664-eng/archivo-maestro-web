const fs = require("fs");
const path = require("path");
const Module = require("module");
const esbuild = require("esbuild");
const XLSX = require("xlsx");

const ROOT = __dirname;
const DOWNLOADS = path.resolve(ROOT, "..");
const OUTPUT = path.join(DOWNLOADS, "ESTATUS_PRODUCTOS_PARA_LLENAR.xlsx");
const JULY_SALES = "C:/Users/X13/Documents/ventas por mes/ventas julio.xlsx";
const JUNE_CLOSE = "C:/Users/X13/Documents/ventas por mes/ventas junio.xlsx";

const SOURCES = [
  "Venta de diciembre 2024.xlsx",
  "VENTA AÑO 2025 - ANGEL.xlsx",
  "VENTA ENERO 2026.xlsx",
  "VENTAS FEEEBRERO 2026.xlsx",
  "MARZO VENTAS.xlsx",
  "venta abril 2026.xlsx",
  "VENTAS DE MAYO Y JUNIO - ANGEL.xlsx",
];

const SEASONAL_PATTERNS = [
  /MUERTO/, /CALAVERA/, /BRUJA/, /CALABAZA/, /HUESITOS/, /ROSCA/, /REYES/,
  /NAVID/, /SANTA/, /GUADALUPE/, /MADRES/, /\bMAMA\b/, /\bPADRE\b/, /CORAZON/,
];

const ACCESSORY_PATTERNS = [
  /BENGALA/, /\bVELA\b/, /VELADORA/, /SERPENTINA/, /KIT DE PLATOS/, /LETRERO/,
  /BOLSA ECOLOGICA/, /CHUNCHES/, /ACETATO/, /OTROS VARIOS/, /\bMONEDA\b/,
  /ENCENDEDOR/, /GLOBO/, /MOÑO/, /LISTON/, /VELA MUSICAL/,
];

function normalized(value) {
  return String(value || "")
    .toUpperCase()
    .normalize("NFD")
    .replace(/[\u0300-\u036f]/g, "")
    .replace(/[.,/\\_-]+/g, " ")
    .replace(/\s+/g, " ")
    .trim();
}

function matchKey(value) {
  return normalized(value)
    .replace(/\b(PASTEL|TARTA|PANQUE|PAY|DE|DEL|LA|EL)\b/g, " ")
    .replace(/\bGRANDE\b/g, "GDE")
    .replace(/\bMEDIANO\b/g, "MED")
    .replace(/\bCHICO\b/g, "CH")
    .replace(/\s+/g, " ")
    .trim();
}

function findDownload(search) {
  const target = normalized(search);
  const match = fs.readdirSync(DOWNLOADS).find((name) => normalized(name) === target);
  if (!match) throw new Error(`No se encontro ${search}`);
  return path.join(DOWNLOADS, match);
}

function monthKey(value) {
  const date = value instanceof Date ? value : new Date(value);
  if (Number.isNaN(date.getTime())) return "";
  return `${date.getFullYear()}-${String(date.getMonth() + 1).padStart(2, "0")}`;
}

function isExcluded(product) {
  const value = normalized(product);
  return /\b(REBANADA|REBANADAS|REB|RBN)\b/.test(value) || /\b(PROMO|PROMOCION|PROMOCIONAL)\b/.test(value);
}

function isAccessory(product) {
  const value = normalized(product);
  return ACCESSORY_PATTERNS.some((pattern) => pattern.test(value));
}

async function loadAppFunctions() {
  const built = await esbuild.build({
    entryPoints: [path.join(ROOT, "App.jsx")],
    bundle: true, platform: "node", format: "cjs", write: false,
    loader: { ".css": "text" },
    define: { "import.meta.env.VITE_API_URL": JSON.stringify("") },
    logLevel: "silent",
  });
  const appModule = new Module("status-template");
  appModule.filename = path.join(ROOT, "status-template.bundle.cjs");
  appModule.paths = module.paths;
  appModule._compile(built.outputFiles[0].text, appModule.filename);
  return appModule.exports;
}

function styleSheet(sheet, widths) {
  sheet["!cols"] = widths.map((wch) => ({ wch }));
  sheet["!freeze"] = { xSplit: 0, ySplit: 1, topLeftCell: "A2", activePane: "bottomRight", state: "frozen" };
  const range = XLSX.utils.decode_range(sheet["!ref"] || "A1");
  for (let col = range.s.c; col <= range.e.c; col += 1) {
    const cell = sheet[XLSX.utils.encode_cell({ r: 0, c: col })];
    if (cell) cell.s = { font: { bold: true, color: { rgb: "FFFFFF" } }, fill: { fgColor: { rgb: "16324F" } } };
  }
  sheet["!autofilter"] = { ref: sheet["!ref"] };
}

async function main() {
  const app = await loadAppFunctions();
  const stockRows = app.parseStock(XLSX.readFile(findDownload("STOCK IDEAL SUCURSALES.xlsx"), { cellDates: true }));
  const catalog = stockRows
    .map((row) => ({ producto: normalized(row.producto), original: row.productoOriginal, orden: row.orden }))
    .filter((row) => row.producto && !isExcluded(row.producto));

  const direct = new Set(catalog.map((row) => row.producto));
  const byMatchKey = new Map();
  for (const row of catalog) {
    const key = matchKey(row.producto);
    if (!byMatchKey.has(key)) byMatchKey.set(key, []);
    byMatchKey.get(key).push(row.producto);
  }
  const toOfficial = (raw) => {
    const source = normalized(raw);
    if (direct.has(source)) return source;
    const candidates = byMatchKey.get(matchKey(source)) || [];
    return candidates.length === 1 ? candidates[0] : "";
  };

  const read = (filePath) =>
    app.parseSalesOrReturns(XLSX.readFile(filePath, { cellDates: true }), "ventas", path.basename(filePath));

  const perFile = [
    ...SOURCES.map((name) => ({ name, rows: read(findDownload(name)) })),
    { name: "ventas junio.xlsx", rows: read(JUNE_CLOSE) },
    { name: "ventas julio.xlsx", rows: read(JULY_SALES) },
  ];

  // Una sola fuente por mes: la de mayor volumen, para no duplicar junio.
  const monthTotalsByFile = new Map();
  for (const file of perFile) {
    const perMonth = new Map();
    for (const row of file.rows) {
      const key = monthKey(row.fecha);
      if (!key) continue;
      perMonth.set(key, (perMonth.get(key) || 0) + Number(row.cantidad || 0));
    }
    monthTotalsByFile.set(file.name, perMonth);
  }
  const monthOwner = new Map();
  for (const [fileName, perMonth] of monthTotalsByFile.entries()) {
    for (const [month, total] of perMonth.entries()) {
      const current = monthOwner.get(month);
      if (!current || total > current.total) monthOwner.set(month, { fileName, total });
    }
  }

  const salesByProductMonth = new Map();
  const rawByProductMonth = new Map();
  for (const file of perFile) {
    for (const row of file.rows) {
      const month = monthKey(row.fecha);
      if (!month || monthOwner.get(month)?.fileName !== file.name) continue;
      const quantity = Number(row.cantidad || 0);
      const raw = normalized(row.producto);
      if (!raw) continue;
      const rawKey = `${raw}|${month}`;
      rawByProductMonth.set(rawKey, (rawByProductMonth.get(rawKey) || 0) + quantity);
      const official = toOfficial(raw);
      if (!official) continue;
      const key = `${official}|${month}`;
      salesByProductMonth.set(key, (salesByProductMonth.get(key) || 0) + quantity);
    }
  }

  const allMonths = [...monthOwner.keys()].sort();
  const recentMonths = allMonths.slice(-3);
  const latestMonth = allMonths.at(-1);

  const statusRows = catalog.map((row) => {
    const monthsWithSales = allMonths.filter((month) => (salesByProductMonth.get(`${row.producto}|${month}`) || 0) > 0);
    const lastSale = monthsWithSales.at(-1) || "";
    const dormant = lastSale ? allMonths.filter((month) => month > lastSale).length : allMonths.length;
    const recent = recentMonths.map((month) => salesByProductMonth.get(`${row.producto}|${month}`) || 0);
    const recentAverage = recent.reduce((sum, value) => sum + value, 0) / Math.max(1, recent.length);

    // Un producto de temporada esta inactivo casi todo el año sin estar dado de baja.
    const sellingCalendarMonths = new Set(monthsWithSales.map((month) => month.slice(5)));
    const currentCalendarMonth = String(latestMonth || "").slice(5);
    const looksSeasonalByPattern =
      sellingCalendarMonths.size > 0 &&
      sellingCalendarMonths.size <= 4 &&
      !sellingCalendarMonths.has(currentCalendarMonth);
    const looksSeasonalByName = SEASONAL_PATTERNS.some((pattern) => pattern.test(row.producto));
    // La temporada solo se confirma si el nombre lo indica o si el producto repitio el mismo
    // mes calendario en dos años distintos. Con un solo año no se distingue de una baja.
    const repeatedCalendarMonth = [...sellingCalendarMonths].some(
      (calendarMonth) =>
        new Set(monthsWithSales.filter((month) => month.endsWith(calendarMonth)).map((month) => month.slice(0, 4))).size >= 2
    );
    const seasonalConfirmed = looksSeasonalByName || repeatedCalendarMonth;

    let suggestion = "ACTIVO";
    let note = "";
    if (!lastSale) {
      suggestion = "SIN COINCIDENCIA";
      note = "Ninguna venta se pudo ligar a este nombre. Puede ser un encabezado del catalogo o un nombre distinto en ventas";
    } else if (dormant >= 2 && looksSeasonalByPattern && seasonalConfirmed) {
      suggestion = "ESTACIONAL";
      note = `Vende solo en los meses ${[...sellingCalendarMonths].sort().join(", ")}; ultima venta ${lastSale}`;
    } else if (dormant >= 2 && looksSeasonalByPattern) {
      suggestion = "REVISAR";
      note = `Solo vendio en los meses ${[...sellingCalendarMonths].sort().join(", ")} y no hay un segundo año para confirmar. Dinos si es temporada o baja`;
    } else if (dormant >= 2) {
      suggestion = "BAJA";
      note = `Sin ventas desde ${lastSale}, ${dormant} meses inactivo`;
    } else if (dormant === 1) {
      suggestion = "REVISAR";
      note = `No vendio en ${latestMonth}; ultima venta ${lastSale}`;
    } else if (recentAverage > 0 && recentAverage < 8) {
      suggestion = "BAJO PEDIDO";
      note = `Baja rotacion: ${recentAverage.toFixed(1)} piezas por mes`;
    }

    return {
      Producto: row.original || row.producto,
      Estatus: suggestion,
      Desde: suggestion === "BAJA" && lastSale ? lastSale : "",
      "Volumen esperado": "",
      "Sugerencia del sistema": suggestion,
      "Ultima venta": lastSale || "nunca",
      "Meses sin vender": dormant,
      [recentMonths[0] || "mes-1"]: Math.round(recent[0] || 0),
      [recentMonths[1] || "mes-2"]: Math.round(recent[1] || 0),
      [recentMonths[2] || "mes-3"]: Math.round(recent[2] || 0),
      "Promedio 3 meses": Number(recentAverage.toFixed(1)),
      Nota: note,
    };
  });

  const orderRank = { REVISAR: 0, BAJA: 1, "SIN COINCIDENCIA": 2, ESTACIONAL: 3, "BAJO PEDIDO": 4, ACTIVO: 5 };
  statusRows.sort(
    (a, b) =>
      (orderRank[a["Sugerencia del sistema"]] ?? 9) - (orderRank[b["Sugerencia del sistema"]] ?? 9) ||
      String(a.Producto).localeCompare(String(b.Producto), "es")
  );

  // Claves que vendieron pero no existen en el catalogo.
  const outsideCatalog = new Map();
  for (const [key, quantity] of rawByProductMonth.entries()) {
    const [raw, month] = key.split("|");
    if (!month.startsWith("2026") || quantity <= 0) continue;
    if (isExcluded(raw) || toOfficial(raw)) continue;
    const current = outsideCatalog.get(raw) || { total: 0, months: new Map() };
    current.total += quantity;
    current.months.set(month, (current.months.get(month) || 0) + quantity);
    outsideCatalog.set(raw, current);
  }

  const newRows = [...outsideCatalog.entries()]
    .filter(([raw]) => !isAccessory(raw))
    .sort((a, b) => b[1].total - a[1].total)
    .map(([raw, data]) => ({
      "Clave en ventas": raw,
      "Entra al catalogo": "",
      "Nombre oficial propuesto": "",
      "Stock ideal sugerido": "",
      "Piezas 2026": Math.round(data.total),
      [recentMonths[0] || "mes-1"]: Math.round(data.months.get(recentMonths[0]) || 0),
      [recentMonths[1] || "mes-2"]: Math.round(data.months.get(recentMonths[1]) || 0),
      [recentMonths[2] || "mes-3"]: Math.round(data.months.get(recentMonths[2]) || 0),
    }));

  const accessoryRows = [...outsideCatalog.entries()]
    .filter(([raw]) => isAccessory(raw))
    .sort((a, b) => b[1].total - a[1].total)
    .map(([raw, data]) => ({
      "Clave en ventas": raw,
      "Piezas 2026": Math.round(data.total),
      Clasificacion: "Accesorio o complemento",
      "Confirmas que queda fuera": "",
    }));

  const instructions = [
    { Campo: "Como se llena", Detalle: "Solo corrige la columna Estatus donde no estes de acuerdo con la sugerencia. Lo demas es informacion de apoyo." },
    { Campo: "ACTIVO", Detalle: "El producto se sigue vendiendo y debe pronosticarse normal." },
    { Campo: "BAJA", Detalle: "Ya no se vende. El sistema dejara de pronosticarlo. Pon el mes en la columna Desde." },
    { Campo: "NUEVO", Detalle: "Producto que apenas arranca y no tiene historial suficiente. Pon el volumen mensual esperado." },
    { Campo: "BAJO PEDIDO", Detalle: "Se produce solo cuando se pide. No entra al calendario de produccion regular." },
    { Campo: "ESTACIONAL", Detalle: "Solo se vende en ciertos meses, como pan de muerto o rosca. No esta dado de baja: el sistema solo debe pronosticarlo en su temporada." },
    { Campo: "REVISAR", Detalle: "El sistema no pudo decidir. Necesita tu confirmacion: puede ser BAJA o ACTIVO con venta irregular." },
    { Campo: "SIN COINCIDENCIA", Detalle: "Ninguna venta se pudo ligar a este nombre. Dime si es un encabezado del catalogo que no es producto, o con que nombre aparece en ventas." },
    { Campo: "Prioridad", Detalle: "Empieza por REVISAR, BAJA y SIN COINCIDENCIA; estan ordenadas primero. ESTACIONAL y ACTIVO normalmente ya estan bien." },
    { Campo: "Hoja Candidatos nuevos", Detalle: "Productos que vendieron en 2026 pero no existen en el stock ideal. Marca SI o NO en Entra al catalogo." },
    { Campo: "Hoja Accesorios", Detalle: "Bengalas, velas y kits. Confirma que quedan fuera del pronostico de produccion." },
    { Campo: "Devolucion", Detalle: "Guarda el archivo con el mismo nombre y avisame; con eso ajusto el pronostico." },
  ];

  const workbook = XLSX.utils.book_new();
  const statusSheet = XLSX.utils.json_to_sheet(statusRows);
  styleSheet(statusSheet, [38, 14, 10, 18, 22, 14, 17, 12, 12, 12, 17, 46]);
  XLSX.utils.book_append_sheet(workbook, statusSheet, "Estatus productos");

  const newSheet = XLSX.utils.json_to_sheet(newRows);
  styleSheet(newSheet, [40, 18, 30, 20, 13, 12, 12, 12]);
  XLSX.utils.book_append_sheet(workbook, newSheet, "Candidatos nuevos");

  const accessorySheet = XLSX.utils.json_to_sheet(accessoryRows);
  styleSheet(accessorySheet, [40, 13, 26, 24]);
  XLSX.utils.book_append_sheet(workbook, accessorySheet, "Accesorios");

  const instructionSheet = XLSX.utils.json_to_sheet(instructions);
  styleSheet(instructionSheet, [24, 110]);
  XLSX.utils.book_append_sheet(workbook, instructionSheet, "Instrucciones");

  XLSX.writeFile(workbook, OUTPUT);

  const counts = statusRows.reduce((map, row) => {
    map[row["Sugerencia del sistema"]] = (map[row["Sugerencia del sistema"]] || 0) + 1;
    return map;
  }, {});
  console.log(JSON.stringify({
    archivo: OUTPUT,
    productosCatalogo: statusRows.length,
    sugerencias: counts,
    candidatosNuevos: newRows.length,
    accesorios: accessoryRows.length,
    mesesDeApoyo: recentMonths,
  }, null, 2));
}

main().catch((error) => { console.error(error.message); process.exit(1); });
