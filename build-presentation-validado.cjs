const fs = require("fs");
const path = require("path");
const Module = require("module");
const esbuild = require("esbuild");
const XLSX = require("xlsx");

const ROOT = __dirname;
const DOWNLOADS = path.resolve(ROOT, "..");
const OUTPUT = path.join(DOWNLOADS, "PRESENTACION_JEFE_MAYO_JUNIO_MODELO_VALIDADO.xlsx");
const TARGET_MONTHS = [
  { key: "2026-05", label: "Mayo 2026" },
  { key: "2026-06", label: "Junio 2026" },
];
const TOLERANCE = 15;

const SOURCE_FILES = [
  "Venta de diciembre 2024.xlsx",
  "VENTA AÑO 2025 - ANGEL.xlsx",
  "VENTA ENERO 2026.xlsx",
  "VENTAS FEEEBRERO 2026.xlsx",
  "MARZO VENTAS.xlsx",
  "venta abril 2026.xlsx",
  "VENTAS DE MAYO Y JUNIO - ANGEL.xlsx",
];

function normalized(value) {
  return String(value)
    .toUpperCase()
    .normalize("NFD")
    .replace(/[\u0300-\u036f]/g, "")
    .replace(/\s+/g, " ")
    .trim();
}

function findDownload(search) {
  const target = normalized(search);
  const match = fs.readdirSync(DOWNLOADS).find((name) => normalized(name) === target);
  if (!match) throw new Error(`No se encontro el archivo de origen: ${search}`);
  return path.join(DOWNLOADS, match);
}

function monthKey(value) {
  const date = value instanceof Date ? value : new Date(value);
  if (Number.isNaN(date.getTime())) return "";
  return `${date.getFullYear()}-${String(date.getMonth() + 1).padStart(2, "0")}`;
}

function category(product) {
  const value = normalized(product);
  if (value.includes("GELATINA")) return "Gelatinas";
  if (value.includes("GALLETA")) return "Galletas";
  if (/\b(GDE|GRANDE)\b/.test(value)) return "Pasteles grandes";
  if (/\b(MED|MEDIANO)\b/.test(value) && !value.includes("MINI")) return "Pasteles medianos";
  if (/\b(CH|CHICO)\b/.test(value)) return "Pasteles chicos";
  if (value.includes("MINI")) return "Mini medianos";
  if (value.includes("BOLLO") || value.includes("PAN")) return "Pan";
  return "Otros";
}

function round(value) {
  return Number(Number(value || 0).toFixed(2));
}

function metrics(rows) {
  const actual = rows.reduce((sum, row) => sum + row.actual, 0);
  const forecast = rows.reduce((sum, row) => sum + row.forecast, 0);
  const absoluteError = rows.reduce((sum, row) => sum + Math.abs(row.actual - row.forecast), 0);
  return {
    products: rows.length,
    actual: round(actual),
    forecast: round(forecast),
    bias: round(actual - forecast),
    wape: actual > 0 ? round((absoluteError / actual) * 100) : null,
    accuracy: actual > 0 ? round(100 - (absoluteError / actual) * 100) : null,
    mae: rows.length ? round(absoluteError / rows.length) : null,
    inside: rows.filter((row) => Math.abs(row.actual - row.forecast) <= TOLERANCE).length,
  };
}

// Reutiliza el modelo real de App.jsx en lugar de reimplementarlo: el comparativo
// debe medir exactamente lo que produce la aplicacion.
async function loadAppFunctions() {
  const built = await esbuild.build({
    entryPoints: [path.join(ROOT, "App.jsx")],
    bundle: true,
    platform: "node",
    format: "cjs",
    write: false,
    loader: { ".css": "text" },
    define: { "import.meta.env.VITE_API_URL": JSON.stringify("") },
    logLevel: "silent",
  });
  const appModule = new Module("presentacion-validada");
  appModule.filename = path.join(ROOT, "presentacion-validada.cjs");
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

function setFormula(sheet, address, formula, value) {
  sheet[address] = { t: typeof value === "string" ? "s" : "n", f: formula };
  if (value !== undefined && value !== "") sheet[address].v = value;
}

function statusFor(actual, forecast) {
  if (actual === 0 && forecast < 1) return "Sin movimiento";
  return Math.abs(actual - forecast) <= TOLERANCE ? "Dentro de +/-15 piezas" : "Fuera de +/-15 piezas";
}

async function main() {
  const app = await loadAppFunctions();
  const stockRows = app.parseStock(XLSX.readFile(findDownload("STOCK IDEAL SUCURSALES.xlsx"), { cellDates: true }));

  const allSales = SOURCE_FILES.flatMap((sourceName) => {
    const filePath = findDownload(sourceName);
    const workbook = XLSX.readFile(filePath, { cellDates: true });
    return app.parseSalesOrReturns(workbook, "ventas", path.basename(filePath));
  });

  const monthResults = TARGET_MONTHS.map(({ key, label }) => {
    const historical = app.filterVentasBeforeMonth(allSales, key);
    const actualMap = new Map();
    for (const sale of allSales.filter((row) => monthKey(row.fecha) === key)) {
      const product = normalized(sale.producto);
      actualMap.set(product, (actualMap.get(product) || 0) + Number(sale.cantidad || 0));
    }

    const forecastRows = app.calculateForecast({
      stockRows,
      historicalVentas: historical,
      bajas: [],
      existencias: [],
      realProduction: [],
      selectedMonth: key,
      dailyBufferPct: 0,
    });

    const rows = forecastRows.map((row) => {
      const product = normalized(row.producto);
      const actual = actualMap.get(product) || 0;
      const forecast = Number(row.pronosticoVenta || 0);
      return {
        product: row.producto,
        category: category(row.producto),
        actual,
        forecast,
        method: row.metodoPronostico,
        history: row.mesesUsados,
      };
    });

    const historyMonths = [...new Set(historical.map((row) => monthKey(row.fecha)).filter(Boolean))].sort();
    return { key, label, rows, overall: metrics(rows), historyMonths };
  });

  const workbook = XLSX.utils.book_new();

  const summaryRows = [
    { Indicador: "Modelo", Valor: "Modelo validado de la aplicacion (categorySeasonal)" },
    { Indicador: "Regla de control", Valor: "Cada mes se pronostica usando unicamente historico anterior a ese mes" },
    { Indicador: "Universo", Valor: "Catalogo de stock ideal, sin rebanadas ni promocionales" },
    { Indicador: "Tolerancia operativa", Valor: `+/- ${TOLERANCE} piezas` },
  ];
  for (const result of monthResults) {
    summaryRows.push(
      { Indicador: `--- ${result.label} ---`, Valor: "" },
      { Indicador: "Historico usado", Valor: result.historyMonths.join(", ") },
      { Indicador: "Productos comparados", Valor: result.overall.products },
      { Indicador: "Venta real (piezas)", Valor: result.overall.actual },
      { Indicador: "Pronostico (piezas)", Valor: result.overall.forecast },
      { Indicador: "Diferencia total (real - pronostico)", Valor: result.overall.bias },
      { Indicador: "WAPE (error ponderado)", Valor: `${result.overall.wape}%` },
      { Indicador: "Exactitud agregada", Valor: `${result.overall.accuracy}%` },
      { Indicador: "MAE (error medio por producto)", Valor: result.overall.mae },
      { Indicador: `Productos dentro de +/-${TOLERANCE} piezas`, Valor: `${result.overall.inside} de ${result.overall.products}` }
    );
  }

  const combined = monthResults.flatMap((result) => result.rows);
  const combinedMetrics = metrics(combined);
  summaryRows.push(
    { Indicador: "--- Mayo + Junio ---", Valor: "" },
    { Indicador: "Venta real total", Valor: combinedMetrics.actual },
    { Indicador: "Pronostico total", Valor: combinedMetrics.forecast },
    { Indicador: "WAPE ponderado", Valor: `${combinedMetrics.wape}%` },
    { Indicador: "Exactitud agregada", Valor: `${combinedMetrics.accuracy}%` },
    { Indicador: `Comparaciones dentro de +/-${TOLERANCE}`, Valor: `${combinedMetrics.inside} de ${combinedMetrics.products}` }
  );

  const summarySheet = XLSX.utils.json_to_sheet(summaryRows);
  styleSheet(summarySheet, [42, 62]);
  XLSX.utils.book_append_sheet(workbook, summarySheet, "Resumen ejecutivo");

  for (const result of monthResults) {
    const detailRows = result.rows
      .slice()
      .sort((a, b) => Math.abs(b.actual - b.forecast) - Math.abs(a.actual - a.forecast))
      .map((row) => ({
        Producto: row.product,
        Categoria: row.category,
        "Venta real": round(row.actual),
        "Pronostico venta": round(row.forecast),
        "Diferencia (real - pronostico)": round(row.actual - row.forecast),
        "Diferencia absoluta": round(Math.abs(row.actual - row.forecast)),
        "Error %": row.actual > 0 ? round((Math.abs(row.actual - row.forecast) / row.actual) * 100) : "",
        "Precision %": row.actual > 0 ? round(100 - (Math.abs(row.actual - row.forecast) / row.actual) * 100) : "",
        Estado: statusFor(row.actual, row.forecast),
        "Metodo elegido": row.method,
        "Meses usados": row.history,
      }));

    const sheet = XLSX.utils.json_to_sheet(detailRows);
    for (let rowNumber = 2; rowNumber <= detailRows.length + 1; rowNumber += 1) {
      const source = detailRows[rowNumber - 2];
      setFormula(sheet, `E${rowNumber}`, `C${rowNumber}-D${rowNumber}`, source["Diferencia (real - pronostico)"]);
      setFormula(sheet, `F${rowNumber}`, `ABS(E${rowNumber})`, source["Diferencia absoluta"]);
      setFormula(sheet, `G${rowNumber}`, `IF(C${rowNumber}>0,F${rowNumber}/C${rowNumber}*100,"")`, source["Error %"]);
      setFormula(sheet, `H${rowNumber}`, `IF(C${rowNumber}>0,100-G${rowNumber},"")`, source["Precision %"]);
      setFormula(
        sheet,
        `I${rowNumber}`,
        `IF(AND(C${rowNumber}=0,D${rowNumber}<1),"Sin movimiento",IF(F${rowNumber}<=${TOLERANCE},"Dentro de +/-${TOLERANCE} piezas","Fuera de +/-${TOLERANCE} piezas"))`,
        source.Estado
      );
    }
    styleSheet(sheet, [38, 20, 14, 18, 26, 20, 12, 14, 26, 30, 30]);
    XLSX.utils.book_append_sheet(workbook, sheet, result.label);
  }

  const categoryRows = [];
  for (const result of monthResults) {
    for (const categoryName of [...new Set(result.rows.map((row) => row.category))].sort()) {
      const stats = metrics(result.rows.filter((row) => row.category === categoryName));
      categoryRows.push({
        Mes: result.label,
        Categoria: categoryName,
        Productos: stats.products,
        "Venta real": stats.actual,
        Pronostico: stats.forecast,
        Diferencia: stats.bias,
        "WAPE %": stats.wape,
        "Exactitud %": stats.accuracy,
        [`Dentro +/-${TOLERANCE}`]: stats.inside,
      });
    }
  }
  const categorySheet = XLSX.utils.json_to_sheet(categoryRows);
  styleSheet(categorySheet, [14, 22, 12, 14, 14, 14, 12, 14, 16]);
  XLSX.utils.book_append_sheet(workbook, categorySheet, "Por categoria");

  const methodRows = [];
  for (const result of monthResults) {
    const counts = result.rows.reduce((map, row) => map.set(row.method, (map.get(row.method) || 0) + 1), new Map());
    for (const [method, count] of [...counts.entries()].sort((a, b) => b[1] - a[1])) {
      methodRows.push({ Mes: result.label, "Metodo elegido": method, Productos: count });
    }
  }
  const methodSheet = XLSX.utils.json_to_sheet(methodRows);
  styleSheet(methodSheet, [14, 38, 12]);
  XLSX.utils.book_append_sheet(workbook, methodSheet, "Metodos usados");

  const notesRows = [
    { Tema: "Como se calcula", Detalle: "Para cada producto se prueban varias bases (ultimo mes, promedios ponderados, dia de semana, mismo mes del año anterior ajustado por crecimiento) y se elige la que menor error tuvo en meses previos." },
    { Tema: "Sin contaminacion", Detalle: "El pronostico de mayo usa solo historico hasta abril. El de junio usa solo historico hasta mayo. La venta real del mes se usa unicamente para medir." },
    { Tema: "WAPE", Detalle: "Suma de errores absolutos dividida entre la venta real total. Es el indicador principal porque pondera por volumen." },
    { Tema: "Exactitud agregada", Detalle: "100% menos WAPE. Es la lectura ejecutiva del acierto del modelo." },
    { Tema: "MAE", Detalle: "Error promedio en piezas por producto. Util para dimensionar el ajuste operativo." },
    { Tema: "Tolerancia", Detalle: `Se considera dentro de rango cuando la diferencia absoluta es de maximo ${TOLERANCE} piezas.` },
    { Tema: "Universo", Detalle: "Solo productos del catalogo de stock ideal. Se excluyen rebanadas y promocionales porque siguen planeacion manual." },
    { Tema: "Reproducible", Detalle: "node build-presentation-validado.cjs" },
  ];
  const notesSheet = XLSX.utils.json_to_sheet(notesRows);
  styleSheet(notesSheet, [24, 120]);
  XLSX.utils.book_append_sheet(workbook, notesSheet, "Metodologia");

  XLSX.writeFile(workbook, OUTPUT);

  console.log(JSON.stringify({
    output: OUTPUT,
    ventasLeidas: allSales.length,
    productosCatalogo: stockRows.length,
    meses: monthResults.map((result) => ({ mes: result.label, ...result.overall })),
    combinado: combinedMetrics,
  }, null, 2));
}

main().catch((error) => {
  console.error(error.message);
  process.exit(1);
});
