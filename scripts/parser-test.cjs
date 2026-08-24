const path = require("path");
const Module = require("module");
const esbuild = require("esbuild");
const XLSX = require("xlsx");

async function loadAppFunctions() {
  const built = await esbuild.build({
    entryPoints: [path.join(__dirname, "..", "App.jsx")],
    bundle: true,
    platform: "node",
    format: "cjs",
    write: false,
    loader: { ".css": "text" },
    define: {
      "import.meta.env.VITE_API_URL": JSON.stringify(""),
      "import.meta.env.DEV": "false",
      "import.meta.env.PROD": "true",
    },
    logLevel: "silent",
  });
  const appModule = new Module("parser-test");
  appModule.filename = path.join(__dirname, "parser-test.bundle.cjs");
  appModule.paths = module.paths;
  appModule._compile(built.outputFiles[0].text, appModule.filename);
  return appModule.exports;
}

function workbookFromSheets(sheets) {
  const workbook = XLSX.utils.book_new();
  for (const [name, rows] of Object.entries(sheets)) {
    XLSX.utils.book_append_sheet(workbook, XLSX.utils.aoa_to_sheet(rows), name);
  }
  return workbook;
}

function assert(condition, message) {
  if (!condition) throw new Error(message);
}

function monthKey(row) {
  if (row.fecha instanceof Date) {
    return `${row.fecha.getFullYear()}-${String(row.fecha.getMonth() + 1).padStart(2, "0")}`;
  }
  return String(row.fecha || "").slice(0, 7);
}

async function main() {
  const {
    monthsNamedInFileName,
    parseBajasReport,
    parseBajasSummaryWorkbook,
    parseExistencias,
    parseMonthlySummaryWorkbook,
    parseSalesOrReturns,
    parseStock,
    resolveCanonicalMonthSources,
    computeAnnualGrowthFactor,
  } = await loadAppFunctions();

  assert(monthsNamedInFileName("VENTAS DE MAYO Y JUNIO 2026.xlsx").join(",") === "2026-05,2026-06", "el combinado debe nombrar mayo y junio");
  assert(monthsNamedInFileName("ventas junio.xlsx")[0] === "2026-06", "el cierre de junio debe ser un mes");

  const combined = { name: "VENTAS DE MAYO Y JUNIO.xlsx", rows: [
    { fecha: "2026-05-03", producto: "BOLILLO", cantidad: 10 },
    { fecha: "2026-06-04", producto: "BOLILLO", cantidad: 20 },
  ] };
  const juneClose = { name: "ventas junio.xlsx", rows: [
    { fecha: "2026-06-04", producto: "BOLILLO", cantidad: 18 },
  ] };
  const juneFirst = resolveCanonicalMonthSources([juneClose, combined]);
  const combinedAfterJuneFirst = juneFirst.entries.find((entry) => entry.name === combined.name);
  const juneAfterJuneFirst = juneFirst.entries.find((entry) => entry.name === juneClose.name);
  assert(combinedAfterJuneFirst.rows.every((row) => monthKey(row) !== "2026-06"), "el cierre dedicado debe quitar junio del combinado aunque se cargue después");
  assert(juneAfterJuneFirst.rows.length === 1, "el cierre dedicado de junio se conserva");
  assert(juneFirst.decisions[0].winner === "ventas junio.xlsx", "el ganador debe ser el archivo de un solo mes");

  const marchClose = resolveCanonicalMonthSources([
    { name: "VENTAS FEBRERO Y MARZO.xlsx", rows: [
      { fecha: "2026-02-01", producto: "BOLILLO", cantidad: 5 },
      { fecha: "2026-03-01", producto: "BOLILLO", cantidad: 9 },
    ] },
    { name: "cierre marzo.xlsx", rows: [{ fecha: "2026-03-01", producto: "BOLILLO", cantidad: 7 }] },
  ]);
  assert(marchClose.decisions[0].month === "2026-03", "el conflicto genérico no es solo junio");
  assert(marchClose.decisions[0].winner === "cierre marzo.xlsx", "marzo dedicado gana al combinado");

  const overridden = resolveCanonicalMonthSources([combined, juneClose], { "2026-06": combined.name });
  const combinedKept = overridden.entries.find((entry) => entry.name === combined.name);
  assert(combinedKept.rows.some((row) => monthKey(row) === "2026-06"), "el override debe conservar junio del combinado");

  const stockWorkbook = workbookFromSheets({
    "OTRA HOJA": [
      ["PRODUCTO", "TOTAL GRAL SUC"],
      ["PAN MAL", 9999],
    ],
    "TOTAL A TENER SUC.(EXIST.+DIST)": [
      ["PRODUCTO", "SUC", "STOCK"],
      ["BOLILLO", 1, 113],
    ],
  });
  const stockRows = parseStock(stockWorkbook);
  assert(stockRows.length === 1 && stockRows[0].stock === 113, "stock debe leer la hoja TOTAL A TENER y la columna STOCK");

  const existenciasWorkbook = workbookFromSheets({
    "EXISTENCIA EN SUCURSALES": [
      ["PRODUCTO", "TOTAL GRAL SUC", "C.F.", "SUMA SUC+CF"],
      ["BOLILLO", 10, 2, 12],
    ],
  });
  const existenciasRows = parseExistencias(existenciasWorkbook);
  assert(existenciasRows.length === 1 && existenciasRows[0].sumaSucCf === 12, "existencias no debe confundirse con stock");
  assert(parseStock(existenciasWorkbook)[0]?.stock !== 12, "parseStock no debe tomar existencias como stock objetivo");

  const bajasWorkbook = workbookFromSheets({
    Resumen: [
      ["PRODUCTO", "CANTIDAD"],
      ["BOLILLO", 500],
    ],
    Reporte: [
      ["PRODUCTO", "CANTIDAD", "FECHA", "SUCURSAL", "MOTIVO"],
      ["BOLILLO", 3, "2026-07-02", "CENTRO", "Merma"],
      ["ERICK", 80, "2026-07-02", "CENTRO", "Subtotal"],
    ],
  });
  const bajasRows = parseSalesOrReturns(bajasWorkbook, "bajas", "bajas julio.xlsx");
  assert(bajasRows.length === 1, "bajas diarias deben salir solo de la hoja Reporte");
  assert(bajasRows[0].cantidad === 3, "no debe colarse el total mensual de Resumen");
  assert(parseBajasReport(bajasWorkbook).every((row) => !/^ERICK/.test(row.producto)), "Erick no es un producto de bajas");

  const summaryWorkbook = workbookFromSheets({
    Totales: [
      ["PRODUCTO", "CANTIDAD"],
      ["BOLILLO", 42],
    ],
  });
  const summaryRows = parseMonthlySummaryWorkbook(summaryWorkbook, { year: 2026, monthIndex: 5 });
  assert(summaryRows.length === 1 && summaryRows[0].monthlyTotal === true, "un resumen sin fechas diarias es total mensual");
  const salesSummary = parseSalesOrReturns(summaryWorkbook, "ventas", "ventas junio.xlsx");
  assert(salesSummary.every((row) => row.monthlyTotal), "parseSalesOrReturns no debe inventar días a partir del total mensual");

  const summaryBajas = parseBajasSummaryWorkbook(workbookFromSheets({
    "BAJAS ERICK": [
      ["ETIQUETAS", "CANTIDAD"],
      ["BOLILLO", 11],
      ["ERICK", 80],
    ],
  }));
  assert(summaryBajas.length === 1 && summaryBajas[0].producto.includes("BOLILLO"), "el subtotal Erick no entra al resumen de bajas");

  const wideWorkbook = workbookFromSheets({
    "JULIO 2026": [
      ["", "LUNES", "MARTES"],
      ["", 1, 2],
      ["BOLILLO", 10, 20],
      ["CONCHA", 4, 5],
    ],
  });
  const remapped = parseSalesOrReturns(wideWorkbook, "ventas", "VENTAS DE MAYO Y JUNIO.xlsx");
  assert(remapped.length >= 2, "la hoja ancha debe producir ventas diarias");
  assert(remapped.every((row) => monthKey(row) === "2026-06"), "hoja JULIO en archivo mayo-junio se interpreta como junio");

  const growthData = new Map([
    ["2025-03", { total: 100 }],
    ["2025-04", { total: 80 }],
    ["2025-05", { total: 100 }],
    ["2026-03", { total: 110 }],
    ["2026-04", { total: 100 }],
    ["2026-05", { total: 90 }],
  ]);
  assert(Math.abs(computeAnnualGrowthFactor(growthData, "2026-06", 1) - 0.9) < 1e-9, "un mes de crecimiento usa solo el mes previo");
  assert(Math.abs(computeAnnualGrowthFactor(growthData, "2026-06", 3) - 1.1) < 1e-9, "tres meses usan la mediana anual");

  console.log("parser-test ok");
}

main().catch((error) => {
  console.error(error);
  process.exit(1);
});
