const fs = require("fs");
const path = require("path");
const Module = require("module");
const esbuild = require("esbuild");
const XLSX = require("xlsx");

const ROOT = __dirname;
const DOWNLOADS = path.resolve(ROOT, "..");
const OUTPUT = path.join(DOWNLOADS, "PLAN_PRODUCCION_AGOSTO_SEPTIEMBRE.xlsx");
const STATUS_FILE = path.join(DOWNLOADS, "ESTATUS_PRODUCTOS_PARA_LLENAR.xlsx");
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

const WEEKDAY_LABELS = ["Domingo", "Lunes", "Martes", "Miercoles", "Jueves", "Viernes", "Sabado"];
const MIN_BATCH = 8;
const WEEKLY_ON_DEMAND_THRESHOLD = 10;

const PERIODS = [
  { label: "Agosto restante", month: "2026-08", from: "2026-08-18", to: "2026-08-31" },
  { label: "Septiembre", month: "2026-09", from: "2026-09-01", to: "2026-09-30" },
];

function normalized(value) {
  return String(value || "")
    .toUpperCase().normalize("NFD").replace(/[\u0300-\u036f]/g, "")
    .replace(/[.,/\\_-]+/g, " ").replace(/\s+/g, " ").trim();
}

function findDownload(search) {
  const target = normalized(search);
  const match = fs.readdirSync(DOWNLOADS).find((name) => normalized(name) === target);
  if (!match) throw new Error(`No se encontro ${search}`);
  return path.join(DOWNLOADS, match);
}

function dateKey(date) {
  return `${date.getFullYear()}-${String(date.getMonth() + 1).padStart(2, "0")}-${String(date.getDate()).padStart(2, "0")}`;
}

function datesBetween(from, to) {
  const [fy, fm, fd] = from.split("-").map(Number);
  const [ty, tm, td] = to.split("-").map(Number);
  const cursor = new Date(fy, fm - 1, fd);
  const end = new Date(ty, tm - 1, td);
  const dates = [];
  while (cursor <= end) {
    dates.push(new Date(cursor));
    cursor.setDate(cursor.getDate() + 1);
  }
  return dates;
}

// Regla operativa de pasteles: minimo 10 piezas y multiplos de 5; menos de 8 no se produce.
function applyBatchRule(value) {
  const amount = Number(value) || 0;
  if (amount < MIN_BATCH) return 0;
  return 10 + Math.floor((amount - MIN_BATCH) / 5) * 5;
}

function loadStatus() {
  const sheet = XLSX.readFile(STATUS_FILE).Sheets["Estatus productos"];
  return new Map(
    XLSX.utils.sheet_to_json(sheet, { defval: "" })
      .map((row) => [normalized(row.Producto), String(row.Estatus || "").trim().toUpperCase()])
  );
}

async function loadAppFunctions() {
  const built = await esbuild.build({
    entryPoints: [path.join(ROOT, "App.jsx")],
    bundle: true, platform: "node", format: "cjs", write: false,
    loader: { ".css": "text" },
    define: { "import.meta.env.VITE_API_URL": JSON.stringify("") },
    logLevel: "silent",
  });
  const appModule = new Module("production-plan");
  appModule.filename = path.join(ROOT, "production-plan.bundle.cjs");
  appModule.paths = module.paths;
  appModule._compile(built.outputFiles[0].text, appModule.filename);
  return appModule.exports;
}

function weekdayAverage(row, weekday) {
  const keys = [
    "promedioDomingo", "promedioLunes", "promedioMartes", "promedioMiercoles",
    "promedioJueves", "promedioViernes", "promedioSabado",
  ];
  return Number(row[keys[weekday]] || 0);
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
  const status = loadStatus();
  const stockRows = app.parseStock(XLSX.readFile(findDownload("STOCK IDEAL SUCURSALES.xlsx"), { cellDates: true }));
  const read = (filePath) =>
    app.parseSalesOrReturns(XLSX.readFile(filePath, { cellDates: true }), "ventas", path.basename(filePath));
  const sales = [
    ...SOURCES.flatMap((name) => read(findDownload(name))),
    ...read(JULY_SALES),
    ...read(JUNE_CLOSE),
  ];

  const planRows = [];
  const productRows = [];
  const onDemandRows = [];
  const summaryRows = [];

  for (const period of PERIODS) {
    const forecastRows = app.calculateForecast({
      stockRows,
      historicalVentas: app.filterVentasBeforeMonth(sales, period.month),
      bajas: [], existencias: [], realProduction: [],
      selectedMonth: period.month,
      dailyBufferPct: 0,
    });
    const dates = datesBetween(period.from, period.to);
    const productionDates = dates.filter((date) => date.getDay() !== 0);

    let periodDemand = 0;
    let periodProduction = 0;
    let excludedProducts = 0;

    for (const row of forecastRows) {
      const product = normalized(row.producto);
      const estatus = status.get(product) || "ACTIVO";
      if (estatus === "BAJA" || estatus === "ESTACIONAL") {
        excludedProducts += 1;
        continue;
      }

      // La demanda del domingo se suma al sabado porque la planta no produce en domingo.
      const demandByProductionDate = new Map();
      let productDemand = 0;
      for (const date of dates) {
        const demand = weekdayAverage(row, date.getDay());
        productDemand += demand;
        if (date.getDay() === 0) continue;
        demandByProductionDate.set(dateKey(date), (demandByProductionDate.get(dateKey(date)) || 0) + demand);
      }
      for (const date of dates) {
        if (date.getDay() !== 0) continue;
        const saturday = new Date(date);
        saturday.setDate(saturday.getDate() - 1);
        const key = demandByProductionDate.has(dateKey(saturday))
          ? dateKey(saturday)
          : dateKey(productionDates.at(-1));
        demandByProductionDate.set(key, (demandByProductionDate.get(key) || 0) + weekdayAverage(row, 0));
      }

      const weeks = Math.max(1, dates.length / 7);
      const weeklyDemand = productDemand / weeks;
      if (weeklyDemand > 0 && weeklyDemand < WEEKLY_ON_DEMAND_THRESHOLD) {
        onDemandRows.push({
          Periodo: period.label,
          Producto: row.producto,
          Estatus: estatus,
          "Demanda del periodo": Number(productDemand.toFixed(1)),
          "Demanda semanal": Number(weeklyDemand.toFixed(1)),
          Motivo: `Menos de ${WEEKLY_ON_DEMAND_THRESHOLD} piezas por semana: producir bajo pedido`,
        });
        continue;
      }
      if (productDemand <= 0) continue;

      // Lo que no alcanza lote minimo se arrastra al siguiente dia de produccion.
      let carry = 0;
      let productProduction = 0;
      const productPlan = [];
      productionDates.forEach((date, index) => {
        const dayDemand = demandByProductionDate.get(dateKey(date)) || 0;
        const available = dayDemand + carry;
        const isLast = index === productionDates.length - 1;
        let production = applyBatchRule(available);
        if (isLast && production === 0 && available > 0) production = 0;
        carry = production === 0 ? available : Math.max(0, available - production);
        productProduction += production;
        productPlan.push({
          Periodo: period.label,
          Fecha: dateKey(date),
          Dia: WEEKDAY_LABELS[date.getDay()],
          Producto: row.producto,
          "Demanda del dia": Number(dayDemand.toFixed(2)),
          "Arrastre previo": Number((available - dayDemand).toFixed(2)),
          "Base de lote": Number(available.toFixed(2)),
          "Produccion sugerida": production,
          Regla: production === 0
            ? `Menor a ${MIN_BATCH}: se acumula al siguiente dia`
            : "Minimo 10 y multiplos de 5",
        });
      });

      planRows.push(...productPlan.filter((entry) => entry["Produccion sugerida"] > 0 || entry["Demanda del dia"] > 0));
      periodDemand += productDemand;
      periodProduction += productProduction;
      productRows.push({
        Periodo: period.label,
        Producto: row.producto,
        Estatus: estatus,
        "Demanda del periodo": Number(productDemand.toFixed(1)),
        "Produccion sugerida": productProduction,
        "Diferencia": Number((productProduction - productDemand).toFixed(1)),
        "Promedio por dia de produccion": Number((productProduction / productionDates.length).toFixed(1)),
        "Saldo sin producir": Number(carry.toFixed(1)),
      });
    }

    summaryRows.push(
      { Indicador: `--- ${period.label} ---`, Valor: `${period.from} al ${period.to}` },
      { Indicador: "Dias del periodo", Valor: dates.length },
      { Indicador: "Dias de produccion (sin domingo)", Valor: productionDates.length },
      { Indicador: "Productos en el plan", Valor: productRows.filter((r) => r.Periodo === period.label).length },
      { Indicador: "Productos excluidos por baja o temporada", Valor: excludedProducts },
      { Indicador: "Productos bajo pedido", Valor: onDemandRows.filter((r) => r.Periodo === period.label).length },
      { Indicador: "Demanda estimada (piezas)", Valor: Number(periodDemand.toFixed(0)) },
      { Indicador: "Produccion sugerida (piezas)", Valor: periodProduction },
      { Indicador: "Diferencia por regla de lotes", Valor: Number((periodProduction - periodDemand).toFixed(0)) }
    );
  }

  const methodology = [
    { Concepto: "Demanda", Detalle: "Promedio historico del mismo dia de semana, del modelo validado, sin usar ventas del mes objetivo." },
    { Concepto: "Domingo", Detalle: "La planta no produce en domingo. La demanda del domingo se suma al sabado anterior." },
    { Concepto: "Regla de lote", Detalle: `Minimo 10 piezas y multiplos de 5. Si el dia no alcanza ${MIN_BATCH} piezas, se acumula al siguiente dia de produccion en lugar de perderse.` },
    { Concepto: "Bajo pedido", Detalle: `Productos con menos de ${WEEKLY_ON_DEMAND_THRESHOLD} piezas por semana salen del calendario regular.` },
    { Concepto: "Estatus", Detalle: "Se excluyen los productos marcados BAJA y ESTACIONAL en ESTATUS_PRODUCTOS_PARA_LLENAR.xlsx." },
    { Concepto: "Existencias", Detalle: "Este plan NO descuenta inventario. Al recibir el archivo de existencias, la produccion recomendada baja." },
    { Concepto: "Septiembre", Detalle: "Se calcula con historia hasta julio porque agosto todavia no cierra. Debe recalcularse al cerrar agosto." },
    { Concepto: "Reproducible", Detalle: "node build-production-plan.cjs" },
  ];

  const workbook = XLSX.utils.book_new();
  const summarySheet = XLSX.utils.json_to_sheet(summaryRows);
  styleSheet(summarySheet, [42, 30]);
  XLSX.utils.book_append_sheet(workbook, summarySheet, "Resumen");

  const planSheet = XLSX.utils.json_to_sheet(planRows);
  styleSheet(planSheet, [17, 12, 12, 36, 16, 15, 14, 20, 34]);
  XLSX.utils.book_append_sheet(workbook, planSheet, "Plan diario");

  const productSheet = XLSX.utils.json_to_sheet(productRows);
  styleSheet(productSheet, [17, 36, 13, 20, 20, 12, 28, 18]);
  XLSX.utils.book_append_sheet(workbook, productSheet, "Por producto");

  const onDemandSheet = XLSX.utils.json_to_sheet(onDemandRows.length ? onDemandRows : [{ Periodo: "", Producto: "sin productos bajo pedido" }]);
  styleSheet(onDemandSheet, [17, 36, 13, 20, 18, 52]);
  XLSX.utils.book_append_sheet(workbook, onDemandSheet, "Bajo pedido");

  const methodologySheet = XLSX.utils.json_to_sheet(methodology);
  styleSheet(methodologySheet, [16, 120]);
  XLSX.utils.book_append_sheet(workbook, methodologySheet, "Metodologia");

  XLSX.writeFile(workbook, OUTPUT);

  console.log(JSON.stringify({
    archivo: OUTPUT,
    renglonesPlan: planRows.length,
    periodos: PERIODS.map((period) => {
      const rows = productRows.filter((row) => row.Periodo === period.label);
      return {
        periodo: period.label,
        productos: rows.length,
        demanda: Number(rows.reduce((sum, row) => sum + row["Demanda del periodo"], 0).toFixed(0)),
        produccion: rows.reduce((sum, row) => sum + row["Produccion sugerida"], 0),
        bajoPedido: onDemandRows.filter((row) => row.Periodo === period.label).length,
      };
    }),
  }, null, 2));
}

main().catch((error) => { console.error(error.message); process.exit(1); });
