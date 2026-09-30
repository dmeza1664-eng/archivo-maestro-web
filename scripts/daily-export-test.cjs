// «Exportar diario» (oct 2026): una sola hoja con exactamente Día, Producto y
// «Mandar a producir», con el mismo número que la tabla diaria (inventario restado
// y lote de pastel incluidos) y solo las filas filtradas de la tabla.
const assert = require("assert");
const fs = require("fs");
const path = require("path");
const Module = require("module");
const esbuild = require("esbuild");

const ROOT = path.join(__dirname, "..");
const XLSX = require(path.join(ROOT, "node_modules", "xlsx"));
const app = fs.readFileSync(path.join(ROOT, "App.jsx"), "utf8");

// El botón exporta las filas filtradas de la tabla y la celda de la tabla usa la misma expresión.
assert.match(app, /onClick=\{\(\) => exportDailyToExcel\(filteredDailyRows\)\}[^>]*>\s*<Download size=\{18\} \/> Exportar diario/, "El botón exporta filteredDailyRows.");
assert.match(app, /<span className="produce-badge">\{row\.aProducirDia \?\? row\.produccionSugeridaDia\}<\/span>/, "La tabla muestra aProducirDia ?? produccionSugeridaDia.");
assert.match(app, /"Mandar a producir": row\.aProducirDia \?\? row\.produccionSugeridaDia,/, "El export usa la misma expresión que la tabla.");
assert.match(app, /onClick=\{\(\) => exportToExcel\(/, "El «Exportar» mensual sigue igual.");

async function loadApp() {
  const built = await esbuild.build({
    entryPoints: [path.join(ROOT, "App.jsx")], bundle: true, platform: "node", format: "cjs", write: false,
    loader: { ".css": "text" },
    define: { "import.meta.env.VITE_API_URL": '""', "import.meta.env.DEV": "false", "import.meta.env.PROD": "true" },
    logLevel: "silent",
  });
  const m = new Module("daily-export-test");
  m.filename = path.join(__dirname, "daily-export-test.bundle.cjs");
  m.paths = module.paths;
  m._compile(built.outputFiles[0].text, m.filename);
  return m.exports;
}

function monthClose(month, producto, cantidad) {
  const [year, monthNumber] = month.split("-").map(Number);
  return { fecha: new Date(year, monthNumber - 1, 1), producto, cantidad, monthlyTotal: true, monthDays: new Date(year, monthNumber, 0).getDate() };
}

(async () => {
  const { calculateForecast, calculateDailyForecast, buildBranchStockTargetMap, branchStockKey, buildDailyExportRows, buildDailyExportWorkbook, DAILY_EXPORT_COLUMNS } = await loadApp();
  assert.deepStrictEqual(DAILY_EXPORT_COLUMNS, ["Día", "Producto", "Mandar a producir"]);

  const products = ["GELATINA IND FRESA", "FRUTAS GDE", "CHOCOLATE GDE"];
  const monthlyRows = calculateForecast({
    stockRows: products.map((producto, i) => ({ producto, stock: 30, orden: i + 1 })),
    historicalVentas: products.flatMap((p, i) => ["2026-07", "2026-08", "2026-09"].map((m, j) => monthClose(m, p, 220 + 40 * i + 10 * j))),
    bajas: [], existencias: [], realProduction: [], selectedMonth: "2026-10", dailyBufferPct: 10,
  });
  const fecha = "2026-10-01";
  const targets = buildBranchStockTargetMap([
    { producto: "FRUTAS GDE", stock: 30, stockSucursales: { [branchStockKey("Suc. Plaza")]: 1 } },
  ]);
  const daily = calculateDailyForecast({
    monthlyRows, ventasReales: [], realProduction: [], selectedMonth: "2026-10", dailyBufferPct: 10, activePromos: [],
    dailyBranchStock: [{ fecha, sucursal: "Suc. Plaza", producto: "FRUTAS GDE", cantidad: 4 }],
    dailyColdRoom: [{ fecha, producto: "FRUTAS GDE", cantidad: 2 }, { fecha, producto: "GELATINA IND FRESA", cantidad: 3 }],
    branchStockTargets: targets,
  });
  const mandar = (row) => row.aProducirDia ?? row.produccionSugeridaDia;
  // Mismo filtro que la tabla (fecha): el export recibe solo esas filas.
  const filtered = daily.filter((row) => row.fecha === fecha);
  assert.strictEqual(filtered.length, products.length, "Una fila por producto el 1/10.");
  const conInventario = filtered.filter((row) => row.hasDailyBranchStock || row.hasDailyColdRoom);
  assert(conInventario.length >= 2, "Hay filas con inventario restado.");
  assert(conInventario.some((row) => mandar(row) !== row.produccionBrutaDia), "El inventario cambia «Mandar a producir» respecto al bruto.");

  const readBack = (rows) => {
    const buf = XLSX.write(buildDailyExportWorkbook(rows), { type: "buffer", bookType: "xlsx" });
    return XLSX.read(buf, { type: "buffer" });
  };
  const wb = readBack(filtered);
  assert.deepStrictEqual(wb.SheetNames, ["Mandar a producir"], "Una sola hoja.");
  const sheet = wb.Sheets["Mandar a producir"];
  const aoa = XLSX.utils.sheet_to_json(sheet, { header: 1, defval: "" });
  assert.deepStrictEqual(aoa[0], ["Día", "Producto", "Mandar a producir"], "Encabezado: exactamente 3 columnas.");
  assert.strictEqual(aoa.length, filtered.length + 1, "Una fila por fila filtrada de la tabla.");
  for (const line of aoa) assert.strictEqual(line.length, 3, "Ninguna fila tiene columnas extra.");
  aoa.slice(1).forEach(([dia, producto, valor], i) => {
    const row = filtered[i];
    assert.strictEqual(dia, fecha, "Día en AAAA-MM-DD con la fecha de la tabla.");
    assert.strictEqual(producto, row.producto);
    assert.strictEqual(valor, mandar(row), `${row.producto}: «Mandar a producir» igual que la tabla.`);
  });
  assert(sheet["!ref"] && XLSX.utils.decode_range(sheet["!ref"]).e.c === 2, "Rango de la hoja: columnas A–C.");

  // Todo el mes (sin filtro de fecha) y filtro por producto/día de semana: mismas filas, mismo orden.
  const mes = buildDailyExportRows(daily);
  assert.strictEqual(mes.length, daily.length);
  mes.forEach((row, i) => {
    assert.deepStrictEqual(Object.keys(row), ["Día", "Producto", "Mandar a producir"]);
    assert.strictEqual(row["Mandar a producir"], mandar(daily[i]));
  });
  const jueves = daily.filter((row) => row.weekday === 4 && row.producto.includes("FRUTAS"));
  assert.deepStrictEqual(buildDailyExportRows(jueves).map((r) => r.Día), jueves.map((r) => r.fecha), "Respeta filtros de producto y día de semana.");
  const domingo = daily.find((row) => row.weekday === 0);
  assert.strictEqual(buildDailyExportRows([domingo])[0]["Mandar a producir"], 0, "Domingo sale en 0 como en la tabla.");

  console.log(`daily-export-test OK (${filtered.length} filas el ${fecha}, ${mes.length} en el mes)`);
})().catch((error) => { console.error(error); process.exit(1); });
