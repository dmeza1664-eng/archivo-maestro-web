/**
 * Regresión con fixture SINTÉTICO (scripts/fixtures/sintetico-pronostico-2024-2025.json).
 * No son ventas reales: corre siempre en CI, sin archivos Pepes, y protege
 *  - arranque en frío: enero lee el año anterior y sin año anterior queda en 0;
 *  - año anterior oculto en feb–dic: con y sin 2024 el pronóstico es idéntico;
 *  - regla de mayo/diciembre para pasteles (impulso frío Madres / Navidad);
 *  - guard de producto intermitente o apagado al cierre del año anterior;
 *  - WAPE por mes del fixture (golden) para detectar cambios no intencionales.
 * Si un cambio del modelo es intencional: SYNTHETIC_UPDATE=1 node scripts/synthetic-cold-start-check.cjs
 * imprime el bloque "expected" nuevo para pegarlo en el fixture.
 */
const fs = require("fs");
const path = require("path");
const Module = require("module");
const esbuild = require("esbuild");

const ROOT = path.join(__dirname, "..");
const FIXTURE = path.join(__dirname, "fixtures", "sintetico-pronostico-2024-2025.json");
const MONTHS = Array.from({ length: 12 }, (_, i) => `2025-${String(i + 1).padStart(2, "0")}`);

async function loadAppFunctions() {
  const built = await esbuild.build({
    entryPoints: [path.join(ROOT, "App.jsx")],
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
  const appModule = new Module("synthetic-cold-start-check");
  appModule.filename = path.join(__dirname, "synthetic-cold-start-check.bundle.cjs");
  appModule.paths = module.paths;
  appModule._compile(built.outputFiles[0].text, appModule.filename);
  return appModule.exports;
}

function assert(condition, message) {
  if (!condition) throw new Error(`[sintético] ${message}`);
}

function monthClose(monthKey, product, quantity) {
  const [year, month] = monthKey.split("-").map(Number);
  return {
    fecha: new Date(year, month - 1, 15, 12),
    producto: product,
    cantidad: quantity,
    monthlyTotal: true,
    monthDays: new Date(year, month, 0).getDate(),
  };
}

function loadFixture() {
  const fixture = JSON.parse(fs.readFileSync(FIXTURE, "utf8"));
  assert(/SINT[ÉE]TICO/i.test(fixture._SINTETICO || ""), "el fixture debe estar marcado como sintético");
  const ventas = [];
  for (const [product, months] of Object.entries(fixture.cierresMensuales)) {
    for (const [monthKey, qty] of Object.entries(months)) ventas.push(monthClose(monthKey, product, qty));
  }
  return { fixture, ventas };
}

function runMonth(app, stockRows, ventas, month) {
  const rows = app.calculateForecast({
    stockRows,
    historicalVentas: app.filterVentasBeforeMonth(ventas, month),
    bajas: [],
    existencias: [],
    realProduction: [],
    selectedMonth: month,
    dailyBufferPct: 10,
  });
  const actual = new Map();
  for (const row of ventas) {
    const key = `${row.fecha.getFullYear()}-${String(row.fecha.getMonth() + 1).padStart(2, "0")}`;
    if (key === month) actual.set(row.producto, (actual.get(row.producto) || 0) + row.cantidad);
  }
  const analysis = app.analyzeForecastProductErrors(rows, actual);
  return { rows, analysis, byProduct: new Map(rows.map((row) => [row.producto, row])) };
}

async function runSyntheticChecks() {
  const app = await loadAppFunctions();
  const { fixture, ventas } = loadFixture();
  const stockRows = fixture.stock;
  const sin2024 = ventas.filter((row) => row.fecha.getFullYear() >= 2025);
  const con = {};
  const sin = {};
  for (const month of MONTHS) {
    con[month] = runMonth(app, stockRows, ventas, month);
    sin[month] = runMonth(app, stockRows, sin2024, month);
  }

  // 1) Arranque en frío: sin año anterior enero queda en 0; con 2024 lo lee.
  assert(sin["2025-01"].rows.every((row) => row.pronosticoVenta === 0), "sin 2024 enero debe quedar en 0");
  const mokaJan = con["2025-01"].byProduct.get("MOKA GDE").pronosticoVenta;
  const mokaJan2024 = fixture.cierresMensuales["MOKA GDE"]["2024-01"];
  assert(
    mokaJan > mokaJan2024 * 0.75 && mokaJan < mokaJan2024 * 1.35,
    `enero en frío debe partir de enero 2024 (${mokaJan2024}); salió ${mokaJan.toFixed(1)}`
  );

  // 2) Guard: intermitente sin venta en enero del año anterior → 0;
  //    apagado al cierre → no pasa del nivel nov–dic.
  assert(con["2025-01"].byProduct.get("BOLLOS C 6").pronosticoVenta === 0, "intermitente sin enero previo debe quedar en 0");
  const rosca = fixture.cierresMensuales["ROSCA C CAJETA"];
  const closeLevel = (rosca["2024-11"] + rosca["2024-12"]) / 2;
  assert(
    con["2025-01"].byProduct.get("ROSCA C CAJETA").pronosticoVenta <= closeLevel + 1e-6,
    `producto apagado al cierre no debe pasar de ${closeLevel}`
  );

  // 3) Feb–dic: el año anterior queda oculto → idéntico con y sin 2024
  //    (BOLLOS C 6 queda fuera: tiene huecos y la reactivación de hueco sí
  //    lee el mismo mes del año anterior a propósito).
  for (const month of MONTHS.slice(1)) {
    for (const row of con[month].rows) {
      if (row.producto === "BOLLOS C 6") continue;
      const other = sin[month].byProduct.get(row.producto);
      assert(
        Math.abs(row.pronosticoVenta - other.pronosticoVenta) < 1e-6,
        `${month} ${row.producto}: con 2024 (${row.pronosticoVenta.toFixed(2)}) debe ser igual que sin 2024 (${other.pronosticoVenta.toFixed(2)})`
      );
    }
  }

  // 4) Regla de mayo y diciembre para pasteles (×1.18 en arranque en frío).
  const mayMoka = con["2025-05"].byProduct.get("MOKA GDE");
  assert(/impulso frío pastel Madres/.test(mayMoka.metodoPronostico || ""), `mayo: MOKA GDE debe llevar impulso Madres (${mayMoka.metodoPronostico})`);
  const mayMed = con["2025-05"].byProduct.get("CHOCOLATE MED");
  assert(/impulso frío pastel Madres/.test(mayMed.metodoPronostico || ""), `mayo: CHOCOLATE MED debe llevar impulso Madres (${mayMed.metodoPronostico})`);
  const decMoka = con["2025-12"].byProduct.get("MOKA GDE");
  assert(/impulso frío pastel Navidad/.test(decMoka.metodoPronostico || ""), `diciembre: MOKA GDE debe llevar impulso Navidad (${decMoka.metodoPronostico})`);
  const decMed = con["2025-12"].byProduct.get("CHOCOLATE MED");
  assert(!/impulso frío pastel Navidad/.test(decMed.metodoPronostico || ""), "diciembre: pasteles MED no llevan impulso Navidad");
  const sepMoka = con["2025-09"].byProduct.get("MOKA GDE");
  assert(!/impulso frío/.test(sepMoka.metodoPronostico || ""), "septiembre no lleva impulso de calendario");

  // 5) Golden: WAPE por mes del fixture.
  const observed = {
    con2024: Object.fromEntries(MONTHS.map((m) => [m, Number(con[m].analysis.wape.toFixed(2))])),
    sin2024: Object.fromEntries(MONTHS.map((m) => [m, Number(sin[m].analysis.wape.toFixed(2))])),
  };
  if (process.env.SYNTHETIC_UPDATE === "1") {
    console.log(JSON.stringify({ expected: observed }, null, 1));
    return observed;
  }
  assert(fixture.expected, "falta el bloque expected en el fixture (SYNTHETIC_UPDATE=1 para generarlo)");
  for (const kind of ["con2024", "sin2024"]) {
    for (const month of MONTHS) {
      const want = fixture.expected[kind][month];
      const got = observed[kind][month];
      assert(Math.abs(want - got) < 0.005 + 1e-9, `${kind} ${month}: WAPE sintético esperado ${want}, salió ${got}`);
    }
  }
  return observed;
}

module.exports = { runSyntheticChecks };

if (require.main === module) {
  runSyntheticChecks()
    .then(() => console.log("synthetic-cold-start-check ok"))
    .catch((error) => {
      console.error(error);
      process.exit(1);
    });
}
