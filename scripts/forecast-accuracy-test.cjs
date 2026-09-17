const path = require("path");
const Module = require("module");
const esbuild = require("esbuild");

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
  const appModule = new Module("forecast-accuracy-test");
  appModule.filename = path.join(__dirname, "forecast-accuracy-test.bundle.cjs");
  appModule.paths = module.paths;
  appModule._compile(built.outputFiles[0].text, appModule.filename);
  return appModule.exports;
}

function assert(condition, message) {
  if (!condition) throw new Error(message);
}

function monthClose(monthKey, product, quantity) {
  const [year, month] = monthKey.split("-").map(Number);
  return {
    fecha: new Date(year, month - 1, 1),
    producto: product,
    cantidad: quantity,
    monthlyTotal: true,
    monthDays: new Date(year, month, 0).getDate(),
  };
}

function evaluateMonth(app, stockRows, ventas, hideMonth) {
  const historical = app.filterVentasBeforeMonth(ventas, hideMonth);
  const forecastRows = app.calculateForecast({
    stockRows,
    historicalVentas: historical,
    bajas: [],
    existencias: [],
    realProduction: [],
    selectedMonth: hideMonth,
    dailyBufferPct: 10,
  });
  const actualMap = new Map();
  for (const row of ventas) {
    const key = `${row.fecha.getFullYear()}-${String(row.fecha.getMonth() + 1).padStart(2, "0")}`;
    if (key !== hideMonth) continue;
    actualMap.set(row.producto, (actualMap.get(row.producto) || 0) + row.cantidad);
  }
  return app.analyzeForecastProductErrors(forecastRows, actualMap);
}

async function main() {
  const app = await loadAppFunctions();

  const stockRows = [
    { producto: "FRUTAS GDE", stock: 40, orden: 1 },
    { producto: "MOKA GDE", stock: 40, orden: 2 },
    { producto: "CHESSECAKE GDE", stock: 20, orden: 3 },
    { producto: "GELATINA IND FRESA", stock: 30, orden: 4 },
    { producto: "GALLETA NUEZ", stock: 15, orden: 5 },
  ];

  // Reproduce el patrón documentado de julio 2026: el mismo mes de 2025 cayó
  // ~12% vs junio, junio YoY quedó casi plano y julio 2026 subió. El modelo
  // anterior copiaba esa baja (WAPE sintético de julio ~8% en GDE; 18% si
  // CHEESECAKE no empataba con el catálogo).
  const series = {
    "FRUTAS GDE": {
      "2025-04": 480, "2025-05": 510, "2025-06": 500, "2025-07": 440, "2025-08": 495,
      "2026-04": 495, "2026-05": 525, "2026-06": 503, "2026-07": 525, "2026-08": 518,
    },
    "MOKA GDE": {
      "2025-04": 600, "2025-05": 640, "2025-06": 620, "2025-07": 545, "2025-08": 615,
      "2026-04": 618, "2026-05": 655, "2026-06": 624, "2026-07": 650, "2026-08": 640,
    },
    "CHESSECAKE GDE": {
      "2025-04": 180, "2025-05": 190, "2025-06": 188, "2025-07": 165, "2025-08": 186,
      "2026-04": 186, "2026-05": 196, "2026-06": 189, "2026-07": 197, "2026-08": 192,
    },
    "GELATINA IND FRESA": {
      "2025-04": 220, "2025-05": 230, "2025-06": 225, "2025-07": 222, "2025-08": 228,
      "2026-04": 226, "2026-05": 234, "2026-06": 228, "2026-07": 232, "2026-08": 230,
    },
    "GALLETA NUEZ": {
      "2025-04": 90, "2025-05": 95, "2025-06": 92, "2025-07": 91, "2025-08": 93,
      "2026-04": 93, "2026-05": 96, "2026-06": 94, "2026-07": 95, "2026-08": 94,
    },
  };

  const ventas = [];
  for (const [product, months] of Object.entries(series)) {
    for (const [month, qty] of Object.entries(months)) {
      ventas.push(monthClose(month, product, qty));
    }
  }

  const june = evaluateMonth(app, stockRows, ventas, "2026-06");
  const july = evaluateMonth(app, stockRows, ventas, "2026-07");
  const august = evaluateMonth(app, stockRows, ventas, "2026-08");

  assert(june.wape < 5, `junio sintético no debe degradarse (WAPE ${june.wape.toFixed(2)}%)`);
  assert(july.wape < 4, `julio tipo crecimiento debe bajar de ~8% a menos de 4% (WAPE ${july.wape.toFixed(2)}%)`);
  assert(august.wape < 4, `agosto estable no debe empeorar (WAPE ${august.wape.toFixed(2)}%)`);

  const julyCakes = july.rows.filter((row) => /GDE$/.test(row.producto));
  assert(
    julyCakes.every((row) => row.absoluteError < 15),
    `los pasteles GDE de julio deben quedar dentro de ±15 (peor ${Math.max(...julyCakes.map((row) => row.absoluteError)).toFixed(1)})`
  );

  const health = app.buildForecastHealth({
    stockRows,
    ventas,
    selectedMonth: "2026-09",
    dailyBufferPct: 10,
  });
  assert(health.backtests.length === 3, "el control debe ocultar los últimos 3 meses cerrados");
  assert(health.latestBacktest.topErrors[0].producto, "el control debe nombrar el producto con más error");
  assert(health.backtests.every((row) => row.wape < 5), "los tres meses sintéticos deben quedar por debajo de 5% WAPE");

  const ventasSinJulio2025 = ventas.filter((row) => {
    const key = `${row.fecha.getFullYear()}-${String(row.fecha.getMonth() + 1).padStart(2, "0")}`;
    return !(key === "2025-07" && row.producto === "FRUTAS GDE");
  });
  const julyWithoutPrior = evaluateMonth(app, stockRows, ventasSinJulio2025, "2026-07");
  const frutas = julyWithoutPrior.rows.find((row) => row.producto === "FRUTAS GDE");
  assert(frutas.forecast > 480, "sin julio 2025 debe usar meses vecinos, no dejar FRUTAS en la caída de 440");
  assert(frutas.absoluteError < 40, "el proxy estacional no debe disparar el error de FRUTAS");

  console.log("forecast-accuracy-test ok");
  console.log(
    JSON.stringify(
      {
        june: { wape: Number(june.wape.toFixed(2)), inside15: june.inside15, top: june.topErrors[0].producto },
        july: { wape: Number(july.wape.toFixed(2)), inside15: july.inside15, top: july.topErrors[0].producto },
        august: { wape: Number(august.wape.toFixed(2)), inside15: august.inside15, top: august.topErrors[0].producto },
      },
      null,
      2
    )
  );
}

main().catch((error) => {
  console.error(error);
  process.exit(1);
});
