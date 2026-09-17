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

  // Magnitudes reales de julio 2026: mayo pico (Día de las Madres), junio YoY
  // plano, julio 2025 con caída suave (~8.5%), junio diario acelerando, y
  // julio 2026 saltó (FRUTAS 519→580, MOKA 617→697). PAY baja en julio.
  function juneDaily(product, firstHalf, secondHalf) {
    const rows = [];
    for (let day = 1; day <= 30; day += 1) {
      const half = day <= 15 ? firstHalf : secondHalf;
      rows.push({
        fecha: new Date(2026, 5, day),
        producto: product,
        cantidad: half / 15,
      });
    }
    return rows;
  }

  const realishStock = [
    { producto: "FRUTAS GDE", stock: 40, orden: 1 },
    { producto: "MOKA GDE", stock: 40, orden: 2 },
    { producto: "M & M GDE", stock: 20, orden: 3 },
    { producto: "PAY DE FRESA GDE", stock: 20, orden: 4 },
    { producto: "FRUTAS MED", stock: 30, orden: 5 },
  ];
  const realishSeries = {
    "FRUTAS GDE": {
      "2025-03": 524, "2025-04": 500, "2025-05": 658, "2025-06": 519, "2025-07": 475, "2025-08": 524, "2025-09": 492, "2025-10": 512,
      "2026-05": 660, "2026-07": 580, "2026-08": 541,
    },
    "MOKA GDE": {
      "2025-03": 605, "2025-04": 605, "2025-05": 755, "2025-06": 624, "2025-07": 612, "2025-08": 639, "2025-09": 604, "2025-10": 652,
      "2026-05": 720, "2026-07": 697, "2026-08": 619,
    },
    "M & M GDE": {
      "2025-03": 302, "2025-04": 302, "2025-05": 320, "2025-06": 284, "2025-07": 308, "2025-08": 324, "2025-09": 298, "2025-10": 317,
      "2026-05": 263, "2026-07": 324, "2026-08": 304,
    },
    "PAY DE FRESA GDE": {
      "2025-03": 200, "2025-04": 175, "2025-05": 308, "2025-06": 264, "2025-07": 199, "2025-08": 222, "2025-09": 168, "2025-10": 194,
      "2026-05": 285, "2026-07": 212, "2026-08": 181,
    },
    "FRUTAS MED": {
      "2025-03": 530, "2025-04": 510, "2025-05": 640, "2025-06": 520, "2025-07": 500, "2025-08": 530, "2025-09": 510, "2025-10": 520,
      "2026-05": 642, "2026-06": 522, "2026-07": 530, "2026-08": 528,
    },
  };
  const realishVentas = [];
  for (const [product, months] of Object.entries(realishSeries)) {
    for (const [month, qty] of Object.entries(months)) {
      realishVentas.push(monthClose(month, product, qty));
    }
  }
  realishVentas.push(...juneDaily("FRUTAS GDE", 223, 296));
  realishVentas.push(...juneDaily("MOKA GDE", 276, 341));
  realishVentas.push(...juneDaily("M & M GDE", 115, 141));
  realishVentas.push(...juneDaily("PAY DE FRESA GDE", 89, 165));

  const realJune = evaluateMonth(app, realishStock, realishVentas, "2026-06");
  const realJuly = evaluateMonth(app, realishStock, realishVentas, "2026-07");
  const realAugust = evaluateMonth(app, realishStock, realishVentas, "2026-08");
  const pick = (analysis, name) => analysis.rows.find((row) => row.producto === name);

  assert(realJune.wape < 8, `junio real-like no debe romperse (WAPE ${realJune.wape.toFixed(2)}%)`);
  assert(realAugust.wape < 12, `agosto real-like no debe dispararse (WAPE ${realAugust.wape.toFixed(2)}%)`);

  const realFrutas = pick(realJuly, "FRUTAS GDE");
  const realMoka = pick(realJuly, "MOKA GDE");
  const realMm = pick(realJuly, "M & M GDE");
  const realPay = pick(realJuly, "PAY DE FRESA GDE");
  const realMed = pick(realJuly, "FRUTAS MED");
  assert(realFrutas.absoluteError < 40, `FRUTAS GDE julio debe bajar del faltante ~70 (abs ${realFrutas.absoluteError.toFixed(1)})`);
  assert(realFrutas.forecast > 540, `FRUTAS GDE debe subir de ~510 por el impulso de junio (fc ${realFrutas.forecast.toFixed(1)})`);
  assert(realMoka.absoluteError < 40, `MOKA GDE julio debe bajar del faltante ~66 (abs ${realMoka.absoluteError.toFixed(1)})`);
  assert(realMm.absoluteError < 40, `M & M GDE no debe recortar el julio 2025 de 308 (abs ${realMm.absoluteError.toFixed(1)})`);
  assert(realPay.forecast < 270, `PAY DE FRESA no debe recibir el impulso (fc ${realPay.forecast.toFixed(1)})`);
  assert(realMed.absoluteError < 45, `FRUTAS MED no es GDE: el cambio no debe dispararlo (abs ${realMed.absoluteError.toFixed(1)})`);

  const augustFrutas = pick(realAugust, "FRUTAS GDE");
  assert(augustFrutas.forecast < 600, `agosto FRUTAS no debe heredar el impulso de junio (fc ${augustFrutas.forecast.toFixed(1)})`);

  console.log("forecast-accuracy-test ok");
  console.log(
    JSON.stringify(
      {
        june: { wape: Number(june.wape.toFixed(2)), inside15: june.inside15, top: june.topErrors[0].producto },
        july: { wape: Number(july.wape.toFixed(2)), inside15: july.inside15, top: july.topErrors[0].producto },
        august: { wape: Number(august.wape.toFixed(2)), inside15: august.inside15, top: august.topErrors[0].producto },
        realish: {
          june: Number(realJune.wape.toFixed(2)),
          july: Number(realJuly.wape.toFixed(2)),
          august: Number(realAugust.wape.toFixed(2)),
          frutas: Number(realFrutas.forecast.toFixed(1)),
          moka: Number(realMoka.forecast.toFixed(1)),
          mm: Number(realMm.forecast.toFixed(1)),
        },
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
