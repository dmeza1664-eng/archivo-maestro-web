const fs = require("fs");
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

  // Misma fixture por la vía de la UI: cierre dedicado + diario combinado
  // pasan por resolveCanonicalMonthSources. Antes el cierre tiraba el diario
  // y el impulso de julio casi no corría.
  const juneClosesForResolve = [
    monthClose("2026-06", "FRUTAS GDE", 519),
    monthClose("2026-06", "MOKA GDE", 617),
    monthClose("2026-06", "M & M GDE", 256),
    monthClose("2026-06", "PAY DE FRESA GDE", 254),
  ];
  const ventasForResolve = [...realishVentas, ...juneClosesForResolve];
  const dailySource = {
    name: "VENTAS DE MAYO Y JUNIO 2026.xlsx",
    rows: ventasForResolve.filter((row) => !row.monthlyTotal),
  };
  const closeSources = [...new Set(
    ventasForResolve.filter((row) => row.monthlyTotal).map((row) => {
      const key = `${row.fecha.getFullYear()}-${String(row.fecha.getMonth() + 1).padStart(2, "0")}`;
      return key;
    })
  )].map((key) => {
    const [year, month] = key.split("-").map(Number);
    const labels = ["enero", "febrero", "marzo", "abril", "mayo", "junio", "julio", "agosto", "septiembre", "octubre", "noviembre", "diciembre"];
    return {
      name: `ventas ${labels[month - 1]} ${year}.xlsx`,
      rows: ventasForResolve.filter((row) => {
        if (!row.monthlyTotal) return false;
        const rowKey = `${row.fecha.getFullYear()}-${String(row.fecha.getMonth() + 1).padStart(2, "0")}`;
        return rowKey === key;
      }),
    };
  });
  const resolvedAppDefault = app.resolveCanonicalMonthSources([dailySource, ...closeSources]);
  const juneDecision = resolvedAppDefault.decisions.find((decision) => decision.month === "2026-06");
  assert(juneDecision?.strategy === "keep-daily-and-close", "resolveCanonical debe conservar diario+cierre de junio");
  assert(juneDecision.kept.includes(dailySource.name), "el diario combinado de junio no se omite");
  const resolvedVentas = resolvedAppDefault.entries.flatMap((entry) => entry.rows);
  const resolvedJuly = evaluateMonth(app, realishStock, resolvedVentas, "2026-07");
  const resolvedAugust = evaluateMonth(app, realishStock, resolvedVentas, "2026-08");
  const resolvedFrutas = pick(resolvedJuly, "FRUTAS GDE");
  const exclusiveJune = app.resolveCanonicalMonthSources([dailySource, ...closeSources], { "2026-06": "ventas junio 2026.xlsx" });
  const exclusiveVentas = exclusiveJune.entries.flatMap((entry) => entry.rows);
  const exclusiveJuly = evaluateMonth(app, realishStock, exclusiveVentas, "2026-07");
  const exclusiveFrutas = pick(exclusiveJuly, "FRUTAS GDE");
  assert(resolvedFrutas.forecast > exclusiveFrutas.forecast + 20, `con resolveCanonical el impulso debe subir FRUTAS vs cierre-solo (${resolvedFrutas.forecast.toFixed(1)} vs ${exclusiveFrutas.forecast.toFixed(1)})`);
  assert(resolvedFrutas.absoluteError < exclusiveFrutas.absoluteError, "el error de FRUTAS GDE julio debe bajar al conservar el diario");
  assert(Math.abs(resolvedJuly.wape - realJuly.wape) < 0.15, "el WAPE de julio por resolveCanonical debe empatar con diario+cierre directo");
  assert(Math.abs(resolvedAugust.wape - realAugust.wape) < 0.15, "agosto no debe empeorar por conservar el diario de junio");

  // Peores SKU de error absoluto tras PR #7: pasteles regulares con huecos de
  // carga (ceros inferidos) y alias (CHEESECAKE / NUTELLA / M&M MEDIANO).
  // Antes la limpieza los marcaba intermitentes y dominaban el WAPE (junio
  // ~10.5%, julio ~12.4% en esta fixture).
  const heavyStock = [
    { producto: "FRUTAS GDE", stock: 40, orden: 1 },
    { producto: "MOKA GDE", stock: 40, orden: 2 },
    { producto: "M & M GDE", stock: 20, orden: 3 },
    { producto: "PAY DE FRESA GDE", stock: 20, orden: 4 },
    { producto: "FRUTAS MED", stock: 30, orden: 5 },
    { producto: "GELATINA IND FRESA", stock: 40, orden: 6 },
    { producto: "GELATINA IND MOSAICO", stock: 40, orden: 7 },
    { producto: "CHEESECAKE GDE", stock: 20, orden: 8 },
    { producto: "NUTELA GDE", stock: 20, orden: 9 },
    { producto: "M & M MED", stock: 20, orden: 10 },
  ];
  const heavySeries = {
    "FRUTAS GDE": {
      "2025-03": 524, "2025-04": 500, "2025-05": 658, "2025-06": 519, "2025-07": 475, "2025-08": 524, "2025-09": 492, "2025-10": 512,
      "2026-03": 510, "2026-04": 500, "2026-05": 660, "2026-07": 580, "2026-08": 541,
    },
    "MOKA GDE": {
      "2025-03": 605, "2025-04": 605, "2025-05": 755, "2025-06": 624, "2025-07": 612, "2025-08": 639, "2025-09": 604, "2025-10": 652,
      "2026-03": 610, "2026-04": 600, "2026-05": 720, "2026-07": 697, "2026-08": 619,
    },
    "M & M GDE": {
      "2025-03": 302, "2025-04": 302, "2025-05": 320, "2025-06": 284, "2025-07": 308, "2025-08": 324, "2025-09": 298, "2025-10": 317,
      "2026-03": 270, "2026-04": 268, "2026-05": 263, "2026-07": 324, "2026-08": 304,
    },
    "PAY DE FRESA GDE": {
      "2025-03": 200, "2025-04": 175, "2025-05": 308, "2025-06": 264, "2025-07": 199, "2025-08": 222, "2025-09": 168, "2025-10": 194,
      "2026-03": 190, "2026-04": 180, "2026-05": 285, "2026-07": 212, "2026-08": 181,
    },
    "FRUTAS MED": {
      "2025-03": 530, "2025-04": 510, "2025-05": 640, "2025-06": 520, "2025-07": 500, "2025-08": 530, "2025-09": 510, "2025-10": 520,
      "2026-03": 525, "2026-04": 512, "2026-05": 642, "2026-06": 522, "2026-07": 530, "2026-08": 528,
    },
    "GELATINA IND FRESA": {
      "2025-03": 780, "2025-04": 790, "2025-05": 980, "2025-06": 800, "2025-07": 810, "2025-08": 805,
      "2026-03": 785, "2026-04": 795, "2026-05": 990, "2026-06": 816, "2026-07": 820, "2026-08": 812,
    },
    "GELATINA IND MOSAICO": {
      "2025-03": 680, "2025-04": 690, "2025-05": 860, "2025-06": 700, "2025-07": 705, "2025-08": 698,
      "2026-03": 685, "2026-04": 692, "2026-05": 870, "2026-06": 710, "2026-07": 708, "2026-08": 702,
    },
    "CHEESECAKE GDE": {
      "2025-04": 180, "2025-05": 230, "2025-06": 188, "2025-07": 165, "2025-08": 186,
      "2026-04": 182, "2026-05": 235, "2026-06": 189, "2026-07": 197, "2026-08": 192,
    },
    "NUTELLA GDE": {
      "2025-04": 270, "2025-05": 340, "2025-06": 275, "2025-07": 250, "2025-08": 280,
      "2026-04": 268, "2026-05": 338, "2026-06": 272, "2026-07": 307, "2026-08": 278,
    },
    "M&M MEDIANO": {
      "2025-04": 270, "2025-05": 330, "2025-06": 280, "2025-07": 275, "2025-08": 285,
      "2026-04": 272, "2026-05": 328, "2026-06": 285, "2026-07": 288, "2026-08": 282,
    },
  };
  const heavyVentas = [];
  for (const [product, months] of Object.entries(heavySeries)) {
    for (const [month, qty] of Object.entries(months)) {
      heavyVentas.push(monthClose(month, product, qty));
    }
  }
  heavyVentas.push(...juneDaily("FRUTAS GDE", 223, 296));
  heavyVentas.push(...juneDaily("MOKA GDE", 276, 341));
  heavyVentas.push(...juneDaily("M & M GDE", 115, 141));
  heavyVentas.push(...juneDaily("PAY DE FRESA GDE", 89, 165));

  const heavyJune = evaluateMonth(app, heavyStock, heavyVentas, "2026-06");
  const heavyJuly = evaluateMonth(app, heavyStock, heavyVentas, "2026-07");
  const heavyAugust = evaluateMonth(app, heavyStock, heavyVentas, "2026-08");
  const heavyPick = (analysis, name) => analysis.rows.find((row) => app.normalizeProduct(row.producto) === app.normalizeProduct(name));
  const heavyWeighted = (() => {
    const months = [heavyJune, heavyJuly, heavyAugust];
    const actual = months.reduce((sum, row) => sum + row.actual, 0);
    const abs = months.reduce((sum, row) => sum + row.rows.reduce((inner, item) => inner + item.absoluteError, 0), 0);
    return actual > 0 ? (abs / actual) * 100 : null;
  })();

  assert(heavyJune.wape < 5.5, `junio de SKU pesados debe bajar del ~10.5% (WAPE ${heavyJune.wape.toFixed(2)}%)`);
  assert(heavyJuly.wape < 6.5, `julio de SKU pesados debe bajar del ~12.4% (WAPE ${heavyJuly.wape.toFixed(2)}%)`);
  assert(heavyAugust.wape < 6, `agosto de SKU pesados no debe dispararse (WAPE ${heavyAugust.wape.toFixed(2)}%)`);
  assert(heavyWeighted < 5.5, `WAPE ponderado jun-ago debe bajar del ~8.9% (${heavyWeighted.toFixed(2)}%)`);
  const sharedWeighted = app.weightedWapeFromBacktests([heavyJune, heavyJuly, heavyAugust]);
  assert(
    sharedWeighted != null && Math.abs(sharedWeighted - heavyWeighted) < 1e-9,
    `Salud debe usar la misma suma de errores que las pruebas (${sharedWeighted} vs ${heavyWeighted})`
  );
  assert(app.describeAccuracyWindow([
    { month: "2026-06" },
    { month: "2026-07" },
    { month: "2026-08" },
  ]).includes("–"), "la ventana jun–ago debe etiquetarse con el primer y último mes");

  const heavyCheeseJune = heavyPick(heavyJune, "CHEESECAKE GDE");
  const heavyNutelaJune = heavyPick(heavyJune, "NUTELA GDE");
  const heavyMmMedJune = heavyPick(heavyJune, "M & M MED");
  const heavyCheeseJuly = heavyPick(heavyJuly, "CHEESECAKE GDE");
  const heavyNutelaJuly = heavyPick(heavyJuly, "NUTELA GDE");
  const heavyFrutasJuly = heavyPick(heavyJuly, "FRUTAS GDE");
  const heavyMedJuly = heavyPick(heavyJuly, "FRUTAS MED");
  assert(heavyCheeseJune.absoluteError < 40, `junio CHEESECAKE no debe aplastarse por huecos (abs ${heavyCheeseJune.absoluteError.toFixed(1)})`);
  assert(heavyNutelaJune.absoluteError < 40, `junio NUTELA/NUTELLA debe empatar y no ir a 0 (abs ${heavyNutelaJune.absoluteError.toFixed(1)})`);
  assert(heavyMmMedJune.absoluteError < 40, `junio M&M MEDIANO debe empatar con M & M MED (abs ${heavyMmMedJune.absoluteError.toFixed(1)})`);
  assert(heavyCheeseJuly.forecast > 150, `julio CHEESECAKE no es intermitente (fc ${heavyCheeseJuly.forecast.toFixed(1)})`);
  assert(heavyNutelaJuly.forecast > 220, `julio NUTELA no es intermitente (fc ${heavyNutelaJuly.forecast.toFixed(1)})`);
  assert(heavyFrutasJuly.forecast > 540, `el impulso GDE de julio debe seguir (fc ${heavyFrutasJuly.forecast.toFixed(1)})`);
  assert(heavyMedJuly.absoluteError < 45, `FRUTAS MED no hereda el impulso GDE (abs ${heavyMedJuly.absoluteError.toFixed(1)})`);
  assert(!/venta intermitente/i.test(heavyCheeseJuly.metodo || ""), "CHEESECAKE no debe marcarse intermitente");

  const uploadDir = process.env.WAPE_UPLOAD_DIR
    || "/home/ubuntu/.cursor/projects/workspace/uploads";
  const excelCandidates = {
    stock: ["stock_ideal_849e.xlsx", "stock_ideal.xlsx"],
    daily: ["ventas_mayo_junio_angel_8999.xlsx", "ventas_mayo_junio_angel.xlsx"],
    juneClose: ["ventas_junio_4303.xlsx", "ventas_junio.xlsx"],
    julyClose: ["ventas_julio_a94f.xlsx", "ventas_julio.xlsx"],
  };
  function findExcel(names) {
    const dirs = [uploadDir, path.join(__dirname, "..", "wape-excel"), process.env.WAPE_EXCEL_DIR].filter(Boolean);
    for (const dir of dirs) {
      for (const name of names) {
        const full = path.join(dir, name);
        if (fs.existsSync(full)) return full;
      }
    }
    return null;
  }
  const excelPaths = {
    stock: findExcel(excelCandidates.stock),
    daily: findExcel(excelCandidates.daily),
    juneClose: findExcel(excelCandidates.juneClose),
    julyClose: findExcel(excelCandidates.julyClose),
  };
  let plantExcel = null;
  if (Object.values(excelPaths).every(Boolean)) {
    const stockRowsPlant = app.parseStock(XLSX.readFile(excelPaths.stock, { cellDates: true }));
    const dailyRows = app.parseSalesOrReturns(XLSX.readFile(excelPaths.daily, { cellDates: true }), "ventas", "VENTAS DE MAYO Y JUNIO 2026.xlsx");
    const juneCloseRows = app.parseSalesOrReturns(XLSX.readFile(excelPaths.juneClose, { cellDates: true }), "ventas", "ventas junio.xlsx");
    const julyCloseRows = app.parseSalesOrReturns(XLSX.readFile(excelPaths.julyClose, { cellDates: true }), "ventas", "ventas julio.xlsx");
    assert(juneCloseRows.some((row) => row.monthlyTotal), "el Excel de junio dedicado debe parsearse como cierre");
    assert(dailyRows.some((row) => !row.monthlyTotal && String(row.fecha).includes("2026-06") || (row.fecha instanceof Date && row.fecha.getMonth() === 5)), "el combinado debe traer diario de junio");
    const plantResolved = app.resolveCanonicalMonthSources([
      { name: "VENTAS DE MAYO Y JUNIO 2026.xlsx", rows: dailyRows },
      { name: "ventas junio.xlsx", rows: juneCloseRows },
      { name: "ventas julio.xlsx", rows: julyCloseRows },
    ]);
    const plantJune = plantResolved.decisions.find((decision) => decision.month === "2026-06");
    assert(plantJune?.strategy === "keep-daily-and-close", "en los Excel de planta junio diario+cierre son complementarios");
    const plantVentas = plantResolved.entries.flatMap((entry) => entry.rows);
    const juneCoverage = app.buildSalesMonthCoverage(plantVentas).find((row) => row.monthKey === "2026-06");
    assert(juneCoverage?.status === "complete", "junio plantilla debe quedar complete (diario + cierre)");
    const plantJuly = evaluateMonth(app, stockRowsPlant, plantVentas, "2026-07");
    const exclusivePlant = app.resolveCanonicalMonthSources([
      { name: "VENTAS DE MAYO Y JUNIO 2026.xlsx", rows: dailyRows },
      { name: "ventas junio.xlsx", rows: juneCloseRows },
      { name: "ventas julio.xlsx", rows: julyCloseRows },
    ], { "2026-06": "ventas junio.xlsx" });
    const exclusivePlantJuly = evaluateMonth(app, stockRowsPlant, exclusivePlant.entries.flatMap((entry) => entry.rows), "2026-07");
    const gdeAbs = (analysis) => analysis.rows.filter((row) => /GDE$/i.test(row.producto)).reduce((sum, row) => sum + row.absoluteError, 0);
    assert(plantJuly.wape < exclusivePlantJuly.wape, `julio WAPE app-default debe mejorar vs cierre-solo (${plantJuly.wape.toFixed(2)} vs ${exclusivePlantJuly.wape.toFixed(2)})`);
    assert(gdeAbs(plantJuly) < gdeAbs(exclusivePlantJuly) - 30, "el error absoluto GDE de julio debe bajar al conservar el diario de junio");
    plantExcel = {
      julyWape: Number(plantJuly.wape.toFixed(2)),
      julyWapeCloseOnly: Number(exclusivePlantJuly.wape.toFixed(2)),
      gdeAbs: Number(gdeAbs(plantJuly).toFixed(2)),
      gdeAbsCloseOnly: Number(gdeAbs(exclusivePlantJuly).toFixed(2)),
    };
  }

  // Limpieza de outliers de catálogo: intermitentes con $ y dormidos.
  const outlierStock = [
    { producto: "PALETA GALLETA $35", stock: 20, orden: 1 },
    { producto: "BOLLOS C 4", stock: 40, orden: 2 },
    { producto: "CAJITA FELIZ", stock: 10, orden: 3 },
    { producto: "FRUTAS GDE", stock: 50, orden: 4 },
  ];
  const outlierVentas = [];
  const pushClose = (monthKey, product, qty) => {
    const [year, month] = monthKey.split("-").map(Number);
    outlierVentas.push({
      fecha: new Date(year, month - 1, 1),
      producto: product,
      cantidad: qty,
      monthlyTotal: true,
      monthDays: new Date(year, month, 0).getDate(),
    });
  };
  for (const [month, qty] of [
    ["2026-01", 30], ["2026-02", 184], ["2026-03", 2], ["2026-04", 0], ["2026-05", 79], ["2026-06", 296], ["2026-07", 35],
  ]) pushClose(month, "PALETA GALLETA $35", qty);
  for (const [month, qty] of [
    ["2026-02", 369], ["2026-03", 0], ["2026-04", 0], ["2026-05", 364], ["2026-06", 214], ["2026-07", 28], ["2026-08", 0],
  ]) pushClose(month, "BOLLOS C 4", qty);
  for (const [month, qty] of [
    ["2025-04", 599], ["2026-04", 773], ["2026-05", 75], ["2026-06", 0], ["2026-07", 0], ["2026-08", 0],
  ]) pushClose(month, "CAJITA FELIZ", qty);
  for (const [month, qty] of [
    ["2025-07", 510], ["2025-08", 495], ["2025-09", 480], ["2025-10", 470],
    ["2025-11", 455], ["2026-01", 490], ["2026-02", 505], ["2026-03", 260],
    ["2026-04", 480], ["2026-05", 500], ["2026-06", 520], ["2026-07", 580], ["2026-08", 540],
  ]) pushClose(month, "FRUTAS GDE", qty);

  const outlierJuly = evaluateMonth(app, outlierStock, outlierVentas, "2026-07");
  const outlierAugust = evaluateMonth(app, outlierStock, outlierVentas, "2026-08");
  const paletaJuly = outlierJuly.rows.find((row) => row.producto === "PALETA GALLETA $35");
  const cajitaJuly = outlierJuly.rows.find((row) => row.producto === "CAJITA FELIZ");
  const bollosAugust = outlierAugust.rows.find((row) => row.producto === "BOLLOS C 4");
  const frutasJuly = outlierJuly.rows.find((row) => row.producto === "FRUTAS GDE");
  assert(paletaJuly.forecast < 200, `julio PALETA $35 debe bajar del overshoot ~422 (fc ${paletaJuly.forecast.toFixed(1)})`);
  assert(cajitaJuly.forecast <= 5, `julio CAJITA FELIZ no debe inventar venta con junio en 0 (fc ${cajitaJuly.forecast.toFixed(1)})`);
  assert(bollosAugust.forecast < 80, `agosto BOLLOS C 4 debe amortiguar la caída a cero (fc ${bollosAugust.forecast.toFixed(1)})`);
  assert(frutasJuly.forecast > 450, `FRUTAS GDE no debe apagarse por la limpieza (fc ${frutasJuly.forecast.toFixed(1)})`);
  assert(frutasJuly.absoluteError < 100, `FRUTAS GDE debe seguir razonable (abs ${frutasJuly.absoluteError.toFixed(1)})`);

  const paletaPromo = app.normalizeActivePromo({
    producto: "PALETA GALLETA $35",
    startDate: "2026-07-01",
    durationPreset: "hasta_desactivar",
    multiplier: 1.3,
    extraPiecesPerDay: 0,
    active: true,
  });
  const paletaPromoJuly = app.calculateForecast({
    stockRows: outlierStock,
    historicalVentas: app.filterVentasBeforeMonth(outlierVentas, "2026-07"),
    bajas: [],
    existencias: [],
    realProduction: [],
    selectedMonth: "2026-07",
    dailyBufferPct: 10,
    activePromos: [paletaPromo],
  });
  const paletaPromoRow = paletaPromoJuly.find((row) => row.producto === "PALETA GALLETA $35");
  const frutasPromoRow = paletaPromoJuly.find((row) => row.producto === "FRUTAS GDE");
  assert(paletaPromoRow.pronosticoVenta > paletaJuly.forecast + 30, "con promo activa PALETA no debe quedar amortiguada");
  assert(Math.abs(frutasPromoRow.pronosticoVenta - frutasJuly.forecast) < 1, "la promo de PALETA no debe mover el WAPE/base de FRUTAS");
  assert(paletaJuly.forecast < 200, "sin promo el WAPE/base de PALETA sigue usando la limpieza");

  const paletaDailyBase = app.calculateDailyForecast({
    monthlyRows: app.calculateForecast({
      stockRows: outlierStock,
      historicalVentas: app.filterVentasBeforeMonth(outlierVentas, "2026-07"),
      bajas: [],
      existencias: [],
      realProduction: [],
      selectedMonth: "2026-07",
      dailyBufferPct: 10,
    }),
    ventasReales: [],
    realProduction: [],
    selectedMonth: "2026-07",
    dailyBufferPct: 10,
  });
  const paletaDailyPromo = app.calculateDailyForecast({
    monthlyRows: paletaPromoJuly,
    ventasReales: [],
    realProduction: [],
    selectedMonth: "2026-07",
    dailyBufferPct: 10,
    activePromos: [paletaPromo],
  });
  const paletaDayBase = paletaDailyBase.find((row) => row.producto === "PALETA GALLETA $35" && row.weekday === 3);
  const paletaDayPromo = paletaDailyPromo.find((row) => row.producto === "PALETA GALLETA $35" && row.weekday === 3);
  assert(paletaDayPromo.produccionSugeridaDia > paletaDayBase.produccionSugeridaDia, "la planta debe subir el miércoles con promo ×1.3");
  const deactivatedDaily = app.calculateDailyForecast({
    monthlyRows: paletaPromoJuly,
    ventasReales: [],
    realProduction: [],
    selectedMonth: "2026-07",
    dailyBufferPct: 10,
    activePromos: [{ ...paletaPromo, active: false }],
  });
  const deactivatedDay = deactivatedDaily.find((row) => row.producto === "PALETA GALLETA $35" && row.weekday === 3);
  assert(deactivatedDay.produccionSugeridaDia < paletaDayPromo.produccionSugeridaDia, "desactivar quita el multiplicador diario");
  assert(!deactivatedDay.promoActiva, "desactivar no debe marcar el día como promo");

  // Residual post-#12: alias que #12 no cubrió (INDIVIDUAL, CHEESE CAKE,
  // TRES LECHES, DE), gelatinas $150 con hueco may–jun y pico julio, e
  // impulso por SKU en Mini/MED. Magnitudes de CONTROL / julio congelado.
  const leftoverStock = [
    { producto: "FRUTAS GDE", stock: 40, orden: 1 },
    { producto: "MOKA GDE", stock: 40, orden: 2 },
    { producto: "PAY DE GUAYABA GDE", stock: 20, orden: 3 },
    { producto: "DURAZNO GDE", stock: 20, orden: 4 },
    { producto: "CHOCOLATE GDE", stock: 20, orden: 5 },
    { producto: "DUO DURAZNO GDE", stock: 20, orden: 6 },
    { producto: "3 LECHES GDE", stock: 20, orden: 7 },
    { producto: "RED VELVET CHEESE CAKE GDE", stock: 20, orden: 8 },
    { producto: "GELATINA IND FRESA", stock: 40, orden: 9 },
    { producto: "GELATINA IND MOSAICO", stock: 40, orden: 10 },
    { producto: "GELATINA DE PINA IND", stock: 20, orden: 11 },
    { producto: "MINI MEDIANO CHOCOLATE", stock: 40, orden: 12 },
    { producto: "MINI MEDIANO PINERO", stock: 40, orden: 13 },
    { producto: "MOKA MED", stock: 30, orden: 14 },
    { producto: "TRES LECHES MEDIANO", stock: 20, orden: 15 },
    { producto: "GELATINA FRESA $150", stock: 20, orden: 16 },
    { producto: "FRUTAS MED", stock: 30, orden: 17 },
  ];
  const leftoverSeries = {
    "FRUTAS GDE": {
      "2025-04": 500, "2025-05": 658, "2025-06": 519, "2025-07": 475, "2025-08": 524,
      "2026-04": 500, "2026-05": 660, "2026-07": 580, "2026-08": 541,
    },
    "MOKA GDE": {
      "2025-04": 605, "2025-05": 755, "2025-06": 624, "2025-07": 612, "2025-08": 639,
      "2026-04": 600, "2026-05": 720, "2026-07": 697, "2026-08": 619,
    },
    "PAY DE GUAYABA GRANDE": {
      "2025-04": 340, "2025-05": 420, "2025-06": 360, "2025-07": 330, "2025-08": 350,
      "2026-04": 338, "2026-05": 418, "2026-07": 413, "2026-08": 352,
    },
    "DURAZNO GRANDE": {
      "2025-04": 240, "2025-05": 310, "2025-06": 250, "2025-07": 230, "2025-08": 245,
      "2026-04": 238, "2026-05": 308, "2026-07": 322, "2026-08": 248,
    },
    "CHOCOLATE GDE": {
      "2025-04": 350, "2025-05": 430, "2025-06": 360, "2025-07": 340, "2025-08": 355,
      "2026-04": 348, "2026-05": 428, "2026-07": 390, "2026-08": 360,
    },
    "DUO DURAZNO GDE": {
      "2025-04": 290, "2025-05": 360, "2025-06": 300, "2025-07": 280, "2025-08": 295,
      "2026-04": 288, "2026-05": 358, "2026-07": 340, "2026-08": 300,
    },
    "TRES LECHES GDE": {
      "2025-04": 250, "2025-05": 320, "2025-06": 260, "2025-07": 240, "2025-08": 255,
      "2026-04": 248, "2026-05": 318, "2026-07": 290, "2026-08": 258,
    },
    "RED VELVET CHEESECAKE GDE": {
      "2025-04": 170, "2025-05": 220, "2025-06": 180, "2025-07": 165, "2025-08": 175,
      "2026-04": 168, "2026-05": 218, "2026-07": 205, "2026-08": 178,
    },
    "GELATINA INDIVIDUAL FRESA": {
      "2025-04": 790, "2025-05": 980, "2025-06": 800, "2025-07": 810, "2025-08": 805,
      "2026-04": 795, "2026-05": 990, "2026-07": 820, "2026-08": 812,
    },
    "GELATINA INDIVIDUAL MOSAICO": {
      "2025-04": 690, "2025-05": 860, "2025-06": 700, "2025-07": 705, "2025-08": 698,
      "2026-04": 692, "2026-05": 870, "2026-06": 710, "2026-07": 708, "2026-08": 702,
    },
    "GELATINA PINA IND": {
      "2025-04": 55, "2025-05": 70, "2025-06": 58, "2025-07": 56, "2025-08": 57,
      "2026-04": 54, "2026-05": 68, "2026-06": 59, "2026-07": 60, "2026-08": 58,
    },
    "MINI MEDIANO CHOCOLATE": {
      "2025-04": 1080, "2025-05": 1320, "2025-06": 1100, "2025-07": 1085, "2025-08": 1110,
      "2026-04": 1075, "2026-05": 1310, "2026-07": 1240, "2026-08": 1120,
    },
    "MINI MEDIANO PINERO": {
      "2025-04": 1000, "2025-05": 1240, "2025-06": 1020, "2025-07": 1010, "2025-08": 1030,
      "2026-04": 998, "2026-05": 1230, "2026-07": 1160, "2026-08": 1040,
    },
    "MOKA MEDIANO": {
      "2025-04": 590, "2025-05": 730, "2025-06": 600, "2025-07": 580, "2025-08": 595,
      "2026-04": 588, "2026-05": 725, "2026-07": 680, "2026-08": 600,
    },
    "3 LECHES MEDIANO": {
      "2025-04": 290, "2025-05": 360, "2025-06": 300, "2025-07": 285, "2025-08": 295,
      "2026-04": 288, "2026-05": 358, "2026-06": 302, "2026-07": 320, "2026-08": 298,
    },
    "GELATINA FRESA $150": {
      "2025-04": 400, "2025-05": 0, "2025-06": 0, "2025-07": 420, "2025-08": 80,
      "2026-04": 410, "2026-05": 0, "2026-06": 0, "2026-07": 438, "2026-08": 70,
    },
    "FRUTAS MED": {
      "2025-04": 510, "2025-05": 640, "2025-06": 520, "2025-07": 500, "2025-08": 530,
      "2026-04": 512, "2026-05": 642, "2026-06": 522, "2026-07": 530, "2026-08": 528,
    },
  };
  const leftoverVentas = [];
  for (const [product, months] of Object.entries(leftoverSeries)) {
    for (const [month, qty] of Object.entries(months)) leftoverVentas.push(monthClose(month, product, qty));
  }
  leftoverVentas.push(...juneDaily("FRUTAS GDE", 223, 296));
  leftoverVentas.push(...juneDaily("MOKA GDE", 276, 341));
  leftoverVentas.push(...juneDaily("PAY DE GUAYABA GRANDE", 155, 205));
  leftoverVentas.push(...juneDaily("DURAZNO GRANDE", 108, 148));
  leftoverVentas.push(...juneDaily("CHOCOLATE GDE", 155, 205));
  leftoverVentas.push(...juneDaily("DUO DURAZNO GDE", 128, 172));
  leftoverVentas.push(...juneDaily("TRES LECHES GDE", 112, 150));
  leftoverVentas.push(...juneDaily("RED VELVET CHEESECAKE GDE", 78, 104));
  leftoverVentas.push(...juneDaily("MINI MEDIANO CHOCOLATE", 470, 640));
  leftoverVentas.push(...juneDaily("MINI MEDIANO PINERO", 440, 590));
  leftoverVentas.push(...juneDaily("MOKA MEDIANO", 255, 345));
  leftoverVentas.push(...juneDaily("GELATINA INDIVIDUAL FRESA", 350, 466));

  const leftoverJune = evaluateMonth(app, leftoverStock, leftoverVentas, "2026-06");
  const leftoverJuly = evaluateMonth(app, leftoverStock, leftoverVentas, "2026-07");
  const leftoverAugust = evaluateMonth(app, leftoverStock, leftoverVentas, "2026-08");
  const leftoverPick = (analysis, name) => analysis.rows.find((row) => app.normalizeProduct(row.producto) === app.normalizeProduct(name));
  const leftoverWeighted = (() => {
    const months = [leftoverJune, leftoverJuly, leftoverAugust];
    const actual = months.reduce((sum, row) => sum + row.actual, 0);
    const abs = months.reduce((sum, row) => sum + row.rows.reduce((inner, item) => inner + item.absoluteError, 0), 0);
    return actual > 0 ? (abs / actual) * 100 : null;
  })();

  const leftGelatinaJuly = leftoverPick(leftoverJuly, "GELATINA IND FRESA");
  const leftMosaicoJuly = leftoverPick(leftoverJuly, "GELATINA IND MOSAICO");
  const leftPinaJuly = leftoverPick(leftoverJuly, "GELATINA DE PINA IND");
  const leftTresJuly = leftoverPick(leftoverJuly, "3 LECHES GDE");
  const leftVelvetJuly = leftoverPick(leftoverJuly, "RED VELVET CHEESE CAKE GDE");
  const leftTresMedJuly = leftoverPick(leftoverJuly, "TRES LECHES MEDIANO");
  const leftDollarJuly = leftoverPick(leftoverJuly, "GELATINA FRESA $150");
  const leftDollarAugust = leftoverPick(leftoverAugust, "GELATINA FRESA $150");
  const leftMiniChocJuly = leftoverPick(leftoverJuly, "MINI MEDIANO CHOCOLATE");
  const leftMiniPinJuly = leftoverPick(leftoverJuly, "MINI MEDIANO PINERO");
  const leftMokaMedJuly = leftoverPick(leftoverJuly, "MOKA MED");
  const leftFrutasMedJuly = leftoverPick(leftoverJuly, "FRUTAS MED");
  const leftFrutasGdeJuly = leftoverPick(leftoverJuly, "FRUTAS GDE");

  assert(leftoverJune.wape < 6, `junio residual no debe romperse (WAPE ${leftoverJune.wape.toFixed(2)}%)`);
  assert(leftoverJuly.wape < 8, `julio residual debe bajar del ~14.2% silencioso / peor con alias (WAPE ${leftoverJuly.wape.toFixed(2)}%)`);
  assert(leftoverAugust.wape < 6, `agosto residual no debe copiar el salto de julio (WAPE ${leftoverAugust.wape.toFixed(2)}%)`);
  assert(leftoverWeighted < 4.5, `WAPE ponderado residual jun-ago debe bajar del ~4.8% post-#14 (${leftoverWeighted.toFixed(2)}%)`);

  assert(leftGelatinaJuly.actual > 700, "GELATINA INDIVIDUAL debe contar en el actual de IND FRESA");
  assert(leftGelatinaJuly.forecast > 700, `julio GELATINA IND FRESA no puede quedar en 0 por alias (fc ${leftGelatinaJuly.forecast.toFixed(1)})`);
  assert(leftMosaicoJuly.forecast > 500, `julio MOSAICO INDIVIDUAL debe empatar (fc ${leftMosaicoJuly.forecast.toFixed(1)})`);
  assert(leftPinaJuly.forecast > 40, `julio GELATINA DE PINA / PINA IND debe empatar (fc ${leftPinaJuly.forecast.toFixed(1)})`);
  assert(leftTresJuly.forecast > 200, `julio TRES LECHES debe empatar con 3 LECHES (fc ${leftTresJuly.forecast.toFixed(1)})`);
  assert(leftVelvetJuly.forecast > 140, `julio CHEESE CAKE debe empatar con CHEESECAKE (fc ${leftVelvetJuly.forecast.toFixed(1)})`);
  assert(leftTresMedJuly.forecast > 200, `julio 3 LECHES MEDIANO debe empatar con TRES LECHES MED (fc ${leftTresMedJuly.forecast.toFixed(1)})`);

  assert(leftDollarJuly.forecast > 300, `julio GELATINA $150 no es baja: debe usar julio 2025 (fc ${leftDollarJuly.forecast.toFixed(1)})`);
  assert(leftDollarJuly.absoluteError < 80, `julio $150 debe acercarse a 438 (abs ${leftDollarJuly.absoluteError.toFixed(1)})`);
  assert(leftDollarAugust.forecast < 150, `agosto $150 no debe corregir el fade 80 como dip (fc ${leftDollarAugust.forecast.toFixed(1)})`);
  assert(leftDollarAugust.absoluteError < 80, `agosto $150 debe quedar cerca de 70 (abs ${leftDollarAugust.absoluteError.toFixed(1)})`);

  assert(leftMiniChocJuly.forecast > 1150, `julio MINI CHOCOLATE debe impulsarse (fc ${leftMiniChocJuly.forecast.toFixed(1)})`);
  assert(leftMiniChocJuly.absoluteError < 90, `julio MINI CHOCOLATE debe bajar del faltante 141 (abs ${leftMiniChocJuly.absoluteError.toFixed(1)})`);
  assert(leftMiniPinJuly.forecast > 1080, `julio MINI PINERO debe impulsarse (fc ${leftMiniPinJuly.forecast.toFixed(1)})`);
  assert(leftMokaMedJuly.forecast > 620, `julio MOKA MED debe impulsarse por su propio diario (fc ${leftMokaMedJuly.forecast.toFixed(1)})`);
  assert(/impulso reciente/i.test(leftMiniChocJuly.metodo || ""), "MINI CHOCOLATE debe anotar impulso reciente");
  assert(/impulso reciente/i.test(leftMokaMedJuly.metodo || ""), "MOKA MED debe anotar impulso reciente");
  assert(!/impulso reciente/i.test(leftFrutasMedJuly.metodo || ""), "FRUTAS MED sin diario no hereda impulso");
  assert(leftFrutasMedJuly.absoluteError < 45, `FRUTAS MED sigue acotado (abs ${leftFrutasMedJuly.absoluteError.toFixed(1)})`);
  assert(leftFrutasGdeJuly.forecast > 540, `el impulso GDE de julio debe seguir (fc ${leftFrutasGdeJuly.forecast.toFixed(1)})`);

  const leftMiniChocAugust = leftoverPick(leftoverAugust, "MINI MEDIANO CHOCOLATE");
  const leftMiniPinAugust = leftoverPick(leftoverAugust, "MINI MEDIANO PINERO");
  const leftMokaMedAugust = leftoverPick(leftoverAugust, "MOKA MED");
  const leftDuraznoAugust = leftoverPick(leftoverAugust, "DURAZNO GDE");
  assert(leftMiniChocAugust.absoluteError < 50, `agosto MINI CHOCOLATE no hereda el salto de julio (abs ${leftMiniChocAugust.absoluteError.toFixed(1)})`);
  assert(leftMiniPinAugust.absoluteError < 50, `agosto MINI PINERO no hereda el salto de julio (abs ${leftMiniPinAugust.absoluteError.toFixed(1)})`);
  assert(leftMokaMedAugust.absoluteError < 35, `agosto MOKA MED no hereda el salto de julio (abs ${leftMokaMedAugust.absoluteError.toFixed(1)})`);
  assert(leftDuraznoAugust.absoluteError < 60, `agosto DURAZNO no hereda el salto de julio (abs ${leftDuraznoAugust.absoluteError.toFixed(1)})`);
  assert(!/impulso reciente/i.test(leftMiniChocAugust.metodo || ""), "agosto Mini no vuelve a aplicar impulso");

  // Backtest parcial 2025 (Mar–Nov, sin año anterior): los picos de un mes
  // se copiaban al siguiente. CAJITA FELIZ abril→mayo, PETIT 3 LECHES y el
  // kilo de galleta de Día de las Madres. Marzo de PETIT se recupera del
  // pronóstico de abril (un mes de historia, último mes por día, calibración 1):
  // 343.55 * 31/30 = 355. El resto son actuals publicados en el top de error.
  const spikeStock = [
    { producto: "MINI MEDIANO CHOCOLATE", stock: 40, orden: 1 },
    { producto: "FRUTAS CH", stock: 30, orden: 2 },
    { producto: "CAJITA FELIZ", stock: 10, orden: 3 },
    { producto: "1 2 KG DE GALLETA", stock: 20, orden: 4 },
    { producto: "PETIT 3 LECHES CHOCOLATE", stock: 20, orden: 5 },
  ];
  const spikeSeries = {
    "MINI MEDIANO CHOCOLATE": {
      "2025-03": 1143, "2025-04": 1030, "2025-05": 1268, "2025-06": 1348,
      "2025-07": 1107, "2025-08": 1204, "2025-09": 1226, "2025-10": 1192, "2025-11": 1245,
    },
    "FRUTAS CH": {
      "2025-03": 410, "2025-04": 450, "2025-05": 656, "2025-06": 529, "2025-07": 451, "2025-08": 519,
    },
    "CAJITA FELIZ": {
      "2025-04": 599, "2025-05": 37, "2025-06": 0, "2025-07": 0,
    },
    "1 2 KG DE GALLETA": {
      "2025-04": 226, "2025-05": 619, "2025-06": 375, "2025-07": 263, "2025-08": 285,
    },
    "PETIT 3 LECHES CHOCOLATE": {
      "2025-03": 355, "2025-04": 895, "2025-05": 525,
    },
  };
  const spikeVentas = [];
  for (const [product, months] of Object.entries(spikeSeries)) {
    for (const [month, qty] of Object.entries(months)) spikeVentas.push(monthClose(month, product, qty));
  }
  const spikeMonths = ["2025-05", "2025-06", "2025-07", "2025-08"];
  const spikeByMonth = Object.fromEntries(
    [...spikeMonths, "2025-11"].map((month) => [month, evaluateMonth(app, spikeStock, spikeVentas, month)])
  );
  const spikePick = (month, name) => spikeByMonth[month].rows.find((row) => row.producto === name);
  const cajitaMay = spikePick("2025-05", "CAJITA FELIZ");
  const petitMay = spikePick("2025-05", "PETIT 3 LECHES CHOCOLATE");
  const galletaJune = spikePick("2025-06", "1 2 KG DE GALLETA");
  const frutasJune = spikePick("2025-06", "FRUTAS CH");
  const miniJuly = spikePick("2025-07", "MINI MEDIANO CHOCOLATE");
  const miniNov = spikePick("2025-11", "MINI MEDIANO CHOCOLATE");
  assert(cajitaMay.forecast <= 250, `mayo CAJITA FELIZ no debe copiar abril 599 (fc ${cajitaMay.forecast.toFixed(1)})`);
  assert(cajitaMay.absoluteError < 400, `mayo CAJITA debe bajar del |e| 582 (abs ${cajitaMay.absoluteError.toFixed(1)})`);
  assert(/pico sin soporte/i.test(cajitaMay.metodo || ""), "CAJITA mayo debe anotar el recorte de pico");
  assert(petitMay.forecast < 650, `mayo PETIT no debe copiar abril 895 (fc ${petitMay.forecast.toFixed(1)})`);
  assert(petitMay.forecast > 300, `mayo PETIT no debe apagarse (fc ${petitMay.forecast.toFixed(1)})`);
  assert(petitMay.absoluteError < 250, `mayo PETIT debe bajar del |e| 504 (abs ${petitMay.absoluteError.toFixed(1)})`);
  assert(galletaJune.forecast < 420, `junio 1/2 kg no debe arrastrar mayo 619 (fc ${galletaJune.forecast.toFixed(1)})`);
  assert(galletaJune.absoluteError < 120, `junio 1/2 kg debe bajar del |e| 185 (abs ${galletaJune.absoluteError.toFixed(1)})`);
  assert(frutasJune.absoluteError < 80, `junio FRUTAS CH no debe perder el resguardo de Madres (abs ${frutasJune.absoluteError.toFixed(1)})`);
  assert(miniJuly.absoluteError < 200, `julio MINI no debe volver al overshoot 373 (abs ${miniJuly.absoluteError.toFixed(1)})`);
  assert(miniNov.absoluteError < 80, `noviembre MINI en régimen no debe empeorar (abs ${miniNov.absoluteError.toFixed(1)})`);

  const miniMay = spikePick("2025-05", "MINI MEDIANO CHOCOLATE");
  assert(miniMay.forecast > 1040, `mayo MINI no debe heredar el recorte de abril (fc ${miniMay.forecast.toFixed(1)})`);
  assert(miniMay.absoluteError < 240, `mayo MINI debe bajar del |e| 286 (abs ${miniMay.absoluteError.toFixed(1)})`);
  assert(/sin recorte pre-Madres/i.test(miniMay.metodo || ""), "mayo MINI sin año anterior debe quitar el recorte");

  const madresAnchored = evaluateMonth(
    app,
    [{ producto: "MINI MEDIANO CHOCOLATE", stock: 40, orden: 1 }],
    [
      monthClose("2025-03", "MINI MEDIANO CHOCOLATE", 1143),
      monthClose("2025-04", "MINI MEDIANO CHOCOLATE", 1080),
      monthClose("2025-05", "MINI MEDIANO CHOCOLATE", 1320),
      monthClose("2026-04", "MINI MEDIANO CHOCOLATE", 1030),
    ],
    "2026-05"
  );
  const miniMayAnchored = madresAnchored.rows.find((row) => row.producto === "MINI MEDIANO CHOCOLATE");
  assert(
    !/sin recorte pre-Madres/i.test(miniMayAnchored.metodo || ""),
    "mayo con el mismo mes del año anterior no usa el resguardo de arranque en frío"
  );

  const debutStock = [
    { producto: "CHESSECAKE PETIT", stock: 20, orden: 1 },
    { producto: "CAJITA 3 LECHES", stock: 20, orden: 2 },
    { producto: "GELATINA IND MOSAICO", stock: 30, orden: 3 },
  ];
  const debutVentas = [
    monthClose("2025-03", "CHESSECAKE PETIT", 438),
    monthClose("2025-04", "CHESSECAKE PETIT", 439),
    monthClose("2025-03", "CAJITA 3 LECHES", 410),
    monthClose("2025-04", "CAJITA 3 LECHES", 328),
    monthClose("2025-03", "GELATINA IND MOSAICO", 610),
    monthClose("2025-04", "GELATINA IND MOSAICO", 532),
    monthClose("2025-05", "GELATINA IND MOSAICO", 685),
  ];
  const debutApril = evaluateMonth(app, debutStock, debutVentas, "2025-04");
  const debutMay = evaluateMonth(app, debutStock, debutVentas, "2025-05");
  const debutPick = (analysis, name) => analysis.rows.find((row) => row.producto === name);
  const cheeseApril = debutPick(debutApril, "CHESSECAKE PETIT");
  const cajita3April = debutPick(debutApril, "CAJITA 3 LECHES");
  const mosaicoMay = debutPick(debutMay, "GELATINA IND MOSAICO");
  assert(cheeseApril.forecast > 350, `abril CHESSECAKE PETIT no debe caer a 35% de marzo (fc ${cheeseApril.forecast.toFixed(1)})`);
  assert(cheeseApril.absoluteError < 80, `abril CHESSECAKE PETIT debe acercarse a 439 (abs ${cheeseApril.absoluteError.toFixed(1)})`);
  assert(!/pico sin soporte/i.test(cheeseApril.metodo || ""), "CHESSECAKE PETIT de marzo es nivel, no estreno");
  assert(cajita3April.forecast > 320, `abril CAJITA 3 LECHES no debe caer a 35% de marzo (fc ${cajita3April.forecast.toFixed(1)})`);
  assert(cajita3April.absoluteError < 120, `abril CAJITA 3 LECHES debe bajar del |e| 185 (abs ${cajita3April.absoluteError.toFixed(1)})`);
  assert(!/pico sin soporte/i.test(cajita3April.metodo || ""), "CAJITA 3 LECHES de marzo es nivel, no estreno");
  assert(mosaicoMay.forecast > 530, `mayo MOSAICO no debe heredar el recorte de abril (fc ${mosaicoMay.forecast.toFixed(1)})`);
  assert(mosaicoMay.absoluteError < 170, `mayo MOSAICO debe bajar del |e| 192 (abs ${mosaicoMay.absoluteError.toFixed(1)})`);
  assert(/sin recorte pre-Madres/i.test(mosaicoMay.metodo || ""), "mayo MOSAICO sin año anterior debe quitar el recorte");

  const scored = [];
  for (const month of spikeMonths) {
    for (const row of spikeByMonth[month].rows) {
      if ((spikeSeries[row.producto] || {})[month] == null) continue;
      scored.push(row);
    }
  }
  const spikeActual = scored.reduce((sum, row) => sum + row.actual, 0);
  const spikeAbs = scored.reduce((sum, row) => sum + row.absoluteError, 0);
  const spikeWape = spikeActual > 0 ? (spikeAbs / spikeActual) * 100 : null;
  // |e| publicado en f923484 para estos mismos meses/SKU (top de error; junio
  // MINI se recupera del acumulado: fc 1343, |e| 4.8).
  const publishedAbs = 286.17 + 4.84 + 372.94 + 139.95
    + 162.1 + 201.06 + 109.18 + 149.38
    + 581.97
    + 423.54 + 185.19 + 110.81 + 106.16
    + 504.25;
  assert(spikeAbs < publishedAbs - 400, `la canasta May–Ago debe bajar el |e| publicado ${publishedAbs.toFixed(0)} (ahora ${spikeAbs.toFixed(0)})`);

  // Pepes 2025, arranque sin el mismo mes del año anterior. Magnitudes de
  // los meses que dominan Σ|e| (informe @ d664fd2). El |e| "antes" es el
  // publicado en ese backtest; aquí se exige que el guardia lo baje.
  const pepesStock = [
    { producto: "MINI MED CHOCOLATE", stock: 40, orden: 1 },
    { producto: "GELATINA IND FRESA", stock: 40, orden: 2 },
    { producto: "PETIT 3 LECHES PINERO", stock: 20, orden: 3 },
    { producto: "PETIT 3 LECHES CHOCOLATE", stock: 20, orden: 4 },
    { producto: "PETIT DECORADO", stock: 20, orden: 5 },
    { producto: "1 4 KG GALLETA", stock: 20, orden: 6 },
    { producto: "1 2 KG GALLETA", stock: 20, orden: 7 },
    { producto: "MINI MED PINERO", stock: 40, orden: 8 },
    { producto: "MOKA CH", stock: 30, orden: 9 },
    { producto: "FRUTAS CH", stock: 30, orden: 10 },
    { producto: "MOKA MED", stock: 30, orden: 11 },
    { producto: "FRUTAS GDE", stock: 40, orden: 12 },
    { producto: "PAY GUAYABA GDE", stock: 30, orden: 13 },
    { producto: "GELATINA FRESA GDE", stock: 30, orden: 14 },
    { producto: "PINA MED", stock: 20, orden: 15 },
  ];
  const pepesSeries = {
    "MINI MED CHOCOLATE": {
      "2025-02": 2300, "2025-03": 2254, "2025-04": 2054, "2025-05": 2614,
      "2025-06": 2775, "2025-07": 2400, "2025-08": 2380, "2025-09": 2400, "2025-10": 2417, "2025-11": 2350,
    },
    "MINI MED PINERO": {
      "2025-02": 2200, "2025-03": 2081, "2025-04": 2100, "2025-05": 2480,
      "2025-06": 2509, "2025-07": 2300, "2025-08": 2241, "2025-09": 2393, "2025-10": 2293, "2025-11": 2300,
    },
    "GELATINA IND FRESA": {
      "2025-02": 2200, "2025-03": 2100, "2025-04": 2080, "2025-05": 2571,
      "2025-06": 2300, "2025-07": 2066, "2025-08": 2100, "2025-09": 2100, "2025-10": 2050, "2025-11": 1945,
    },
    "PETIT 3 LECHES PINERO": {
      "2025-03": 755, "2025-04": 1206, "2025-05": 857, "2025-06": 800, "2025-07": 780,
      "2025-08": 934, "2025-09": 820, "2025-10": 800, "2025-11": 819,
    },
    "PETIT 3 LECHES CHOCOLATE": {
      "2025-03": 819, "2025-04": 1824, "2025-05": 900, "2025-06": 850,
      "2025-07": 819, "2025-08": 984,
    },
    "PETIT DECORADO": {
      "2025-03": 620, "2025-04": 377, "2025-05": 860, "2025-06": 700,
    },
    "1 4 KG GALLETA": {
      "2025-02": 520, "2025-03": 589, "2025-04": 540, "2025-05": 1295,
      "2025-06": 560, "2025-07": 540, "2025-08": 550, "2025-09": 530, "2025-10": 545,
      "2025-11": 560, "2025-12": 1565,
    },
    "1 2 KG GALLETA": {
      "2025-02": 500, "2025-03": 540, "2025-04": 510, "2025-05": 1177,
      "2025-06": 500, "2025-07": 490, "2025-08": 510, "2025-09": 500, "2025-10": 505,
      "2025-11": 520, "2025-12": 1148,
    },
    "MOKA CH": {
      "2025-02": 960, "2025-03": 950, "2025-04": 940, "2025-05": 1405,
      "2025-06": 980, "2025-07": 970, "2025-08": 965, "2025-09": 970, "2025-10": 960, "2025-11": 960,
    },
    "FRUTAS CH": {
      "2025-02": 910, "2025-03": 900, "2025-04": 890, "2025-05": 1326,
      "2025-06": 920, "2025-07": 910, "2025-08": 905, "2025-09": 910, "2025-10": 900, "2025-11": 900,
    },
    "MOKA MED": {
      "2025-02": 1070, "2025-03": 1060, "2025-04": 1050, "2025-05": 1479,
      "2025-06": 1080, "2025-07": 1070, "2025-08": 1065, "2025-09": 1070, "2025-10": 1060, "2025-11": 1060,
    },
    "FRUTAS GDE": {
      "2025-02": 1020, "2025-03": 1010, "2025-04": 1000, "2025-05": 1180,
      "2025-06": 1020, "2025-07": 1015, "2025-08": 1010, "2025-09": 1010, "2025-10": 1020, "2025-11": 1015, "2025-12": 1453,
    },
    "PAY GUAYABA GDE": {
      "2025-02": 660, "2025-03": 650, "2025-04": 640, "2025-05": 750,
      "2025-06": 650, "2025-07": 645, "2025-08": 640, "2025-09": 640, "2025-10": 650, "2025-11": 655, "2025-12": 1221,
    },
    "GELATINA FRESA GDE": {
      "2025-02": 470, "2025-03": 460, "2025-04": 450, "2025-05": 520,
      "2025-06": 460, "2025-07": 455, "2025-08": 450, "2025-09": 450, "2025-10": 455, "2025-11": 460, "2025-12": 897,
    },
    "PINA MED": {
      "2025-02": 510, "2025-03": 500, "2025-04": 495, "2025-05": 580,
      "2025-06": 500, "2025-07": 505, "2025-08": 500, "2025-09": 505, "2025-10": 500, "2025-11": 498, "2025-12": 75,
    },
  };
  const pepesVentas = [];
  for (const [product, months] of Object.entries(pepesSeries)) {
    for (const [month, qty] of Object.entries(months)) pepesVentas.push(monthClose(month, product, qty));
  }
  const pepesMonths = ["2025-04", "2025-05", "2025-06", "2025-09", "2025-11", "2025-12"];
  const pepesByMonth = Object.fromEntries(
    pepesMonths.map((month) => [month, evaluateMonth(app, pepesStock, pepesVentas, month)])
  );
  const pepesPick = (month, name) => pepesByMonth[month].rows.find((row) => row.producto === name);
  const pepesMiniMay = pepesPick("2025-05", "MINI MED CHOCOLATE");
  const pepesPinMay = pepesPick("2025-05", "MINI MED PINERO");
  const pepesGelMay = pepesPick("2025-05", "GELATINA IND FRESA");
  const pepesPetitPinApr = pepesPick("2025-04", "PETIT 3 LECHES PINERO");
  const pepesPetitChocApr = pepesPick("2025-04", "PETIT 3 LECHES CHOCOLATE");
  const pepesDecoradoApr = pepesPick("2025-04", "PETIT DECORADO");
  const pepesPetitPinMay = pepesPick("2025-05", "PETIT 3 LECHES PINERO");
  const pepesQuarterMay = pepesPick("2025-05", "1 4 KG GALLETA");
  const pepesHalfMay = pepesPick("2025-05", "1 2 KG GALLETA");
  const pepesQuarterDec = pepesPick("2025-12", "1 4 KG GALLETA");
  const pepesHalfDec = pepesPick("2025-12", "1 2 KG GALLETA");
  const pepesMiniJun = pepesPick("2025-06", "MINI MED CHOCOLATE");
  const pepesPinJun = pepesPick("2025-06", "MINI MED PINERO");
  const pepesGelJun = pepesPick("2025-06", "GELATINA IND FRESA");
  const pepesMiniSep = pepesPick("2025-09", "MINI MED CHOCOLATE");
  const pepesMiniNov = pepesPick("2025-11", "MINI MED PINERO");
  const pepesGelNov = pepesPick("2025-11", "GELATINA IND FRESA");
  const pepesMokaChMay = pepesPick("2025-05", "MOKA CH");
  const pepesFrutasChMay = pepesPick("2025-05", "FRUTAS CH");
  const pepesMokaMedMay = pepesPick("2025-05", "MOKA MED");
  const pepesFrutasGdeMay = pepesPick("2025-05", "FRUTAS GDE");
  const pepesFrutasGdeDec = pepesPick("2025-12", "FRUTAS GDE");
  const pepesPayDec = pepesPick("2025-12", "PAY GUAYABA GDE");
  const pepesGelGdeDec = pepesPick("2025-12", "GELATINA FRESA GDE");
  const pepesPinaDec = pepesPick("2025-12", "PINA MED");
  const pepesMokaSep = pepesPick("2025-09", "MOKA CH");
  const pepesFrutasGdeSep = pepesPick("2025-09", "FRUTAS GDE");
  const pepesMokaJun = pepesPick("2025-06", "MOKA CH");

  assert(pepesMiniMay.absoluteError < 400, `mayo MINI MED CHOCOLATE debe bajar del |e| 557 (abs ${pepesMiniMay.absoluteError.toFixed(1)}, fc ${pepesMiniMay.forecast.toFixed(1)})`);
  assert(/impulso frío Día de las Madres/i.test(pepesMiniMay.metodo || ""), "mayo MINI debe anotar el impulso frío");
  assert(pepesPinMay.absoluteError < 350, `mayo MINI PINERO debe acercarse (abs ${pepesPinMay.absoluteError.toFixed(1)})`);
  assert(pepesGelMay.absoluteError < 250, `mayo GELATINA IND FRESA debe bajar del |e| 395 (abs ${pepesGelMay.absoluteError.toFixed(1)}, fc ${pepesGelMay.forecast.toFixed(1)})`);
  assert(/impulso frío Día de las Madres/i.test(pepesGelMay.metodo || ""), "mayo gelatina individual debe anotar el impulso");

  assert(pepesPetitPinApr.forecast > 900, `abril PETIT PINERO debe subir de ~731 (fc ${pepesPetitPinApr.forecast.toFixed(1)})`);
  assert(pepesPetitPinApr.absoluteError < 350, `abril PETIT PINERO debe bajar del |e| 475 (abs ${pepesPetitPinApr.absoluteError.toFixed(1)})`);
  assert(/Semana Santa/i.test(pepesPetitPinApr.metodo || ""), "abril PETIT 3 LECHES debe anotar Semana Santa");
  assert(pepesPetitChocApr.forecast > 1000, `abril PETIT CHOCOLATE debe subir de ~793 (fc ${pepesPetitChocApr.forecast.toFixed(1)})`);
  assert(pepesPetitChocApr.absoluteError < 900, `abril PETIT CHOCOLATE debe bajar del |e| 1031 (abs ${pepesPetitChocApr.absoluteError.toFixed(1)})`);
  assert(!/Semana Santa/i.test(pepesDecoradoApr.metodo || ""), "PETIT DECORADO no hereda el impulso de 3 leches");
  assert(pepesPetitPinMay.forecast < 1100, `mayo PETIT PINERO no debe copiar abril 1206 (fc ${pepesPetitPinMay.forecast.toFixed(1)})`);
  assert(pepesPetitPinMay.absoluteError < 350, `mayo PETIT PINERO debe bajar del |e| 530 (abs ${pepesPetitPinMay.absoluteError.toFixed(1)})`);

  assert(pepesQuarterMay.absoluteError < 500, `mayo 1/4 KG debe bajar del |e| 690 (abs ${pepesQuarterMay.absoluteError.toFixed(1)}, fc ${pepesQuarterMay.forecast.toFixed(1)})`);
  assert(pepesHalfMay.absoluteError < 450, `mayo 1/2 KG debe bajar del |e| 549 (abs ${pepesHalfMay.absoluteError.toFixed(1)}, fc ${pepesHalfMay.forecast.toFixed(1)})`);
  assert(/kilo Madres/i.test(pepesQuarterMay.metodo || ""), "mayo kilo debe anotar el impulso");
  assert(pepesQuarterDec.absoluteError < 750, `diciembre 1/4 KG debe bajar del |e| 988 (abs ${pepesQuarterDec.absoluteError.toFixed(1)}, fc ${pepesQuarterDec.forecast.toFixed(1)})`);
  assert(pepesHalfDec.absoluteError < 450, `diciembre 1/2 KG debe bajar del |e| 538 (abs ${pepesHalfDec.absoluteError.toFixed(1)}, fc ${pepesHalfDec.forecast.toFixed(1)})`);
  assert(/kilo Navidad/i.test(pepesHalfDec.metodo || ""), "diciembre kilo debe anotar Navidad");

  assert(pepesMiniJun.forecast > 2500, `junio MINI MED CHOCOLATE debe subir del hombro post-Madres (fc ${pepesMiniJun.forecast.toFixed(1)})`);
  assert(pepesMiniJun.forecast <= 2614, `junio MINI no debe pasar el mayo observado (fc ${pepesMiniJun.forecast.toFixed(1)})`);
  assert(/Día del Padre/i.test(pepesMiniJun.metodo || ""), "junio MINI debe anotar el impulso frío de Padre");
  assert(/Día del Padre/i.test(pepesPinJun.metodo || ""), "junio MINI PINERO debe anotar el impulso");
  assert(pepesPinJun.absoluteError < 120, `junio MINI PINERO debe acercarse a 2509 (abs ${pepesPinJun.absoluteError.toFixed(1)})`);
  assert(!/Día del Padre/i.test(pepesGelJun.metodo || ""), "junio gelatina no hereda el impulso de Padre");
  assert(!/impulso frío/i.test(pepesMiniSep.metodo || ""), "septiembre no debe llevar impulso de evento");
  assert(pepesMiniSep.absoluteError < 180, `septiembre MINI en régimen no debe empeorar (abs ${pepesMiniSep.absoluteError.toFixed(1)})`);
  assert(pepesMiniNov.absoluteError < 200, `noviembre MINI PINERO no debe empeorar (abs ${pepesMiniNov.absoluteError.toFixed(1)})`);
  assert(pepesGelNov.absoluteError < 220, `noviembre GELATINA no debe empeorar (abs ${pepesGelNov.absoluteError.toFixed(1)})`);

  assert(/pastel Madres/i.test(pepesMokaChMay.metodo || ""), "mayo MOKA CH debe anotar el impulso de pastel");
  assert(pepesMokaChMay.forecast > 1050, `mayo MOKA CH debe subir del hombro (~940) (fc ${pepesMokaChMay.forecast.toFixed(1)})`);
  assert(pepesMokaChMay.absoluteError < 380, `mayo MOKA CH debe bajar del |e| 444 (abs ${pepesMokaChMay.absoluteError.toFixed(1)})`);
  assert(/pastel Madres/i.test(pepesFrutasChMay.metodo || ""), "mayo FRUTAS CH debe anotar el impulso de pastel");
  assert(pepesFrutasChMay.absoluteError < 360, `mayo FRUTAS CH debe bajar del |e| 413 (abs ${pepesFrutasChMay.absoluteError.toFixed(1)})`);
  assert(/pastel Madres/i.test(pepesMokaMedMay.metodo || ""), "mayo MOKA MED debe anotar el impulso de pastel");
  assert(pepesMokaMedMay.absoluteError < 360, `mayo MOKA MED debe bajar del |e| 409 (abs ${pepesMokaMedMay.absoluteError.toFixed(1)})`);
  assert(/pastel Madres/i.test(pepesFrutasGdeMay.metodo || ""), "mayo FRUTAS GDE también entra en Madres");
  assert(!/pastel Madres|Navidad/i.test(pepesMokaJun.metodo || ""), "junio MOKA no hereda el impulso de Madres");
  assert(!/impulso frío/i.test(pepesMokaSep.metodo || ""), "septiembre MOKA no lleva impulso de evento");
  assert(pepesMokaSep.absoluteError < 80, `septiembre MOKA en régimen no debe empeorar (abs ${pepesMokaSep.absoluteError.toFixed(1)})`);
  assert(!/impulso frío/i.test(pepesFrutasGdeSep.metodo || ""), "septiembre FRUTAS GDE no lleva impulso de evento");
  assert(pepesFrutasGdeSep.absoluteError < 80, `septiembre FRUTAS GDE en régimen no debe empeorar (abs ${pepesFrutasGdeSep.absoluteError.toFixed(1)})`);

  assert(/pastel Navidad/i.test(pepesFrutasGdeDec.metodo || ""), "diciembre FRUTAS GDE debe anotar Navidad");
  assert(pepesFrutasGdeDec.forecast > 1100, `diciembre FRUTAS GDE debe subir del hombro (fc ${pepesFrutasGdeDec.forecast.toFixed(1)})`);
  assert(pepesFrutasGdeDec.absoluteError < 380, `diciembre FRUTAS GDE debe bajar del |e| 432 (abs ${pepesFrutasGdeDec.absoluteError.toFixed(1)})`);
  assert(/pastel Navidad/i.test(pepesPayDec.metodo || ""), "diciembre PAY GUAYABA debe anotar Navidad");
  assert(pepesPayDec.absoluteError < 520, `diciembre PAY GUAYABA debe bajar del |e| 564 (abs ${pepesPayDec.absoluteError.toFixed(1)})`);
  assert(/gelatina Navidad/i.test(pepesGelGdeDec.metodo || ""), "diciembre gelatina GDE debe anotar Navidad");
  assert(pepesGelGdeDec.absoluteError < 400, `diciembre gelatina GDE debe bajar del |e| 430 (abs ${pepesGelGdeDec.absoluteError.toFixed(1)})`);
  assert(!/Navidad/i.test(pepesPinaDec.metodo || ""), "diciembre PINA MED no hereda el impulso de GDE");
  assert(!/Navidad/i.test(pepesMiniSep.metodo || ""), "septiembre no es Navidad");

  const focusRows = [
    pepesMiniMay, pepesGelMay, pepesPetitPinApr, pepesPetitChocApr, pepesPetitPinMay,
    pepesQuarterMay, pepesHalfMay, pepesQuarterDec, pepesHalfDec,
  ];
  const focusAbs = focusRows.reduce((sum, row) => sum + row.absoluteError, 0);
  const focusBefore = 557.3 + 395.12 + 475.35 + 1031.42 + 529.9 + 689.89 + 548.83 + 988.35 + 537.63;
  assert(focusAbs < focusBefore - 1500, `la canasta foco debe bajar el |e| publicado ${focusBefore.toFixed(0)} (ahora ${focusAbs.toFixed(0)})`);

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
        resolveCanonical: {
          july: Number(resolvedJuly.wape.toFixed(2)),
          julyCloseOnly: Number(exclusiveJuly.wape.toFixed(2)),
          frutas: Number(resolvedFrutas.forecast.toFixed(1)),
          frutasCloseOnly: Number(exclusiveFrutas.forecast.toFixed(1)),
          august: Number(resolvedAugust.wape.toFixed(2)),
        },
        plantExcel,
        catalogCleanup: {
          paletaJuly: Number(paletaJuly.forecast.toFixed(1)),
          cajitaJuly: Number(cajitaJuly.forecast.toFixed(1)),
          bollosAugust: Number(bollosAugust.forecast.toFixed(1)),
        },
        heavySkus: {
          june: Number(heavyJune.wape.toFixed(2)),
          july: Number(heavyJuly.wape.toFixed(2)),
          august: Number(heavyAugust.wape.toFixed(2)),
          weighted: Number(heavyWeighted.toFixed(2)),
          cheeseJuneAbs: Number(heavyCheeseJune.absoluteError.toFixed(1)),
          nutelaJuneAbs: Number(heavyNutelaJune.absoluteError.toFixed(1)),
          mmMedJuneAbs: Number(heavyMmMedJune.absoluteError.toFixed(1)),
          frutasJuly: Number(heavyFrutasJuly.forecast.toFixed(1)),
        },
        spike2025: {
          wapeMayAug: spikeWape != null ? Number(spikeWape.toFixed(2)) : null,
          absMayAug: Number(spikeAbs.toFixed(1)),
          publishedAbs: Number(publishedAbs.toFixed(1)),
          cajitaMay: Number(cajitaMay.forecast.toFixed(1)),
          petitMay: Number(petitMay.forecast.toFixed(1)),
          galletaJune: Number(galletaJune.forecast.toFixed(1)),
          frutasJuneAbs: Number(frutasJune.absoluteError.toFixed(1)),
          miniJulyAbs: Number(miniJuly.absoluteError.toFixed(1)),
          miniNovAbs: Number(miniNov.absoluteError.toFixed(1)),
          miniMay: Number(miniMay.forecast.toFixed(1)),
          miniMayAbs: Number(miniMay.absoluteError.toFixed(1)),
          cheeseApril: Number(cheeseApril.forecast.toFixed(1)),
          cheeseAprilAbs: Number(cheeseApril.absoluteError.toFixed(1)),
          cajita3April: Number(cajita3April.forecast.toFixed(1)),
          cajita3AprilAbs: Number(cajita3April.absoluteError.toFixed(1)),
          mosaicoMay: Number(mosaicoMay.forecast.toFixed(1)),
          mosaicoMayAbs: Number(mosaicoMay.absoluteError.toFixed(1)),
        },
        pepesColdStart: {
          focusAbs: Number(focusAbs.toFixed(1)),
          focusBefore: Number(focusBefore.toFixed(1)),
          miniMay: Number(pepesMiniMay.forecast.toFixed(1)),
          miniMayAbs: Number(pepesMiniMay.absoluteError.toFixed(1)),
          gelMay: Number(pepesGelMay.forecast.toFixed(1)),
          gelMayAbs: Number(pepesGelMay.absoluteError.toFixed(1)),
          petitPinApr: Number(pepesPetitPinApr.forecast.toFixed(1)),
          petitPinAprAbs: Number(pepesPetitPinApr.absoluteError.toFixed(1)),
          petitChocApr: Number(pepesPetitChocApr.forecast.toFixed(1)),
          petitChocAprAbs: Number(pepesPetitChocApr.absoluteError.toFixed(1)),
          petitPinMay: Number(pepesPetitPinMay.forecast.toFixed(1)),
          petitPinMayAbs: Number(pepesPetitPinMay.absoluteError.toFixed(1)),
          quarterMayAbs: Number(pepesQuarterMay.absoluteError.toFixed(1)),
          halfMayAbs: Number(pepesHalfMay.absoluteError.toFixed(1)),
          quarterDecAbs: Number(pepesQuarterDec.absoluteError.toFixed(1)),
          halfDecAbs: Number(pepesHalfDec.absoluteError.toFixed(1)),
          miniSepAbs: Number(pepesMiniSep.absoluteError.toFixed(1)),
          miniJun: Number(pepesMiniJun.forecast.toFixed(1)),
          miniJunAbs: Number(pepesMiniJun.absoluteError.toFixed(1)),
          pinJun: Number(pepesPinJun.forecast.toFixed(1)),
          pinJunAbs: Number(pepesPinJun.absoluteError.toFixed(1)),
          mokaChMay: Number(pepesMokaChMay.forecast.toFixed(1)),
          mokaChMayAbs: Number(pepesMokaChMay.absoluteError.toFixed(1)),
          frutasChMay: Number(pepesFrutasChMay.forecast.toFixed(1)),
          frutasChMayAbs: Number(pepesFrutasChMay.absoluteError.toFixed(1)),
          mokaMedMayAbs: Number(pepesMokaMedMay.absoluteError.toFixed(1)),
          frutasGdeDec: Number(pepesFrutasGdeDec.forecast.toFixed(1)),
          frutasGdeDecAbs: Number(pepesFrutasGdeDec.absoluteError.toFixed(1)),
          payDecAbs: Number(pepesPayDec.absoluteError.toFixed(1)),
          gelGdeDecAbs: Number(pepesGelGdeDec.absoluteError.toFixed(1)),
          pinaDecAbs: Number(pepesPinaDec.absoluteError.toFixed(1)),
          mokaSepAbs: Number(pepesMokaSep.absoluteError.toFixed(1)),
        },
        leftoverSkus: {
          june: Number(leftoverJune.wape.toFixed(2)),
          july: Number(leftoverJuly.wape.toFixed(2)),
          august: Number(leftoverAugust.wape.toFixed(2)),
          weighted: Number(leftoverWeighted.toFixed(2)),
          gelatinaIndJuly: Number(leftGelatinaJuly.forecast.toFixed(1)),
          dollarJuly: Number(leftDollarJuly.forecast.toFixed(1)),
          dollarAugust: Number(leftDollarAugust.forecast.toFixed(1)),
          miniChocJuly: Number(leftMiniChocJuly.forecast.toFixed(1)),
          mokaMedJuly: Number(leftMokaMedJuly.forecast.toFixed(1)),
          miniChocAugustAbs: Number(leftMiniChocAugust.absoluteError.toFixed(1)),
          miniPinAugustAbs: Number(leftMiniPinAugust.absoluteError.toFixed(1)),
          mokaMedAugustAbs: Number(leftMokaMedAugust.absoluteError.toFixed(1)),
          duraznoAugustAbs: Number(leftDuraznoAugust.absoluteError.toFixed(1)),
        },
        promoActiva: {
          paletaJuly: Number(paletaPromoRow.pronosticoVenta.toFixed(1)),
          paletaPlantWed: paletaDayPromo.produccionSugeridaDia,
          paletaPlantWedBase: paletaDayBase.produccionSugeridaDia,
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
