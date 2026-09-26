/**
 * Serie oficial de demanda = venta de SUCURSALES al público (decisión del 26-sep-2026).
 * DATOS SINTÉTICOS (fixture sintético del repo + un canal "Planta León" inventado;
 * no son ventas reales de Pastelería Pepes). Protege:
 *  - el surtido de Planta León (registrado como venta) no entra al pronóstico ni a sus
 *    reglas: con o sin filas de Planta, pronóstico y producción sugerida son iguales;
 *  - la suma de canales sigue disponible solo como referencia (demandSeries: "total");
 *  - filas sin sucursal (totales mensuales consolidados) pasan tal cual;
 *  - Suc. Amado Nervo antes de ago-2024 (entonces también surtía) queda fuera, como en la serie oficial;
 *  - el reparto semanal/sucursal ignora a Planta (no es sucursal, no recibe reparto).
 */
const fs = require("fs");
const path = require("path");
const Module = require("module");
const esbuild = require("esbuild");
const { disaggregateMonthlyForecast, isInternalSupplyChannel: weeklyIsInternal } = require("./lib/weekly-branch-disaggregation.cjs");

const ROOT = path.join(__dirname, "..");
const FIXTURE = path.join(__dirname, "fixtures", "sintetico-pronostico-2024-2025.json");

async function loadApp() {
  const built = await esbuild.build({
    entryPoints: [path.join(ROOT, "App.jsx")], bundle: true, platform: "node", format: "cjs", write: false,
    loader: { ".css": "text" },
    define: { "import.meta.env.VITE_API_URL": JSON.stringify(""), "import.meta.env.DEV": "false", "import.meta.env.PROD": "true" },
    logLevel: "silent",
  });
  const m = new Module("demand-series-sucursal-test");
  m.filename = path.join(__dirname, "demand-series-sucursal-test.bundle.cjs");
  m.paths = module.paths;
  m._compile(built.outputFiles[0].text, m.filename);
  return m.exports;
}

function assert(condition, message) {
  if (!condition) throw new Error(`[serie oficial sucursales] ${message}`);
}

function close(monthKey, producto, cantidad, extra = {}) {
  const [y, mo] = monthKey.split("-").map(Number);
  return { fecha: new Date(y, mo - 1, 15, 12), producto, cantidad, monthlyTotal: true, monthDays: new Date(y, mo, 0).getDate(), ...extra };
}

function run(app, stockRows, ventas, month, extra = {}) {
  return app.calculateForecast({
    stockRows, historicalVentas: app.filterVentasBeforeMonth(ventas, month), bajas: [], existencias: [], realProduction: [],
    selectedMonth: month, dailyBufferPct: 10, ...extra,
  });
}

(async () => {
  const app = await loadApp();
  assert(app.DEMAND_SERIES_DEFAULT === "sucursal", "la serie oficial por defecto es la de sucursales");

  // Reconocimiento del canal interno por la palabra PLANTA (no por SKU).
  for (const name of ["Planta León", "Planta León · Piso de venta", "PLANTA LEON", "planta"]) {
    assert(app.isInternalSupplyChannel(name) && weeklyIsInternal(name), `${name} es surtido de planta`);
  }
  for (const name of ["Suc. León", "Suc. Amado Nervo", "San Cayetano", "Villa Hidalgo", "", null, "IMPLANTACION"]) {
    assert(!app.isInternalSupplyChannel(name) && !weeklyIsInternal(name), `${name} no es surtido de planta`);
  }

  // Suc. Amado Nervo antes de ago-2024 surtía (filas mezcladas): fuera; desde ago-2024 es sucursal normal.
  const amado = [close("2024-05", "X", 100, { sucursal: "Suc. Amado Nervo" }), close("2024-08", "X", 100, { sucursal: "Suc. Amado Nervo" })];
  const amadoKept = app.filterDemandSales(amado);
  assert(amadoKept.length === 1 && amadoKept[0].fecha.getMonth() === 7, "Amado Nervo sale antes de ago-2024 y entra desde ago-2024");
  const amadoDaily = [{ fecha: "2024-07-31", sucursal: "Suc. Amado Nervo", producto: "X", cantidad: 500 }];

  const fixture = JSON.parse(fs.readFileSync(FIXTURE, "utf8"));
  assert(/SINT[ÉE]TICO/i.test(fixture._SINTETICO || ""), "el fixture debe estar marcado como sintético");
  const stockRows = fixture.stock;
  const sucursal = [];
  const planta = [];
  for (const [product, months] of Object.entries(fixture.cierresMensuales)) {
    for (const [monthKey, qty] of Object.entries(months)) {
      sucursal.push(close(monthKey, product, qty, { sucursal: "Suc. Centro", canal: "Suc. Centro" }));
      // Surtido inventado: más que la venta y con un diciembre inflado (×3).
      const factor = monthKey.endsWith("-12") ? 3 : 1.1;
      planta.push(close(monthKey, product, Math.round(qty * factor), { sucursal: "Planta León", canal: "Planta León · Piso de venta" }));
    }
  }
  const mixed = [...sucursal, ...planta];

  // Filas sin sucursal pasan tal cual; filas de Planta salen.
  const noBranch = sucursal.map(({ sucursal: _s, canal: _c, ...row }) => row);
  assert(app.filterDemandSales(noBranch).length === noBranch.length, "filas sin sucursal (totales consolidados) pasan tal cual");
  assert(app.filterDemandSales(mixed).length === sucursal.length, "las filas de Planta León salen de la demanda");
  assert(app.filterDemandSales(mixed, "total").length === mixed.length, "demandSeries total conserva la suma de canales (referencia)");

  let changedWithTotal = 0;
  for (const month of ["2025-05", "2025-06", "2025-10", "2025-12"]) {
    const suc = run(app, stockRows, sucursal, month);
    const mix = run(app, stockRows, mixed, month);
    const plain = run(app, stockRows, noBranch, month);
    const tot = run(app, stockRows, mixed, month, { demandSeries: "total" });
    assert(suc.length === mix.length && suc.length > 0, `${month}: mismas filas`);
    for (let i = 0; i < suc.length; i += 1) {
      const a = suc[i];
      const b = mix[i];
      const c = plain[i];
      assert(a.producto === b.producto, `${month}: mismo orden de productos`);
      assert(a.pronosticoVenta === b.pronosticoVenta && a.pronosticoVenta === c.pronosticoVenta,
        `${month} ${a.producto}: Planta no cambia el pronóstico (${a.pronosticoVenta} vs ${b.pronosticoVenta})`);
      assert(a.produccionSugerida === b.produccionSugerida,
        `${month} ${a.producto}: la producción sugerida sale de la venta de sucursales + colchón`);
      if (tot[i].pronosticoVenta !== a.pronosticoVenta) changedWithTotal += 1;
    }
  }
  assert(changedWithTotal > 0, "con demandSeries total el pronóstico sí cambia (la referencia sigue disponible)");

  // Reparto semanal/sucursal: Planta no es sucursal ni cambia perfiles/factores.
  const daily = [];
  const DAY = 86400000;
  for (let t = Date.UTC(2024, 8, 1); t < Date.UTC(2025, 9, 1); t += DAY) {
    const fecha = new Date(t).toISOString().slice(0, 10);
    const wd = new Date(t).getUTCDay();
    daily.push({ fecha, sucursal: "A", producto: "PASTEL X", cantidad: wd === 6 ? 30 : 10 });
    daily.push({ fecha, sucursal: "B", producto: "PASTEL X", cantidad: 5 });
  }
  const withPlanta = daily.concat(amadoDaily).concat(daily.filter((_, i) => i % 2 === 0).map((r) => ({ ...r, sucursal: "Planta León", cantidad: (new Date(r.fecha).getUTCDay() === 2 ? 400 : 50) })));
  const base = disaggregateMonthlyForecast({ month: "2025-10", monthlyForecast: { "PASTEL X": 1000 }, dailySales: daily });
  const leak = disaggregateMonthlyForecast({ month: "2025-10", monthlyForecast: { "PASTEL X": 1000 }, dailySales: withPlanta });
  assert(base.sucursales.join(",") === "A,B" && leak.sucursales.join(",") === "A,B", `Planta no recibe reparto (${leak.sucursales})`);
  assert(JSON.stringify(base.porSucursalSkuSemana) === JSON.stringify(leak.porSucursalSkuSemana), "Planta no cambia el reparto semanal/sucursal");

  console.log("demand-series-sucursal-test ok");
})().catch((error) => {
  console.error(error);
  process.exit(1);
});
