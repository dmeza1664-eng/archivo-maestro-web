const path = require('path');
const Module = require('module');
const esbuild = require('esbuild');

async function loadAppFunctions() {
  const built = await esbuild.build({
    entryPoints: [path.join(__dirname, 'App.jsx')],
    bundle: true,
    platform: 'node',
    format: 'cjs',
    write: false,
    loader: { '.css': 'text' },
    define: { 'import.meta.env.VITE_API_URL': JSON.stringify('') },
    logLevel: 'silent',
  });
  const appModule = new Module('frozen-forecast-test');
  appModule.filename = path.join(__dirname, 'frozen-forecast-test.bundle.cjs');
  appModule.paths = module.paths;
  appModule._compile(built.outputFiles[0].text, appModule.filename);
  return appModule.exports;
}

function assert(condition, message) {
  if (!condition) throw new Error(message);
}

async function main() {
  const { buildOperationalForecastScenario } = await loadAppFunctions();
  const forecastRows = [
    {
      producto: 'PASTEL A GDE',
      orden: 1,
      pronosticoVenta: 100,
      baseConColchon: 150,
      metodoPronostico: 'Prueba',
      tendenciaAplicada: 1,
      mesesUsados: '2026-06, 2026-07',
    },
    {
      producto: 'GELATINA FRESA',
      orden: 2,
      pronosticoVenta: 50,
      baseConColchon: 90,
      metodoPronostico: 'Prueba',
      tendenciaAplicada: 1,
      mesesUsados: '2026-07',
    },
  ];
  const original = JSON.stringify(forecastRows);
  const scenario = buildOperationalForecastScenario(forecastRows);

  assert(scenario.length === forecastRows.length, 'El congelamiento debe incluir todos los productos');
  assert(scenario[0].pronosticoBase === 100, 'Debe usar el pronóstico de venta sin colchón');
  assert(scenario[0].margenOperativoPiezas === 12, 'El margen de 12% es incorrecto');
  assert(scenario[0].pronosticoOperativo === 112, 'El escenario operativo es incorrecto');
  assert(scenario[1].pronosticoOperativo === 56, 'Debe aplicar el margen a cada producto');
  assert(JSON.stringify(forecastRows) === original, 'No debe modificar las filas originales');

  console.log(JSON.stringify({ ok: true, products: scenario.length, base: 150, operational: 168 }));
}

main().catch((error) => {
  console.error(error.message);
  process.exit(1);
});
