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
  const appModule = new Module('weekly-progress-test');
  appModule.filename = path.join(__dirname, 'weekly-progress-test.bundle.cjs');
  appModule.paths = module.paths;
  appModule._compile(built.outputFiles[0].text, appModule.filename);
  return appModule.exports;
}

function assert(condition, message) {
  if (!condition) throw new Error(message);
}

async function main() {
  const { buildWeeklyProgress } = await loadAppFunctions();
  const rows = [];
  for (let day = 1; day <= 31; day += 1) {
    const fecha = `2026-07-${String(day).padStart(2, '0')}`;
    rows.push({
      fecha,
      producto: 'PRODUCTO A',
      pronosticoVentaDia: 10,
      hasVentaReal: day <= 8,
      ventaRealDia: day <= 8 ? 12 : null,
    });
    rows.push({
      fecha,
      producto: 'PRODUCTO B',
      pronosticoVentaDia: 20,
      hasVentaReal: day <= 8,
      ventaRealDia: day <= 8 ? 18 : null,
    });
  }

  const progress = buildWeeklyProgress(rows, '2026-07');
  assert(progress.month.cutoffDate === '2026-07-08', 'El corte mensual debe usar la última venta real');
  assert(progress.month.pronosticoPeriodo === 930, 'El pronóstico mensual debe conservarse completo');
  assert(progress.month.pronosticoCorte === 240, 'El pronóstico al corte debe incluir ocho días');
  assert(progress.month.ventaReal === 240, 'La venta real acumulada es incorrecta');
  assert(progress.month.proyeccionPeriodo === 930, 'La proyección debe sumar real y pronóstico pendiente');
  assert(progress.month.status.label === 'En objetivo', 'El avance mensual debe quedar en objetivo');

  const selectedWeek = progress.weeks.find((week) => week.key === progress.suggestedWeekKey);
  assert(selectedWeek?.comparedDays === 3, 'La semana de corte debe comparar tres días');
  assert(selectedWeek?.pronosticoCorte === 90, 'El pronóstico semanal al corte es incorrecto');
  assert(selectedWeek?.ventaReal === 90, 'La venta semanal real es incorrecta');
  assert(selectedWeek?.products[0].producto === 'PRODUCTO A', 'Debe ordenar productos por desviación absoluta');

  const empty = buildWeeklyProgress(rows.map((row) => ({ ...row, hasVentaReal: false, ventaRealDia: null })), '2026-07');
  assert(!empty.hasRealData, 'Un mes sin venta real no debe marcarse como cargado');
  assert(empty.month.proyeccionPeriodo === 930, 'Sin ventas, la proyección debe conservar el pronóstico base');

  console.log(JSON.stringify({ ok: true, weeks: progress.weeks.length, cutoff: progress.month.cutoffDate }));
}

main().catch((error) => {
  console.error(error.message);
  process.exit(1);
});
