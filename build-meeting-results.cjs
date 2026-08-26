const XLSX = require('xlsx');

const JULY_PATH = 'C:/Users/X13/Documents/COMPARATIVO_JULIO_REAL_VS_PRONOSTICADO.xlsx';
const AUGUST_PATH = 'C:/Users/X13/Documents/PRONOSTICO_AGOSTO_2026_BASE_Y_ESCENARIO_12.xlsx';
const OUTPUT_PATH = 'C:/Users/X13/Documents/RESULTADOS_JUNTA_PRONOSTICO_2026.xlsx';

function appendReportSheet(workbook, title, rows, name, widths) {
  const sheet = XLSX.utils.aoa_to_sheet([[title], []]);
  XLSX.utils.sheet_add_json(sheet, rows, { origin: 'A3' });
  const columnCount = Math.max(1, Object.keys(rows[0] || {}).length);
  sheet['!merges'] = [{ s: { r: 0, c: 0 }, e: { r: 0, c: columnCount - 1 } }];
  sheet['!cols'] = widths.map((wch) => ({ wch }));
  if (sheet['!ref']) {
    const range = XLSX.utils.decode_range(sheet['!ref']);
    sheet['!autofilter'] = {
      ref: XLSX.utils.encode_range({ s: { r: 2, c: 0 }, e: range.e }),
    };
  }
  XLSX.utils.book_append_sheet(workbook, sheet, name);
}

function readRows(workbook, sheetName) {
  const sheet = workbook.Sheets[sheetName];
  if (!sheet) throw new Error(`No existe la hoja ${sheetName}`);
  return XLSX.utils.sheet_to_json(sheet, { defval: '' });
}

function main() {
  const julyWorkbook = XLSX.readFile(JULY_PATH);
  const augustWorkbook = XLSX.readFile(AUGUST_PATH);
  const julyProducts = readRows(julyWorkbook, 'Por producto');
  const julyCategories = readRows(julyWorkbook, 'Por categoria');
  const augustProducts = readRows(augustWorkbook, 'Por producto');
  const augustCategories = readRows(augustWorkbook, 'Por categoría');

  const executiveSummary = [
    {
      Resultado: 'Reducción del error histórico',
      Valor: '14.7%',
      Comparación: 'WAPE de 10.27% a 8.76%',
      Interpretación: 'El modelo actual reduce 1.51 puntos porcentuales de error.',
    },
    {
      Resultado: 'Error promedio por producto',
      Valor: '19.33 piezas',
      Comparación: 'Antes: 22.63 piezas',
      Interpretación: 'Reducción promedio de 3.30 piezas de error por producto.',
    },
    {
      Resultado: 'Productos dentro de +/-15',
      Valor: '207 de 333',
      Comparación: 'Antes: 196 de 333',
      Interpretación: 'Once comparaciones adicionales quedaron dentro del rango.',
    },
    {
      Resultado: 'Cierre agregado de julio',
      Valor: '0.96% de desviación',
      Comparación: '26,061 pronosticadas vs 26,312 reales',
      Interpretación: 'Diferencia agregada de 251 piezas; no representa el error individual.',
    },
    {
      Resultado: 'Precisión por producto en julio',
      Valor: 'WAPE 12.42%',
      Comparación: 'MAE 29.44 piezas',
      Interpretación: 'Resultado reconstruido con el modelo actual por producto.',
    },
    {
      Resultado: 'Escenario retrospectivo +12% en julio',
      Valor: 'WAPE 8.48%',
      Comparación: '26,164.81 estimadas vs 26,312 reales',
      Interpretación: 'Escenario calculado después del cierre; no sustituye al pronóstico original.',
    },
    {
      Resultado: 'Pronóstico estadístico de agosto',
      Valor: '27,235.24 piezas',
      Comparación: '111 productos',
      Interpretación: 'Calculado con ventas históricas disponibles hasta julio.',
    },
    {
      Resultado: 'Escenario operativo de agosto',
      Valor: '30,503.47 piezas',
      Comparación: 'Margen separado de 12%',
      Interpretación: 'Referencia de capacidad; no es el pronóstico estadístico oficial.',
    },
    {
      Resultado: 'Control del pronóstico',
      Valor: 'Congelamiento por producto',
      Comparación: 'MySQL + Excel',
      Interpretación: 'Permite conservar la versión emitida y compararla al cierre.',
    },
  ];

  const historicalResults = [
    {
      Mes: 'Abril 2026',
      Productos: 111,
      'Venta real': 21825,
      'Pronóstico anterior': 20346.09,
      'WAPE anterior': '10.53%',
      'MAE anterior': 20.70,
      'Dentro +/-15 anterior': 66,
      'Pronóstico actual': 21480.01,
      'WAPE actual': '8.36%',
      'MAE actual': 16.43,
      'Dentro +/-15 actual': 71,
      'Mejora WAPE': '2.17 puntos',
    },
    {
      Mes: 'Mayo 2026',
      Productos: 111,
      'Venta real': 27127,
      'Pronóstico anterior': 27511.87,
      'WAPE anterior': '9.71%',
      'MAE anterior': 23.73,
      'Dentro +/-15 anterior': 62,
      'Pronóstico actual': 27231.41,
      'WAPE actual': '8.75%',
      'MAE actual': 21.39,
      'Dentro +/-15 actual': 69,
      'Mejora WAPE': '0.96 puntos',
    },
    {
      Mes: 'Junio 2026',
      Productos: 111,
      'Venta real': 24469,
      'Pronóstico anterior': 23785.73,
      'WAPE anterior': '10.65%',
      'MAE anterior': 23.47,
      'Dentro +/-15 anterior': 68,
      'Pronóstico actual': 23442.59,
      'WAPE actual': '9.14%',
      'MAE actual': 20.16,
      'Dentro +/-15 actual': 67,
      'Mejora WAPE': '1.51 puntos',
    },
    {
      Mes: 'Resultado combinado',
      Productos: 333,
      'Venta real': 73421,
      'Pronóstico anterior': 71643.69,
      'WAPE anterior': '10.27%',
      'MAE anterior': 22.63,
      'Dentro +/-15 anterior': 196,
      'Pronóstico actual': 72154.01,
      'WAPE actual': '8.76%',
      'MAE actual': 19.33,
      'Dentro +/-15 actual': 207,
      'Mejora WAPE': '1.51 puntos / 14.7%',
    },
  ];

  const julyResults = [
    {
      Referencia: 'Pronóstico agregado mostrado en página',
      Pronóstico: 26061,
      'Venta real comparable': 26312,
      'Diferencia real - pronóstico': 251,
      'Desviación agregada': '0.96%',
      Cumplimiento: '100.96%',
      WAPE: 'No disponible por producto',
      MAE: 'No disponible',
      'Dentro de +/-15': 'No disponible',
      Nota: 'No existe la exportación congelada por producto del total 26,061.',
    },
    {
      Referencia: 'Modelo actual reconstruido',
      Pronóstico: 23361.44,
      'Venta real comparable': 26312,
      'Diferencia real - pronóstico': 2950.56,
      'Desviación agregada': '12.63%',
      Cumplimiento: '112.63%',
      WAPE: '12.42%',
      MAE: 29.44,
      'Dentro de +/-15': '59 de 111',
      Nota: 'Detalle reconstruido usando historia anterior a julio.',
    },
    {
      Referencia: 'Escenario retrospectivo +12%',
      Pronóstico: 26164.81,
      'Venta real comparable': 26312,
      'Diferencia real - pronóstico': 147.19,
      'Desviación agregada': '0.56%',
      Cumplimiento: '100.56%',
      WAPE: '8.48%',
      MAE: 20.10,
      'Dentro de +/-15': '63 de 111',
      Nota: 'Calculado después del cierre; no debe presentarse como pronóstico original.',
    },
  ];

  const methodology = [
    {
      Concepto: 'WAPE',
      Explicación: 'Suma de los errores absolutos por producto dividida entre la venta real total comparable.',
    },
    {
      Concepto: 'MAE',
      Explicación: 'Promedio de piezas de error absoluto por producto.',
    },
    {
      Concepto: 'Desviación agregada',
      Explicación: 'Compara únicamente los totales; un total cercano puede ocultar errores entre productos.',
    },
    {
      Concepto: 'Validación histórica',
      Explicación: 'Cada mes se oculta y se pronostica usando solamente información anterior.',
    },
    {
      Concepto: 'Margen de 12%',
      Explicación: 'Escenario operativo separado. Empeoró el WAPE en abril, mayo y junio, por lo que no se incorporó al modelo oficial.',
    },
    {
      Concepto: 'Pronóstico de agosto',
      Explicación: 'Usa información histórica hasta julio y no utiliza ventas reales de agosto.',
    },
    {
      Concepto: 'Próximo control',
      Explicación: 'Congelar el pronóstico antes del cierre y comparar por producto contra la venta real.',
    },
    {
      Concepto: 'Redondeo',
      Explicación: 'La suma de filas mostradas con dos decimales puede diferir algunas centésimas del total calculado sin redondear.',
    },
  ];

  const workbook = XLSX.utils.book_new();
  appendReportSheet(workbook, 'Resultados ejecutivos del sistema de pronóstico', executiveSummary, 'Resumen ejecutivo', [38, 24, 45, 80]);
  appendReportSheet(workbook, 'Comparativo histórico del modelo', historicalResults, 'Mejora histórica', [23, 12, 15, 22, 17, 15, 24, 20, 15, 13, 22, 26]);
  appendReportSheet(workbook, 'Resultados del cierre de julio 2026', julyResults, 'Resultado julio', [42, 16, 23, 28, 22, 18, 25, 16, 22, 75]);
  appendReportSheet(workbook, 'Julio por categoría con escenario de 12%', julyCategories, 'Julio categorías', [25, 12, 16, 24, 28, 28, 20, 24, 24, 30]);
  appendReportSheet(workbook, 'Julio por producto con escenario de 12%', julyProducts, 'Julio productos', [40, 25, 16, 24, 20, 28, 27, 27, 27, 25, 34, 28, 34]);
  appendReportSheet(workbook, 'Pronóstico de agosto por categoría', augustCategories, 'Agosto categorías', [25, 12, 24, 22, 30]);
  appendReportSheet(workbook, 'Pronóstico de agosto por producto', augustProducts, 'Agosto productos', [40, 25, 24, 22, 30, 32, 38]);
  appendReportSheet(workbook, 'Notas para interpretar los resultados', methodology, 'Metodología', [30, 110]);
  XLSX.writeFile(workbook, OUTPUT_PATH);

  console.log(JSON.stringify({
    ok: true,
    outputPath: OUTPUT_PATH,
    sheets: workbook.SheetNames,
    julyProducts: julyProducts.length,
    augustProducts: augustProducts.length,
  }, null, 2));
}

try {
  main();
} catch (error) {
  console.error(error.message);
  process.exit(1);
}
