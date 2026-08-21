const createOperationalRouter = require('./operationalImport');

module.exports = createOperationalRouter({
  type: 'bajas',
  label: 'bajas',
  table: 'bajas_diarias',
  bodyKeys: ['bajas', 'mermas', 'rows', 'data'],
  quantityAliases: ['cantidad', 'baja', 'bajas', 'merma', 'unidades'],
  dimensions: [
    { name: 'sucursal', aliases: ['sucursal', 'canal', 'zona', 'tienda'], maxLength: 120 },
    { name: 'motivo', aliases: ['motivo', 'causa', 'tipo_baja'], maxLength: 160 },
  ],
});
