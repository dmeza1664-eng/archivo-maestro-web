const createOperationalRouter = require('./operationalImport');

module.exports = createOperationalRouter({
  type: 'produccion',
  label: 'producción',
  table: 'produccion_real',
  bodyKeys: ['produccion', 'produccionReal', 'rows', 'data'],
  quantityAliases: ['cantidad', 'produccion', 'produccion_real', 'unidades'],
  dimensions: [
    { name: 'turno', aliases: ['turno', 'shift'], maxLength: 80 },
  ],
});
