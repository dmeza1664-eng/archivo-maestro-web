const { query } = require('../db');

async function listVentas() {
  return query(
    `SELECT
       v.id,
       v.fecha,
       DATE_FORMAT(v.fecha, '%Y-%m-%d') AS fecha_iso,
       DATE_FORMAT(v.fecha, '%Y-%m') AS mes,
       p.codigo AS producto_codigo,
       p.nombre AS producto_nombre,
       v.cantidad,
       v.importe,
       v.canal,
       v.cliente
     FROM ventas_diarias v
     INNER JOIN productos p ON p.id = v.producto_id
     ORDER BY v.fecha, p.codigo`
  );
}

async function listStock() {
  return query(
    `SELECT
       s.id,
       s.mes,
       p.codigo AS producto_codigo,
       p.nombre AS producto_nombre,
       s.cantidad
     FROM stock_fijo s
     INNER JOIN productos p ON p.id = s.producto_id
     ORDER BY s.mes DESC, p.codigo`
  );
}

async function listProduccion() {
  return query(
    `SELECT
       pr.id,
       pr.fecha,
       DATE_FORMAT(pr.fecha, '%Y-%m-%d') AS fecha_iso,
       DATE_FORMAT(pr.fecha, '%Y-%m') AS mes,
       p.codigo AS producto_codigo,
       p.nombre AS producto_nombre,
       pr.cantidad,
       pr.turno
     FROM produccion_real pr
     INNER JOIN productos p ON p.id = pr.producto_id
     ORDER BY pr.fecha, p.codigo`
  );
}

module.exports = {
  listProduccion,
  listStock,
  listVentas,
};
