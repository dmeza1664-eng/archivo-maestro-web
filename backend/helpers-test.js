const { limitClause, paginationFromQuery } = require('./routes/helpers');

function assert(condition, message) {
  if (!condition) throw new Error(message);
}

function rejects(fn, message) {
  try {
    fn();
  } catch (error) {
    return error;
  }
  throw new Error(message);
}

function main() {
  // TiDB responde "Incorrect arguments to LIMIT" si el limite viaja como
  // marcador en una sentencia preparada, asi que va interpolado y tiene que
  // quedar blindado contra cualquier valor que no sea un entero acotado.
  assert(limitClause(4001) === 'LIMIT 4001', 'un entero valido debe producir la clausula');
  assert(limitClause(1) === 'LIMIT 1', 'el limite minimo es 1');

  rejects(() => limitClause(0), 'un limite de cero debe rechazarse');
  rejects(() => limitClause(-5), 'un limite negativo debe rechazarse');
  rejects(() => limitClause(1.5), 'un limite fraccionario debe rechazarse');
  rejects(() => limitClause(99999), 'un limite fuera de rango debe rechazarse');
  rejects(() => limitClause('10'), 'una cadena numerica no es un entero valido');
  rejects(() => limitClause('1; DROP TABLE ventas_diarias'), 'no debe aceptarse SQL en el limite');
  rejects(() => limitClause(undefined), 'sin limite debe rechazarse');
  rejects(() => limitClause(NaN), 'NaN debe rechazarse');

  // El limite que produce paginationFromQuery siempre debe pasar el blindaje,
  // incluido el +1 que usan los endpoints para detectar si hay mas paginas.
  const defaults = paginationFromQuery({});
  assert(limitClause(defaults.limit + 1) === `LIMIT ${defaults.limit + 1}`, 'el limite por defecto debe ser utilizable');

  const clamped = paginationFromQuery({ limit: '999999' });
  assert(clamped.limit === 10000, 'el limite pedido debe recortarse al maximo');
  assert(limitClause(clamped.limit + 1) === 'LIMIT 10001', 'el maximo recortado mas uno sigue siendo valido');

  console.log('helpers-test ok');
}

main();
