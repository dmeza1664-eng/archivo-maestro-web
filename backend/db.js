const mysql = require('mysql2/promise');
require('dotenv').config();

const tidbConfigured = Boolean(process.env.TIDB_HOST);
const connectionUrl = tidbConfigured ? '' : process.env.DATABASE_URL || process.env.MYSQL_URL;
const required = ['DB_HOST', 'DB_USER', 'DB_NAME'];
const missing = tidbConfigured || connectionUrl ? [] : required.filter((key) => !process.env[key]);

if (missing.length > 0) {
  throw new Error(`Faltan variables de entorno requeridas: ${missing.join(', ')}`);
}

const pool = mysql.createPool({
  ...(tidbConfigured
    ? {
        host: process.env.TIDB_HOST,
        user: process.env.TIDB_USER,
        password: process.env.TIDB_PASSWORD || '',
        database: process.env.TIDB_DATABASE,
        port: Number(process.env.TIDB_PORT || 4000),
      }
    : connectionUrl
    ? { uri: connectionUrl }
    : {
        host: process.env.DB_HOST,
        user: process.env.DB_USER,
        password: process.env.DB_PASSWORD || '',
        database: process.env.DB_NAME,
        port: Number(process.env.DB_PORT || 3306),
      }),
  ...((tidbConfigured || process.env.DB_SSL === 'true')
    ? {
        ssl: {
          minVersion: 'TLSv1.2',
          rejectUnauthorized: process.env.DB_SSL_REJECT_UNAUTHORIZED !== 'false',
        },
      }
    : {}),
  waitForConnections: true,
  connectionLimit: Number(process.env.DB_CONNECTION_LIMIT || (process.env.VERCEL ? 2 : 5)),
  queueLimit: 0,
  decimalNumbers: true,
  enableKeepAlive: true,
  keepAliveInitialDelay: 0,
});

async function query(sql, params = []) {
  const [rows] = await pool.execute(sql, params);
  return rows;
}

async function transaction(callback) {
  const connection = await pool.getConnection();

  try {
    await connection.beginTransaction();
    const result = await callback(connection);
    await connection.commit();
    return result;
  } catch (error) {
    await connection.rollback();
    throw error;
  } finally {
    connection.release();
  }
}

module.exports = {
  pool,
  query,
  transaction,
};
