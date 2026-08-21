const fs = require('fs');
const path = require('path');
const mysql = require('mysql2/promise');
require('dotenv').config();

async function main() {
  const tidbConfigured = Boolean(process.env.TIDB_HOST);
  const connectionUrl = tidbConfigured ? '' : process.env.DATABASE_PUBLIC_URL || process.env.MYSQL_PUBLIC_URL || process.env.DATABASE_URL || process.env.MYSQL_URL;
  const config = tidbConfigured
    ? {
        host: process.env.TIDB_HOST,
        user: process.env.TIDB_USER,
        password: process.env.TIDB_PASSWORD || '',
        database: process.env.TIDB_DATABASE,
        port: Number(process.env.TIDB_PORT || 4000),
        ssl: { minVersion: 'TLSv1.2', rejectUnauthorized: true },
        multipleStatements: true,
      }
    : connectionUrl
    ? { uri: connectionUrl, multipleStatements: true }
    : {
        host: process.env.DB_HOST,
        user: process.env.DB_USER,
        password: process.env.DB_PASSWORD || '',
        database: process.env.DB_NAME,
        port: Number(process.env.DB_PORT || 3306),
        multipleStatements: true,
      };
  const connection = await mysql.createConnection(config);
  const schema = fs.readFileSync(path.join(__dirname, 'schema.sql'), 'utf8');
  await connection.query(schema);
  await connection.end();
  console.log('Esquema aplicado correctamente');
}

main().catch((error) => {
  console.error(error.code || error.message);
  process.exit(1);
});
