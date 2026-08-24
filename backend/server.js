const express = require('express');
const cors = require('cors');
require('dotenv').config();

const ventasRoutes = require('./routes/ventas');
const stockRoutes = require('./routes/stock');
const produccionRoutes = require('./routes/produccion');
const bajasRoutes = require('./routes/bajas');
const pronosticoRoutes = require('./routes/pronostico');
const authRoutes = require('./routes/auth');
const snapshotsRoutes = require('./routes/snapshots');
const auditRoutes = require('./routes/audit');
const { query } = require('./db');

const app = express();
const port = Number(process.env.PORT || 4000);

const allowedOrigins = String(process.env.CORS_ORIGIN || '')
  .split(',')
  .map((origin) => origin.trim())
  .filter(Boolean);
app.use(cors({
  origin(origin, callback) {
    if (!origin) return callback(null, true);
    if (!allowedOrigins.length || allowedOrigins.includes(origin)) return callback(null, origin);
    return callback(new Error('Origen no permitido'));
  },
  credentials: true,
}));
app.use(express.json({ limit: '4mb' }));

app.get('/api/health', async (_req, res, next) => {
  try {
    await query('SELECT 1 AS ok');
    res.json({ ok: true, service: 'archivo-maestro-backend', database: 'connected' });
  } catch (error) {
    next(error);
  }
});

app.use('/api/auth', authRoutes);
app.use('/api/snapshots', snapshotsRoutes);
app.use('/api/audit', auditRoutes);
app.use('/api/ventas', ventasRoutes);
app.use('/api/stock', stockRoutes);
app.use('/api/produccion-real', produccionRoutes);
app.use('/api/bajas', bajasRoutes);
app.use('/api/pronostico', pronosticoRoutes);

app.use((req, res) => {
  res.status(404).json({ error: `Ruta no encontrada: ${req.method} ${req.originalUrl}` });
});

app.use((error, _req, res, _next) => {
  console.error(error);
  res.status(error.status || 500).json({
    error: error.message || 'Error interno del servidor',
  });
});

if (require.main === module) {
  app.listen(port, () => {
    console.log(`API Archivo Maestro escuchando en http://localhost:${port}`);
  });
}

module.exports = app;
