const express = require('express');
const cors = require('cors');
require('dotenv').config();

const ventasRoutes = require('./routes/ventas');
const stockRoutes = require('./routes/stock');
const produccionRoutes = require('./routes/produccion');
const pronosticoRoutes = require('./routes/pronostico');

const app = express();
const port = Number(process.env.PORT || 4000);

app.use(cors({
  origin: process.env.CORS_ORIGIN || true,
}));
app.use(express.json({ limit: '50mb' }));

app.get('/api/health', (_req, res) => {
  res.json({ ok: true, service: 'archivo-maestro-backend' });
});

app.use('/api/ventas', ventasRoutes);
app.use('/api/stock', stockRoutes);
app.use('/api/produccion-real', produccionRoutes);
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

app.listen(port, () => {
  console.log(`API Archivo Maestro escuchando en http://localhost:${port}`);
});
