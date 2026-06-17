# Backend Archivo Maestro

API Node.js + Express + MySQL para guardar ventas, stock fijo, produccion real, homologaciones y pronosticos.

## Instalacion

```bash
cd backend
npm install
cp .env.example .env
```

Edita `.env` con tus variables reales:

```bash
DB_HOST=localhost
DB_USER=root
DB_PASSWORD=
DB_NAME=archivo_maestro
DB_PORT=3306
PORT=4000
CORS_ORIGIN=http://localhost:5173
```

No hay credenciales dentro del codigo.

## Crear tablas

Primero crea la base de datos en MySQL y luego ejecuta:

```bash
mysql -u TU_USUARIO -p archivo_maestro < schema.sql
```

## Ejecutar

```bash
npm run dev
```

La API queda en:

```text
http://localhost:4000
```

## Endpoints

```text
POST /api/ventas/bulk
GET  /api/ventas?mes=2026-04
POST /api/stock/bulk
GET  /api/stock?mes=2026-04
POST /api/produccion-real/bulk
GET  /api/produccion-real?mes=2026-04
POST /api/pronostico/calcular
GET  /api/pronostico?mes=2026-04
```

## Formato JSON desde el frontend

Los endpoints `bulk` aceptan un arreglo directo o un objeto con `ventas`, `stock`, `produccion`, `rows` o `data`.

Ejemplo desde React/Vite:

```js
const API_URL = import.meta.env.VITE_API_URL || 'http://localhost:4000';

export async function guardarVentas(ventasProcesadas) {
  const response = await fetch(`${API_URL}/api/ventas/bulk`, {
    method: 'POST',
    headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify({ ventas: ventasProcesadas }),
  });

  if (!response.ok) {
    throw new Error(await response.text());
  }

  return response.json();
}
```

En el frontend puedes usar:

```bash
VITE_API_URL=http://localhost:4000
```

Ejemplo ventas:

```json
{
  "ventas": [
    {
      "fecha": "2026-04-01",
      "producto_codigo": "SKU-001",
      "producto_nombre": "Producto 1",
      "cantidad": 25,
      "importe": 1500,
      "canal": "Retail",
      "cliente": "Cliente A",
      "codigo_origen": "EXCEL-001",
      "nombre_origen": "Producto Excel 1"
    }
  ]
}
```

Ejemplo stock:

```json
{
  "stock": [
    {
      "mes": "2026-04",
      "producto_codigo": "SKU-001",
      "producto_nombre": "Producto 1",
      "cantidad": 100
    }
  ]
}
```

Ejemplo produccion real:

```json
{
  "produccion": [
    {
      "fecha": "2026-04-01",
      "producto_codigo": "SKU-001",
      "producto_nombre": "Producto 1",
      "cantidad": 40,
      "turno": "Matutino"
    }
  ]
}
```

Ejemplo calcular pronostico:

```json
{
  "mes": "2026-04",
  "mesesHistoricos": 3,
  "metodo": "promedio_ventas"
}
```

El calculo toma el promedio diario de ventas de los meses historicos previos y genera un pronostico diario para cada producto activo.
