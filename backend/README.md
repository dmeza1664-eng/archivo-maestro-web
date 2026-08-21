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
APP_SETUP_KEY=una-clave-temporal-de-instalacion
```

En Railway se puede usar `DATABASE_URL` con referencia a `MySQL.MYSQL_URL` en lugar de las variables `DB_*`.

No hay credenciales dentro del codigo.

## Despliegue en Railway

El repositorio contiene el frontend en la raiz y la API en `backend/`. El servicio de Railway debe apuntar a esa subcarpeta, de lo contrario Nixpacks construye el proyecto de Vite y el dominio queda sin proceso que responda.

Configuracion del servicio:

```text
Root Directory: backend
Variables: DATABASE_URL (o DB_*), PORT, CORS_ORIGIN, APP_SETUP_KEY
```

`CORS_ORIGIN` debe incluir el dominio publicado del frontend, por ejemplo `https://archivo-maestro-web.vercel.app`.

El archivo `railway.json` fija el comando de arranque y el health check en `/api/health`, de modo que un despliegue sin base de datos accesible se marca como fallido en lugar de quedar servido a medias.

Verificacion despues de cada despliegue:

```bash
curl https://TU-API.up.railway.app/api/health
```

La respuesta esperada es `{"ok":true,"service":"archivo-maestro-backend","database":"connected"}`. Si el dominio responde 404 con la cabecera `x-railway-fallback: true`, el dominio existe pero no hay servicio activo detras: revisar que el servicio no este eliminado o detenido.

En el frontend de Vercel debe existir `VITE_API_URL` con la URL publica de la API. Vite congela ese valor durante la compilacion, por lo que un cambio de URL exige un nuevo despliegue.

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
POST /api/ventas/importar
GET  /api/ventas?mes=2026-04
GET  /api/ventas/importaciones
GET  /api/ventas/sync
POST /api/stock/bulk
GET  /api/stock?mes=2026-04
POST /api/produccion-real/bulk
POST /api/produccion-real/importar
GET  /api/produccion-real?mes=2026-04
GET  /api/produccion-real/importaciones
GET  /api/produccion-real/sync
POST /api/bajas/importar
GET  /api/bajas?mes=2026-04
GET  /api/bajas/importaciones
GET  /api/bajas/sync
POST /api/pronostico/calcular
GET  /api/pronostico?mes=2026-04
GET  /api/auth/status
POST /api/auth/setup
POST /api/auth/login
GET  /api/auth/me
POST /api/auth/logout
GET  /api/auth/users
POST /api/auth/users
GET  /api/snapshots
GET  /api/snapshots/:type
GET  /api/snapshots/:type/history
POST /api/snapshots/:type
GET  /api/audit
```

Los respaldos son inmutables y versionados. Cada guardado registra usuario, fecha, tamaño y acción en la bitácora.

La importación de ventas consolida filas repetidas por fecha, producto, sucursal y cliente. Si el registro ya existe, lo actualiza en lugar de duplicarlo. Los totales mensuales se rechazan en esta ruta porque no representan ventas diarias.

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
