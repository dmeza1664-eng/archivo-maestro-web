# Archivo Maestro Web — Pronóstico Pastelería Pepes

App web (React + Vite) y scripts Node para pronosticar la demanda mensual por producto de Pastelería Pepes, repartirla por semana y sucursal, y medir el error contra la venta real.

- **Serie oficial de demanda**: venta de las sucursales al público, por producto y mes. El surtido de *Planta León · Piso de venta* a sucursales **no** entra al pronóstico ni al error, y Suc. Amado Nervo solo cuenta desde ago-2024 (PR #31).
- **Modelo vigente**: `d0e61f7` (PR #32), función `calculateForecast` en `App.jsx`. La bitácora de cambios del modelo está en `CONTROL_MODELO_PRONOSTICO.md`.
- **Datos**: extractos de la base SQL Server `pepes_devBI` (tablas `dbo`: `Ventas`, `VentaDet`, `Produccion`, `ProduccionDet` y los catálogos `Productos`, `Bodegas`, `Sucursales`) y el catálogo `stock_ideal.xlsx` (113 SKUs; 111 se pronostican, las 2 gelatinas de promoción quedan fuera).

## Métricas vigentes (modelo `d0e61f7`, serie de sucursales)

Error WAPE por producto-mes, walk-forward (cada mes se pronostica solo con meses anteriores):

| Año | Producto-mes | Sin enero | Total mensual |
|---|---|---|---|
| 2025 (ene–dic) | **11.19%** | 11.24% | 2.45% |
| 2024 (ene–dic) | **13.12%** | 12.66% | 4.58% |
| 2026 (ene–ago, fuera de muestra) | **11.75%** | 11.45% | 3.53% |

- 2025 por mes: ene 10.46, feb 9.99, mar 15.33, abr 13.60, may 12.96, jun 8.22, jul 15.48, ago 9.61, sep 7.17, oct 6.60, nov 6.27, dic 17.20.
- 2026 por mes: ene 13.81, feb 18.20, mar 11.54, abr 13.22, may 9.51, jun 10.02, jul 10.02, ago 8.06.
- Semanal (SKU × semana): 2025 16.98, 2026 17.73, 2024 (mar–dic sin Amado Nervo) 21.54. Detalle en `PRONOSTICO_SEMANAL_SUCURSAL.md`.
- Plan de producción 2025 (pronóstico + 10% de colchón): 333,085 piezas, contra 305,199 vendidas y 320,556 producidas.

Estas cifras están en `datos/referencia/metricas-modelo.json`; el pronóstico automático las compara en cada corrida.

## Pronóstico automático (un solo comando)

```
npm ci
npm run pronostico:auto
```

Qué hace (`scripts/pronostico-automatico.cjs`):

1. **Lee los datos** de la fuente configurada:
   - **modo archivos** (`PEPES_FUENTE=archivos`, por defecto): los extractos CSV actuales (`ventas-AAAA-mensual-producto-bodega.csv`, `part-ventas-bodega-AAAA-MM.csv`, `produccion-*.csv`, `extract-meta*.json`) en las carpetas de `PEPES_DATOS_DIRS`;
   - **modo base** (`PEPES_FUENTE=base`): la copia espejo de `pepes_devBI` en SQL Server, con consultas **solo de lectura** (`SELECT`) a las mismas tablas `dbo`.
2. **Arma la serie oficial** (venta de sucursales por producto y mes) con la misma regla de la app (`filterDemandSales`) y el mapeo `datos/catalogo/mapeo-productos.csv`.
3. **Pronostica** el mes siguiente al último mes completo con `calculateForecast` de `App.jsx`, sin tocar el modelo.
4. **Cruza producción contra pronóstico**: recalcula walk-forward el pronóstico de cada mes cerrado y lo compara con la venta y la producción real, por producto y mes; además total por mes y por producto-año, y el plan con colchón (`PEPES_COLCHON_PCT`, 10% por defecto).
5. **Deja salidas y log** en `salidas/pronostico-automatico/corrida-AAAAMMDD-HHMMSS/` (carpeta ignorada por git):

| Archivo | Contenido |
|---|---|
| `resumen.md` | Resumen legible: fecha de corte, pronóstico del mes, métricas por año contra la referencia, cruce mensual, avisos |
| `pronostico-AAAA-MM.csv` | Pronóstico y plan por producto del mes siguiente (con venta y producción al corte si el mes ya empezó) |
| `cruce-producto-mes.csv` | Por producto y mes: pronóstico, plan, venta, producción, sobrante/faltante contra venta y error |
| `cruce-mensual.csv` | Totales por mes y WAPE del mes |
| `cruce-producto-anual.csv` | Por producto y año |
| `metricas.json` | WAPE por año, sin enero y total mensual; comparación con `datos/referencia/metricas-modelo.json` |
| `serie-sucursal.json` | Serie oficial usada |
| `pk-sin-mapeo.csv` | Productos de la base sin mapeo al catálogo (para revisar) |
| `corrida.log` | Qué corrió, con qué fuente, fecha de corte, commit del código y tiempos |

En `salidas/pronostico-automatico/` quedan también `historial-corridas.log` (una línea por corrida) y `ultima-corrida.txt`.

Si las métricas de 2024, 2025 o 2026 ene–ago no coinciden con la referencia, el resumen y el log lo marcan (sirve para detectar datos cambiados o un modelo distinto).

### Modo archivos (hoy)

```
PEPES_DATOS_DIRS=/ruta/extract-pepes-2023:/ruta/extract-pepes-2024:/ruta/pepes-devBI-extract-2025:/ruta/validacion-2026/extract \
PEPES_STOCK_IDEAL=/ruta/stock_ideal.xlsx \
npm run pronostico:auto
```

Con los extractos al 20-sep-2026 (corrida de prueba, ~22 s): corte en el último mes completo (ago-2026), pronóstico de sep-2026 = 25,224 piezas (plan 27,749) y métricas idénticas a la tabla de arriba. La historia de 2023 solo se usa para poder evaluar 2024; la serie oficial empieza en `PEPES_DESDE` (2024-01).

### Modo base (cuando TI entregue la copia espejo)

1. Pedir a TI: servidor, puerto, nombre de la base (`pepes_devBI`) y un usuario **de solo lectura**.
2. Copiar `config/pronostico-automatico.env.example` como `.env.pronostico` en la raíz (git lo ignora) o definir las variables en el programador de tareas:

```
PEPES_FUENTE=base
PEPES_DB_SERVER=servidor-o-ip
PEPES_DB_NAME=pepes_devBI
PEPES_DB_USER=usuario_lectura
PEPES_DB_PASSWORD=********
# opcionales: PEPES_DB_PORT=1433, PEPES_DB_ENCRYPT=true, PEPES_DB_TRUST_SERVER_CERT=false,
#             PEPES_DB_TIMEOUT_MS=300000, PEPES_ENV_FILE=/ruta/otro.env
PEPES_STOCK_IDEAL=/ruta/stock_ideal.xlsx
```

3. Correr `npm run pronostico:auto`. La primera vez conviene comparar el `resumen.md` con una corrida en modo archivos: las métricas deben salir iguales a la referencia.

Detalles del modo base:

- **Nunca** poner credenciales en el repo. La contraseña no se escribe en el log ni en los errores.
- Solo se permiten consultas `SELECT` de una sentencia sobre `dbo.Ventas`, `dbo.VentaDet`, `dbo.Productos`, `dbo.Bodegas`, `dbo.Sucursales`, `dbo.Produccion` y `dbo.ProduccionDet`; cualquier otra cosa se rechaza antes de mandarse. La conexión se abre con intención de solo lectura.
- Las ventas se leen mes por mes desde `PEPES_DESDE` (2024-01, igual que la copia espejo). La producción se lee **completa** desde esa fecha, sin tope, porque las órdenes cambian de estado.
- Cada corrida guarda en `extracto/` los CSV leídos de la base, con el mismo formato del modo archivos, para poder repetirla sin conexión.
- Avisos en el log: órdenes de producción abiertas en meses cerrados y producción que termina antes que la venta.
- Driver: `mssql` 11 (compatible con Node 20, el de CI).

**Recopia de la espejo**: `Produccion` y `ProduccionDet` deben recopiarse **completas** en cada actualización (o al menos todas las órdenes abiertas y las modificadas desde la última copia), porque una orden cambia de estado y de cantidades después de creada. `Ventas`/`VentaDet` pueden recopiarse de forma incremental si incluyen cancelaciones y cambios.

**Frecuencia sugerida**: una corrida diaria de madrugada, después de la recopia nocturna de la espejo (por ejemplo 06:00). El pronóstico del mes siguiente queda listo desde el primer día hábil del mes; las corridas a mitad de mes actualizan la venta y la producción al corte en `pronostico-AAAA-MM.csv`. Si no hay recopia diaria, basta una corrida semanal y una el día 1 de cada mes.

Otras variables: `PEPES_SALIDA_DIR` (carpeta de salidas), `PEPES_FECHA_CORTE=AAAA-MM-DD` (repetir un corte pasado), `PEPES_MESES_OBJETIVO=AAAA-MM[,AAAA-MM]` (forzar el mes a pronosticar), `PEPES_MAPEO_PRODUCTOS`, `PEPES_MAPEO_CODIGOS`, `PEPES_REFERENCIA_METRICAS`, `PEPES_COLCHON_PCT`.

## Estructura

- `App.jsx`: app y modelo de pronóstico (`calculateForecast`, `filterDemandSales`, `analyzeForecastProductErrors`).
- `scripts/pronostico-automatico.cjs` y `scripts/lib/pepes-*.cjs`: pronóstico automático (fuentes archivos/base, serie, cruce).
- `scripts/weekly-branch-forecast.cjs`, `scripts/lib/weekly-branch-disaggregation.cjs`: reparto semanal y por sucursal (`PRONOSTICO_SEMANAL_SUCURSAL.md`).
- `scripts/*-test.cjs`: pruebas sintéticas (fixture `scripts/fixtures/sintetico-pronostico-2024-2025.json`); todas corren en `npm test`.
- `datos/catalogo/`: mapeo de productos de la base al catálogo; `datos/referencia/`: métricas vigentes.
- `backend/`, `api/`: servidor y funciones de la app web (ver `backend/README.md`).

## Pruebas y build

```
npm ci
npm test
npx vite build
```

`scripts/pronostico-automatico-test.cjs` cubre el modo archivos con datos sintéticos y el modo base con una base simulada en memoria (consultas, bloqueo de escrituras, variables faltantes, producción con órdenes abiertas, contraseña oculta).
