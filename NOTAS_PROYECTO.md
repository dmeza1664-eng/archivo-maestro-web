# Archivo Maestro

Proyecto de pronóstico de ventas y producción para pastelería.

## Ubicación del proyecto

C:\Users\X13\Downloads\archivo-maestro-web

## Archivo fijo de stock ideal

C:\Users\X13\Downloads\STOCK IDEAL SUCURSALES.xlsx

Este archivo debe usarse como fuente fija del stock ideal y debe respetarse el orden de productos.

## Metodología principal

- Pronóstico por producto.
- Backtesting por producto: se prueba cada método contra el mes anterior.
- Selección automática del método con menor error absoluto previo.
- Candidatos: último mes, promedios ponderados, tendencia, día de semana y mismo mes del año anterior.
- Calibración con el error del mes anterior, limitada entre 85% y 115%.
- Calendario del mes objetivo.
- Colchón operativo de 10% a 15%.
- Usar todo el histórico disponible, no solo el último mes.
- Ajuste con existencias y stock ideal.
- Comparación real vs pronóstico.
- Exportación clara a Excel para jefe.
- Nunca usar el mes objetivo ni meses futuros como histórico del pronóstico.

## Histórico de ventas disponible

- `C:\Users\X13\Downloads\VENTA ENERO 2026.xlsx`
- `C:\Users\X13\Downloads\VENTAS FEEEBRERO 2026.xlsx`
- `C:\Users\X13\Downloads\MARZO VENTAS.xlsx`
- `C:\Users\X13\Downloads\venta abril 2026.xlsx`
- `C:\Users\X13\Downloads\VENTAS DE MAYO Y JUNIO - ANGEL.xlsx`

La carga de Ventas acepta varios archivos. Enero y febrero son resúmenes mensuales; marzo contiene hojas diarias del 1 al 15 y abril-junio contienen calendarios diarios.

## Fórmulas base

Existencia actual:

```text
Total Sucursales + Cuarto Frio
```

Pronóstico con colchón:

```text
Promedio historico del dia de semana * 1.10 a 1.15
```

Producción balanceada:

```text
(Stock ideal - Existencia actual + Produccion sugerida) / 2
```

Producción sugerida final:

```text
max(0, produccion balanceada)
```

## Archivos importantes del sistema web

- App.jsx
- style.css
- backend/server.js
- backend/db.js
- backend/routes/ventas.js
- backend/routes/stock.js
- backend/routes/pronostico.js
- backend/routes/produccion.js

## Reglas importantes

- No cambiar el nombre del proyecto: Archivo Maestro.
- No cambiar la metodología principal.
- Usar el stock ideal desde el Excel fijo.
- Mantener comparación real vs pronóstico.
- Exportar resultados a Excel.
