# Pronóstico semanal y por sucursal (desagregación del mensual)

Módulo aparte: **no cambia el pronóstico mensual oficial**, solo lo reparte.

- `scripts/lib/weekly-branch-disaggregation.cjs`: función `disaggregateMonthlyForecast({ month, monthlyForecast, dailySales, options })`.
- `scripts/weekly-branch-forecast.cjs`: CLI.
- `scripts/weekly-branch-disaggregation-test.cjs`: prueba sintética, corre dentro de `npm test`.

## Cómo reparte
Todo con ventas diarias **anteriores** al día 1 del mes. Las filas del mes pronosticado se ignoran.

1. **Semana**: semanas ISO (lunes a domingo) recortadas al mes. Peso de cada día =
   - perfil de día de la semana del SKU en las últimas 12 semanas, encogido hacia el perfil de todos los SKU (β = 50 piezas), ×
   - factor de fecha especial: 6-ene, 14-feb, 30-abr, 10-may, 15-sep, 2-nov, 24/31-dic y Día del Padre (3er domingo de junio). El factor es la venta de ese día el año anterior entre el promedio del mismo día de la semana de ese mes, por SKU, encogido hacia el factor de todos los SKU (k = 20) y acotado a [0.2, 8].
2. **Sucursal**: participación del SKU en cada sucursal en las últimas 12 semanas, encogida hacia la participación general de la sucursal (α = 30 piezas). Solo reciben reparto las sucursales que vendieron en los últimos 14 días.
3. La suma de semanas y de sucursales es exactamente el pronóstico mensual del SKU.

## Uso
```
node scripts/weekly-branch-forecast.cjs --month 2025-10 \
  --forecast pronostico-mensual.json \
  --daily ventas-diarias-sucursal.csv \
  --out reparto-2025-10.csv [--level sku]
```
- `pronostico-mensual.json`: `{ "PRODUCTO": cantidad }` o `[{ "producto", "cantidad" }]`.
- `ventas-diarias-sucursal.csv`: columnas `fecha,sucursal,producto,cantidad`, con **ventas de piso de las sucursales**. No incluir "Planta León · Piso de venta", que es surtido a sucursales.
- Para los factores de fechas especiales hace falta al menos el mismo mes del año anterior en el archivo diario. Sin esa historia el factor es 1.

## Backtest 2025 (pepes_devBI, modelo mensual 91a2462 corrido sobre la venta de sucursales)
WAPE %, meses ene–dic 2025. "4 semanas" = promedio diario de las 4 semanas previas al mes. "Año anterior" = misma semana de 2024 (364 días antes).

| Nivel | Este reparto | 4 semanas | Año anterior | Mensual real repartido (piso del reparto) |
|---|---|---|---|---|
| SKU × semana | **17.37** | 23.23 | 27.90 | 12.94 |
| Sucursal × SKU × semana | **42.18** | 46.62 | 63.10 | 40.48 |
| Sucursal × SKU × mes | **24.21** | 29.36 | 44.70 | 20.59 |
| Semana total | **6.67** | 13.84 | 20.37 | 6.14 |
