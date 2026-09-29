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
   - víspera (1 y 2 días antes de esas fechas) **solo para los SKU que históricamente suben la víspera**. Con la venta de los 730 días previos al mes se compara, por SKU, la venta de cada víspera contra el promedio del mismo día de la semana en días normales de ese mes (sin evento, víspera ni días 15–17). La razón se encoge hacia 1 con 30 piezas; el SKU entra si queda en ≥ 1.5 con al menos 4 vísperas observadas, y su víspera se multiplica por esa razón (tope 3). Los demás SKU no se tocan (un multiplicador igual para todos empeoraba 2026). Sin nombres de SKU. `--no-vispera` (u `options.visperas = false`) lo apaga.
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
- `ventas-diarias-sucursal.csv`: columnas `fecha,sucursal,producto,cantidad`, con **ventas de piso de las sucursales**. "Planta León · Piso de venta" es surtido a sucursales: desde el 26-sep-2026 el módulo descarta solo las filas cuya sucursal dice PLANTA (y Suc. Amado Nervo antes de ago-2024), igual que la serie oficial de demanda del mensual.
- Para los factores de fechas especiales hace falta al menos el mismo mes del año anterior en el archivo diario. Sin esa historia el factor es 1.

## Backtest (pepes_devBI, modelo mensual `d0e61f7` sobre la venta de sucursales)
WAPE %, con la víspera por SKU activa. 2025 = ene–dic; 2026 = ene–ago (fuera de muestra); 2024 = mar–dic sin Suc. Amado Nervo. El reparto no cambia el mensual: solo lo distribuye.

| Nivel | 2025 | 2026 | 2024 |
|---|---|---|---|
| SKU × semana | **16.98** (sin enero 17.10) | **17.73** | **21.54** |
| Sucursal × SKU × semana | **42.21** | **42.89** | **45.14** |
| Semana total | **6.48** | **6.86** | **12.81** |

Bases ingenuas (mismo universo y misma venta diaria): "4 semanas" = promedio diario de las 4 semanas previas al mes; "Año anterior" = misma semana 364 días antes.

| Nivel | 2025: 4 semanas / año anterior | 2026: 4 semanas / año anterior |
|---|---|---|
| SKU × semana | 23.23 / 27.90 | 25.54 / 24.73 |
| Sucursal × SKU × semana | 46.62 / 63.10 | 49.14 / 59.98 |

Fuente: `backtest-2025-completo-pepes-sucursal/semanal/wk-suc-{2024,2025,2026}.json` (arnés `wk.cjs`) y `wk-naive.cjs`. La víspera por SKU (#30) bajó el error semanal en los tres años cuando se integró; el detalle de esa medición está en el PR #30.
