# Scripts de experimento (no usar para presentaciones)

Estos scripts se conservan como registro historico. **No producen pronosticos validos.**

## Por que estan aislados

`add-tolerance-adjustment`, `make-visible-adjustment` y `add-presentation-scenario` calculan el
"pronostico" a partir de la venta real del mes que estan midiendo:

```js
const adjusted = Math.abs(difference) > tolerance
  ? actual - Math.sign(difference) * tolerance
  : base;
```

Eso fuerza por construccion que el error caiga dentro de la tolerancia. Contradice la regla 1 de
`CONTROL_MODELO_PRONOSTICO.md`: *"No usar ventas reales del mes objetivo para calcular su
pronostico"*.

`make-visible-adjustment` ademas sobreescribia el archivo de presentacion en su lugar, sin
respaldo. Esa escritura fue eliminada; ahora solo genera un archivo de salida separado.

Los tres estan bloqueados por `guard-lookahead.cjs` y requieren `PERMITIR_LOOKAHEAD=1` para
ejecutarse.

## restore-real-forecast.cjs

Es la remediacion, no un experimento. Devuelve la columna visible `Pronostico venta` al valor base
original guardado en la columna `S`. Se conserva por si alguien vuelve a contaminar el archivo.

## Alternativa correcta

Para el comparativo real contra ventas de mayo y junio:

```bash
node build-presentation-validado.cjs
```

Usa el modelo validado fuera de muestra y solo histórico anterior a cada mes objetivo.
