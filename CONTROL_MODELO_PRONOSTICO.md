# Control del modelo de pronostico

Ultima actualizacion: 2026-08-26

> **Aviso de reproducibilidad.** Las cifras de este documento anteriores al
> 2026-08-21 se midieron sobre un catalogo distinto al actual. Ver la seccion
> "El catalogo cambio de hoja" antes de comparar contra cualquier corrida nueva.

## Reglas de control

- No usar ventas reales del mes objetivo para calcular su pronostico.
- Congelar cada pronostico antes de compararlo con el resultado real.
- No modificar retroactivamente valores congelados.
- Medir MAE, WAPE, diferencia total y productos dentro de +/-15 piezas.
- Mantener promociones fuera de los modelos regulares.
- Excluir meses diarios con menos de 70% de cobertura; los archivos declarados como resumen mensual se consideran completos.

## Reglas operativas confirmadas

- La planta produce 6 dias por semana.
- El domingo no hay produccion.
- El producto dura aproximadamente 7 dias antes de convertirse en baja.
- La produccion del sabado debe cubrir tambien la demanda del domingo.
- La produccion sugerida debe descontar existencias en sucursales y cuarto frio.
- El stock ideal es el objetivo; las existencias actuales son el inventario inicial.
- No existe una capacidad maxima diaria declarada; la prioridad es evitar sobreproduccion.
- Los lotes semanales de baja rotacion pueden repartirse entre lunes y sabado.
- Mientras no existan inventarios, no se agrega margen fijo; solo redondeo operativo.

## Homologaciones confirmadas

- PINA GDE = PIÑA GDE
- CHESSECAKE = CHEESECAKE
- MED = MEDIANO

## Segmentos

- Pasteles grandes: modelo estacional y comportamiento reciente.
- Gelatinas regulares: modelo independiente por tamano y presentacion.
- Galletas: seleccion automatica por producto.
- Pan: comportamiento reciente y referencia operativa.
- Promociones: planeacion manual por fechas, cantidad y sucursales.

## Modelo congelado para cinco pasteles

- 90% mismo mes del ano anterior ajustado por crecimiento.
- 10% comportamiento reciente.
- Usar 50% y 50% si ambas referencias difieren mas de 50%.

## Modelo web validado

- Seleccionar una base por producto con backtesting sin usar ventas del mes objetivo.
- Registrar como cero un mes historico completo cuando el producto pertenece al catalogo pero no tuvo ventas en ese mes.
- Si la referencia estacional difiere al menos 8% del comportamiento reciente, combinarla con la base seleccionada.
- Usar 50% de referencia estacional en Otros y Mini medianos, y 75% en las demas categorias regulares.
- Mantener galletas y pan en el selector base por su historial irregular y confianza baja.
- Excluir productos promocionales y productos por rebanada del universo regular.
- Omitir marzo de 2026 del entrenamiento porque el archivo disponible solo contiene los dias 1 al 15.
- No aplicar calibraciones por categoria basadas en el ultimo mes: aumentaron el error fuera de muestra.
- Benchmark reproducible: `node .\benchmark-app-forecast.cjs`.
- Congelar cada mes desde la pagina antes de conocer sus ventas; el respaldo conserva el detalle por producto y descarga el mismo contenido a Excel.

| Mes oculto | Productos | Venta real | Pronostico | WAPE | MAE | Dentro de +/-15 |
|---|---:|---:|---:|---:|---:|---:|
| Abril 2026 | 111 | 21,825 | 21,480.01 | 8.36% | 16.43 | 71 |
| Mayo 2026 | 111 | 27,127 | 27,231.41 | 8.75% | 21.39 | 69 |
| Junio 2026 | 111 | 24,469 | 23,442.59 | 9.14% | 20.16 | 67 |

- Resultado ponderado abril-junio: WAPE aproximado de 8.76%, equivalente a 91.24% de exactitud agregada.
- 207 de 333 comparaciones por producto quedaron dentro de +/-15 piezas.
- Cambios abruptos de alta, baja o reactivacion de productos requieren un estatus operativo conocido antes del mes; no deben inferirse usando ventas futuras.

## Escenario operativo de 12%

- Mantenerlo separado del pronostico estadistico; no cambia el modelo validado.
- El backtest confirma que no debe aplicarse como aumento automatico permanente.

| Mes oculto | WAPE modelo base | WAPE con +12% |
|---|---:|---:|
| Abril 2026 | 8.36% | 12.39% |
| Mayo 2026 | 8.75% | 14.39% |
| Junio 2026 | 9.14% | 11.29% |

- En julio reconstruido, el escenario bajo de 12.42% a 8.48%, pero se calculo despues del cierre y no sustituye al pronostico original.
- Pronostico agosto 2026 generado con historia hasta julio: 27,235.24 piezas base y 30,503.47 piezas con +12%.
- Archivo: `C:\Users\X13\Documents\PRONOSTICO_AGOSTO_2026_BASE_Y_ESCENARIO_12.xlsx`.

## Comparativo mayo-junio para direccion

- Generado con el modelo validado de la aplicacion, no con un modelo paralelo.
- Reproducible: `node build-presentation-validado.cjs`.
- Archivo: `C:\Users\X13\Downloads\PRESENTACION_JEFE_MAYO_JUNIO_MODELO_VALIDADO.xlsx`.
- Reproduce exactamente el benchmark: mayo WAPE 8.75% y junio WAPE 9.14%.
- Exactitud agregada mayo-junio: 91.06%; 136 de 222 comparaciones dentro de +/-15 piezas.
- Sustituye a `PRESENTACION_JEFE_MAYO_JUNIO_AJUSTADO.xlsx`, que usaba un modelo propio
  no validado y reportaba WAPE de 45.67% en mayo y 32.81% en junio sobre 157 productos.

## Aislamiento de scripts con look-ahead

- `add-tolerance-adjustment`, `make-visible-adjustment` y `add-presentation-scenario` calculaban
  el pronostico a partir de la venta real del mes medido, violando la regla 1 de control.
- Movidos a `scripts/experimentos/` y bloqueados por `guard-lookahead.cjs`; requieren
  `PERMITIR_LOOKAHEAD=1` para ejecutarse.
- Se elimino la escritura en sitio que sobreescribia el archivo de presentacion sin respaldo.

## Cierre de julio 2026

Reproducible: `SUMMARY_ONLY=1 WRITE_REPORT=1 node audit-july-close.cjs`.
Reporte: `C:\Users\X13\Documents\COMPARATIVO_JULIO_REAL_VS_PRONOSTICADO.xlsx`.

### Cinco pasteles congelados el 15 de julio

Comparados sin modificar formula ni pesos, como indica `PRONOSTICO_JULIO_CONGELADO.md`.

| Producto | Congelado | Real | Diferencia |
|---|---:|---:|---:|
| FRUTAS GDE | 476 | 580 | +104 |
| MOKA GDE | 594 | 697 | +103 |
| PAY DE GUAYABA GDE | 353 | 413 | +60 |
| DURAZNO GDE | 238 | 322 | +84 |
| CHEESECAKE GDE | 180 | 233 | +53 |
| **Total** | **1,841** | **2,245** | **+404** |

- WAPE 18.00%, exactitud 82.00%, MAE 80.80.
- 0 de 5 productos dentro de +/-15 piezas.
- Los cinco quedaron por debajo de la venta real: el error es sistematico, no aleatorio.

### Catalogo regular completo

| Medida | Modelo estadistico | Escenario operativo +12% |
|---|---:|---:|
| Pronostico | 23,361.44 | 26,164.81 |
| Venta real comparable | 26,312 | 26,312 |
| Diferencia | +2,950.56 | +147.19 |
| WAPE | 12.42% | 8.48% |
| MAE | 29.44 | 20.10 |
| Dentro de +/-15 | 59 de 111 | 63 de 111 |

- Cifra agregada mostrada en la pagina: 26,061 contra 26,312 reales, desviacion de 0.96%.
- Produccion real del catalogo regular: 27,146 piezas, 834 por encima de la venta.
- Las ocho categorias quedaron por debajo de la venta real. Julio fue un cambio de nivel
  generalizado, no un error de productos aislados.
- El escenario operativo de 12% absorbio casi exactamente ese desplazamiento, pero el backtest de
  abril a junio sigue mostrando que no debe volverse un aumento automatico permanente.

### Limitaciones detectadas en el cierre

- El archivo de ventas de julio no trae fechas diarias: 923 renglones, todos totales mensuales.
  No permite evaluar avance semanal ni promedio por dia de semana.
- 17,741 piezas de venta no entraron al catalogo. El desglose es:
  14,585 piezas de accesorios en 15 claves (bengalas 8,314, vela magica 3,358, kits de platos,
  letreros, bolsas, serpentinas) que la planta no produce y quedan fuera de alcance por naturaleza;
  y 3,156 piezas de pasteleria real en 51 claves.
- La pasteleria fuera del catalogo no es un problema de nombres: son productos que no existen en el
  stock ideal. Son lineas nuevas: PETIT 3 LECHES CHOCOLATE y PINERO 1,156 piezas, galletas por sabor
  en bolsa y version tiendita 1,361 piezas, nieves en vaso, linea Dubai y alfajor mundialista.
- **Decision de direccion, 2026-08-17: no entran al catalogo, ni los accesorios.** El universo del
  pronostico se queda en los 111 productos del stock ideal. Las 3,156 piezas de julio y las 14,585
  de accesorios quedan formalmente fuera de alcance, no son una brecha por cerrar. Todas las
  mediciones publicadas ya usaban este universo, asi que ninguna cifra cambia.
- BOLLOS C/4 no existe en el stock ideal, por lo que nunca entro a la comparacion de 111 productos.
  Vendio 28 piezas en julio contra las 215 congeladas. La referencia de pan quedo sin base comparable.
- ~~La categoria Pan del catalogo son productos de PAN DE MUERTO, que venden cero en julio por
  temporada. No hay error ahi.~~ **Corregido el 2026-08-26:** la categoria Pan del catalogo son
  BOLLOS C 4 y BOLLOS C 6, bollos regulares. No hay ningun pan de muerto en el catalogo. La
  afirmacion de que el cero de julio era estacional no se sostiene y el faltante de BOLLOS queda
  sin explicar. Depende ademas de que hoja se lea: ver "El catalogo cambio de hoja".
- Julio 2026 aparece en un solo archivo, pero junio 2026 esta en dos con cifras distintas:
  24,469 piezas en el archivo diario y 25,220 en el cierre mensual. Falta definir cual es la oficial.

### Defecto detectado en el modelo

Cuando un mes tiene total mensual y detalle diario al mismo tiempo, `buildMonthlyForecastData`
descarta el detalle diario y lo reemplaza por un reparto uniforme del total entre todos los dias.
El mes pierde su forma por dia de semana sin ninguna advertencia.

Impacto medido en el pronostico de julio:

| Historico de junio usado | Pronostico julio | WAPE | Dentro de +/-15 |
|---|---:|---:|---:|
| Total mensual, 25,220 piezas, reparto plano | 23,361.44 | 12.42% | 59 de 111 |
| Detalle diario, 24,469 piezas, forma real | 22,616.22 | 15.13% | 54 de 111 |

Las cifras publicadas del cierre de julio corresponden a la primera fila. Contradice la metodologia
de promedio por dia de semana y debe corregirse antes del proximo cierre.

## Diagnostico del crecimiento de julio

Reproducible: `WRITE_REPORT=1 node diagnose-july-growth.cjs`.
Reporte: `C:\Users\X13\Downloads\DIAGNOSTICO_JULIO_CRECIMIENTO_Y_CATALOGO.xlsx`.

### El crecimiento fue real y generalizado

- Julio 2026 vendio 26,312 piezas contra 22,081 de julio 2025: **19.2% de crecimiento anual**.
- 62 de 69 productos comparables crecieron. Mediana por producto x1.22, cuartiles x1.10 y x1.29.
- No fue un producto ni una categoria: fue el mes completo.

### Por que el modelo se quedo corto

El modelo estima el crecimiento anual del mes objetivo usando el crecimiento del mes previo:
`crecimiento = junio 2026 / junio 2025`, limitado a [0.80, 1.20]. Junio 2026 crecio apenas 0.6%
contra junio 2025, asi que el modelo asumio un año plano y aplico ese factor a la referencia
estacional de julio 2025.

El tope de 1.20 no fue la causa: solo recorto 213 piezas y afecto 7 de 68 productos, con razon
mediana de 1.027.

La causa fue el cambio de forma estacional. En 2025 julio cayo 11.9% respecto a junio; en 2026 julio
subio 4.3%. La referencia apuntaba a una baja y ocurrio un alza, un giro de 16 puntos.

### Crecimiento anual mes por mes

| Mes | Año anterior | 2026 | Crecimiento |
|---|---:|---:|---:|
| Abril | 21,043 | 21,825 | 3.7% |
| Mayo | 28,442 | 27,127 | -4.6% |
| Junio | 25,073 | 25,220 | 0.6% |
| Julio | 22,081 | 26,312 | 19.2% |

Julio es el unico mes con crecimiento fuerte. Los tres previos fueron planos o negativos, lo que
explica por que el escenario de +12% empeoro abril, mayo y junio y acerto en julio.

### Conclusion para agosto

El +12% no debe aplicarse como regla fija: solo un mes de cuatro lo justifica. Antes de decidir
agosto hay que revisar si agosto 2025 tuvo una caida parecida a la de julio 2025, porque el modelo
volveria a arrastrar esa forma. La correccion de fondo es estimar el crecimiento anual con varios
meses en lugar de solo el mes previo.

## Busqueda de mayor precision

Reproducible: `node benchmark-model-tuning.cjs`. Incluye julio como cuarto mes oculto y usa el
archivo diario de junio para conservar la forma por dia de semana.

### Ajustar el peso estacional es un camino cerrado

WAPE por version, ocho configuraciones contra cuatro meses ocultos:

| Version | Abril | Mayo | Junio | Julio | Ponderado |
|---|---:|---:|---:|---:|---:|
| legacy | 10.03 | 16.96 | 10.78 | 14.05 | 13.16 |
| rolling | 11.62 | 21.48 | 9.26 | 14.41 | 14.46 |
| seasonal25 | 8.39 | 13.37 | 9.76 | 13.84 | 11.52 |
| seasonal50 | 8.08 | 10.39 | 9.28 | 14.35 | 10.66 |
| **categorySeasonal (actual)** | 8.36 | 8.75 | 9.14 | 15.13 | **10.44** |
| seasonal75 | 8.69 | 9.21 | 9.01 | 15.29 | 10.65 |
| seasonal100 | 10.12 | 10.51 | 9.14 | 16.52 | 11.67 |

La configuracion en produccion es la mejor de las ocho. No hay ganancia disponible moviendo el peso
estacional. En julio ninguna version bajo de 13.84%: fue un quiebre que el modelo no podia anticipar
con solo historia de ventas.

### Composicion del error

| Mes | Error total | Descontinuado | Nuevo | Estimacion | Evitable |
|---|---:|---:|---:|---:|---:|
| Abril | 1,824 | 118 | 21 | 1,685 | 7.6% |
| Mayo | 2,374 | 529 | 0 | 1,845 | 22.3% |
| Junio | 2,237 | 117 | 249 | 1,871 | 16.4% |
| Julio | 3,981 | 0 | 439 | 3,542 | 11.0% |

- Descontinuado: el producto no vendio nada pero el modelo si lo pronostico.
- Nuevo: el producto vendio pero no tenia historico para pronosticarlo.
- **14.1% del error acumulado es estatus de producto**, no estimacion de cantidad. Se elimina
  capturando activo, baja o nuevo antes del mes, sin tocar el modelo.
- El error de estimacion se mantuvo entre 1,685 y 1,871 piezas en abril, mayo y junio, y salto a
  3,542 en julio. El quiebre de julio es estimacion pura.

### Mejora encontrada: usar el cierre mensual junto con el detalle diario

Junio 2026 existe en dos archivos: 24,469 piezas en el diario y 25,220 en el cierre mensual.
Alimentar ambos, conservando la forma por dia de semana del diario y ajustando su nivel al total
del cierre, mejora el pronostico de julio de forma medible.

| Historico de junio | Abril | Mayo | Junio | Julio | Ponderado |
|---|---:|---:|---:|---:|---:|
| Solo archivo diario, 24,469 | 8.36 | 8.75 | 9.14 | 15.13 | 10.44 |
| Diario + cierre mensual, nivel 25,220 | 8.36 | 8.75 | 9.14 | **12.48** | **9.74** |

- Ganancia de 0.70 puntos en el ponderado y 2.65 puntos en julio.
- La ganancia viene del nivel, no de la forma: el cierre mensual es 751 piezas mas alto, 3.1%.
- Conservar la forma diaria cuesta 0.06 puntos en el WAPE mensual y a cambio devuelve el promedio
  por dia de semana, que es lo que usa la planta para producir. Se conserva.
- Regla operativa: alimentar siempre el cierre mensual ademas del detalle diario.
- Queda pendiente definir cual es la cifra oficial de junio. Que 25,220 mejore el pronostico de
  julio fuera de muestra es evidencia a su favor, pero la definicion es del area de datos.

### Estatus por producto: medido, el beneficio es futuro y no historico

Archivo llenado por direccion: `C:\Users\X13\Downloads\ESTATUS_PRODUCTOS_PARA_LLENAR.xlsx`.
Resultado: 25 BAJA, 14 ESTACIONAL, 1 BAJO PEDIDO, 71 ACTIVO.
Reproducible: `node measure-status-impact.cjs` y `node measure-status-august.cjs`.

| Medicion | Resultado |
|---|---|
| Ponderado abril-julio sin estatus | 9.74% |
| Ponderado abril-julio con estatus | 9.63% |
| Ganancia historica | 0.12 puntos, 118 piezas |
| Efecto en el pronostico de agosto | **0 piezas** |

El modelo ya se autocorrige: `fillCompleteZeroMonths` registra ceros en los meses sin venta, asi que
tras dos o tres meses inactivo el pronostico decae a cero por si solo. Los 39 productos marcados
BAJA o ESTACIONAL ya recibian cero. La estimacion previa de 14% de error evitable sobreestimaba el
beneficio accionable.

**Matiz del 2026-08-26:** ese cero no prueba lo que parecia. Con el catalogo actual, solo 77 de 113
productos cruzan por nombre con el archivo de estatus, y de los 23 marcados BAJA unos 15 ni siquiera
estan en el catalogo. Suprimirlos no podia cambiar nada porque nunca recibieron pronostico. La
medicion cruzo dos listas que en su mayoria no se cruzan. Hay que rehacerla cuando se defina cual es
la hoja oficial del stock ideal.

El valor real del estatus es otro:

- Avisar una baja **el mismo mes** en que ocurre, en lugar de esperar dos o tres meses a que el
  modelo lo deduzca. Ese es el ahorro tipo 500 piezas por evento.
- Distinguir temporada de baja para septiembre y octubre: los 14 productos de pan de muerto y
  calabaza deben volver a pronosticarse en su temporada, no quedar apagados.

### Correccion: GELATINA FRESA $150 no esta descontinuada

Se afirmo antes que este producto estaba dado de baja y que explicaba 21% del error de mayo. La
contribucion al error es correcta, la causa no. Vendio 0 en mayo, 0 en junio y **438 piezas en
julio**. Direccion lo marco ACTIVO.

Es demanda intermitente, no una baja. Ningun estatus lo habria resuelto: apagarlo en mayo habria
acertado y en julio habria fallado. Los productos con este patron necesitan tratamiento aparte.

### Caminos cerrados, medidos y descartados

- **Ajustar el peso estacional:** ocho configuraciones probadas, la actual es la mejor. Sin ganancia.
- **Peso estacional adaptativo por producto:** elegirlo por backtesting de los ultimos tres meses
  sobreajusta. Ponderado de 12.89% contra 10.44% del fijo. Disponible como `seasonalAdaptive` solo
  para benchmark. Confirma la regla ya documentada de no calibrar con los ultimos meses.
- **Deteccion automatica de descontinuados:** de las 764 piezas de error por producto sin venta,
  solo 235 son detectables por inactividad previa. GELATINA FRESA $150 vendio en abril y paro en
  mayo; ningun dato historico lo anticipa. El estatus tiene que capturarlo una persona.

### Sesgo por mes

| Mes | Sesgo |
|---|---|
| Abril | quedo corto 1.6% |
| Mayo | se paso 0.4% |
| Junio | quedo corto 4.2% |
| Julio | quedo corto 14.0% |

Tres de cuatro meses quedaron cortos y la magnitud crece. Aun asi, corregir julio con el sesgo
promedio de abril a junio, que es 1.8%, habria movido el pronostico de 22,616 a 23,023 contra 26,312
reales. Una correccion automatica de sesgo no resuelve el quiebre y vuelve el modelo reactivo.

## El catalogo cambio de hoja

Reproducible: `node scripts/measure-seasonal-reactivation.cjs`.

`parseStock` no lee el archivo de stock ideal completo: elige **una hoja** de una lista de
candidatas. El archivo guarda varias hojas que son fotos de fechas distintas, asi que la hoja que se
elija define el universo del pronostico.

La lista nacio el 2026-08-21 en `d0ec122` con `EXIST. SUCURSALES Y RESTANTE CF` primero. Ese mismo
dia `973e29f` invirtio el orden y dejo `TOTAL A TENER SUC.(EXIST.+DIST)` primero. El commit se
titulaba "Corrige la carga inicial de datos historicos" y no menciona el cambio de universo.

Las dos hojas comparten solo **79 productos**; cada una aporta 34 propios:

| | `EXIST. SUCURSALES Y RESTANTE CF` | `TOTAL A TENER SUC.(EXIST.+DIST)` |
|---|---:|---:|
| Productos | 113 | 113 |
| De temporada de muertos | 16 | 0 |
| Cruzan con el archivo de estatus | 111 de 111 | 77 de 111 |

- La hoja anterior trae la linea de pan de muerto y productos que luego se dieron de baja
  (IMPERIAL CHOCOLATE, SELVA NEGRA, GELATINA PIÑA COLADA): es una foto de octubre 2025.
- La hoja actual trae la linea de San Valentin, BOLLOS C 4 y C 6, y las lineas nuevas con stock real
  (PETIT 3 LECHES CHOCOLATE 80, CAPIROTADA 70, GALLETA ATE BOLSA 150): es una foto mas reciente.
- Ninguna de las dos es el catalogo completo.

Consecuencias:

- **Las cifras publicadas no se reproducen.** El WAPE de 8.36, 8.75 y 9.14 y el cierre de julio se
  midieron antes del 2026-08-21, es decir sobre la hoja anterior. Correr el benchmark hoy mide otro
  universo.
- **La temporada de pan de muerto es invisible al sistema.** Los 16 productos ESTACIONAL no existen
  en el catalogo actual, asi que `calculateForecast` nunca emite un renglon para ellos. No es un
  problema de pesos ni de estatus. La temporada pasada fueron 3,889 piezas: 399 en septiembre,
  3,490 en octubre y 0 en noviembre. Un solo producto, PAN MUERTO IND AZUCAR 50GR, son 2,174.
- La produccion de principios de octubre se decide a finales de septiembre.

Mitigacion aplicada el 2026-08-26: `assessStockSheetSelection` compara la hoja elegida contra las
otras candidatas y advierte, al cargar el stock ideal y en las alertas de validacion, cuando alguna
trae productos que la elegida no incluye. No decide cual hoja es la correcta; solo evita que el
universo vuelva a cambiar en silencio.

Pendiente de direccion: definir cual hoja es el stock ideal oficial, o si el catalogo debe ser la
union de todas, y si la temporada de pan de muerto se planea este año en el sistema.

## Avance semanal

- Comparar venta real contra el pronostico solo hasta la ultima fecha diaria cargada.
- Mantener intacto el pronostico mensual usado como referencia.
- Calcular cumplimiento semanal como `Venta real / Pronostico al corte`.
- Proyectar el cierre como `Venta real cargada + Pronostico aun no medido`.
- Usar semaforo verde dentro de +/-10%, amarillo entre 10% y 20%, y rojo por encima de 20%.
- Mostrar cobertura de fechas para advertir cuando la carga diaria esta incompleta.
- Permitir desglose por producto, categoria y exportacion a Excel.

## Factores de evento

- Dia de las Madres: conservar 90% del incremento historico.
- Dia del Padre: conservar 35% del incremento historico.
- Las ventanas se alinean por evento, no solo por numero de dia.

## Pronostico julio congelado

| Producto | Pronostico piezas |
|---|---:|
| FRUTAS GDE | 476 |
| MOKA GDE | 594 |
| PAY DE GUAYABA GDE | 353 |
| DURAZNO GDE | 238 |
| CHEESECAKE GDE | 180 |
| **Total** | **1,841** |

Fuente detallada: `PRONOSTICO_JULIO_CONGELADO.md`.

## Estado

- [x] Homologar nombres principales.
- [x] Separar promociones del modelo regular.
- [x] Validar cinco pasteles grandes en mayo y junio.
- [x] Calcular factores de Dia de las Madres y Dia del Padre.
- [x] Congelar pronostico de julio para cinco pasteles.
- [x] Extender julio al resto de pasteles grandes: 18 productos, total 4,491 piezas.
- [x] Calcular julio para pasteles medianos: 11 productos, total 2,851 piezas.
- [x] Calcular julio para pasteles chicos: 14 productos, total 3,640 piezas.
- [x] Calcular julio para mini medianos: 7 productos, total 2,670 piezas.
- [x] Calcular julio para gelatinas regulares: 11 productos, total redondeado 2,800 piezas.
- [x] Definir y congelar modelo de julio para galletas: 9 productos, total 1,496 piezas, confianza baja.
- [x] Definir y congelar referencia de julio para pan: BOLLOS C/4, total 215 piezas, confianza baja.
- [x] Congelar calendario diario de produccion base sin existencias: total 19,361 piezas.
- [x] Generar detalle diario por producto: 1,070 renglones en `C:\Users\X13\Downloads\PRODUCCION_JULIO_DETALLE_POR_PRODUCTO.csv`.
- [x] Marcar baja rotacion como bajo pedido: 51 renglones, demanda acumulada 175.32 piezas.
- [x] Generar plan restante del 15 al 30 de julio con saldo virtual: `C:\Users\X13\Downloads\PRODUCCION_RESTANTE_JULIO_CON_SALDO_VIRTUAL.csv`.
- [x] Reducir produccion repetida con saldo virtual: ahorro calculado de 530 piezas.
- [x] Produccion restante base: 9,311 piezas; demanda bajo pedido: 233.93 piezas.
- [x] El plan restante con saldo virtual sustituye al calendario bruto para las fechas del 15 al 30 de julio.
- [x] Generar comparativo retrospectivo mayo-junio: `C:\Users\X13\Downloads\COMPARATIVO_MAYO_JUNIO_REAL_VS_PRONOSTICO.xlsx`.
- [x] Generar control diario de julio con captura de ventas, produccion, bajas, saldos y formulas: `C:\Users\X13\Downloads\CONTROL_DIARIO_JULIO_PRODUCCION_VENTAS.xlsx`.
- [x] Comparar contra ventas reales al cierre de julio: modelo estadistico WAPE 12.42%, escenario
      operativo +12% WAPE 8.48%, cinco pasteles congelados WAPE 18.00%.
- [x] Diagnosticar el crecimiento de julio: 19.2% anual generalizado, causado por un cambio de forma
      estacional que el estimador de crecimiento de un solo mes no capta.
- [x] Homologar CHEESECAKE y CHESSECAKE en codigo; el benchmark se mantiene en 8.36, 8.75 y 9.14.
- [x] Decidir si las 3,156 piezas de pasteleria fuera del catalogo entran al stock ideal: no entran.
      Los accesorios tampoco. El universo se queda en 111 productos.
- [x] Capturar estatus operativo por producto: 25 BAJA, 14 ESTACIONAL, 1 BAJO PEDIDO, 71 ACTIVO.
- [ ] Confirmar si PAN DE MUERTO CHOCOLATE GDE y PAN DE MUERTO IND CHOCOLATE son BAJA o ESTACIONAL;
      como BAJA nunca se produciran en octubre.
- [ ] Verificar que las lineas fuera de alcance se planeen por otra via antes de septiembre.
- [x] Corregir el descarte silencioso del detalle diario cuando existe total mensual del mismo mes:
      ahora se conserva la forma diaria y se ajusta su nivel al total declarado.
- [x] Medir el efecto de alimentar el cierre mensual: ponderado de 10.44% a 9.74%.
- [ ] Definir la cifra oficial de junio 2026: 24,469 del archivo diario o 25,220 del cierre mensual.
- [ ] Cargar siempre el cierre mensual ademas del detalle diario en la pagina.
- [x] Estimar el crecimiento anual con varios meses en lugar de solo el mes previo: mediana de tres
      meses, `f2bb0d9`.
- [ ] Pedir el archivo de ventas mensual con fecha diaria si se requiere avance semanal.
- [ ] Registrar estatus operativo por producto antes del mes para evitar pronosticar descontinuados.
- [x] Advertir cuando la hoja elegida del stock ideal omite productos que otra hoja si trae, para que
      el universo no vuelva a cambiar en silencio.
- [ ] **Vencido.** Congelar agosto antes de conocer sus ventas no ocurrio; agosto ya esta corrido.
      Aplica ahora a septiembre, y depende de definir la hoja oficial del catalogo.
- [ ] Definir con direccion cual hoja del stock ideal es el catalogo oficial.
- [ ] Decidir si la temporada de pan de muerto se planea en el sistema: 3,889 piezas el año pasado,
      concentradas en octubre. Hoy esos 16 productos no existen en el catalogo.
- [ ] Rehacer la medicion del efecto del estatus cuando el catalogo quede definido; la de agosto
      cruzo listas que no empatan.
- [ ] Revisar si septiembre 2025 tuvo una caida como la de julio 2025 antes de fijar el plan de
      septiembre. Sustituye a la revision de agosto, que ya no aplica.
