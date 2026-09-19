const assert = require("assert");
const fs = require("fs");
const path = require("path");

const app = fs.readFileSync(path.join(__dirname, "..", "App.jsx"), "utf8");
const css = fs.readFileSync(path.join(__dirname, "..", "style.css"), "utf8");

assert.match(app, /function SectionDisclosure/, "El rediseño reutiliza un disclosure nativo.");
assert.match(app, /className=\{`promo-section compact-analytics-section/, "Promo activa usa accordion, no un bloque siempre abierto.");
assert.match(app, /Más archivos/, "Los uploads complementarios quedan detrás de Más archivos.");
assert.match(app, /className="notes-section compact-analytics-section"/, "Las notas de interpretación se colapsan.");
assert.match(app, /className="validation-section compact-analytics-section"/, "La validación de cálculos se colapsa.");
assert.match(app, /Más opciones/, "Usuarios y filtros del Excel mensual viven en Más opciones.");
assert.match(app, /secondary-tools-heading/, "El seguimiento queda agrupado como secundario.");
assert.match(app, /Paso 1 · Datos/, "La carga de datos es el primer paso visible.");
assert.match(app, /Paso 2 · Salud/, "La salud del pronóstico es el segundo paso visible.");
assert.match(app, /Paso 3 · Planta/, "La tabla diaria es el tercer paso visible.");
assert.match(app, /className="daily-stock-panel"/, "El inventario diario vive en Planta, no enterrado en Más archivos.");
assert.match(app, /Inventario del día/, "La captura de stock diario usa etiqueta en español.");
assert.match(app, /Pedido planta/, "La tabla de planta muestra el pedido neto.");
assert.doesNotMatch(
  app,
  /<FreezeReadinessStrip[\s\S]*<\/header>/,
  "Las tiras de freeze/salud ya no hinchan el header."
);

assert.match(css, /\.main > \.loaded-files-section \{ order: 2; \}/, "Los datos van primero en el orden visual.");
assert.match(css, /\.main > \.decision-section \{ order: 6; \}/, "La salud del pronóstico sigue a la carga.");
assert.match(css, /\.main > \.daily-section \{ order: 8; \}/, "La tabla de planta queda en el flujo primario.");
assert.match(css, /\.main > \.weekly-progress-section \{ order: 10; \}/, "El avance semanal queda después de planta.");

console.log("ui-hierarchy-test: jerarquía de divulgación progresiva intacta");
