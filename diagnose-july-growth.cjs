const fs = require("fs");
const path = require("path");
const Module = require("module");
const esbuild = require("esbuild");
const XLSX = require("xlsx");

const ROOT = __dirname;
const DOWNLOADS = path.resolve(ROOT, "..");
const JULY_SALES = "C:/Users/X13/Documents/ventas por mes/ventas julio.xlsx";
const FINAL_JUNE_SALES = "C:/Users/X13/Documents/ventas por mes/ventas junio.xlsx";

const DOWNLOAD_SOURCES = [
  "Venta de diciembre 2024.xlsx",
  "VENTA AÑO 2025 - ANGEL.xlsx",
  "VENTA ENERO 2026.xlsx",
  "VENTAS FEEEBRERO 2026.xlsx",
  "MARZO VENTAS.xlsx",
  "venta abril 2026.xlsx",
  "VENTAS DE MAYO Y JUNIO - ANGEL.xlsx",
];

const ACCESSORY_PATTERNS = [
  /BENGALA/, /\bVELA\b/, /VELADORA/, /SERPENTINA/, /KIT DE PLATOS/, /LETRERO/,
  /BOLSA ECOLOGICA/, /CHUNCHES/, /ACETATO/, /OTROS VARIOS/, /\bMONEDA\b/,
  /ENCENDEDOR/, /GLOBO/, /MOÑO/, /LISTON/, /CHAROLA/, /BASE\b/, /TAPA\b/,
];

function normalized(value) {
  return String(value || "")
    .toUpperCase()
    .normalize("NFD")
    .replace(/[\u0300-\u036f]/g, "")
    .replace(/[.,/\\_-]+/g, " ")
    .replace(/\s+/g, " ")
    .trim();
}

function matchKey(value) {
  return normalized(value)
    .replace(/\b(PASTEL|TARTA|PANQUE|PAY|DE|DEL|LA|EL)\b/g, " ")
    .replace(/\bGRANDE\b/g, "GDE")
    .replace(/\bMEDIANO\b/g, "MED")
    .replace(/\bCHICO\b/g, "CH")
    .replace(/\s+/g, " ")
    .trim();
}

function findDownload(search) {
  const target = normalized(search);
  const match = fs.readdirSync(DOWNLOADS).find((name) => normalized(name) === target);
  if (!match) throw new Error(`No se encontro ${search}`);
  return path.join(DOWNLOADS, match);
}

function monthKey(value) {
  const date = value instanceof Date ? value : new Date(value);
  if (Number.isNaN(date.getTime())) return "";
  return `${date.getFullYear()}-${String(date.getMonth() + 1).padStart(2, "0")}`;
}

function isExcluded(product) {
  const value = normalized(product);
  return /\b(REBANADA|REBANADAS|REB|RBN)\b/.test(value) || /\b(PROMO|PROMOCION|PROMOCIONAL)\b/.test(value);
}

function isAccessory(product) {
  const value = normalized(product);
  return ACCESSORY_PATTERNS.some((pattern) => pattern.test(value));
}

function pct(value) {
  return `${(value * 100).toFixed(1)}%`;
}

async function loadAppFunctions() {
  const built = await esbuild.build({
    entryPoints: [path.join(ROOT, "App.jsx")],
    bundle: true,
    platform: "node",
    format: "cjs",
    write: false,
    loader: { ".css": "text" },
    define: { "import.meta.env.VITE_API_URL": JSON.stringify("") },
    logLevel: "silent",
  });
  const appModule = new Module("july-growth-diagnosis");
  appModule.filename = path.join(ROOT, "july-growth-diagnosis.bundle.cjs");
  appModule.paths = module.paths;
  appModule._compile(built.outputFiles[0].text, appModule.filename);
  return appModule.exports;
}

async function main() {
  const app = await loadAppFunctions();
  const stockRows = app.parseStock(XLSX.readFile(findDownload("STOCK IDEAL SUCURSALES.xlsx"), { cellDates: true }));
  const officialProducts = stockRows.map((row) => normalized(row.producto)).filter((p) => !isExcluded(p));
  const byMatchKey = new Map();
  for (const product of officialProducts) {
    const key = matchKey(product);
    if (!byMatchKey.has(key)) byMatchKey.set(key, []);
    byMatchKey.get(key).push(product);
  }
  const direct = new Set(officialProducts);
  const toOfficial = (raw) => {
    const source = normalized(raw);
    if (direct.has(source)) return source;
    const candidates = byMatchKey.get(matchKey(source)) || [];
    return candidates.length === 1 ? candidates[0] : "";
  };

  // Carga cada archivo por separado para detectar meses duplicados entre fuentes.
  const files = [];
  for (const sourceName of DOWNLOAD_SOURCES) {
    const filePath = findDownload(sourceName);
    files.push({
      name: path.basename(filePath),
      rows: app.parseSalesOrReturns(XLSX.readFile(filePath, { cellDates: true }), "ventas", path.basename(filePath)),
    });
  }
  for (const extra of [FINAL_JUNE_SALES, JULY_SALES]) {
    files.push({
      name: path.basename(extra),
      rows: app.parseSalesOrReturns(XLSX.readFile(extra, { cellDates: true }), "ventas", path.basename(extra)),
    });
  }

  console.log("=== APORTE POR ARCHIVO (piezas del catalogo regular, por mes) ===");
  const monthsByFile = new Map();
  for (const file of files) {
    const perMonth = new Map();
    for (const row of file.rows) {
      const official = toOfficial(row.producto);
      if (!official) continue;
      const key = monthKey(row.fecha) || "sin-fecha";
      perMonth.set(key, (perMonth.get(key) || 0) + Number(row.cantidad || 0));
    }
    monthsByFile.set(file.name, perMonth);
    const detail = [...perMonth.entries()].sort().map(([m, v]) => `${m}:${Math.round(v)}`).join("  ");
    console.log(`${file.name}\n   ${detail || "(sin coincidencias)"}`);
  }

  console.log("\n=== MESES APORTADOS POR MAS DE UN ARCHIVO ===");
  const monthSources = new Map();
  for (const [fileName, perMonth] of monthsByFile.entries()) {
    for (const month of perMonth.keys()) {
      if (!monthSources.has(month)) monthSources.set(month, []);
      monthSources.get(month).push(fileName);
    }
  }
  let duplicates = 0;
  for (const [month, sources] of [...monthSources.entries()].sort()) {
    if (sources.length > 1) {
      duplicates += 1;
      const totals = sources.map((s) => `${s}=${Math.round(monthsByFile.get(s).get(month))}`).join("  |  ");
      console.log(`${month}: ${totals}`);
    }
  }
  if (!duplicates) console.log("(ninguno)");

  // Serie mensual usando una sola fuente por mes: la de mayor volumen.
  const monthlyTotal = new Map();
  const monthlyByProduct = new Map();
  for (const [month, sources] of monthSources.entries()) {
    if (month === "sin-fecha") continue;
    const best = sources.reduce((a, b) =>
      (monthsByFile.get(a).get(month) || 0) >= (monthsByFile.get(b).get(month) || 0) ? a : b
    );
    const file = files.find((f) => f.name === best);
    const perProduct = new Map();
    let total = 0;
    for (const row of file.rows) {
      if ((monthKey(row.fecha) || "sin-fecha") !== month) continue;
      const official = toOfficial(row.producto);
      if (!official) continue;
      const quantity = Number(row.cantidad || 0);
      perProduct.set(official, (perProduct.get(official) || 0) + quantity);
      total += quantity;
    }
    monthlyTotal.set(month, total);
    monthlyByProduct.set(month, perProduct);
  }

  console.log("\n=== SERIE MENSUAL DEL CATALOGO REGULAR ===");
  const months = [...monthlyTotal.keys()].sort();
  let previous = null;
  for (const month of months) {
    const total = monthlyTotal.get(month);
    const mom = previous ? ` | vs mes previo: ${pct(total / previous - 1)}` : "";
    console.log(`${month}  ${String(Math.round(total)).padStart(7)}${mom}`);
    previous = total;
  }

  console.log("\n=== MISMO MES, AÑO CONTRA AÑO ===");
  for (const month of months.filter((m) => m.startsWith("2026"))) {
    const priorYear = `${Number(month.slice(0, 4)) - 1}-${month.slice(5)}`;
    if (!monthlyTotal.has(priorYear)) continue;
    const current = monthlyTotal.get(month);
    const prior = monthlyTotal.get(priorYear);
    console.log(
      `${month} vs ${priorYear}:  ${String(Math.round(prior)).padStart(6)} -> ${String(Math.round(current)).padStart(6)}  crecimiento ${pct(current / prior - 1)}`
    );
  }

  // El modelo calcula crecimiento por producto como jun2026/jun2025 y lo limita a [0.8, 1.2].
  console.log("\n=== TOPE DE CRECIMIENTO DEL MODELO (jun2026 / jun2025, limite 1.20) ===");
  const june2026 = monthlyByProduct.get("2026-06") || new Map();
  const june2025 = monthlyByProduct.get("2025-06") || new Map();
  const ratios = [];
  for (const product of officialProducts) {
    const current = june2026.get(product) || 0;
    const prior = june2025.get(product) || 0;
    if (current > 0 && prior > 0) ratios.push({ product, ratio: current / prior, current, prior });
  }
  const above = ratios.filter((r) => r.ratio > 1.2);
  const below = ratios.filter((r) => r.ratio < 0.8);
  console.log(`Productos comparables: ${ratios.length}`);
  console.log(`Con crecimiento arriba de 1.20 (el tope los recorta): ${above.length}`);
  console.log(`Con caida abajo de 0.80 (el tope los sube): ${below.length}`);
  if (ratios.length) {
    const sorted = ratios.map((r) => r.ratio).sort((a, b) => a - b);
    const median = sorted[Math.floor(sorted.length / 2)];
    console.log(`Razon mediana: ${median.toFixed(3)}`);
    const suppressed = above.reduce((sum, r) => sum + r.prior * (r.ratio - 1.2), 0);
    console.log(`Piezas de crecimiento que el tope descarto en junio: ${Math.round(suppressed)}`);
  }
  console.log("Mayores recortes por el tope:");
  above.sort((a, b) => b.ratio - a.ratio).slice(0, 8)
    .forEach((r) => console.log(`   ${r.product.padEnd(30)} ${Math.round(r.prior)} -> ${Math.round(r.current)}  x${r.ratio.toFixed(2)}`));

  console.log("\n=== CRECIMIENTO DE JULIO POR PRODUCTO (jul2026 vs jul2025) ===");
  const july2026 = monthlyByProduct.get("2026-07") || new Map();
  const july2025 = monthlyByProduct.get("2025-07") || new Map();
  const julyRatios = officialProducts
    .map((product) => ({ product, current: july2026.get(product) || 0, prior: july2025.get(product) || 0 }))
    .filter((r) => r.current > 0 && r.prior > 0)
    .map((r) => ({ ...r, ratio: r.current / r.prior }));
  if (julyRatios.length) {
    const sorted = julyRatios.map((r) => r.ratio).sort((a, b) => a - b);
    const q = (p) => sorted[Math.min(sorted.length - 1, Math.floor(sorted.length * p))];
    console.log(`Productos comparables: ${julyRatios.length}`);
    console.log(`Cuartil 1: x${q(0.25).toFixed(2)} | Mediana: x${q(0.5).toFixed(2)} | Cuartil 3: x${q(0.75).toFixed(2)}`);
    console.log(`Productos que crecieron: ${julyRatios.filter((r) => r.ratio > 1).length} de ${julyRatios.length}`);
  } else {
    console.log("Sin comparables: falta julio 2025 o julio 2026 en el catalogo regular.");
  }

  console.log("\n=== VENTA DE JULIO QUE NO ENTRA AL CATALOGO ===");
  const julyFile = files.find((f) => f.name === path.basename(JULY_SALES));
  const unmatched = new Map();
  for (const row of julyFile.rows) {
    const source = normalized(row.producto);
    const quantity = Number(row.cantidad || 0);
    if (!source || quantity === 0 || isExcluded(source)) continue;
    if (toOfficial(source)) continue;
    unmatched.set(source, (unmatched.get(source) || 0) + quantity);
  }
  const accessories = [...unmatched.entries()].filter(([p]) => isAccessory(p));
  const realProducts = [...unmatched.entries()].filter(([p]) => !isAccessory(p));
  const sum = (list) => Math.round(list.reduce((total, [, quantity]) => total + quantity, 0));
  console.log(`Accesorios y complementos: ${sum(accessories)} piezas en ${accessories.length} claves`);
  console.log(`Producto de pasteleria fuera del catalogo: ${sum(realProducts)} piezas en ${realProducts.length} claves`);
  console.log("\nPasteleria fuera del catalogo (top 25):");
  realProducts.sort((a, b) => b[1] - a[1]).slice(0, 25)
    .forEach(([product, quantity]) => console.log(`   ${String(Math.round(quantity)).padStart(6)}  ${product}`));

  if (process.env.WRITE_REPORT !== "1") return;

  const outputPath = path.join(DOWNLOADS, "DIAGNOSTICO_JULIO_CRECIMIENTO_Y_CATALOGO.xlsx");
  const workbook = XLSX.utils.book_new();
  const append = (rows, name, widths) => {
    const sheet = XLSX.utils.json_to_sheet(rows);
    sheet["!cols"] = widths.map((wch) => ({ wch }));
    sheet["!autofilter"] = { ref: sheet["!ref"] };
    sheet["!freeze"] = { xSplit: 0, ySplit: 1, topLeftCell: "A2", activePane: "bottomRight", state: "frozen" };
    XLSX.utils.book_append_sheet(workbook, sheet, name);
  };

  append(
    months.map((month, index) => {
      const total = Math.round(monthlyTotal.get(month));
      const priorMonth = index > 0 ? monthlyTotal.get(months[index - 1]) : null;
      const priorYear = `${Number(month.slice(0, 4)) - 1}-${month.slice(5)}`;
      return {
        Mes: month,
        "Piezas catalogo regular": total,
        "Cambio vs mes previo": priorMonth ? pct(total / priorMonth - 1) : "",
        "Mismo mes año anterior": monthlyTotal.has(priorYear) ? Math.round(monthlyTotal.get(priorYear)) : "",
        "Crecimiento anual": monthlyTotal.has(priorYear) ? pct(total / monthlyTotal.get(priorYear) - 1) : "",
      };
    }),
    "Serie mensual",
    [12, 24, 22, 24, 20]
  );

  append(
    julyRatios.sort((a, b) => b.ratio - a.ratio).map((row) => ({
      Producto: row.product,
      "Julio 2025": Math.round(row.prior),
      "Julio 2026": Math.round(row.current),
      Diferencia: Math.round(row.current - row.prior),
      Crecimiento: pct(row.ratio - 1),
    })),
    "Crecimiento julio",
    [38, 14, 14, 14, 14]
  );

  append(
    [
      ...realProducts.map(([product, quantity]) => ({
        Clave: product,
        "Piezas julio": Math.round(quantity),
        Clasificacion: "Pasteleria fuera del catalogo",
        Accion: "Decidir si entra al stock ideal o se declara fuera de alcance",
      })),
      ...accessories.sort((a, b) => b[1] - a[1]).map(([product, quantity]) => ({
        Clave: product,
        "Piezas julio": Math.round(quantity),
        Clasificacion: "Accesorio o complemento",
        Accion: "Fuera de alcance del pronostico de produccion",
      })),
    ],
    "Fuera de catalogo",
    [40, 14, 32, 56]
  );

  append(
    [
      { Tema: "Crecimiento de julio", Detalle: `Julio 2026 vendio ${pct(monthlyTotal.get("2026-07") / monthlyTotal.get("2025-07") - 1)} mas que julio 2025, con mediana por producto de x${julyRatios.length ? [...julyRatios].map((r) => r.ratio).sort((a, b) => a - b)[Math.floor(julyRatios.length / 2)].toFixed(2) : "?"}.` },
      { Tema: "Por que el modelo se quedo corto", Detalle: "El modelo estima el crecimiento anual del mes objetivo usando el crecimiento del mes previo. Junio 2026 crecio solo 0.6% contra junio 2025, asi que el modelo asumio un año plano y aplico ese factor a la referencia de julio 2025." },
      { Tema: "El tope no fue la causa", Detalle: "El limite de crecimiento de 1.20 solo recorto 213 piezas en junio y afecto 7 de 68 productos. La razon mediana fue 1.027." },
      { Tema: "Cambio de forma estacional", Detalle: "En 2025 julio cayo 11.9% respecto a junio. En 2026 julio subio 4.3% respecto a junio. La referencia estacional apuntaba a una baja y ocurrio un alza." },
      { Tema: "Reproducible", Detalle: "WRITE_REPORT=1 node diagnose-july-growth.cjs" },
    ],
    "Conclusiones",
    [34, 120]
  );

  XLSX.writeFile(workbook, outputPath);
  console.log(`\nReporte: ${outputPath}`);
}

main().catch((error) => {
  console.error(error.message);
  process.exit(1);
});
