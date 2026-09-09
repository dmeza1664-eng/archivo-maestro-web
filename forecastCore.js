export function norm(value) {
  return String(value ?? "")
    .trim()
    .toUpperCase()
    .normalize("NFD")
    .replace(/[\u0300-\u036f]/g, "");
}

export function normalizeProduct(value) {
  const normalized = norm(value)
    .replace(/[.,/\\_-]+/g, " ")
    .replace(/\s+/g, " ")
    .trim();
  const compact = normalized.replace(/[^A-Z0-9]/g, "");
  if (compact === "PINAGDE" || compact === "PINAGRANDE") return "PINA GDE";
  return normalized;
}

export function isSliceProduct(value) {
  const normalized = normalizeProduct(value);
  return /\b(REBANADA|REBANADAS|REB|RBN)\b/.test(normalized);
}

export function isOperationalCakeProduct(value) {
  const normalized = normalizeProduct(value);
  return /\b(GDE|GRANDE|MED|MEDIANO|CH|CHICO)\b/.test(normalized);
}

export function getProduccionSugeridaPastel(value) {
  const numericValue = Number(value) || 0;
  if (numericValue < 8) return 0;
  return 10 + Math.floor((numericValue - 8) / 5) * 5;
}

export function getProduccionSugerida(producto, value) {
  if (isOperationalCakeProduct(producto)) {
    return getProduccionSugeridaPastel(value);
  }
  return Math.max(0, Math.ceil(Number(value) || 0));
}

export function getReglaOperativaLabel(producto, value) {
  if (!isOperationalCakeProduct(producto)) return "Redondeo normal";
  const produccionSugerida = getProduccionSugeridaPastel(value);
  if (produccionSugerida === 0) return "Menor a 8: no producir";
  return `Mínimo 10 y múltiplos de 5: ${produccionSugerida}`;
}

export function weekendMultiplier(weekday, weekendBoost) {
  const boost = Number(weekendBoost);
  const factor = Number.isFinite(boost) && boost > 0 ? boost : 1;
  return weekday === 0 || weekday === 6 ? factor : 1;
}

export function mergeMonthlyFromDaily(productRows, dailyRows) {
  const byProduct = new Map();
  for (const row of dailyRows) {
    const current = byProduct.get(row.producto) || {
      pronosticoVenta: 0,
      colchonOperativo: 0,
      baseConColchon: 0,
      produccionSugerida: 0,
      produccionReal: 0,
      hasRealData: false,
    };
    current.pronosticoVenta += Number(row.pronosticoVentaDia) || 0;
    current.colchonOperativo += Number(row.colchonDiario) || 0;
    current.baseConColchon += Number(row.baseConColchonDia) || 0;
    current.produccionSugerida += Number(row.produccionSugeridaDia) || 0;
    if (row.hasRealData) {
      current.hasRealData = true;
      current.produccionReal += Number(row.produccionRealDia) || 0;
    }
    byProduct.set(row.producto, current);
  }

  return productRows.map((row) => {
    const monthly = byProduct.get(row.producto);
    if (!monthly) return row;

    const produccionSugerida = monthly.produccionSugerida;
    const produccionReal = monthly.produccionReal;
    const hasRealData = monthly.hasRealData;
    const diferenciaReal = produccionReal - produccionSugerida;
    const bajasEsperadas = monthly.pronosticoVenta * Number(row.tasaBajas || 0);
    const precision =
      hasRealData && produccionReal > 0
        ? (1 - Math.abs(produccionSugerida - produccionReal) / produccionReal) * 100
        : null;

    const baseProduccionRecomendada = Math.max(
      0,
      Number(row.inventarioObjetivo || 0) + produccionSugerida - Number(row.sumaSucCf || 0)
    );
    const produccionRecomendada = getProduccionSugerida(row.producto, baseProduccionRecomendada);

    let estatus = "Sin dato real";
    if (!hasRealData) estatus = "Sin dato real";
    else if (produccionSugerida === 0 && produccionReal === 0) estatus = "No producir";
    else if (produccionReal < produccionSugerida) estatus = "Riesgo faltante";
    else if (produccionReal > produccionSugerida) estatus = "Sobreproduccion";
    else if (precision !== null && precision < 80) estatus = "Revisar";
    else estatus = "Dentro de rango";

    return {
      ...row,
      demandaPronosticada: monthly.pronosticoVenta,
      pronosticoVenta: monthly.pronosticoVenta,
      bajasEsperadas,
      colchonOperativo: monthly.colchonOperativo,
      baseConColchon: monthly.baseConColchon,
      reglaOperativa: getReglaOperativaLabel(row.producto, monthly.baseConColchon),
      produccionSugerida,
      baseProduccionRecomendada,
      produccionRecomendada,
      produccionReal,
      hasRealData,
      diferenciaReal,
      precision,
      estatus,
    };
  });
}
