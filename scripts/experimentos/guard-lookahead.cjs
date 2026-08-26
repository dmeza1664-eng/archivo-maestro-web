module.exports = function guardLookahead(scriptName) {
  if (process.env.PERMITIR_LOOKAHEAD === "1") return;

  console.error(
    [
      `Bloqueado: ${scriptName} calcula el pronostico a partir de la venta real del mismo mes.`,
      "",
      'Viola la regla 1 de CONTROL_MODELO_PRONOSTICO.md: "No usar ventas reales del mes objetivo',
      'para calcular su pronostico".',
      "",
      "El resultado no es un pronostico y no debe presentarse como tal.",
      "Para el comparativo real usa: node build-presentation-validado.cjs",
      "",
      `Solo para experimentos: PERMITIR_LOOKAHEAD=1 node ${scriptName}.cjs`,
    ].join("\n")
  );
  process.exit(1);
};
