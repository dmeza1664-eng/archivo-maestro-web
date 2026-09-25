const {
  POST22_EXPECTED,
  TARGET_MONTHS,
  salesInputsExist,
  runPepes2025Backtest,
} = require("./run-backtest-2025-cold-start-2024.cjs");
const { runSyntheticChecks } = require("./synthetic-cold-start-check.cjs");

function assert(condition, message) {
  if (!condition) throw new Error(message);
}

function nearly(actual, expected, digits = 2) {
  if (actual == null || expected == null) return false;
  return Math.abs(Number(actual) - Number(expected)) < 0.5 * 10 ** -digits + 1e-9;
}

async function main() {
  // Siempre corre (CI incluido) con el fixture sintético versionado.
  const synthetic = await runSyntheticChecks();
  console.log(`post22-regression-test sintético ok (enero con 2024 ${synthetic.con2024["2025-01"]}%, sin 2024 ${synthetic.sin2024["2025-01"]}%)`);
  if (!salesInputsExist()) {
    console.log("post22-regression-test: ventas Pepes 2024-2025 no presentes; se omite solo la parte con datos reales");
    return;
  }

  const report = await runPepes2025Backtest();
  for (const [cut, expected] of Object.entries(POST22_EXPECTED.cuts)) {
    const actual = report.cutsSin2024[cut]?.weightedWapePct;
    assert(
      nearly(actual, expected),
      `sin 2024 el corte ${cut} debe ser ${expected} (a050fa9), salió ${actual}`
    );
  }
  for (const month of TARGET_MONTHS) {
    const expected = POST22_EXPECTED.monthly[month];
    const actual = report.monthlySin2024[month]?.wape;
    assert(
      nearly(actual, expected),
      `sin 2024 ${month} debe ser ${expected} (a050fa9), salió ${actual}`
    );
  }

  for (const month of TARGET_MONTHS) {
    if (month === "2025-01") continue;
    const delta = report.monthlyCon2024[month]?.deltaPts;
    if (delta == null) continue;
    assert(
      delta <= 0.3 + 1e-9,
      `con 2024 ${month} no debe empeorar más de 0.3 pts vs post22 (Δ ${delta})`
    );
  }

  assert(
    report.gapReactivationSin2024.length === 0,
    `sin 2024 no debe haber reactivación de hueco (${report.gapReactivationSin2024.length})`
  );

  console.log("post22-regression-test ok");
  console.log(JSON.stringify({
    cutsSin2024: Object.fromEntries(
      Object.entries(report.cutsSin2024).map(([k, v]) => [k, v.weightedWapePct])
    ),
    cutsCon2024: Object.fromEntries(
      Object.entries(report.cutsCon2024).map(([k, v]) => [k, v.weightedWapePct])
    ),
    gapReactivations: report.gapReactivationCon2024.length,
  }));
}

main().catch((error) => {
  console.error(error);
  process.exit(1);
});
