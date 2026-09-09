import assert from "node:assert/strict";
import test from "node:test";
import {
  getProduccionSugerida,
  mergeMonthlyFromDaily,
  weekendMultiplier,
} from "./forecastCore.js";

test("pasteles menores a 8 no se producen", () => {
  assert.equal(getProduccionSugerida("PINA GDE", 7.9), 0);
});

test("pasteles usan minimo 10 y multiplos de 5", () => {
  assert.equal(getProduccionSugerida("CHOC MED", 10), 10);
  assert.equal(getProduccionSugerida("FRESA CH", 12.9), 10);
  assert.equal(getProduccionSugerida("TRES LECHES GRANDE", 14), 15);
});

test("productos que no son pastel se redondean hacia arriba", () => {
  assert.equal(getProduccionSugerida("NAPOLITANO", 12.1), 13);
});

test("el factor de fin de semana solo aplica sabado y domingo", () => {
  assert.equal(weekendMultiplier(1, 1.15), 1);
  assert.equal(weekendMultiplier(6, 1.15), 1.15);
  assert.equal(weekendMultiplier(0, 1.2), 1.2);
});

test("el mensual es la suma de los dias", () => {
  const monthly = mergeMonthlyFromDaily(
    [
      {
        producto: "PINA GDE",
        inventarioObjetivo: 0,
        sumaSucCf: 0,
        confianza: 90,
      },
    ],
    [
      {
        producto: "PINA GDE",
        pronosticoVentaDia: 10,
        colchonDiario: 1,
        baseConColchonDia: 11,
        produccionSugeridaDia: 10,
        hasRealData: true,
        produccionRealDia: 12,
      },
      {
        producto: "PINA GDE",
        pronosticoVentaDia: 8,
        colchonDiario: 0.8,
        baseConColchonDia: 8.8,
        produccionSugeridaDia: 10,
        hasRealData: true,
        produccionRealDia: 8,
      },
    ]
  );

  assert.equal(monthly[0].pronosticoVenta, 18);
  assert.equal(monthly[0].produccionSugerida, 20);
  assert.equal(monthly[0].produccionReal, 20);
  assert.equal(monthly[0].hasRealData, true);
});
