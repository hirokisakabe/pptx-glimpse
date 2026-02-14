import { describe, it, expect } from "vitest";
import { evaluateFormula, evaluateGuides, resolveValue } from "./geometry-formula.js";

describe("evaluateFormula", () => {
  it("handles val", () => {
    expect(evaluateFormula("val 50000", {})).toBe(50000);
  });

  it("handles +- (a + b - c)", () => {
    expect(evaluateFormula("+- 100 50 30", {})).toBe(120);
  });

  it("handles */ (a * b / c)", () => {
    expect(evaluateFormula("*/ 1000 50000 100000", {})).toBe(500);
  });

  it("handles +/ ((a + b) / c)", () => {
    expect(evaluateFormula("+/ 100 200 2", {})).toBe(150);
  });

  it("handles sin with 60000th degree", () => {
    // sin(90°) = 1, 90° = 5400000 in 60000ths
    expect(evaluateFormula("sin 10000 5400000", {})).toBe(10000);
  });

  it("handles cos with 60000th degree", () => {
    // cos(0°) = 1
    expect(evaluateFormula("cos 10000 0", {})).toBe(10000);
  });

  it("handles at2 (atan2)", () => {
    // atan2(1, 0) = 90° = 5400000 in 60000ths
    const result = evaluateFormula("at2 0 1", {});
    expect(result).toBe(5400000);
  });

  it("handles sqrt", () => {
    expect(evaluateFormula("sqrt 10000", {})).toBe(100);
  });

  it("handles min", () => {
    expect(evaluateFormula("min 100 200", {})).toBe(100);
  });

  it("handles max", () => {
    expect(evaluateFormula("max 100 200", {})).toBe(200);
  });

  it("handles abs", () => {
    expect(evaluateFormula("abs -100", {})).toBe(100);
  });

  it("handles pin (clamp)", () => {
    expect(evaluateFormula("pin 10 5 20", {})).toBe(10);
    expect(evaluateFormula("pin 10 15 20", {})).toBe(15);
    expect(evaluateFormula("pin 10 25 20", {})).toBe(20);
  });

  it("handles mod (modulus = sqrt(a² + b² + c²))", () => {
    expect(evaluateFormula("mod 3 4 0", {})).toBe(5);
  });

  it("handles ?: (conditional)", () => {
    expect(evaluateFormula("?: 1 100 200", {})).toBe(100);
    expect(evaluateFormula("?: 0 100 200", {})).toBe(200);
    expect(evaluateFormula("?: -1 100 200", {})).toBe(200);
  });

  it("resolves guide names from vars", () => {
    const vars = { adj: 50000, w: 1000 };
    expect(evaluateFormula("*/ w adj 100000", vars)).toBe(500);
  });

  it("handles division by zero gracefully (falls back to divisor=1)", () => {
    expect(evaluateFormula("*/ 100 200 0", {})).toBe(20000);
  });

  it("returns 0 for unknown operators", () => {
    expect(evaluateFormula("unknown 1 2 3", {})).toBe(0);
  });
});

describe("evaluateGuides", () => {
  it("evaluates avLst then gdLst in order", () => {
    const avLst = [{ name: "adj", fmla: "val 50000" }];
    const gdLst = [{ name: "midX", fmla: "*/ w adj 100000" }];
    const result = evaluateGuides(avLst, gdLst, 1000, 500);
    expect(result["adj"]).toBe(50000);
    expect(result["midX"]).toBe(500);
  });

  it("makes builtin variables available", () => {
    const result = evaluateGuides([], [], 1000, 500);
    expect(result["w"]).toBe(1000);
    expect(result["h"]).toBe(500);
    expect(result["wd2"]).toBe(500);
    expect(result["hd2"]).toBe(250);
    expect(result["l"]).toBe(0);
    expect(result["t"]).toBe(0);
    expect(result["r"]).toBe(1000);
    expect(result["b"]).toBe(500);
    expect(result["cd4"]).toBe(5400000);
  });

  it("allows gdLst to reference avLst values", () => {
    const avLst = [{ name: "adj1", fmla: "val 25000" }];
    const gdLst = [
      { name: "x1", fmla: "*/ w adj1 100000" },
      { name: "x2", fmla: "+- w 0 x1" },
    ];
    const result = evaluateGuides(avLst, gdLst, 1000, 1000);
    expect(result["x1"]).toBe(250);
    expect(result["x2"]).toBe(750);
  });
});

describe("resolveValue", () => {
  it("resolves numeric string", () => {
    expect(resolveValue("500", {})).toBe(500);
  });

  it("resolves number", () => {
    expect(resolveValue(42, {})).toBe(42);
  });

  it("resolves guide name", () => {
    expect(resolveValue("midX", { midX: 300 })).toBe(300);
  });

  it("returns 0 for unknown name", () => {
    expect(resolveValue("unknown", {})).toBe(0);
  });
});
