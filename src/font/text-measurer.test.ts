import { describe, it, expect, afterEach } from "vitest";
import {
  DefaultTextMeasurer,
  setTextMeasurer,
  getTextMeasurer,
  resetTextMeasurer,
  type TextMeasurer,
} from "./text-measurer.js";
import { measureTextWidth, getLineHeightRatio } from "../utils/text-measure.js";

afterEach(() => {
  resetTextMeasurer();
});

describe("DefaultTextMeasurer", () => {
  it("measureTextWidth は既存関数と同じ結果を返す", () => {
    const measurer = new DefaultTextMeasurer();
    expect(measurer.measureTextWidth("Hello", 18, false, "Calibri")).toBe(
      measureTextWidth("Hello", 18, false, "Calibri"),
    );
  });

  it("getLineHeightRatio は既存関数と同じ結果を返す", () => {
    const measurer = new DefaultTextMeasurer();
    expect(measurer.getLineHeightRatio("Calibri")).toBe(getLineHeightRatio("Calibri"));
  });
});

describe("setTextMeasurer / getTextMeasurer / resetTextMeasurer", () => {
  it("デフォルトでは DefaultTextMeasurer が使われる", () => {
    expect(getTextMeasurer()).toBeInstanceOf(DefaultTextMeasurer);
  });

  it("setTextMeasurer でカスタム実装に差し替えられる", () => {
    const custom: TextMeasurer = {
      measureTextWidth: () => 42,
      getLineHeightRatio: () => 1.5,
    };
    setTextMeasurer(custom);
    expect(getTextMeasurer().measureTextWidth("x", 12, false)).toBe(42);
    expect(getTextMeasurer().getLineHeightRatio()).toBe(1.5);
  });

  it("resetTextMeasurer でデフォルトに戻る", () => {
    const custom: TextMeasurer = {
      measureTextWidth: () => 42,
      getLineHeightRatio: () => 1.5,
    };
    setTextMeasurer(custom);
    resetTextMeasurer();
    expect(getTextMeasurer()).toBeInstanceOf(DefaultTextMeasurer);
  });
});
