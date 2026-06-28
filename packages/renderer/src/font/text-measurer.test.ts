import { afterEach, describe, expect, it } from "vitest";

import { getAscenderRatio, getLineHeightRatio, measureTextWidth } from "../utils/text-measure.js";
import {
  DefaultTextMeasurer,
  getTextMeasurer,
  resetTextMeasurer,
  setTextMeasurer,
  type TextMeasurer,
} from "./text-measurer.js";

afterEach(() => {
  resetTextMeasurer();
});

describe("DefaultTextMeasurer", () => {
  it("covers text-measurer behavior 1", () => {
    const measurer = new DefaultTextMeasurer();
    expect(measurer.measureTextWidth("Hello", 18, false, "Calibri")).toBe(
      measureTextWidth("Hello", 18, false, "Calibri"),
    );
  });

  it("covers text-measurer behavior 2", () => {
    const measurer = new DefaultTextMeasurer();
    expect(measurer.getLineHeightRatio("Calibri")).toBe(getLineHeightRatio("Calibri"));
  });

  it("covers text-measurer behavior 3", () => {
    const measurer = new DefaultTextMeasurer();
    expect(measurer.getAscenderRatio("Calibri")).toBe(getAscenderRatio("Calibri"));
  });
});

describe("setTextMeasurer / getTextMeasurer / resetTextMeasurer", () => {
  it("covers text-measurer behavior 4", () => {
    expect(getTextMeasurer()).toBeInstanceOf(DefaultTextMeasurer);
  });

  it("covers text-measurer behavior 5", () => {
    const custom: TextMeasurer = {
      measureTextWidth: () => 42,
      getLineHeightRatio: () => 1.5,
      getAscenderRatio: () => 0.9,
    };
    setTextMeasurer(custom);
    expect(getTextMeasurer().measureTextWidth("x", 12, false)).toBe(42);
    expect(getTextMeasurer().getLineHeightRatio()).toBe(1.5);
    expect(getTextMeasurer().getAscenderRatio()).toBe(0.9);
  });

  it("covers text-measurer behavior 6", () => {
    const custom: TextMeasurer = {
      measureTextWidth: () => 42,
      getLineHeightRatio: () => 1.5,
      getAscenderRatio: () => 0.9,
    };
    setTextMeasurer(custom);
    resetTextMeasurer();
    expect(getTextMeasurer()).toBeInstanceOf(DefaultTextMeasurer);
  });
});
