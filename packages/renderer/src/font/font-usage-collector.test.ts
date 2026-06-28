import { afterEach, describe, expect, it } from "vitest";

import {
  FontUsageCollector,
  getFontUsageCollector,
  resetFontUsageCollector,
  setFontUsageCollector,
} from "./font-usage-collector.js";

describe("FontUsageCollector", () => {
  it("Collect characters using the beginning of the font name as a key", () => {
    const collector = new FontUsageCollector();
    collector.record(["Arial", null], "AB");
    collector.record(["Arial", "Noto Sans JP"], "BC");

    const usages = collector.getUsages();
    expect(usages.size).toBe(1);
    const usage = usages.get("Arial")!;
    expect([...usage.chars].sort()).toEqual(["A", "B", "C"]);
    // The first recorded fonts list is retained.
    expect(usage.fonts).toEqual(["Arial", null]);
  });

  it("Skip leading null and use first non-null font name as key", () => {
    const collector = new FontUsageCollector();
    collector.record([null, "Noto Sans JP"], "あ");

    expect(collector.getUsages().has("Noto Sans JP")).toBe(true);
  });

  it("Do not record if all font names are null", () => {
    const collector = new FontUsageCollector();
    collector.record([null, null], "A");

    expect(collector.getUsages().size).toBe(0);
  });

  it("Do not record empty strings", () => {
    const collector = new FontUsageCollector();
    collector.record(["Arial"], "");

    expect(collector.getUsages().size).toBe(0);
  });

  it("Collect surrogate pairs as one character", () => {
    const collector = new FontUsageCollector();
    collector.record(["Test"], "𠮷野");

    const usage = collector.getUsages().get("Test")!;
    expect(usage.chars.has("𠮷")).toBe(true);
    expect(usage.chars.has("野")).toBe(true);
    expect(usage.chars.size).toBe(2);
  });

  it("Clear the collected contents with reset", () => {
    const collector = new FontUsageCollector();
    collector.record(["Arial"], "A");
    collector.reset();

    expect(collector.getUsages().size).toBe(0);
  });
});

describe("font usage collector context", () => {
  afterEach(() => {
    resetFontUsageCollector();
  });

  it("Context can be managed with set / get / reset", () => {
    expect(getFontUsageCollector()).toBeNull();

    const collector = new FontUsageCollector();
    setFontUsageCollector(collector);
    expect(getFontUsageCollector()).toBe(collector);

    resetFontUsageCollector();
    expect(getFontUsageCollector()).toBeNull();
  });
});
