import { afterEach, describe, expect, it } from "vitest";

import {
  FontUsageCollector,
  getFontUsageCollector,
  resetFontUsageCollector,
  setFontUsageCollector,
} from "./font-usage-collector.js";

describe("FontUsageCollector", () => {
  it("covers font-usage-collector behavior 1", () => {
    const collector = new FontUsageCollector();
    collector.record(["Arial", null], "AB");
    collector.record(["Arial", "Noto Sans JP"], "BC");

    const usages = collector.getUsages();
    expect(usages.size).toBe(1);
    const usage = usages.get("Arial")!;
    expect([...usage.chars].sort()).toEqual(["A", "B", "C"]);
    // Test note.
    expect(usage.fonts).toEqual(["Arial", null]);
  });

  it("covers font-usage-collector behavior 2", () => {
    const collector = new FontUsageCollector();
    collector.record([null, "Noto Sans JP"], "あ");

    expect(collector.getUsages().has("Noto Sans JP")).toBe(true);
  });

  it("covers font-usage-collector behavior 3", () => {
    const collector = new FontUsageCollector();
    collector.record([null, null], "A");

    expect(collector.getUsages().size).toBe(0);
  });

  it("covers font-usage-collector behavior 4", () => {
    const collector = new FontUsageCollector();
    collector.record(["Arial"], "");

    expect(collector.getUsages().size).toBe(0);
  });

  it("covers font-usage-collector behavior 5", () => {
    const collector = new FontUsageCollector();
    collector.record(["Test"], "𠮷野");

    const usage = collector.getUsages().get("Test")!;
    expect(usage.chars.has("𠮷")).toBe(true);
    expect(usage.chars.has("野")).toBe(true);
    expect(usage.chars.size).toBe(2);
  });

  it("covers font-usage-collector behavior 6", () => {
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

  it("covers font-usage-collector behavior 7", () => {
    expect(getFontUsageCollector()).toBeNull();

    const collector = new FontUsageCollector();
    setFontUsageCollector(collector);
    expect(getFontUsageCollector()).toBe(collector);

    resetFontUsageCollector();
    expect(getFontUsageCollector()).toBeNull();
  });
});
