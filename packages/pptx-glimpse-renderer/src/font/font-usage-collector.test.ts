import { afterEach, describe, expect, it } from "vitest";

import {
  FontUsageCollector,
  getFontUsageCollector,
  resetFontUsageCollector,
  setFontUsageCollector,
} from "./font-usage-collector.js";

describe("FontUsageCollector", () => {
  it("フォント名先頭をキーに文字を収集する", () => {
    const collector = new FontUsageCollector();
    collector.record(["Arial", null], "AB");
    collector.record(["Arial", "Noto Sans JP"], "BC");

    const usages = collector.getUsages();
    expect(usages.size).toBe(1);
    const usage = usages.get("Arial")!;
    expect([...usage.chars].sort()).toEqual(["A", "B", "C"]);
    // 最初に記録された fonts リストが保持される
    expect(usage.fonts).toEqual(["Arial", null]);
  });

  it("先頭の null をスキップして最初の非 null フォント名をキーにする", () => {
    const collector = new FontUsageCollector();
    collector.record([null, "Noto Sans JP"], "あ");

    expect(collector.getUsages().has("Noto Sans JP")).toBe(true);
  });

  it("フォント名がすべて null の場合は記録しない", () => {
    const collector = new FontUsageCollector();
    collector.record([null, null], "A");

    expect(collector.getUsages().size).toBe(0);
  });

  it("空文字列は記録しない", () => {
    const collector = new FontUsageCollector();
    collector.record(["Arial"], "");

    expect(collector.getUsages().size).toBe(0);
  });

  it("サロゲートペアを 1 文字として収集する", () => {
    const collector = new FontUsageCollector();
    collector.record(["Test"], "𠮷野");

    const usage = collector.getUsages().get("Test")!;
    expect(usage.chars.has("𠮷")).toBe(true);
    expect(usage.chars.has("野")).toBe(true);
    expect(usage.chars.size).toBe(2);
  });

  it("reset で収集内容をクリアする", () => {
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

  it("set / get / reset でコンテキストを管理できる", () => {
    expect(getFontUsageCollector()).toBeNull();

    const collector = new FontUsageCollector();
    setFontUsageCollector(collector);
    expect(getFontUsageCollector()).toBe(collector);

    resetFontUsageCollector();
    expect(getFontUsageCollector()).toBeNull();
  });
});
