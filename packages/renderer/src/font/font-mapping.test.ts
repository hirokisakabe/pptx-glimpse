import { describe, expect, it } from "vitest";

import {
  createFontMapping,
  DEFAULT_FONT_MAPPING,
  type FontMapping,
  getMappedFont,
} from "./font-mapping.js";

describe("DEFAULT_FONT_MAPPING", () => {
  it("covers font-mapping behavior 1", () => {
    expect(DEFAULT_FONT_MAPPING["Calibri"]).toBe("Carlito");
    expect(DEFAULT_FONT_MAPPING["Arial"]).toBe("Arimo");
    expect(DEFAULT_FONT_MAPPING["Times New Roman"]).toBe("Tinos");
    expect(DEFAULT_FONT_MAPPING["Courier New"]).toBe("Cousine");
    expect(DEFAULT_FONT_MAPPING["Cambria"]).toBe("Caladea");
  });

  it("covers font-mapping behavior 2", () => {
    expect(DEFAULT_FONT_MAPPING["メイリオ"]).toBe("Noto Sans JP");
    expect(DEFAULT_FONT_MAPPING["Meiryo"]).toBe("Noto Sans JP");
    expect(DEFAULT_FONT_MAPPING["游ゴシック"]).toBe("Noto Sans JP");
    expect(DEFAULT_FONT_MAPPING["Yu Gothic"]).toBe("Noto Sans JP");
    expect(DEFAULT_FONT_MAPPING["MS ゴシック"]).toBe("Noto Sans JP");
    expect(DEFAULT_FONT_MAPPING["MS Gothic"]).toBe("Noto Sans JP");
    expect(DEFAULT_FONT_MAPPING["MS Pゴシック"]).toBe("Noto Sans JP");
    expect(DEFAULT_FONT_MAPPING["MS PGothic"]).toBe("Noto Sans JP");
  });

  it("covers font-mapping behavior 3", () => {
    expect(DEFAULT_FONT_MAPPING["MS 明朝"]).toBe("Noto Serif CJK JP");
    expect(DEFAULT_FONT_MAPPING["MS Mincho"]).toBe("Noto Serif CJK JP");
    expect(DEFAULT_FONT_MAPPING["MS P明朝"]).toBe("Noto Serif CJK JP");
    expect(DEFAULT_FONT_MAPPING["MS PMincho"]).toBe("Noto Serif CJK JP");
    expect(DEFAULT_FONT_MAPPING["游明朝"]).toBe("Noto Serif CJK JP");
    expect(DEFAULT_FONT_MAPPING["Yu Mincho"]).toBe("Noto Serif CJK JP");
  });
});

describe("createFontMapping", () => {
  it("covers font-mapping behavior 4", () => {
    const mapping = createFontMapping();
    expect(mapping["Calibri"]).toBe("Carlito");
    expect(mapping["Arial"]).toBe("Arimo");
  });

  it("covers font-mapping behavior 5", () => {
    const mapping = createFontMapping({ Calibri: "Custom Font" });
    expect(mapping["Calibri"]).toBe("Custom Font");
    // Test note.
    expect(mapping["Arial"]).toBe("Arimo");
  });

  it("covers font-mapping behavior 6", () => {
    const mapping = createFontMapping({ "My Custom Font": "Noto Sans" });
    expect(mapping["My Custom Font"]).toBe("Noto Sans");
    expect(mapping["Calibri"]).toBe("Carlito");
  });
});

describe("getMappedFont", () => {
  const mapping: FontMapping = {
    Calibri: "Carlito",
    Arial: "Arimo",
    "MS Gothic": "Noto Sans JP",
  };

  it("covers font-mapping behavior 7", () => {
    expect(getMappedFont("Calibri", mapping)).toBe("Carlito");
    expect(getMappedFont("Arial", mapping)).toBe("Arimo");
  });

  it("covers font-mapping behavior 8", () => {
    expect(getMappedFont("calibri", mapping)).toBe("Carlito");
    expect(getMappedFont("ARIAL", mapping)).toBe("Arimo");
    expect(getMappedFont("ms gothic", mapping)).toBe("Noto Sans JP");
  });

  it("covers font-mapping behavior 9", () => {
    expect(getMappedFont("Unknown Font", mapping)).toBeNull();
  });

  it("covers font-mapping behavior 10", () => {
    expect(getMappedFont(null, mapping)).toBeNull();
    expect(getMappedFont(undefined, mapping)).toBeNull();
  });

  it("covers font-mapping behavior 11", () => {
    const fullMapping = createFontMapping();
    // Test note.
    expect(getMappedFont("ＭＳ Ｐゴシック", fullMapping)).toBe("Noto Sans JP");
    // Test note.
    expect(getMappedFont("ＭＳ Ｐ明朝", fullMapping)).toBe("Noto Serif CJK JP");
    // Test note.
    expect(getMappedFont("ＭＳ\u3000Ｐゴシック", fullMapping)).toBe("Noto Sans JP");
  });
});
