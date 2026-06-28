import { describe, expect, it } from "vitest";

import {
  createFontMapping,
  DEFAULT_FONT_MAPPING,
  type FontMapping,
  getMappedFont,
} from "./font-mapping.js";

describe("DEFAULT_FONT_MAPPING", () => {
  it("Latin font mapping is defined", () => {
    expect(DEFAULT_FONT_MAPPING["Calibri"]).toBe("Carlito");
    expect(DEFAULT_FONT_MAPPING["Arial"]).toBe("Arimo");
    expect(DEFAULT_FONT_MAPPING["Times New Roman"]).toBe("Tinos");
    expect(DEFAULT_FONT_MAPPING["Courier New"]).toBe("Cousine");
    expect(DEFAULT_FONT_MAPPING["Cambria"]).toBe("Caladea");
  });

  it("Japanese Gothic font is mapped to Noto Sans JP", () => {
    expect(DEFAULT_FONT_MAPPING["メイリオ"]).toBe("Noto Sans JP");
    expect(DEFAULT_FONT_MAPPING["Meiryo"]).toBe("Noto Sans JP");
    expect(DEFAULT_FONT_MAPPING["游ゴシック"]).toBe("Noto Sans JP");
    expect(DEFAULT_FONT_MAPPING["Yu Gothic"]).toBe("Noto Sans JP");
    expect(DEFAULT_FONT_MAPPING["MS ゴシック"]).toBe("Noto Sans JP");
    expect(DEFAULT_FONT_MAPPING["MS Gothic"]).toBe("Noto Sans JP");
    expect(DEFAULT_FONT_MAPPING["MS Pゴシック"]).toBe("Noto Sans JP");
    expect(DEFAULT_FONT_MAPPING["MS PGothic"]).toBe("Noto Sans JP");
  });

  it("Japanese Mincho fonts are mapped to Noto Serif CJK JP", () => {
    expect(DEFAULT_FONT_MAPPING["MS 明朝"]).toBe("Noto Serif CJK JP");
    expect(DEFAULT_FONT_MAPPING["MS Mincho"]).toBe("Noto Serif CJK JP");
    expect(DEFAULT_FONT_MAPPING["MS P明朝"]).toBe("Noto Serif CJK JP");
    expect(DEFAULT_FONT_MAPPING["MS PMincho"]).toBe("Noto Serif CJK JP");
    expect(DEFAULT_FONT_MAPPING["游明朝"]).toBe("Noto Serif CJK JP");
    expect(DEFAULT_FONT_MAPPING["Yu Mincho"]).toBe("Noto Serif CJK JP");
  });
});

describe("createFontMapping", () => {
  it("Return default copy without user mapping", () => {
    const mapping = createFontMapping();
    expect(mapping["Calibri"]).toBe("Carlito");
    expect(mapping["Arial"]).toBe("Arimo");
  });

  it("User mapping can override defaults", () => {
    const mapping = createFontMapping({ Calibri: "Custom Font" });
    expect(mapping["Calibri"]).toBe("Custom Font");
    // other defaults are retained
    expect(mapping["Arial"]).toBe("Arimo");
  });

  it("New entries can be added with user mapping", () => {
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

  it("Return mapping with exact match", () => {
    expect(getMappedFont("Calibri", mapping)).toBe("Carlito");
    expect(getMappedFont("Arial", mapping)).toBe("Arimo");
  });

  it("Match ignoring case", () => {
    expect(getMappedFont("calibri", mapping)).toBe("Carlito");
    expect(getMappedFont("ARIAL", mapping)).toBe("Arimo");
    expect(getMappedFont("ms gothic", mapping)).toBe("Noto Sans JP");
  });

  it("Fonts not present in the mapping return null", () => {
    expect(getMappedFont("Unknown Font", mapping)).toBeNull();
  });

  it("null or undefined returns null", () => {
    expect(getMappedFont(null, mapping)).toBeNull();
    expect(getMappedFont(undefined, mapping)).toBeNull();
  });

  it("Normalize full-width alphanumeric characters to half-width and match", () => {
    const fullMapping = createFontMapping();
    // MS P Gothic (full-width P) -> Match MS P Gothic (half-width P)
    expect(getMappedFont("ＭＳ Ｐゴシック", fullMapping)).toBe("Noto Sans JP");
    // MS P Mincho (full-width P) -> Match MS P Mincho (half-width P)
    expect(getMappedFont("ＭＳ Ｐ明朝", fullMapping)).toBe("Noto Serif CJK JP");
    // Full-width spaces (\u3000) are also normalized to half-width spaces.
    expect(getMappedFont("ＭＳ\u3000Ｐゴシック", fullMapping)).toBe("Noto Sans JP");
  });
});
