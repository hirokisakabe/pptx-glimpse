import { describe, it, expect } from "vitest";
import {
  DEFAULT_FONT_MAPPING,
  createFontMapping,
  getMappedFont,
  type FontMapping,
} from "./font-mapping.js";

describe("DEFAULT_FONT_MAPPING", () => {
  it("ラテン文字フォントのマッピングが定義されている", () => {
    expect(DEFAULT_FONT_MAPPING["Calibri"]).toBe("Carlito");
    expect(DEFAULT_FONT_MAPPING["Arial"]).toBe("Arimo");
    expect(DEFAULT_FONT_MAPPING["Times New Roman"]).toBe("Tinos");
    expect(DEFAULT_FONT_MAPPING["Courier New"]).toBe("Cousine");
    expect(DEFAULT_FONT_MAPPING["Cambria"]).toBe("Caladea");
  });

  it("日本語ゴシック系フォントが Noto Sans JP にマッピングされている", () => {
    expect(DEFAULT_FONT_MAPPING["メイリオ"]).toBe("Noto Sans JP");
    expect(DEFAULT_FONT_MAPPING["Meiryo"]).toBe("Noto Sans JP");
    expect(DEFAULT_FONT_MAPPING["游ゴシック"]).toBe("Noto Sans JP");
    expect(DEFAULT_FONT_MAPPING["Yu Gothic"]).toBe("Noto Sans JP");
    expect(DEFAULT_FONT_MAPPING["MS ゴシック"]).toBe("Noto Sans JP");
    expect(DEFAULT_FONT_MAPPING["MS Gothic"]).toBe("Noto Sans JP");
    expect(DEFAULT_FONT_MAPPING["MS Pゴシック"]).toBe("Noto Sans JP");
    expect(DEFAULT_FONT_MAPPING["MS PGothic"]).toBe("Noto Sans JP");
  });

  it("日本語明朝系フォントが Noto Serif JP にマッピングされている", () => {
    expect(DEFAULT_FONT_MAPPING["MS 明朝"]).toBe("Noto Serif JP");
    expect(DEFAULT_FONT_MAPPING["MS Mincho"]).toBe("Noto Serif JP");
    expect(DEFAULT_FONT_MAPPING["MS P明朝"]).toBe("Noto Serif JP");
    expect(DEFAULT_FONT_MAPPING["MS PMincho"]).toBe("Noto Serif JP");
    expect(DEFAULT_FONT_MAPPING["游明朝"]).toBe("Noto Serif JP");
    expect(DEFAULT_FONT_MAPPING["Yu Mincho"]).toBe("Noto Serif JP");
  });
});

describe("createFontMapping", () => {
  it("ユーザーマッピングなしでデフォルトのコピーを返す", () => {
    const mapping = createFontMapping();
    expect(mapping["Calibri"]).toBe("Carlito");
    expect(mapping["Arial"]).toBe("Arimo");
  });

  it("ユーザーマッピングでデフォルトを上書きできる", () => {
    const mapping = createFontMapping({ Calibri: "Custom Font" });
    expect(mapping["Calibri"]).toBe("Custom Font");
    // 他のデフォルトは維持される
    expect(mapping["Arial"]).toBe("Arimo");
  });

  it("ユーザーマッピングで新しいエントリを追加できる", () => {
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

  it("完全一致でマッピングを返す", () => {
    expect(getMappedFont("Calibri", mapping)).toBe("Carlito");
    expect(getMappedFont("Arial", mapping)).toBe("Arimo");
  });

  it("大文字小文字を無視してマッチする", () => {
    expect(getMappedFont("calibri", mapping)).toBe("Carlito");
    expect(getMappedFont("ARIAL", mapping)).toBe("Arimo");
    expect(getMappedFont("ms gothic", mapping)).toBe("Noto Sans JP");
  });

  it("マッピングに存在しないフォントは null を返す", () => {
    expect(getMappedFont("Unknown Font", mapping)).toBeNull();
  });

  it("null または undefined は null を返す", () => {
    expect(getMappedFont(null, mapping)).toBeNull();
    expect(getMappedFont(undefined, mapping)).toBeNull();
  });
});
