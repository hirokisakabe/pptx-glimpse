import { afterEach, describe, expect, it, vi } from "vitest";

import {
  clearFontCache,
  createOpentypeSetupFromBuffers,
  createOpentypeSetupFromSystem,
  createOpentypeTextMeasurerFromBuffers,
} from "./opentype-helpers.js";
import { buildTtcFromTtfs } from "./ttc-test-helper.js";

/**
 * opentype.js を使って最小限の有効な TTF バッファを作成する。
 */
async function createTestFontBuffer(familyName: string): Promise<ArrayBuffer> {
  // eslint-disable-next-line @typescript-eslint/no-unsafe-assignment
  const opentype: {
    Glyph: new (opts: Record<string, unknown>) => unknown;
    Path: new () => {
      moveTo(x: number, y: number): void;
      lineTo(x: number, y: number): void;
      close(): void;
    };
    Font: new (opts: Record<string, unknown>) => { toArrayBuffer(): ArrayBuffer };
  } = await import("opentype.js");

  const notdefGlyph = new opentype.Glyph({
    name: ".notdef",
    unicode: 0,
    advanceWidth: 650,
    path: new opentype.Path(),
  });

  const pathA = new opentype.Path();
  pathA.moveTo(0, 0);
  pathA.lineTo(300, 800);
  pathA.lineTo(600, 0);
  pathA.close();

  const glyphA = new opentype.Glyph({
    name: "A",
    unicode: 65,
    advanceWidth: 600,
    path: pathA,
  });

  const spaceGlyph = new opentype.Glyph({
    name: "space",
    unicode: 32,
    advanceWidth: 250,
    path: new opentype.Path(),
  });

  const font = new opentype.Font({
    familyName,
    styleName: "Regular",
    unitsPerEm: 1000,
    ascender: 800,
    descender: -200,
    glyphs: [notdefGlyph, spaceGlyph, glyphA],
  });

  return font.toArrayBuffer();
}

describe("createOpentypeTextMeasurerFromBuffers", () => {
  it("空配列で null を返す", async () => {
    const result = await createOpentypeTextMeasurerFromBuffers([]);
    expect(result).toBeNull();
  });

  it("不正バッファで null を返す", async () => {
    const result = await createOpentypeTextMeasurerFromBuffers([
      { name: "Invalid", data: new Uint8Array([0, 0, 0, 0]) },
    ]);
    expect(result).toBeNull();
  });

  it("フォントバッファから OpentypeTextMeasurer を構築する", async () => {
    const buffer = await createTestFontBuffer("TestFont");
    const measurer = await createOpentypeTextMeasurerFromBuffers([
      { name: "TestFont", data: buffer },
    ]);
    expect(measurer).not.toBeNull();

    // "A" の幅を計測（advanceWidth=600, unitsPerEm=1000）
    const width = measurer!.measureTextWidth("A", 18, false, "TestFont");
    const expected = (600 / 1000) * 18 * (96 / 72);
    expect(width).toBeCloseTo(expected, 1);
  });

  it("行高さを正しく計算する", async () => {
    const buffer = await createTestFontBuffer("TestFont");
    const measurer = await createOpentypeTextMeasurerFromBuffers([
      { name: "TestFont", data: buffer },
    ]);
    expect(measurer).not.toBeNull();

    // ascender=800, descender=-200 → (800+200)/1000 = 1.0
    const ratio = measurer!.getLineHeightRatio("TestFont");
    expect(ratio).toBeCloseTo(1.0, 5);
  });

  it("フォントマッピング逆引きで PPTX 名でも解決できる", async () => {
    const buffer = await createTestFontBuffer("Carlito");
    const measurer = await createOpentypeTextMeasurerFromBuffers([
      { name: "Carlito", data: buffer },
    ]);
    expect(measurer).not.toBeNull();

    // デフォルトマッピング: Calibri → Carlito
    // 逆引きで "Calibri" でも Carlito のフォントが使える
    const widthByOss = measurer!.measureTextWidth("A", 18, false, "Carlito");
    const widthByPptx = measurer!.measureTextWidth("A", 18, false, "Calibri");
    expect(widthByPptx).toBe(widthByOss);
  });

  it("カスタムフォントマッピングで逆引きが機能する", async () => {
    const buffer = await createTestFontBuffer("MyFont");
    const measurer = await createOpentypeTextMeasurerFromBuffers(
      [{ name: "MyFont", data: buffer }],
      { "Custom Name": "MyFont" },
    );
    expect(measurer).not.toBeNull();

    const widthByOss = measurer!.measureTextWidth("A", 18, false, "MyFont");
    const widthByCustom = measurer!.measureTextWidth("A", 18, false, "Custom Name");
    expect(widthByCustom).toBe(widthByOss);
  });

  it("name なしのバッファは defaultFont として使われる", async () => {
    const buffer = await createTestFontBuffer("Unnamed");
    const measurer = await createOpentypeTextMeasurerFromBuffers([{ data: buffer }]);
    expect(measurer).not.toBeNull();

    // 未知のフォント名でも defaultFont が使われるため opentype ベースの計算になる
    const width = measurer!.measureTextWidth("A", 18, false, "UnknownFont");
    const expected = (600 / 1000) * 18 * (96 / 72);
    expect(width).toBeCloseTo(expected, 1);
  });

  it("Uint8Array バッファでも動作する", async () => {
    const arrayBuffer = await createTestFontBuffer("TestFont");
    const uint8 = new Uint8Array(arrayBuffer);
    const measurer = await createOpentypeTextMeasurerFromBuffers([
      { name: "TestFont", data: uint8 },
    ]);
    expect(measurer).not.toBeNull();

    const width = measurer!.measureTextWidth("A", 18, false, "TestFont");
    expect(width).toBeGreaterThan(0);
  });

  it("複数フォントを登録できる", async () => {
    const buffer1 = await createTestFontBuffer("Font1");
    const buffer2 = await createTestFontBuffer("Font2");
    const measurer = await createOpentypeTextMeasurerFromBuffers([
      { name: "Font1", data: buffer1 },
      { name: "Font2", data: buffer2 },
    ]);
    expect(measurer).not.toBeNull();

    const width1 = measurer!.measureTextWidth("A", 18, false, "Font1");
    const width2 = measurer!.measureTextWidth("A", 18, false, "Font2");
    expect(width1).toBeGreaterThan(0);
    expect(width2).toBeGreaterThan(0);
  });
});

describe("createOpentypeSetupFromBuffers (TTC)", () => {
  it("TTC バッファから OpentypeTextMeasurer を構築できる", async () => {
    const ttf1 = await createTestFontBuffer("FontAlpha");
    const ttf2 = await createTestFontBuffer("FontBeta");
    const ttc = buildTtcFromTtfs([ttf1, ttf2]);

    const measurer = await createOpentypeTextMeasurerFromBuffers([{ data: ttc }]);
    expect(measurer).not.toBeNull();
  });

  it("TTC 内の各フォント名で幅計測ができる", async () => {
    const ttf1 = await createTestFontBuffer("FontAlpha");
    const ttf2 = await createTestFontBuffer("FontBeta");
    const ttc = buildTtcFromTtfs([ttf1, ttf2]);

    const measurer = await createOpentypeTextMeasurerFromBuffers([{ data: ttc }]);
    expect(measurer).not.toBeNull();

    const widthAlpha = measurer!.measureTextWidth("A", 18, false, "FontAlpha");
    const widthBeta = measurer!.measureTextWidth("A", 18, false, "FontBeta");
    expect(widthAlpha).toBeGreaterThan(0);
    expect(widthBeta).toBeGreaterThan(0);
  });

  it("TTC と通常 TTF を混在して渡せる", async () => {
    const ttfSingle = await createTestFontBuffer("SingleFont");
    const ttf1 = await createTestFontBuffer("TtcFont1");
    const ttf2 = await createTestFontBuffer("TtcFont2");
    const ttc = buildTtcFromTtfs([ttf1, ttf2]);

    const measurer = await createOpentypeTextMeasurerFromBuffers([
      { name: "SingleFont", data: ttfSingle },
      { data: ttc },
    ]);
    expect(measurer).not.toBeNull();

    expect(measurer!.measureTextWidth("A", 18, false, "SingleFont")).toBeGreaterThan(0);
    expect(measurer!.measureTextWidth("A", 18, false, "TtcFont1")).toBeGreaterThan(0);
    expect(measurer!.measureTextWidth("A", 18, false, "TtcFont2")).toBeGreaterThan(0);
  });

  it("TTC バッファで fontResolver も構築される", async () => {
    const ttf1 = await createTestFontBuffer("FontAlpha");
    const ttc = buildTtcFromTtfs([ttf1]);

    const setup = await createOpentypeSetupFromBuffers([{ data: ttc }]);
    expect(setup).not.toBeNull();
    expect(setup!.fontResolver).toBeDefined();

    const font = setup!.fontResolver.resolveFont("FontAlpha", null);
    expect(font).not.toBeNull();
  });

  it("TTC 内のフォントにフォントマッピング逆引きが適用される", async () => {
    const ttf = await createTestFontBuffer("Carlito");
    const ttc = buildTtcFromTtfs([ttf]);

    const measurer = await createOpentypeTextMeasurerFromBuffers([{ data: ttc }]);
    expect(measurer).not.toBeNull();

    // Calibri → Carlito マッピングの逆引き
    const widthByOss = measurer!.measureTextWidth("A", 18, false, "Carlito");
    const widthByPptx = measurer!.measureTextWidth("A", 18, false, "Calibri");
    expect(widthByPptx).toBe(widthByOss);
  });
});

describe("createOpentypeSetupFromSystem キャッシュ", () => {
  let testFontDir: string | null = null;
  let testFontPath: string | null = null;

  /**
   * テスト用の一時フォントファイルを作成し、そのパスを返す。
   * collectFontFilePaths をモックしてこのファイルのみを返すようにすることで、
   * システムフォントの全読み込みによる OOM を防ぐ。
   */
  async function setupTestFont(): Promise<void> {
    const { mkdtemp, writeFile } = await import("node:fs/promises");
    const { join } = await import("node:path");
    const { tmpdir } = await import("node:os");
    testFontDir = await mkdtemp(join(tmpdir(), "pptx-glimpse-test-fonts-"));
    testFontPath = join(testFontDir, "CacheTestFont.ttf");
    const buffer = await createTestFontBuffer("CacheTestFont");
    await writeFile(testFontPath, Buffer.from(buffer));
  }

  afterEach(async () => {
    clearFontCache();
    vi.restoreAllMocks();
    if (testFontDir) {
      const { rm } = await import("node:fs/promises");
      await rm(testFontDir, { recursive: true, force: true });
      testFontDir = null;
      testFontPath = null;
    }
  });

  /**
   * collectFontFilePaths をモックしてテスト用フォントのみ返すようにする。
   */
  async function mockCollectFontFilePaths(): Promise<void> {
    const systemFontLoader = await import("./system-font-loader.js");
    vi.spyOn(systemFontLoader, "collectFontFilePaths").mockReturnValue([testFontPath!]);
  }

  it("同じ引数での2回目の呼び出しはキャッシュを返す", async () => {
    await setupTestFont();
    await mockCollectFontFilePaths();

    const setup1 = await createOpentypeSetupFromSystem();
    const setup2 = await createOpentypeSetupFromSystem();
    expect(setup1).not.toBeNull();
    // キャッシュが効いていれば同一オブジェクト参照
    expect(setup1).toBe(setup2);
  });

  it("fontDirs が異なる場合はキャッシュを無効化する", async () => {
    await setupTestFont();
    await mockCollectFontFilePaths();

    const setup1 = await createOpentypeSetupFromSystem();
    const setup2 = await createOpentypeSetupFromSystem(["/some-dir"]);
    expect(setup1).not.toBeNull();
    expect(setup2).not.toBeNull();
    // fontDirs が異なるのでキャッシュキーが変わり別オブジェクト
    expect(setup1).not.toBe(setup2);
  });

  it("fontMapping が異なる場合はキャッシュを無効化する", async () => {
    await setupTestFont();
    await mockCollectFontFilePaths();

    const setup1 = await createOpentypeSetupFromSystem();
    const setup2 = await createOpentypeSetupFromSystem(undefined, { Custom: "Font" });
    expect(setup1).not.toBeNull();
    expect(setup2).not.toBeNull();
    expect(setup1).not.toBe(setup2);
  });

  it("clearFontCache() でキャッシュがクリアされる", async () => {
    await setupTestFont();
    await mockCollectFontFilePaths();

    const setup1 = await createOpentypeSetupFromSystem();
    clearFontCache();
    const setup2 = await createOpentypeSetupFromSystem();
    expect(setup1).not.toBeNull();
    expect(setup2).not.toBeNull();
    // クリア後は再構築されるので別オブジェクト
    expect(setup1).not.toBe(setup2);
  });

  it("キャッシュにより collectFontFilePaths が再呼び出しされない", async () => {
    await setupTestFont();
    const systemFontLoader = await import("./system-font-loader.js");
    const spy = vi.spyOn(systemFontLoader, "collectFontFilePaths").mockReturnValue([testFontPath!]);

    // 1回目: collectFontFilePaths が呼ばれる
    await createOpentypeSetupFromSystem();
    expect(spy).toHaveBeenCalledTimes(1);

    // 2回目: キャッシュが効いて collectFontFilePaths が呼ばれない
    await createOpentypeSetupFromSystem();
    expect(spy).toHaveBeenCalledTimes(1);
  });

  it("skipSystemFonts が異なる場合はキャッシュを無効化する", async () => {
    await setupTestFont();
    await mockCollectFontFilePaths();

    const setup1 = await createOpentypeSetupFromSystem(undefined, undefined, false);
    const setup2 = await createOpentypeSetupFromSystem(undefined, undefined, true);
    expect(setup1).not.toBeNull();
    expect(setup2).not.toBeNull();
    // skipSystemFonts が異なるのでキャッシュキーが変わり別オブジェクト
    expect(setup1).not.toBe(setup2);
  });

  it("skipSystemFonts が collectFontFilePaths に正しく伝播する", async () => {
    await setupTestFont();
    const systemFontLoader = await import("./system-font-loader.js");
    const spy = vi.spyOn(systemFontLoader, "collectFontFilePaths").mockReturnValue([testFontPath!]);

    await createOpentypeSetupFromSystem(undefined, undefined, true);
    expect(spy).toHaveBeenCalledWith(undefined, true);
  });
});
