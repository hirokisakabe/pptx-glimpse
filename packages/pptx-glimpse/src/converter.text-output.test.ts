import { mkdtempSync, rmSync, writeFileSync } from "node:fs";
import { tmpdir } from "node:os";
import { join } from "node:path";

import JSZip from "jszip";
import { clearFontCache } from "pptx-glimpse-renderer";
import { afterAll, beforeAll, describe, expect, it } from "vitest";

import { convertPptxToPng, convertPptxToSvg } from "./converter.js";

/**
 * textOutput オプション (ネイティブ <text> + 埋め込みフォント出力モード) の統合テスト。
 *
 * テスト用フォント (EmbedTestFont: A, B, space / BulletTestFont: C, space) を
 * 一時ディレクトリに書き出し、fontDirs + skipSystemFonts で環境非依存にフォントを解決させる。
 */

const contentTypes = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/ppt/presentation.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml"/>
  <Override PartName="/ppt/slides/slide1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/>
  <Override PartName="/ppt/slideLayouts/slideLayout1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slideLayout+xml"/>
  <Override PartName="/ppt/slideMasters/slideMaster1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slideMaster+xml"/>
  <Override PartName="/ppt/theme/theme1.xml" ContentType="application/vnd.openxmlformats-officedocument.theme+xml"/>
</Types>`;

const rootRels = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="ppt/presentation.xml"/>
</Relationships>`;

const presentationXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:presentation xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">
  <p:sldMasterIdLst>
    <p:sldMasterId r:id="rId1"/>
  </p:sldMasterIdLst>
  <p:sldIdLst>
    <p:sldId id="256" r:id="rId2"/>
  </p:sldIdLst>
  <p:sldSz cx="9144000" cy="5143500" type="screen16x9"/>
</p:presentation>`;

const presentationRels = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideMaster" Target="slideMasters/slideMaster1.xml"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" Target="slides/slide1.xml"/>
  <Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme" Target="theme/theme1.xml"/>
</Relationships>`;

function buildSlideXml(text: string, options?: { pPr?: string; typeface?: string }): string {
  const typeface = options?.typeface ?? "EmbedTestFont";
  const pPr = options?.pPr ?? "";
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:sld xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">
  <p:cSld>
    <p:spTree>
      <p:nvGrpSpPr>
        <p:cNvPr id="1" name=""/>
        <p:cNvGrpSpPr/>
        <p:nvPr/>
      </p:nvGrpSpPr>
      <p:grpSpPr>
        <a:xfrm>
          <a:off x="0" y="0"/>
          <a:ext cx="0" cy="0"/>
          <a:chOff x="0" y="0"/>
          <a:chExt cx="0" cy="0"/>
        </a:xfrm>
      </p:grpSpPr>
      <p:sp>
        <p:nvSpPr>
          <p:cNvPr id="2" name="TextBox 1"/>
          <p:cNvSpPr/>
          <p:nvPr/>
        </p:nvSpPr>
        <p:spPr>
          <a:xfrm>
            <a:off x="457200" y="274638"/>
            <a:ext cx="3048000" cy="1143000"/>
          </a:xfrm>
          <a:prstGeom prst="rect">
            <a:avLst/>
          </a:prstGeom>
        </p:spPr>
        <p:txBody>
          <a:bodyPr/>
          <a:lstStyle/>
          <a:p>
            ${pPr}
            <a:r>
              <a:rPr lang="en-US" sz="1600">
                <a:latin typeface="${typeface}"/>
              </a:rPr>
              <a:t>${text}</a:t>
            </a:r>
          </a:p>
        </p:txBody>
      </p:sp>
    </p:spTree>
  </p:cSld>
</p:sld>`;
}

const slide1Rels = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout" Target="../slideLayouts/slideLayout1.xml"/>
</Relationships>`;

const slideMaster1 = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:sldMaster xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">
  <p:cSld>
    <p:spTree>
      <p:nvGrpSpPr>
        <p:cNvPr id="1" name=""/>
        <p:cNvGrpSpPr/>
        <p:nvPr/>
      </p:nvGrpSpPr>
      <p:grpSpPr/>
    </p:spTree>
  </p:cSld>
  <p:clrMap bg1="lt1" tx1="dk1" bg2="lt2" tx2="dk2" accent1="accent1" accent2="accent2" accent3="accent3" accent4="accent4" accent5="accent5" accent6="accent6" hlink="hlink" folHlink="folHlink"/>
  <p:sldLayoutIdLst>
    <p:sldLayoutId r:id="rId1"/>
  </p:sldLayoutIdLst>
</p:sldMaster>`;

const slideMaster1Rels = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout" Target="../slideLayouts/slideLayout1.xml"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme" Target="../theme/theme1.xml"/>
</Relationships>`;

const slideLayout1 = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:sldLayout xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" type="blank">
  <p:cSld>
    <p:spTree>
      <p:nvGrpSpPr>
        <p:cNvPr id="1" name=""/>
        <p:cNvGrpSpPr/>
        <p:nvPr/>
      </p:nvGrpSpPr>
      <p:grpSpPr/>
    </p:spTree>
  </p:cSld>
</p:sldLayout>`;

const slideLayout1Rels = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideMaster" Target="../slideMasters/slideMaster1.xml"/>
</Relationships>`;

const theme1 = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" name="Office Theme">
  <a:themeElements>
    <a:clrScheme name="Office">
      <a:dk1><a:sysClr val="windowText" lastClr="000000"/></a:dk1>
      <a:lt1><a:sysClr val="window" lastClr="FFFFFF"/></a:lt1>
      <a:dk2><a:srgbClr val="44546A"/></a:dk2>
      <a:lt2><a:srgbClr val="E7E6E6"/></a:lt2>
      <a:accent1><a:srgbClr val="4472C4"/></a:accent1>
      <a:accent2><a:srgbClr val="ED7D31"/></a:accent2>
      <a:accent3><a:srgbClr val="A5A5A5"/></a:accent3>
      <a:accent4><a:srgbClr val="FFC000"/></a:accent4>
      <a:accent5><a:srgbClr val="5B9BD5"/></a:accent5>
      <a:accent6><a:srgbClr val="70AD47"/></a:accent6>
      <a:hlink><a:srgbClr val="0563C1"/></a:hlink>
      <a:folHlink><a:srgbClr val="954F72"/></a:folHlink>
    </a:clrScheme>
    <a:fontScheme name="Office">
      <a:majorFont>
        <a:latin typeface="Calibri Light"/>
      </a:majorFont>
      <a:minorFont>
        <a:latin typeface="Calibri"/>
      </a:minorFont>
    </a:fontScheme>
    <a:fmtScheme name="Office"/>
  </a:themeElements>
</a:theme>`;

async function createTestPptx(
  text: string,
  slideOptions?: { pPr?: string; typeface?: string },
): Promise<Buffer> {
  const zip = new JSZip();
  zip.file("[Content_Types].xml", contentTypes);
  zip.file("_rels/.rels", rootRels);
  zip.file("ppt/presentation.xml", presentationXml);
  zip.file("ppt/_rels/presentation.xml.rels", presentationRels);
  zip.file("ppt/slides/slide1.xml", buildSlideXml(text, slideOptions));
  zip.file("ppt/slides/_rels/slide1.xml.rels", slide1Rels);
  zip.file("ppt/slideMasters/slideMaster1.xml", slideMaster1);
  zip.file("ppt/slideMasters/_rels/slideMaster1.xml.rels", slideMaster1Rels);
  zip.file("ppt/slideLayouts/slideLayout1.xml", slideLayout1);
  zip.file("ppt/slideLayouts/_rels/slideLayout1.xml.rels", slideLayout1Rels);
  zip.file("ppt/theme/theme1.xml", theme1);
  return zip.generateAsync({ type: "nodebuffer" });
}

interface OpentypeTestModule {
  Glyph: new (opts: Record<string, unknown>) => unknown;
  Path: new () => {
    moveTo(x: number, y: number): void;
    lineTo(x: number, y: number): void;
    close(): void;
  };
  Font: new (opts: Record<string, unknown>) => { toArrayBuffer(): ArrayBuffer };
  parse: (buffer: ArrayBuffer) => { charToGlyph(char: string): { index: number } };
}

async function loadOpentype(): Promise<OpentypeTestModule> {
  // eslint-disable-next-line @typescript-eslint/no-unsafe-assignment
  const mod: OpentypeTestModule = await import("opentype.js");
  return mod;
}

/**
 * 指定した文字 (+ space) のグリフのみを持つテスト用 TTF を生成する。
 */
async function createTestFontBuffer(familyName: string, letters: string[]): Promise<ArrayBuffer> {
  const opentype = await loadOpentype();

  const notdefGlyph = new opentype.Glyph({
    name: ".notdef",
    advanceWidth: 650,
    path: new opentype.Path(),
  });

  const makeTrianglePath = () => {
    const path = new opentype.Path();
    path.moveTo(0, 0);
    path.lineTo(300, 800);
    path.lineTo(600, 0);
    path.close();
    return path;
  };

  const spaceGlyph = new opentype.Glyph({
    name: "space",
    unicode: 32,
    advanceWidth: 250,
    path: new opentype.Path(),
  });

  const letterGlyphs = letters.map(
    (letter) =>
      new opentype.Glyph({
        name: letter,
        unicode: letter.codePointAt(0)!,
        advanceWidth: 600,
        path: makeTrianglePath(),
      }),
  );

  const font = new opentype.Font({
    familyName,
    styleName: "Regular",
    unitsPerEm: 1000,
    ascender: 800,
    descender: -200,
    glyphs: [notdefGlyph, spaceGlyph, ...letterGlyphs],
  });

  return font.toArrayBuffer();
}

/**
 * SVG の @font-face から指定ファミリー名の埋め込みフォントを取り出してパースする。
 */
async function parseEmbeddedFont(
  svg: string,
  familyName: string,
): Promise<{ charToGlyph(char: string): { index: number } }> {
  const pattern = new RegExp(
    `@font-face\\{font-family:"${familyName}";src:url\\(data:font/otf;base64,([^)]+)\\)`,
  );
  const match = svg.match(pattern);
  expect(match).not.toBeNull();
  const bytes = Buffer.from(match![1], "base64");
  const opentype = await loadOpentype();
  return opentype.parse(bytes.buffer.slice(bytes.byteOffset, bytes.byteOffset + bytes.byteLength));
}

let fontDir: string;

beforeAll(async () => {
  clearFontCache();
  fontDir = mkdtempSync(join(tmpdir(), "pptx-glimpse-text-output-"));
  writeFileSync(
    join(fontDir, "EmbedTestFont.ttf"),
    Buffer.from(await createTestFontBuffer("EmbedTestFont", ["A", "B"])),
  );
  writeFileSync(
    join(fontDir, "BulletTestFont.ttf"),
    Buffer.from(await createTestFontBuffer("BulletTestFont", ["C"])),
  );
});

afterAll(() => {
  clearFontCache();
  rmSync(fontDir, { recursive: true, force: true });
});

function convertOptions(extra?: Record<string, unknown>) {
  return { fontDirs: [fontDir], skipSystemFonts: true, ...extra } as const;
}

describe('textOutput: "text" (ネイティブ <text> + 埋め込みフォント)', () => {
  it("<text> 要素と @font-face を含み、グリフのアウトラインパスを含まない", async () => {
    const pptx = await createTestPptx("ABA AB");
    const results = await convertPptxToSvg(pptx, convertOptions({ textOutput: "text" }));
    const svg = results[0].svg;

    expect(svg).toContain("<text");
    expect(svg).toContain("@font-face");
    expect(svg).toContain('font-family:"EmbedTestFont"');
    expect(svg).toContain("data:font/otf;base64,");
    // テキスト内容がそのまま含まれる (アウトライン化されていない)
    expect(svg).toContain("ABA AB");
    // グリフのアウトラインパスが含まれない (スライド上の図形は rect のみ)
    expect(svg).not.toContain("<path");
  });

  it("フォント未収録の文字のみの場合は @font-face を埋め込まず <text> のみ出力する", async () => {
    const pptx = await createTestPptx("XYZ");
    const results = await convertPptxToSvg(pptx, convertOptions({ textOutput: "text" }));
    const svg = results[0].svg;

    expect(svg).toContain("<text");
    expect(svg).toContain("XYZ");
    expect(svg).not.toContain("@font-face");
  });

  it("bulletFont 未指定の箇条書きはテキストランのフォントにフォールバックして埋め込まれる", async () => {
    // buChar="B" / buFont なし。箇条書き記号 B はランのフォント (EmbedTestFont) で描画・埋め込みされる
    const pptx = await createTestPptx("A", { pPr: '<a:pPr><a:buChar char="B"/></a:pPr>' });
    const results = await convertPptxToSvg(pptx, convertOptions({ textOutput: "text" }));
    const svg = results[0].svg;

    expect(svg).toContain("@font-face");
    // 箇条書き tspan の font-family にランのフォントが含まれる
    expect(svg).toMatch(/<tspan[^>]*text-anchor="start"[^>]*font-family="EmbedTestFont/);

    // 埋め込みサブセットに箇条書き記号 B のグリフが含まれる
    const embedded = await parseEmbeddedFont(svg, "EmbedTestFont");
    expect(embedded.charToGlyph("B").index).toBeGreaterThan(0);
    expect(embedded.charToGlyph("A").index).toBeGreaterThan(0);
  });

  it("buFont 指定の箇条書きは指定フォントで埋め込まれる", async () => {
    const pptx = await createTestPptx("A", {
      pPr: '<a:pPr><a:buFont typeface="BulletTestFont"/><a:buChar char="C"/></a:pPr>',
    });
    const results = await convertPptxToSvg(pptx, convertOptions({ textOutput: "text" }));
    const svg = results[0].svg;

    // ラン用と箇条書き用で 2 つの @font-face が埋め込まれる
    expect(svg).toContain('font-family:"EmbedTestFont"');
    expect(svg).toContain('font-family:"BulletTestFont"');
    expect(svg).toMatch(/<tspan[^>]*font-family="BulletTestFont/);

    const bulletFont = await parseEmbeddedFont(svg, "BulletTestFont");
    expect(bulletFont.charToGlyph("C").index).toBeGreaterThan(0);
  });

  it("fontMapping で解決されたフォントが PPTX フォント名で埋め込まれる", async () => {
    const pptx = await createTestPptx("ABA", { typeface: "MappedCorpFont" });
    const results = await convertPptxToSvg(
      pptx,
      convertOptions({
        textOutput: "text",
        fontMapping: { MappedCorpFont: "EmbedTestFont" },
      }),
    );
    const svg = results[0].svg;

    // @font-face は tspan が参照する PPTX フォント名で宣言される
    expect(svg).toContain('font-family:"MappedCorpFont"');
    expect(svg).toMatch(/<tspan[^>]*font-family="MappedCorpFont/);

    const embedded = await parseEmbeddedFont(svg, "MappedCorpFont");
    expect(embedded.charToGlyph("A").index).toBeGreaterThan(0);
  });
});

describe("textOutput デフォルト (パス出力)", () => {
  it("デフォルトでは従来どおりパス出力で @font-face を含まない", async () => {
    const pptx = await createTestPptx("ABA AB");
    const results = await convertPptxToSvg(pptx, convertOptions());
    const svg = results[0].svg;

    expect(svg).toContain("<path");
    expect(svg).not.toContain("<text");
    expect(svg).not.toContain("@font-face");
  });

  it('textOutput: "path" 明示指定でも同じ出力になる', async () => {
    const pptx = await createTestPptx("ABA AB");
    const defaultResults = await convertPptxToSvg(pptx, convertOptions());
    const pathResults = await convertPptxToSvg(pptx, convertOptions({ textOutput: "path" }));

    expect(pathResults[0].svg).toBe(defaultResults[0].svg);
  });
});

describe("convertPptxToPng と textOutput", () => {
  it('textOutput: "text" を指定しても PNG 変換はパス出力で行われ成功する', async () => {
    const pptx = await createTestPptx("ABA AB");
    const results = await convertPptxToPng(pptx, convertOptions({ textOutput: "text" }));

    expect(results).toHaveLength(1);
    // PNG magic bytes
    expect(results[0].png[0]).toBe(0x89);
    expect(results[0].png[1]).toBe(0x50);
  });
});
