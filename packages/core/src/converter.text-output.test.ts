import { mkdtempSync, rmSync, writeFileSync } from "node:fs";
import { tmpdir } from "node:os";
import { join } from "node:path";

import { clearFontCache } from "@pptx-glimpse/renderer";
import JSZip from "jszip";
import { afterAll, afterEach, beforeAll, describe, expect, it, vi } from "vitest";

import { convertPptxToPng, convertPptxToSvg } from "./converter.js";

/**
 * Integration test for textOutput option (native <text> + embedded font output mode).
 *
 * Test font (EmbedTestFont: A, B, space / BulletTestFont: C, space)
 * Export to a temporary directory and use fontDirs + skipSystemFonts to resolve fonts independently of the environment.
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
 * Generates a test TTF with only glyphs of the specified character (+space).
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
 * Extracts and parses the embedded font of the specified family name from @font-face in SVG.
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

afterEach(() => {
  vi.restoreAllMocks();
});

function convertOptions(extra?: Record<string, unknown>) {
  return { fontDirs: [fontDir], skipSystemFonts: true, ...extra } as const;
}

describe("textOutput: text SVG output", () => {
  it("Contains the <text> element and @font-face, but does not include the glyph outline path", async () => {
    const pptx = await createTestPptx("ABA AB");
    const { slides } = await convertPptxToSvg(pptx, convertOptions({ textOutput: "text" }));
    const svg = slides[0].svg;

    expect(svg).toContain("<text");
    expect(svg).toContain("@font-face");
    expect(svg).toContain('font-family:"EmbedTestFont"');
    expect(svg).toContain("data:font/otf;base64,");
    // Contains the text content as is (not outlined)
    expect(svg).toContain("ABA AB");
    // Glyph outline path not included (shapes on slide are rect only)
    expect(svg).not.toContain("<path");
  });

  it("If there are only characters that are not included in the font, do not embed @font-face and output only <text>", async () => {
    const pptx = await createTestPptx("XYZ");
    const { slides } = await convertPptxToSvg(pptx, convertOptions({ textOutput: "text" }));
    const svg = slides[0].svg;

    expect(svg).toContain("<text");
    expect(svg).toContain("XYZ");
    expect(svg).not.toContain("@font-face");
  });

  it("embeds bullet glyphs using the text run font when buFont is omitted", async () => {
    // buChar="B" / no buFont. Bullet B is drawn and embedded in the run font (EmbedTestFont).
    const pptx = await createTestPptx("A", { pPr: '<a:pPr><a:buChar char="B"/></a:pPr>' });
    const { slides } = await convertPptxToSvg(pptx, convertOptions({ textOutput: "text" }));
    const svg = slides[0].svg;

    expect(svg).toContain("@font-face");
    // Bullet tspan's font-family includes run's font
    expect(svg).toMatch(/<tspan[^>]*text-anchor="start"[^>]*font-family="EmbedTestFont/);

    // Embedded subset contains bullet B glyph
    const embedded = await parseEmbeddedFont(svg, "EmbedTestFont");
    expect(embedded.charToGlyph("B").index).toBeGreaterThan(0);
    expect(embedded.charToGlyph("A").index).toBeGreaterThan(0);
  });

  it("embeds bullet glyphs using the specified buFont", async () => {
    const pptx = await createTestPptx("A", {
      pPr: '<a:pPr><a:buFont typeface="BulletTestFont"/><a:buChar char="C"/></a:pPr>',
    });
    const { slides } = await convertPptxToSvg(pptx, convertOptions({ textOutput: "text" }));
    const svg = slides[0].svg;

    // Two @font-faces are embedded, one for runs and one for bullet points.
    expect(svg).toContain('font-family:"EmbedTestFont"');
    expect(svg).toContain('font-family:"BulletTestFont"');
    expect(svg).toMatch(/<tspan[^>]*font-family="BulletTestFont/);

    const bulletFont = await parseEmbeddedFont(svg, "BulletTestFont");
    expect(bulletFont.charToGlyph("C").index).toBeGreaterThan(0);
  });

  it("Fonts resolved with fontMapping are embedded with PPTX font names", async () => {
    const pptx = await createTestPptx("ABA", { typeface: "MappedCorpFont" });
    const { slides } = await convertPptxToSvg(
      pptx,
      convertOptions({
        textOutput: "text",
        fontMapping: { MappedCorpFont: "EmbedTestFont" },
      }),
    );
    const svg = slides[0].svg;

    // @font-face is declared with the PPTX font name referenced by tspan
    expect(svg).toContain('font-family:"MappedCorpFont"');
    expect(svg).toMatch(/<tspan[^>]*font-family="MappedCorpFont/);

    const embedded = await parseEmbeddedFont(svg, "MappedCorpFont");
    expect(embedded.charToGlyph("A").index).toBeGreaterThan(0);
  });

  it("uses direct font buffers without system font scanning", async () => {
    const rendererNode = await import("@pptx-glimpse/renderer/node");
    const systemSetupSpy = vi.spyOn(rendererNode, "createOpentypeSetupFromSystem");
    const directFont = new Uint8Array(await createTestFontBuffer("DirectBufferFont", ["A", "B"]));
    const pptx = await createTestPptx("ABA", { typeface: "DirectBufferFont" });

    const { slides } = await convertPptxToSvg(pptx, {
      fonts: [{ name: "DirectBufferFont", data: directFont }],
      textOutput: "text",
    });
    const svg = slides[0].svg;

    expect(systemSetupSpy).not.toHaveBeenCalled();
    expect(svg).toContain("@font-face");
    expect(svg).toContain('font-family:"DirectBufferFont"');
    expect(svg).toContain("ABA");
  });
});

describe("textOutput: path SVG output", () => {
  it("omits @font-face from the default path output", async () => {
    const pptx = await createTestPptx("ABA AB");
    const { slides } = await convertPptxToSvg(pptx, convertOptions());
    const svg = slides[0].svg;

    expect(svg).toContain("<path");
    expect(svg).not.toContain("<text");
    expect(svg).not.toContain("@font-face");
  });

  it('matches the default output when textOutput: "path" is specified explicitly', async () => {
    const pptx = await createTestPptx("ABA AB");
    const { slides: defaultResults } = await convertPptxToSvg(pptx, convertOptions());
    const { slides: pathResults } = await convertPptxToSvg(
      pptx,
      convertOptions({ textOutput: "path" }),
    );

    expect(pathResults[0].svg).toBe(defaultResults[0].svg);
  });
});

describe("textOutput: PNG conversion", () => {
  it('converts PNG through path output even when textOutput: "text" is specified', async () => {
    const pptx = await createTestPptx("ABA AB");
    const { slides } = await convertPptxToPng(pptx, convertOptions({ textOutput: "text" }));

    expect(slides).toHaveLength(1);
    // PNG magic bytes
    expect(slides[0].png[0]).toBe(0x89);
    expect(slides[0].png[1]).toBe(0x50);
  });

  it("passes direct font buffers to PNG conversion without loading Node font files", async () => {
    const nodeFontLoader = await import("./node-font-loader.js");
    const systemFontBufferSpy = vi.spyOn(nodeFontLoader, "loadFontBuffersFromSystem");
    const directFont = new Uint8Array(await createTestFontBuffer("DirectPngFont", ["A", "B"]));
    const pptx = await createTestPptx("ABA", { typeface: "DirectPngFont" });

    const { slides } = await convertPptxToPng(pptx, {
      fonts: [{ name: "DirectPngFont", data: directFont }],
      textOutput: "text",
    });

    expect(systemFontBufferSpy).not.toHaveBeenCalled();
    expect(slides).toHaveLength(1);
    expect(slides[0].png[0]).toBe(0x89);
    expect(slides[0].png[1]).toBe(0x50);
  });
});
