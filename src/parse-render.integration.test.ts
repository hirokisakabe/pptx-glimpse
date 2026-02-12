import { describe, it, expect } from "vitest";
import { parseShapeTree } from "./parser/slide-parser.js";
import { parseXml } from "./parser/xml-parser.js";
import type { XmlNode } from "./parser/xml-parser.js";
import type { PptxArchive } from "./parser/pptx-reader.js";
import type { ShapeElement } from "./model/shape.js";
import type { Relationship } from "./parser/relationship-parser.js";
import { ColorResolver } from "./color/color-resolver.js";
import { renderShape } from "./renderer/shape-renderer.js";

// ---------------------------------------------------------------------------
// Helpers
// ---------------------------------------------------------------------------

function createColorResolver() {
  return new ColorResolver(
    {
      dk1: "#000000",
      lt1: "#FFFFFF",
      dk2: "#44546A",
      lt2: "#E7E6E6",
      accent1: "#4472C4",
      accent2: "#ED7D31",
      accent3: "#A5A5A5",
      accent4: "#FFC000",
      accent5: "#5B9BD5",
      accent6: "#70AD47",
      hlink: "#0563C1",
      folHlink: "#954F72",
    },
    {
      bg1: "lt1",
      tx1: "dk1",
      bg2: "lt2",
      tx2: "dk2",
      accent1: "accent1",
      accent2: "accent2",
      accent3: "accent3",
      accent4: "accent4",
      accent5: "accent5",
      accent6: "accent6",
      hlink: "hlink",
      folHlink: "folHlink",
    },
  );
}

function createEmptyArchive(): PptxArchive {
  return { files: new Map(), media: new Map() };
}

function buildSpTreeXml(spContent: string): string {
  return `
    <p:spTree xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"
               xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
               xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
      ${spContent}
    </p:spTree>
  `;
}

function buildShapeXml(txBodyContent: string, spPrExtra: string = ""): string {
  return buildSpTreeXml(`
    <p:sp>
      <p:nvSpPr>
        <p:cNvPr id="2" name="TextBox 1"/>
        <p:cNvSpPr/>
        <p:nvPr/>
      </p:nvSpPr>
      <p:spPr>
        <a:xfrm>
          <a:off x="914400" y="914400"/>
          <a:ext cx="4572000" cy="2743200"/>
        </a:xfrm>
        <a:prstGeom prst="rect"/>
        ${spPrExtra}
      </p:spPr>
      ${txBodyContent}
    </p:sp>
  `);
}

function parseAndRenderShape(
  xml: string,
  rels?: Map<string, Relationship>,
): { shape: ShapeElement; svg: string } {
  const parsed = parseXml(xml);
  const elements = parseShapeTree(
    parsed.spTree as XmlNode | undefined,
    rels ?? new Map<string, Relationship>(),
    "ppt/slides/slide1.xml",
    createEmptyArchive(),
    createColorResolver(),
  );
  expect(elements).toHaveLength(1);
  expect(elements[0].type).toBe("shape");
  const shape = elements[0] as ShapeElement;
  const svg = renderShape(shape);
  return { shape, svg };
}

function extractDyValues(svg: string): number[] {
  return [...svg.matchAll(/dy="([^"]+)"/g)].map((m) => parseFloat(m[1]));
}

// ---------------------------------------------------------------------------
// Priority 1: paragraph spacing
// ---------------------------------------------------------------------------

describe("parse-render integration: paragraph spacing", () => {
  it("spaceBefore (pts) で段落間隔が広がる", () => {
    const xmlWith = buildShapeXml(`
      <p:txBody>
        <a:bodyPr/>
        <a:p>
          <a:r><a:rPr lang="en-US" sz="1800"/><a:t>First</a:t></a:r>
        </a:p>
        <a:p>
          <a:pPr>
            <a:spcBef><a:spcPts val="1200"/></a:spcBef>
          </a:pPr>
          <a:r><a:rPr lang="en-US" sz="1800"/><a:t>Second</a:t></a:r>
        </a:p>
      </p:txBody>
    `);
    const xmlWithout = buildShapeXml(`
      <p:txBody>
        <a:bodyPr/>
        <a:p>
          <a:r><a:rPr lang="en-US" sz="1800"/><a:t>First</a:t></a:r>
        </a:p>
        <a:p>
          <a:r><a:rPr lang="en-US" sz="1800"/><a:t>Second</a:t></a:r>
        </a:p>
      </p:txBody>
    `);
    const dyWith = extractDyValues(parseAndRenderShape(xmlWith).svg);
    const dyWithout = extractDyValues(parseAndRenderShape(xmlWithout).svg);
    expect(dyWith[1]).toBeGreaterThan(dyWithout[1]);
  });

  it("spaceAfter (pts) で次段落への間隔が広がる", () => {
    const xmlWith = buildShapeXml(`
      <p:txBody>
        <a:bodyPr/>
        <a:p>
          <a:pPr>
            <a:spcAft><a:spcPts val="1200"/></a:spcAft>
          </a:pPr>
          <a:r><a:rPr lang="en-US" sz="1800"/><a:t>First</a:t></a:r>
        </a:p>
        <a:p>
          <a:r><a:rPr lang="en-US" sz="1800"/><a:t>Second</a:t></a:r>
        </a:p>
      </p:txBody>
    `);
    const xmlWithout = buildShapeXml(`
      <p:txBody>
        <a:bodyPr/>
        <a:p>
          <a:r><a:rPr lang="en-US" sz="1800"/><a:t>First</a:t></a:r>
        </a:p>
        <a:p>
          <a:r><a:rPr lang="en-US" sz="1800"/><a:t>Second</a:t></a:r>
        </a:p>
      </p:txBody>
    `);
    const dyWith = extractDyValues(parseAndRenderShape(xmlWith).svg);
    const dyWithout = extractDyValues(parseAndRenderShape(xmlWithout).svg);
    expect(dyWith[1]).toBeGreaterThan(dyWithout[1]);
  });

  it("spaceBefore (pct) でフォントサイズに比例した間隔が適用される", () => {
    const xmlWith = buildShapeXml(`
      <p:txBody>
        <a:bodyPr/>
        <a:p>
          <a:r><a:rPr lang="en-US" sz="1800"/><a:t>First</a:t></a:r>
        </a:p>
        <a:p>
          <a:pPr>
            <a:spcBef><a:spcPct val="100000"/></a:spcBef>
          </a:pPr>
          <a:r><a:rPr lang="en-US" sz="1800"/><a:t>Second</a:t></a:r>
        </a:p>
      </p:txBody>
    `);
    const xmlWithout = buildShapeXml(`
      <p:txBody>
        <a:bodyPr/>
        <a:p>
          <a:r><a:rPr lang="en-US" sz="1800"/><a:t>First</a:t></a:r>
        </a:p>
        <a:p>
          <a:r><a:rPr lang="en-US" sz="1800"/><a:t>Second</a:t></a:r>
        </a:p>
      </p:txBody>
    `);
    const dyWith = extractDyValues(parseAndRenderShape(xmlWith).svg);
    const dyWithout = extractDyValues(parseAndRenderShape(xmlWithout).svg);
    expect(dyWith[1]).toBeGreaterThan(dyWithout[1]);
  });
});

// ---------------------------------------------------------------------------
// Priority 1: text alignment
// ---------------------------------------------------------------------------

describe("parse-render integration: text alignment", () => {
  it('algn="ctr" → text-anchor="middle"', () => {
    const xml = buildShapeXml(`
      <p:txBody>
        <a:bodyPr/>
        <a:p>
          <a:pPr algn="ctr"/>
          <a:r><a:rPr lang="en-US" sz="1800"/><a:t>Centered</a:t></a:r>
        </a:p>
      </p:txBody>
    `);
    const { svg } = parseAndRenderShape(xml);
    expect(svg).toContain('text-anchor="middle"');
  });

  it('algn="r" → text-anchor="end"', () => {
    const xml = buildShapeXml(`
      <p:txBody>
        <a:bodyPr/>
        <a:p>
          <a:pPr algn="r"/>
          <a:r><a:rPr lang="en-US" sz="1800"/><a:t>Right</a:t></a:r>
        </a:p>
      </p:txBody>
    `);
    const { svg } = parseAndRenderShape(xml);
    expect(svg).toContain('text-anchor="end"');
  });

  it('デフォルト (algn="l") → text-anchor="start"', () => {
    const xml = buildShapeXml(`
      <p:txBody>
        <a:bodyPr/>
        <a:p>
          <a:r><a:rPr lang="en-US" sz="1800"/><a:t>Left</a:t></a:r>
        </a:p>
      </p:txBody>
    `);
    const { svg } = parseAndRenderShape(xml);
    expect(svg).toContain('text-anchor="start"');
    expect(svg).not.toContain('text-anchor="middle"');
    expect(svg).not.toContain('text-anchor="end"');
  });
});

// ---------------------------------------------------------------------------
// Priority 1: fontSize
// ---------------------------------------------------------------------------

describe("parse-render integration: fontSize", () => {
  it('sz="2400" → font-size="24pt"', () => {
    const xml = buildShapeXml(`
      <p:txBody>
        <a:bodyPr/>
        <a:p>
          <a:r><a:rPr lang="en-US" sz="2400"/><a:t>Large</a:t></a:r>
        </a:p>
      </p:txBody>
    `);
    const { svg } = parseAndRenderShape(xml);
    expect(svg).toContain('font-size="24pt"');
  });

  it('sz="1200" → font-size="12pt"', () => {
    const xml = buildShapeXml(`
      <p:txBody>
        <a:bodyPr/>
        <a:p>
          <a:r><a:rPr lang="en-US" sz="1200"/><a:t>Small</a:t></a:r>
        </a:p>
      </p:txBody>
    `);
    const { svg } = parseAndRenderShape(xml);
    expect(svg).toContain('font-size="12pt"');
  });
});

// ---------------------------------------------------------------------------
// Priority 2: text formatting (bold, italic, underline, strikethrough)
// ---------------------------------------------------------------------------

describe("parse-render integration: text formatting", () => {
  it('b="1" → font-weight="bold"', () => {
    const xml = buildShapeXml(`
      <p:txBody>
        <a:bodyPr/>
        <a:p>
          <a:r><a:rPr lang="en-US" sz="1800" b="1"/><a:t>Bold</a:t></a:r>
        </a:p>
      </p:txBody>
    `);
    const { svg } = parseAndRenderShape(xml);
    expect(svg).toContain('font-weight="bold"');
  });

  it('i="1" → font-style="italic"', () => {
    const xml = buildShapeXml(`
      <p:txBody>
        <a:bodyPr/>
        <a:p>
          <a:r><a:rPr lang="en-US" sz="1800" i="1"/><a:t>Italic</a:t></a:r>
        </a:p>
      </p:txBody>
    `);
    const { svg } = parseAndRenderShape(xml);
    expect(svg).toContain('font-style="italic"');
  });

  it('u="sng" → text-decoration に underline を含む', () => {
    const xml = buildShapeXml(`
      <p:txBody>
        <a:bodyPr/>
        <a:p>
          <a:r><a:rPr lang="en-US" sz="1800" u="sng"/><a:t>Underline</a:t></a:r>
        </a:p>
      </p:txBody>
    `);
    const { svg } = parseAndRenderShape(xml);
    expect(svg).toMatch(/text-decoration="[^"]*underline[^"]*"/);
  });

  it('strike="sngStrike" → text-decoration に line-through を含む', () => {
    const xml = buildShapeXml(`
      <p:txBody>
        <a:bodyPr/>
        <a:p>
          <a:r><a:rPr lang="en-US" sz="1800" strike="sngStrike"/><a:t>Strike</a:t></a:r>
        </a:p>
      </p:txBody>
    `);
    const { svg } = parseAndRenderShape(xml);
    expect(svg).toMatch(/text-decoration="[^"]*line-through[^"]*"/);
  });

  it("bold + underline + strikethrough の複合書式", () => {
    const xml = buildShapeXml(`
      <p:txBody>
        <a:bodyPr/>
        <a:p>
          <a:r><a:rPr lang="en-US" sz="1800" b="1" u="sng" strike="sngStrike"/><a:t>All</a:t></a:r>
        </a:p>
      </p:txBody>
    `);
    const { svg } = parseAndRenderShape(xml);
    expect(svg).toContain('font-weight="bold"');
    expect(svg).toMatch(/text-decoration="[^"]*underline[^"]*"/);
    expect(svg).toMatch(/text-decoration="[^"]*line-through[^"]*"/);
  });
});

// ---------------------------------------------------------------------------
// Priority 2: baseline (superscript / subscript)
// ---------------------------------------------------------------------------

describe("parse-render integration: baseline", () => {
  it('baseline="30000" → baseline-shift="super"', () => {
    const xml = buildShapeXml(`
      <p:txBody>
        <a:bodyPr/>
        <a:p>
          <a:r><a:rPr lang="en-US" sz="1800"/><a:t>H</a:t></a:r>
          <a:r><a:rPr lang="en-US" sz="1200" baseline="30000"/><a:t>2</a:t></a:r>
          <a:r><a:rPr lang="en-US" sz="1800"/><a:t>O</a:t></a:r>
        </a:p>
      </p:txBody>
    `);
    const { svg } = parseAndRenderShape(xml);
    expect(svg).toContain('baseline-shift="super"');
  });

  it('baseline="-25000" → baseline-shift="sub"', () => {
    const xml = buildShapeXml(`
      <p:txBody>
        <a:bodyPr/>
        <a:p>
          <a:r><a:rPr lang="en-US" sz="1800"/><a:t>x</a:t></a:r>
          <a:r><a:rPr lang="en-US" sz="1200" baseline="-25000"/><a:t>i</a:t></a:r>
        </a:p>
      </p:txBody>
    `);
    const { svg } = parseAndRenderShape(xml);
    expect(svg).toContain('baseline-shift="sub"');
  });
});

// ---------------------------------------------------------------------------
// Priority 2: text color
// ---------------------------------------------------------------------------

describe("parse-render integration: text color", () => {
  it("srgbClr → fill 属性に直接反映", () => {
    const xml = buildShapeXml(`
      <p:txBody>
        <a:bodyPr/>
        <a:p>
          <a:r>
            <a:rPr lang="en-US" sz="1800">
              <a:solidFill><a:srgbClr val="FF0000"/></a:solidFill>
            </a:rPr>
            <a:t>Red</a:t>
          </a:r>
        </a:p>
      </p:txBody>
    `);
    const { svg } = parseAndRenderShape(xml);
    expect(svg).toContain('fill="#FF0000"');
  });

  it("schemeClr → テーマ色に解決されて fill 属性に反映", () => {
    const xml = buildShapeXml(`
      <p:txBody>
        <a:bodyPr/>
        <a:p>
          <a:r>
            <a:rPr lang="en-US" sz="1800">
              <a:solidFill><a:schemeClr val="accent1"/></a:solidFill>
            </a:rPr>
            <a:t>Themed</a:t>
          </a:r>
        </a:p>
      </p:txBody>
    `);
    const { svg } = parseAndRenderShape(xml);
    expect(svg).toContain('fill="#4472C4"');
  });
});

// ---------------------------------------------------------------------------
// Priority 2: bullets
// ---------------------------------------------------------------------------

describe("parse-render integration: bullets", () => {
  it("buChar → SVG に箇条書き文字が出力される", () => {
    const xml = buildShapeXml(`
      <p:txBody>
        <a:bodyPr/>
        <a:p>
          <a:pPr marL="342900" indent="-342900">
            <a:buChar char="\u2022"/>
          </a:pPr>
          <a:r><a:rPr lang="en-US" sz="1800"/><a:t>Item</a:t></a:r>
        </a:p>
      </p:txBody>
    `);
    const { svg } = parseAndRenderShape(xml);
    expect(svg).toContain("\u2022");
    expect(svg).toContain("Item");
  });

  it("buFont → 箇条書きの font-family 属性に反映", () => {
    const xml = buildShapeXml(`
      <p:txBody>
        <a:bodyPr/>
        <a:p>
          <a:pPr marL="342900" indent="-342900">
            <a:buFont typeface="Wingdings"/>
            <a:buChar char="v"/>
          </a:pPr>
          <a:r><a:rPr lang="en-US" sz="1800"/><a:t>Item</a:t></a:r>
        </a:p>
      </p:txBody>
    `);
    const { svg } = parseAndRenderShape(xml);
    expect(svg).toContain("Wingdings");
  });

  it("buClr → 箇条書きの fill 色に反映", () => {
    const xml = buildShapeXml(`
      <p:txBody>
        <a:bodyPr/>
        <a:p>
          <a:pPr marL="342900" indent="-342900">
            <a:buClr><a:srgbClr val="FF0000"/></a:buClr>
            <a:buChar char="\u2022"/>
          </a:pPr>
          <a:r><a:rPr lang="en-US" sz="1800"/><a:t>Item</a:t></a:r>
        </a:p>
      </p:txBody>
    `);
    const { svg } = parseAndRenderShape(xml);
    expect(svg).toContain('fill="#FF0000"');
    expect(svg).toContain("\u2022");
  });

  it("buSzPct → 箇条書きの font-size がスケーリングされる", () => {
    const xml = buildShapeXml(`
      <p:txBody>
        <a:bodyPr/>
        <a:p>
          <a:pPr marL="342900" indent="-342900">
            <a:buSzPct val="75000"/>
            <a:buChar char="\u2022"/>
          </a:pPr>
          <a:r><a:rPr lang="en-US" sz="1800"/><a:t>Item</a:t></a:r>
        </a:p>
      </p:txBody>
    `);
    const { svg } = parseAndRenderShape(xml);
    // 18pt * 75000/100000 = 13.5pt
    expect(svg).toContain('font-size="13.5pt"');
  });
});

// ---------------------------------------------------------------------------
// Priority 3: hyperlink
// ---------------------------------------------------------------------------

describe("parse-render integration: hyperlink", () => {
  it("hlinkClick → <a> タグで tspan がラップされる", () => {
    const xml = buildSpTreeXml(`
      <p:sp>
        <p:nvSpPr>
          <p:cNvPr id="2" name="TextBox 1"/>
          <p:cNvSpPr/>
          <p:nvPr/>
        </p:nvSpPr>
        <p:spPr>
          <a:xfrm>
            <a:off x="914400" y="914400"/>
            <a:ext cx="4572000" cy="2743200"/>
          </a:xfrm>
          <a:prstGeom prst="rect"/>
        </p:spPr>
        <p:txBody>
          <a:bodyPr/>
          <a:p>
            <a:r>
              <a:rPr lang="en-US" sz="1800">
                <a:hlinkClick r:id="rId1"/>
              </a:rPr>
              <a:t>Click me</a:t>
            </a:r>
          </a:p>
        </p:txBody>
      </p:sp>
    `);
    const rels = new Map<string, Relationship>([
      [
        "rId1",
        {
          id: "rId1",
          type: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink",
          target: "https://example.com",
          targetMode: "External",
        },
      ],
    ]);
    const { svg } = parseAndRenderShape(xml, rels);
    expect(svg).toContain('<a href="https://example.com">');
    expect(svg).toContain("Click me");
  });
});

// ---------------------------------------------------------------------------
// Priority 3: anchor (vertical alignment)
// ---------------------------------------------------------------------------

describe("parse-render integration: anchor", () => {
  it('anchor="ctr" → y 座標がデフォルト (top) より大きい', () => {
    const xmlCtr = buildShapeXml(`
      <p:txBody>
        <a:bodyPr anchor="ctr"/>
        <a:p>
          <a:r><a:rPr lang="en-US" sz="1800"/><a:t>Center</a:t></a:r>
        </a:p>
      </p:txBody>
    `);
    const xmlTop = buildShapeXml(`
      <p:txBody>
        <a:bodyPr anchor="t"/>
        <a:p>
          <a:r><a:rPr lang="en-US" sz="1800"/><a:t>Top</a:t></a:r>
        </a:p>
      </p:txBody>
    `);
    const { svg: svgCtr } = parseAndRenderShape(xmlCtr);
    const { svg: svgTop } = parseAndRenderShape(xmlTop);

    const yCtr = Number(svgCtr.match(/<text[^>]*y="([^"]+)"/)?.[1]);
    const yTop = Number(svgTop.match(/<text[^>]*y="([^"]+)"/)?.[1]);
    expect(yCtr).toBeGreaterThan(yTop);
  });
});

// ---------------------------------------------------------------------------
// Priority 3: fontScale (normAutofit)
// ---------------------------------------------------------------------------

describe("parse-render integration: fontScale", () => {
  it("normAutofit fontScale でフォントサイズがスケーリングされる", () => {
    const xml = buildShapeXml(`
      <p:txBody>
        <a:bodyPr>
          <a:normAutofit fontScale="62500"/>
        </a:bodyPr>
        <a:p>
          <a:r><a:rPr lang="en-US" sz="2400"/><a:t>Scaled</a:t></a:r>
        </a:p>
      </p:txBody>
    `);
    const { svg } = parseAndRenderShape(xml);
    // 24pt * 62500/100000 = 15pt
    expect(svg).toContain('font-size="15pt"');
  });
});

// ---------------------------------------------------------------------------
// Priority 3: transform (rotation, flip)
// ---------------------------------------------------------------------------

describe("parse-render integration: transform", () => {
  it('rot="5400000" (90度) → SVG に rotate(90,...) が含まれる', () => {
    const xml = buildSpTreeXml(`
      <p:sp>
        <p:nvSpPr>
          <p:cNvPr id="2" name="Shape 1"/>
          <p:cNvSpPr/>
          <p:nvPr/>
        </p:nvSpPr>
        <p:spPr>
          <a:xfrm rot="5400000">
            <a:off x="914400" y="914400"/>
            <a:ext cx="1828800" cy="914400"/>
          </a:xfrm>
          <a:prstGeom prst="rect"/>
          <a:solidFill><a:srgbClr val="FF0000"/></a:solidFill>
        </p:spPr>
      </p:sp>
    `);
    const { svg } = parseAndRenderShape(xml);
    expect(svg).toContain("rotate(90");
  });

  it('flipH="1" → SVG に scale(-1, 1) が含まれる', () => {
    const xml = buildSpTreeXml(`
      <p:sp>
        <p:nvSpPr>
          <p:cNvPr id="2" name="Shape 1"/>
          <p:cNvSpPr/>
          <p:nvPr/>
        </p:nvSpPr>
        <p:spPr>
          <a:xfrm flipH="1">
            <a:off x="914400" y="914400"/>
            <a:ext cx="1828800" cy="914400"/>
          </a:xfrm>
          <a:prstGeom prst="rect"/>
          <a:solidFill><a:srgbClr val="FF0000"/></a:solidFill>
        </p:spPr>
      </p:sp>
    `);
    const { svg } = parseAndRenderShape(xml);
    expect(svg).toContain("scale(-1, 1)");
  });
});

// ---------------------------------------------------------------------------
// Priority 3: fill types
// ---------------------------------------------------------------------------

describe("parse-render integration: fill types", () => {
  it("solidFill srgbClr → fill 属性に反映", () => {
    const xml = buildSpTreeXml(`
      <p:sp>
        <p:nvSpPr>
          <p:cNvPr id="2" name="Shape 1"/>
          <p:cNvSpPr/>
          <p:nvPr/>
        </p:nvSpPr>
        <p:spPr>
          <a:xfrm>
            <a:off x="914400" y="914400"/>
            <a:ext cx="1828800" cy="914400"/>
          </a:xfrm>
          <a:prstGeom prst="rect"/>
          <a:solidFill><a:srgbClr val="4472C4"/></a:solidFill>
        </p:spPr>
      </p:sp>
    `);
    const { svg } = parseAndRenderShape(xml);
    expect(svg).toContain('fill="#4472C4"');
  });

  it("gradFill → linearGradient + stop-color に反映", () => {
    const xml = buildSpTreeXml(`
      <p:sp>
        <p:nvSpPr>
          <p:cNvPr id="2" name="Shape 1"/>
          <p:cNvSpPr/>
          <p:nvPr/>
        </p:nvSpPr>
        <p:spPr>
          <a:xfrm>
            <a:off x="914400" y="914400"/>
            <a:ext cx="1828800" cy="914400"/>
          </a:xfrm>
          <a:prstGeom prst="rect"/>
          <a:gradFill>
            <a:gsLst>
              <a:gs pos="0"><a:srgbClr val="FF0000"/></a:gs>
              <a:gs pos="100000"><a:srgbClr val="0000FF"/></a:gs>
            </a:gsLst>
            <a:lin ang="5400000"/>
          </a:gradFill>
        </p:spPr>
      </p:sp>
    `);
    const { svg } = parseAndRenderShape(xml);
    expect(svg).toContain("<linearGradient");
    expect(svg).toContain("#FF0000");
    expect(svg).toContain("#0000FF");
  });
});
