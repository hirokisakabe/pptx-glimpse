import { bench, describe, beforeAll } from "vitest";
import JSZip from "jszip";
import { convertPptxToSvg, convertPptxToPng } from "../src/converter.js";
import { readPptx } from "../src/parser/pptx-reader.js";
import { parsePresentation } from "../src/parser/presentation-parser.js";
import { parseTheme } from "../src/parser/theme-parser.js";
import { parseSlideMasterColorMap } from "../src/parser/slide-master-parser.js";
import { parseSlide } from "../src/parser/slide-parser.js";
import {
  parseRelationships,
  resolveRelationshipTarget,
} from "../src/parser/relationship-parser.js";
import { ColorResolver } from "../src/color/color-resolver.js";
import { renderSlideToSvg } from "../src/renderer/svg-renderer.js";
import type { Slide } from "../src/model/slide.js";
import type { SlideSize } from "../src/model/presentation.js";
import type { ColorMap } from "../src/model/theme.js";

// ---------------------------------------------------------------------------
// XML templates (minimal PPTX structure)
// ---------------------------------------------------------------------------

const NAMESPACES = {
  a: "http://schemas.openxmlformats.org/drawingml/2006/main",
  r: "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
  p: "http://schemas.openxmlformats.org/presentationml/2006/main",
  rels: "http://schemas.openxmlformats.org/package/2006/relationships",
  ct: "http://schemas.openxmlformats.org/package/2006/content-types",
};

const rootRels = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="${NAMESPACES.rels}">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="ppt/presentation.xml"/>
</Relationships>`;

const theme1Xml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<a:theme xmlns:a="${NAMESPACES.a}" name="Office Theme">
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
      <a:majorFont><a:latin typeface="Calibri Light"/></a:majorFont>
      <a:minorFont><a:latin typeface="Calibri"/></a:minorFont>
    </a:fontScheme>
    <a:fmtScheme name="Office"/>
  </a:themeElements>
</a:theme>`;

const slideMasterXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:sldMaster xmlns:a="${NAMESPACES.a}" xmlns:r="${NAMESPACES.r}" xmlns:p="${NAMESPACES.p}">
  <p:cSld>
    <p:spTree>
      <p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>
      <p:grpSpPr/>
    </p:spTree>
  </p:cSld>
  <p:clrMap bg1="lt1" tx1="dk1" bg2="lt2" tx2="dk2" accent1="accent1" accent2="accent2" accent3="accent3" accent4="accent4" accent5="accent5" accent6="accent6" hlink="hlink" folHlink="folHlink"/>
  <p:sldLayoutIdLst><p:sldLayoutId r:id="rId1"/></p:sldLayoutIdLst>
</p:sldMaster>`;

const slideMasterRels = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="${NAMESPACES.rels}">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout" Target="../slideLayouts/slideLayout1.xml"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme" Target="../theme/theme1.xml"/>
</Relationships>`;

const slideLayoutXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:sldLayout xmlns:a="${NAMESPACES.a}" xmlns:r="${NAMESPACES.r}" xmlns:p="${NAMESPACES.p}" type="blank">
  <p:cSld>
    <p:spTree>
      <p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>
      <p:grpSpPr/>
    </p:spTree>
  </p:cSld>
</p:sldLayout>`;

const slideLayoutRels = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="${NAMESPACES.rels}">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideMaster" Target="../slideMasters/slideMaster1.xml"/>
</Relationships>`;

// ---------------------------------------------------------------------------
// Helper: defaults (copied from converter.ts private functions)
// ---------------------------------------------------------------------------

function defaultColorScheme() {
  return {
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
  };
}

function defaultColorMap(): ColorMap {
  return {
    bg1: "lt1" as const,
    tx1: "dk1" as const,
    bg2: "lt2" as const,
    tx2: "dk2" as const,
    accent1: "accent1" as const,
    accent2: "accent2" as const,
    accent3: "accent3" as const,
    accent4: "accent4" as const,
    accent5: "accent5" as const,
    accent6: "accent6" as const,
    hlink: "hlink" as const,
    folHlink: "folHlink" as const,
  };
}

// ---------------------------------------------------------------------------
// Helper: shape XML generation
// ---------------------------------------------------------------------------

const COLORS = [
  "4472C4",
  "ED7D31",
  "A5A5A5",
  "FFC000",
  "5B9BD5",
  "70AD47",
  "264478",
  "9B57A0",
  "FF6384",
  "36A2EB",
];

const PRESETS = ["rect", "ellipse", "roundRect", "triangle", "diamond"];

function shapeXml(
  id: number,
  name: string,
  opts: {
    x: number;
    y: number;
    cx: number;
    cy: number;
    preset: string;
    color: string;
    text?: string;
    outlineWidth?: number;
    outlineColor?: string;
  },
): string {
  const textBody = opts.text
    ? `<p:txBody>
        <a:bodyPr anchor="ctr"/>
        <a:lstStyle/>
        <a:p>
          <a:pPr algn="ctr"/>
          <a:r>
            <a:rPr lang="en-US" sz="1200">
              <a:solidFill><a:srgbClr val="FFFFFF"/></a:solidFill>
            </a:rPr>
            <a:t>${opts.text}</a:t>
          </a:r>
        </a:p>
      </p:txBody>`
    : "";

  const outline =
    opts.outlineWidth && opts.outlineColor
      ? `<a:ln w="${opts.outlineWidth}">
          <a:solidFill><a:srgbClr val="${opts.outlineColor}"/></a:solidFill>
        </a:ln>`
      : "";

  return `<p:sp>
    <p:nvSpPr>
      <p:cNvPr id="${id}" name="${name}"/>
      <p:cNvSpPr/>
      <p:nvPr/>
    </p:nvSpPr>
    <p:spPr>
      <a:xfrm>
        <a:off x="${opts.x}" y="${opts.y}"/>
        <a:ext cx="${opts.cx}" cy="${opts.cy}"/>
      </a:xfrm>
      <a:prstGeom prst="${opts.preset}"><a:avLst/></a:prstGeom>
      <a:solidFill><a:srgbClr val="${opts.color}"/></a:solidFill>
      ${outline}
    </p:spPr>
    ${textBody}
  </p:sp>`;
}

function wrapSlideXml(shapes: string): string {
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:sld xmlns:a="${NAMESPACES.a}" xmlns:r="${NAMESPACES.r}" xmlns:p="${NAMESPACES.p}">
  <p:cSld>
    <p:spTree>
      <p:nvGrpSpPr>
        <p:cNvPr id="1" name=""/>
        <p:cNvGrpSpPr/>
        <p:nvPr/>
      </p:nvGrpSpPr>
      <p:grpSpPr>
        <a:xfrm>
          <a:off x="0" y="0"/><a:ext cx="0" cy="0"/>
          <a:chOff x="0" y="0"/><a:chExt cx="0" cy="0"/>
        </a:xfrm>
      </p:grpSpPr>
      ${shapes}
    </p:spTree>
  </p:cSld>
</p:sld>`;
}

// ---------------------------------------------------------------------------
// PPTX builder
// ---------------------------------------------------------------------------

interface SlideEntry {
  xml: string;
  rels: string;
}

async function buildPptx(slides: SlideEntry[]): Promise<Buffer> {
  const zip = new JSZip();

  // Content types
  const slideOverrides = slides
    .map(
      (_, i) =>
        `<Override PartName="/ppt/slides/slide${i + 1}.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/>`,
    )
    .join("\n  ");

  zip.file(
    "[Content_Types].xml",
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="${NAMESPACES.ct}">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/ppt/presentation.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml"/>
  ${slideOverrides}
  <Override PartName="/ppt/slideLayouts/slideLayout1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slideLayout+xml"/>
  <Override PartName="/ppt/slideMasters/slideMaster1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slideMaster+xml"/>
  <Override PartName="/ppt/theme/theme1.xml" ContentType="application/vnd.openxmlformats-officedocument.theme+xml"/>
</Types>`,
  );

  zip.file("_rels/.rels", rootRels);

  // Presentation
  const sldIdList = slides
    .map((_, i) => `<p:sldId id="${256 + i}" r:id="rId${i + 2}"/>`)
    .join("\n    ");

  zip.file(
    "ppt/presentation.xml",
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:presentation xmlns:a="${NAMESPACES.a}" xmlns:r="${NAMESPACES.r}" xmlns:p="${NAMESPACES.p}">
  <p:sldMasterIdLst><p:sldMasterId r:id="rId1"/></p:sldMasterIdLst>
  <p:sldIdLst>
    ${sldIdList}
  </p:sldIdLst>
  <p:sldSz cx="9144000" cy="5143500" type="screen16x9"/>
</p:presentation>`,
  );

  // Presentation rels
  const slideRels = slides
    .map(
      (_, i) =>
        `<Relationship Id="rId${i + 2}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" Target="slides/slide${i + 1}.xml"/>`,
    )
    .join("\n  ");

  zip.file(
    "ppt/_rels/presentation.xml.rels",
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="${NAMESPACES.rels}">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideMaster" Target="slideMasters/slideMaster1.xml"/>
  ${slideRels}
  <Relationship Id="rId${slides.length + 2}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme" Target="theme/theme1.xml"/>
</Relationships>`,
  );

  // Slides
  for (let i = 0; i < slides.length; i++) {
    zip.file(`ppt/slides/slide${i + 1}.xml`, slides[i].xml);
    zip.file(`ppt/slides/_rels/slide${i + 1}.xml.rels`, slides[i].rels);
  }

  // Shared resources
  zip.file("ppt/slideMasters/slideMaster1.xml", slideMasterXml);
  zip.file("ppt/slideMasters/_rels/slideMaster1.xml.rels", slideMasterRels);
  zip.file("ppt/slideLayouts/slideLayout1.xml", slideLayoutXml);
  zip.file("ppt/slideLayouts/_rels/slideLayout1.xml.rels", slideLayoutRels);
  zip.file("ppt/theme/theme1.xml", theme1Xml);

  return zip.generateAsync({ type: "nodebuffer" });
}

const defaultSlideRels = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="${NAMESPACES.rels}">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout" Target="../slideLayouts/slideLayout1.xml"/>
</Relationships>`;

// ---------------------------------------------------------------------------
// Fixture generators
// ---------------------------------------------------------------------------

function createSimpleSlide(): SlideEntry {
  const shapes = shapeXml(2, "Rectangle 1", {
    x: 457200,
    y: 274638,
    cx: 3048000,
    cy: 1143000,
    preset: "rect",
    color: "4472C4",
    text: "Hello World",
    outlineWidth: 25400,
    outlineColor: "2F528F",
  });
  return { xml: wrapSlideXml(shapes), rels: defaultSlideRels };
}

function createComplexSlide(): SlideEntry {
  const SLIDE_W = 9144000;
  const SLIDE_H = 5143500;
  const COLS = 10;
  const ROWS = 5;
  const MARGIN = 40000;
  const cellW = Math.floor((SLIDE_W - MARGIN * (COLS + 1)) / COLS);
  const cellH = Math.floor((SLIDE_H - MARGIN * (ROWS + 1)) / ROWS);

  const shapes: string[] = [];
  for (let i = 0; i < 50; i++) {
    const col = i % COLS;
    const row = Math.floor(i / COLS);
    const x = MARGIN + col * (cellW + MARGIN);
    const y = MARGIN + row * (cellH + MARGIN);
    const preset = PRESETS[i % PRESETS.length];
    const color = COLORS[i % COLORS.length];
    const text = i % 3 === 0 ? `S${i}` : undefined;
    const outlineWidth = i % 2 === 0 ? 12700 : undefined;
    const outlineColor = i % 2 === 0 ? "000000" : undefined;

    shapes.push(
      shapeXml(i + 2, `Shape ${i}`, {
        x,
        y,
        cx: cellW,
        cy: cellH,
        preset,
        color,
        text,
        outlineWidth,
        outlineColor,
      }),
    );
  }

  return { xml: wrapSlideXml(shapes.join("\n")), rels: defaultSlideRels };
}

function createMultiSlideEntries(count: number): SlideEntry[] {
  return Array.from({ length: count }, (_, i) => {
    const shapes = [
      shapeXml(2, "Rect", {
        x: 200000,
        y: 200000,
        cx: 2000000,
        cy: 1000000,
        preset: "rect",
        color: COLORS[i % COLORS.length],
      }),
      shapeXml(3, "Ellipse", {
        x: 3000000,
        y: 200000,
        cx: 1500000,
        cy: 1500000,
        preset: "ellipse",
        color: COLORS[(i + 1) % COLORS.length],
      }),
      shapeXml(4, "TextBox", {
        x: 200000,
        y: 2500000,
        cx: 4000000,
        cy: 1000000,
        preset: "rect",
        color: COLORS[(i + 2) % COLORS.length],
        text: `Slide ${i + 1}`,
      }),
    ];
    return { xml: wrapSlideXml(shapes.join("\n")), rels: defaultSlideRels };
  });
}

// ---------------------------------------------------------------------------
// Parse pipeline helper (reproduces converter.ts internal logic)
// ---------------------------------------------------------------------------

async function parsePptx(input: Buffer): Promise<{ slides: Slide[]; slideSize: SlideSize }> {
  const archive = await readPptx(input);

  const presentationXml = archive.files.get("ppt/presentation.xml");
  if (!presentationXml) throw new Error("Missing presentation.xml");
  const presInfo = parsePresentation(presentationXml);

  const presRelsXml = archive.files.get("ppt/_rels/presentation.xml.rels");
  const presRels = presRelsXml ? parseRelationships(presRelsXml) : new Map();

  let theme = {
    colorScheme: defaultColorScheme(),
    fontScheme: {
      majorFont: "Calibri",
      minorFont: "Calibri",
      majorFontEa: null as string | null,
      minorFontEa: null as string | null,
    },
  };
  for (const [, rel] of presRels) {
    if (rel.type.includes("theme")) {
      const themePath = resolveRelationshipTarget("ppt/presentation.xml", rel.target);
      const themeXml = archive.files.get(themePath);
      if (themeXml) theme = parseTheme(themeXml);
      break;
    }
  }

  let colorMap: ColorMap = defaultColorMap();
  for (const [, rel] of presRels) {
    if (rel.type.includes("slideMaster")) {
      const masterPath = resolveRelationshipTarget("ppt/presentation.xml", rel.target);
      const masterXml = archive.files.get(masterPath);
      if (masterXml) colorMap = parseSlideMasterColorMap(masterXml);
      break;
    }
  }

  const colorResolver = new ColorResolver(theme.colorScheme, colorMap);

  const slides: Slide[] = [];
  for (let i = 0; i < presInfo.slideRIds.length; i++) {
    const rId = presInfo.slideRIds[i];
    const rel = presRels.get(rId);
    if (!rel) continue;
    const path = resolveRelationshipTarget("ppt/presentation.xml", rel.target);
    const slideXml = archive.files.get(path);
    if (!slideXml) continue;
    slides.push(parseSlide(slideXml, path, i + 1, archive, colorResolver, theme.fontScheme));
  }

  return { slides, slideSize: presInfo.slideSize };
}

// ---------------------------------------------------------------------------
// Fixture buffers (populated in beforeAll)
// ---------------------------------------------------------------------------

let simplePptx: Buffer;
let complexPptx: Buffer;
let multiSlide10Pptx: Buffer;
let multiSlide50Pptx: Buffer;

let parsedSimpleSlide: Slide;
let parsedComplexSlide: Slide;
let slideSize: SlideSize;

beforeAll(async () => {
  [simplePptx, complexPptx, multiSlide10Pptx, multiSlide50Pptx] = await Promise.all([
    buildPptx([createSimpleSlide()]),
    buildPptx([createComplexSlide()]),
    buildPptx(createMultiSlideEntries(10)),
    buildPptx(createMultiSlideEntries(50)),
  ]);

  // Pre-parse slides for renderer-only benchmarks
  const simpleResult = await parsePptx(simplePptx);
  parsedSimpleSlide = simpleResult.slides[0];
  slideSize = simpleResult.slideSize;

  const complexResult = await parsePptx(complexPptx);
  parsedComplexSlide = complexResult.slides[0];
});

// ---------------------------------------------------------------------------
// Benchmarks
// ---------------------------------------------------------------------------

describe("E2E conversion", () => {
  bench("simple slide (1 shape) → SVG", async () => {
    await convertPptxToSvg(simplePptx);
  });

  bench("complex slide (50+ shapes) → SVG", async () => {
    await convertPptxToSvg(complexPptx);
  });

  bench("10 slides → SVG", async () => {
    await convertPptxToSvg(multiSlide10Pptx);
  });

  bench("50 slides → SVG", async () => {
    await convertPptxToSvg(multiSlide50Pptx);
  });
});

describe("PNG conversion", () => {
  bench("simple slide → PNG", async () => {
    await convertPptxToPng(simplePptx);
  });

  bench("complex slide → PNG", async () => {
    await convertPptxToPng(complexPptx);
  });
});

describe("parser standalone", () => {
  bench("parse simple slide", async () => {
    await parsePptx(simplePptx);
  });

  bench("parse complex slide", async () => {
    await parsePptx(complexPptx);
  });
});

describe("renderer standalone", () => {
  bench("render simple slide → SVG", () => {
    renderSlideToSvg(parsedSimpleSlide, slideSize);
  });

  bench("render complex slide → SVG", () => {
    renderSlideToSvg(parsedComplexSlide, slideSize);
  });
});
