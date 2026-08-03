import { Buffer } from "node:buffer";

import {
  addChart,
  addShape,
  asEmu,
  createPptx,
  readPptx,
  type SourceConnector,
  type SourceShape,
  writePptx,
} from "@pptx-glimpse/document";
import JSZip from "jszip";
import { expect } from "vitest";

import { isPptxEditorError, PptxEditorError } from "./index.js";
import { type PptxEditorErrorCode } from "./pptx-editor-session.js";

const encoder = new TextEncoder();

export const RED_PNG = pngBytes(
  "iVBORw0KGgoAAAANSUhEUgAAAAQAAAAECAIAAAAmkwkpAAAAEUlEQVR4nGP8z4AATEhsPBwAM9EBBzDn4UwAAAAASUVORK5CYII=",
);

export const BLUE_PNG = pngBytes(
  "iVBORw0KGgoAAAANSUhEUgAAAAQAAAAECAIAAAAmkwkpAAAAE0lEQVR4nGNkYPjPAANMcBZeDgAx0wEH1s7nlgAAAABJRU5ErkJggg==",
);

export function buildGroupCommandFixture(): Uint8Array {
  let source = createPptx();
  const slideHandle = source.slides[0]?.handle;
  if (slideHandle === undefined) throw new Error("group command slide handle is missing");
  for (const offsetX of [914400, 2743200, 4572000]) {
    source = addShape(source, slideHandle, {
      geometry: { kind: "preset", preset: "rect" },
      offsetX: asEmu(offsetX),
      offsetY: asEmu(914400),
      width: asEmu(1371600),
      height: asEmu(914400),
    });
  }
  return writePptx(source);
}

export async function buildScatterChartFixture(): Promise<Uint8Array> {
  let source = createPptx();
  const slideHandle = source.slides[0]?.handle;
  if (slideHandle === undefined) throw new Error("scatter fixture slide handle is missing");
  source = addChart(source, slideHandle, {
    chartType: "bar",
    offsetX: asEmu(500000),
    offsetY: asEmu(500000),
    width: asEmu(5000000),
    height: asEmu(3000000),
    series: [
      { name: "Original 1", categories: ["A", "B"], values: [1, 2] },
      { name: "Original 2", categories: ["A", "B"], values: [3, 4] },
    ],
  });
  const pptx = await JSZip.loadAsync(writePptx(source));
  const chartFile = pptx.file("ppt/charts/chart1.xml");
  const workbookFile = pptx.file("ppt/embeddings/Microsoft_Excel_Worksheet1.xlsx");
  if (chartFile === null || workbookFile === null) throw new Error("scatter fixture parts missing");
  const chart = await chartFile.async("string");
  const scatterSeries = (
    index: number,
    name: string,
    headerRow: number,
    x: number[],
    y: number[],
  ) => {
    const lastRow = headerRow + x.length;
    const points = (values: number[]) =>
      values.map((value, point) => `<c:pt idx="${point}"><c:v>${value}</c:v></c:pt>`).join("");
    return `<c:ser><c:idx val="${index}"/><c:order val="${index}"/><c:tx><c:strRef><c:f>Sheet1!$B$${headerRow}</c:f><c:strCache><c:ptCount val="1"/><c:pt idx="0"><c:v>${name}</c:v></c:pt></c:strCache></c:strRef></c:tx><c:xVal><c:numRef><c:f>Sheet1!$A$${headerRow + 1}:$A$${lastRow}</c:f><c:numCache><c:formatCode>General</c:formatCode><c:ptCount val="${x.length}"/>${points(x)}</c:numCache></c:numRef></c:xVal><c:yVal><c:numRef><c:f>Sheet1!$B$${headerRow + 1}:$B$${lastRow}</c:f><c:numCache><c:formatCode>General</c:formatCode><c:ptCount val="${y.length}"/>${points(y)}</c:numCache></c:numRef></c:yVal></c:ser>`;
  };
  pptx.file(
    "ppt/charts/chart1.xml",
    replaceCategoryAxisWithValueAxis(
      chart.replace(
        /<c:barChart>.*?<\/c:barChart>/,
        `<c:scatterChart><c:scatterStyle val="lineMarker"/><c:varyColors val="0"/>${scatterSeries(0, "Original 1", 1, [1, 2], [1, 2])}${scatterSeries(1, "Original 2", 5, [3, 4], [3, 4])}<c:axId val="100002"/><c:axId val="100003"/></c:scatterChart>`,
      ),
    ),
  );
  const workbook = await JSZip.loadAsync(await workbookFile.async("uint8array"));
  workbook.file(
    "xl/worksheets/sheet1.xml",
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><dimension ref="A1:B7"/><sheetViews><sheetView workbookViewId="0"/></sheetViews><sheetFormatPr defaultRowHeight="15"/><sheetData><row r="1"><c r="B1" t="inlineStr"><is><t>Original 1</t></is></c></row><row r="2"><c r="A2"><v>1</v></c><c r="B2"><v>1</v></c></row><row r="3"><c r="A3"><v>2</v></c><c r="B3"><v>2</v></c></row><row r="5"><c r="B5" t="inlineStr"><is><t>Original 2</t></is></c></row><row r="6"><c r="A6"><v>3</v></c><c r="B6"><v>3</v></c></row><row r="7"><c r="A7"><v>4</v></c><c r="B7"><v>4</v></c></row></sheetData></worksheet>`,
  );
  pptx.file(
    "ppt/embeddings/Microsoft_Excel_Worksheet1.xlsx",
    await workbook.generateAsync({ type: "uint8array" }),
  );
  return pptx.generateAsync({ type: "uint8array" });
}

export async function buildBubbleChartFixture(): Promise<Uint8Array> {
  const pptx = await JSZip.loadAsync(await buildScatterChartFixture());
  const chartFile = pptx.file("ppt/charts/chart1.xml");
  const workbookFile = pptx.file("ppt/embeddings/Microsoft_Excel_Worksheet1.xlsx");
  if (chartFile === null || workbookFile === null) throw new Error("bubble fixture parts missing");
  let chart = (await chartFile.async("string"))
    .replace(
      '<c:scatterChart><c:scatterStyle val="lineMarker"/><c:varyColors val="0"/>',
      '<c:bubbleChart><c:varyColors val="0"/>',
    )
    .replace("</c:scatterChart>", '<c:bubbleScale val="100"/></c:bubbleChart>');
  for (const [range, values] of [
    ["2:$C$3", [4, 8]],
    ["6:$C$7", [6, 9]],
  ] as const) {
    const points = values
      .map((value, point) => `<c:pt idx="${point}"><c:v>${value}</c:v></c:pt>`)
      .join("");
    chart = chart.replace(
      "</c:yVal></c:ser>",
      `</c:yVal><c:bubbleSize><c:numRef><c:f>Sheet1!$C$${range}</c:f><c:numCache><c:formatCode>General</c:formatCode><c:ptCount val="${values.length}"/>${points}</c:numCache></c:numRef></c:bubbleSize></c:ser>`,
    );
  }
  pptx.file("ppt/charts/chart1.xml", chart);
  const workbook = await JSZip.loadAsync(await workbookFile.async("uint8array"));
  workbook.file(
    "xl/worksheets/sheet1.xml",
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><dimension ref="A1:C7"/><sheetViews><sheetView workbookViewId="0"/></sheetViews><sheetFormatPr defaultRowHeight="15"/><sheetData><row r="1"><c r="B1" t="inlineStr"><is><t>Original 1</t></is></c><c r="C1" t="inlineStr"><is><t>Size</t></is></c></row><row r="2"><c r="A2"><v>1</v></c><c r="B2"><v>1</v></c><c r="C2"><v>4</v></c></row><row r="3"><c r="A3"><v>2</v></c><c r="B3"><v>2</v></c><c r="C3"><v>8</v></c></row><row r="5"><c r="B5" t="inlineStr"><is><t>Original 2</t></is></c><c r="C5" t="inlineStr"><is><t>Size</t></is></c></row><row r="6"><c r="A6"><v>3</v></c><c r="B6"><v>3</v></c><c r="C6"><v>6</v></c></row><row r="7"><c r="A7"><v>4</v></c><c r="B7"><v>4</v></c><c r="C7"><v>9</v></c></row></sheetData></worksheet>`,
  );
  pptx.file(
    "ppt/embeddings/Microsoft_Excel_Worksheet1.xlsx",
    await workbook.generateAsync({ type: "uint8array" }),
  );
  return pptx.generateAsync({ type: "uint8array" });
}

function replaceCategoryAxisWithValueAxis(chartXml: string): string {
  const categoryAxis = /<c:catAx>.*?<\/c:catAx>/.exec(chartXml)?.[0];
  if (categoryAxis === undefined) throw new Error("fixture category axis not found");
  const valueAxis = categoryAxis
    .replace("<c:catAx>", "<c:valAx>")
    .replace("</c:catAx>", '<c:crossBetween val="midCat"/></c:valAx>')
    .replace(/<c:(?:auto|lblAlgn|lblOffset|noMultiLvlLbl)\b[^>]*\/>/g, "");
  return chartXml.replace(categoryAxis, valueAxis);
}

export async function buildLayoutCatalogFixture(): Promise<Uint8Array> {
  const zip = new JSZip();
  zip.file(
    "[Content_Types].xml",
    xml(
      `<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">` +
        `<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>` +
        `<Default Extension="xml" ContentType="application/xml"/>` +
        `<Override PartName="/ppt/presentation.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml"/>` +
        [1, 2, 3]
          .map(
            (number) =>
              `<Override PartName="/ppt/slides/slide${String(number)}.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/>`,
          )
          .join("") +
        [1, 2, 3]
          .map(
            (number) =>
              `<Override PartName="/ppt/slideLayouts/slideLayout${String(number)}.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slideLayout+xml"/>`,
          )
          .join("") +
        [1, 2]
          .map(
            (number) =>
              `<Override PartName="/ppt/slideMasters/slideMaster${String(number)}.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slideMaster+xml"/>`,
          )
          .join("") +
        `</Types>`,
    ),
  );
  zip.file(
    "_rels/.rels",
    xml(
      `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">` +
        `<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="ppt/presentation.xml"/>` +
        `</Relationships>`,
    ),
  );
  zip.file(
    "ppt/presentation.xml",
    xml(
      `<p:presentation xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">` +
        `<p:sldMasterIdLst><p:sldMasterId id="2147483648" r:id="rIdMaster2"/><p:sldMasterId id="2147483649" r:id="rIdMaster1"/></p:sldMasterIdLst>` +
        `<p:sldIdLst><p:sldId id="256" r:id="rIdSlide1"/><p:sldId id="257" r:id="rIdSlide2"/><p:sldId id="258" r:id="rIdSlide3"/></p:sldIdLst>` +
        `<p:sldSz cx="9144000" cy="5143500"/>` +
        `</p:presentation>`,
    ),
  );
  zip.file(
    "ppt/_rels/presentation.xml.rels",
    xml(
      `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">` +
        `<Relationship Id="rIdMaster1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideMaster" Target="slideMasters/slideMaster1.xml"/>` +
        `<Relationship Id="rIdMaster2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideMaster" Target="slideMasters/slideMaster2.xml"/>` +
        [1, 2, 3]
          .map(
            (number) =>
              `<Relationship Id="rIdSlide${String(number)}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" Target="slides/slide${String(number)}.xml"/>`,
          )
          .join("") +
        `</Relationships>`,
    ),
  );

  zip.file("ppt/slideMasters/slideMaster1.xml", slideMasterXml("First Master", [[1, 1]]));
  zip.file(
    "ppt/slideMasters/slideMaster2.xml",
    slideMasterXml("Second Master", [
      [3, 3],
      [2, 2],
    ]),
  );
  zip.file("ppt/slideMasters/_rels/slideMaster1.xml.rels", masterLayoutRels([1]));
  zip.file("ppt/slideMasters/_rels/slideMaster2.xml.rels", masterLayoutRels([2, 3]));

  zip.file("ppt/slideLayouts/slideLayout1.xml", slideLayoutXml("Visible by Default", "blank"));
  zip.file("ppt/slideLayouts/slideLayout2.xml", slideLayoutXml("Popular Layout", "twoObj", true));
  zip.file("ppt/slideLayouts/slideLayout3.xml", slideLayoutXml("Hidden Layout", "title", false));
  for (const [layoutNumber, masterNumber] of [
    [1, 1],
    [2, 2],
    [3, 2],
  ] as const) {
    zip.file(
      `ppt/slideLayouts/_rels/slideLayout${String(layoutNumber)}.xml.rels`,
      xml(
        `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">` +
          `<Relationship Id="rIdMaster" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideMaster" Target="../slideMasters/slideMaster${String(masterNumber)}.xml"/>` +
          `</Relationships>`,
      ),
    );
  }

  for (const [slideNumber, layoutNumber] of [
    [1, 1],
    [2, 2],
    [3, 2],
  ] as const) {
    zip.file(
      `ppt/slides/slide${String(slideNumber)}.xml`,
      xml(
        `<p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"><p:cSld><p:spTree/></p:cSld></p:sld>`,
      ),
    );
    zip.file(
      `ppt/slides/_rels/slide${String(slideNumber)}.xml.rels`,
      xml(
        `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">` +
          `<Relationship Id="rIdLayout" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout" Target="../slideLayouts/slideLayout${String(layoutNumber)}.xml"/>` +
          `</Relationships>`,
      ),
    );
  }

  return zip.generateAsync({ type: "uint8array" });
}

function slideMasterXml(name: string, layouts: readonly (readonly [number, number])[]): Uint8Array {
  return xml(
    `<p:sldMaster xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">` +
      `<p:cSld name="${name}"><p:spTree/></p:cSld>` +
      `<p:sldLayoutIdLst>${layouts
        .map(
          ([id, relationshipNumber]) =>
            `<p:sldLayoutId id="${String(2147483648 + id)}" r:id="rIdLayout${String(relationshipNumber)}"/>`,
        )
        .join("")}</p:sldLayoutIdLst>` +
      `</p:sldMaster>`,
  );
}

function masterLayoutRels(layoutNumbers: readonly number[]): Uint8Array {
  return xml(
    `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">${layoutNumbers
      .map(
        (number) =>
          `<Relationship Id="rIdLayout${String(number)}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout" Target="../slideLayouts/slideLayout${String(number)}.xml"/>`,
      )
      .join("")}</Relationships>`,
  );
}

function slideLayoutXml(name: string, type: string, show?: boolean): Uint8Array {
  return xml(
    `<p:sldLayout xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" type="${type}"${show === undefined ? "" : ` show="${show ? "1" : "0"}"`}>` +
      `<p:cSld name="${name}"><p:spTree/></p:cSld>` +
      `</p:sldLayout>`,
  );
}

export async function capturePptxEditorError(promise: Promise<unknown>): Promise<PptxEditorError> {
  try {
    await promise;
  } catch (error) {
    if (isPptxEditorError(error)) return error;
    throw error;
  }
  throw new Error("expected PptxEditorError");
}

export function captureSynchronousPptxEditorError(operation: () => unknown): PptxEditorError {
  try {
    operation();
  } catch (error) {
    if (isPptxEditorError(error)) return error;
    throw error;
  }
  throw new Error("expected PptxEditorError");
}

export function expectErrorCodeAndCause(
  error: PptxEditorError,
  code: PptxEditorErrorCode,
  cause?: unknown,
): void {
  expect(error.code).toBe(code);
  expect(error.name).toBe("PptxEditorError");
  if (cause === undefined) {
    expect(error.cause).toBeDefined();
  } else {
    expect(error.cause).toBe(cause);
  }
}

function xml(content: string): Uint8Array {
  return encoder.encode(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n${content}`);
}

function pngBytes(base64: string): Uint8Array {
  return new Uint8Array(Buffer.from(base64, "base64"));
}

export async function buildShapeFixture(
  options: { includeNoTransformShape?: boolean; includeConnector?: boolean } = {},
): Promise<Uint8Array> {
  const zip = new JSZip();
  zip.file(
    "[Content_Types].xml",
    xml(
      `<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">` +
        `<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>` +
        `<Default Extension="xml" ContentType="application/xml"/>` +
        `<Override PartName="/ppt/presentation.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml"/>` +
        `<Override PartName="/ppt/slides/slide1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/>` +
        `</Types>`,
    ),
  );
  zip.file(
    "_rels/.rels",
    xml(
      `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">` +
        `<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="ppt/presentation.xml"/>` +
        `</Relationships>`,
    ),
  );
  zip.file(
    "ppt/presentation.xml",
    xml(
      `<p:presentation xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">` +
        `<p:sldIdLst><p:sldId id="256" r:id="rIdSlide1"/></p:sldIdLst>` +
        `<p:sldSz cx="9144000" cy="5143500"/>` +
        `</p:presentation>`,
    ),
  );
  zip.file(
    "ppt/_rels/presentation.xml.rels",
    xml(
      `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">` +
        `<Relationship Id="rIdSlide1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" Target="slides/slide1.xml"/>` +
        `</Relationships>`,
    ),
  );
  zip.file(
    "ppt/slides/slide1.xml",
    xml(
      `<p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">` +
        `<p:cSld><p:spTree>` +
        `<p:sp><p:nvSpPr><p:cNvPr id="10" name="Box"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>` +
        `<p:spPr><a:xfrm><a:off x="914400" y="1828800"/><a:ext cx="2743200" cy="914400"/></a:xfrm><a:prstGeom prst="rect"/></p:spPr>` +
        `<p:txBody><a:bodyPr/><a:lstStyle/>` +
        `<a:p><a:r><a:rPr sz="2400"><a:latin typeface="Aptos"/></a:rPr><a:t>Original</a:t></a:r></a:p>` +
        `</p:txBody>` +
        `</p:sp>` +
        (options.includeNoTransformShape
          ? `<p:sp><p:nvSpPr><p:cNvPr id="11" name="No Transform"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>` +
            `<p:spPr><a:prstGeom prst="rect"/></p:spPr>` +
            `<p:txBody><a:bodyPr/><a:lstStyle/>` +
            `<a:p><a:r><a:t>No transform</a:t></a:r></a:p>` +
            `</p:txBody>` +
            `</p:sp>`
          : "") +
        (options.includeConnector
          ? `<p:cxnSp><p:nvCxnSpPr><p:cNvPr id="12" name="Connector"/><p:cNvCxnSpPr>` +
            `<a:stCxn id="10" idx="0"/></p:cNvCxnSpPr><p:nvPr/></p:nvCxnSpPr>` +
            `<p:spPr><a:xfrm><a:off x="914400" y="914400"/><a:ext cx="914400" cy="914400"/></a:xfrm>` +
            `<a:prstGeom prst="straightConnector1"/></p:spPr></p:cxnSp>`
          : "") +
        `</p:spTree></p:cSld>` +
        `</p:sld>`,
    ),
  );

  return zip.generateAsync({ type: "uint8array" });
}

export async function buildImageFixture(
  options: {
    readonly includeImageFill?: boolean;
    readonly includeUnusedImageRelationship?: boolean;
    readonly includeSecondSlide?: boolean;
  } = {},
): Promise<Uint8Array> {
  const zip = new JSZip();
  zip.file(
    "[Content_Types].xml",
    xml(
      `<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">` +
        `<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>` +
        `<Default Extension="xml" ContentType="application/xml"/>` +
        `<Default Extension="png" ContentType="image/png"/>` +
        `<Override PartName="/ppt/presentation.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml"/>` +
        `<Override PartName="/ppt/slides/slide1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/>` +
        (options.includeSecondSlide === true
          ? `<Override PartName="/ppt/slides/slide2.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/>`
          : "") +
        `</Types>`,
    ),
  );
  zip.file(
    "_rels/.rels",
    xml(
      `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">` +
        `<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="ppt/presentation.xml"/>` +
        `</Relationships>`,
    ),
  );
  zip.file(
    "ppt/presentation.xml",
    xml(
      `<p:presentation xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">` +
        `<p:sldIdLst><p:sldId id="256" r:id="rIdSlide1"/>` +
        (options.includeSecondSlide === true ? `<p:sldId id="257" r:id="rIdSlide2"/>` : "") +
        `</p:sldIdLst>` +
        `<p:sldSz cx="9144000" cy="5143500"/>` +
        `</p:presentation>`,
    ),
  );
  zip.file(
    "ppt/_rels/presentation.xml.rels",
    xml(
      `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">` +
        `<Relationship Id="rIdSlide1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" Target="slides/slide1.xml"/>` +
        (options.includeSecondSlide === true
          ? `<Relationship Id="rIdSlide2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" Target="slides/slide2.xml"/>`
          : "") +
        `</Relationships>`,
    ),
  );
  zip.file(
    "ppt/slides/slide1.xml",
    xml(
      `<p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">` +
        `<p:cSld><p:spTree>` +
        `<p:pic><p:nvPicPr><p:cNvPr id="20" name="Shared Picture A"/><p:cNvPicPr/><p:nvPr/></p:nvPicPr>` +
        `<p:blipFill><a:blip r:embed="rIdImage"/><a:stretch><a:fillRect/></a:stretch></p:blipFill>` +
        `<p:spPr><a:xfrm><a:off x="914400" y="914400"/><a:ext cx="914400" cy="914400"/></a:xfrm><a:prstGeom prst="rect"/></p:spPr></p:pic>` +
        `<p:pic><p:nvPicPr><p:cNvPr id="21" name="Shared Picture B"/><p:cNvPicPr/><p:nvPr/></p:nvPicPr>` +
        `<p:blipFill><a:blip r:embed="rIdImage"/><a:stretch><a:fillRect/></a:stretch></p:blipFill>` +
        `<p:spPr><a:xfrm><a:off x="1828800" y="914400"/><a:ext cx="914400" cy="914400"/></a:xfrm><a:prstGeom prst="rect"/></p:spPr></p:pic>` +
        (options.includeImageFill === true
          ? `<p:sp><p:nvSpPr><p:cNvPr id="22" name="Shared Fill"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>` +
            `<p:spPr><a:blipFill><a:blip r:embed="rIdImage"/><a:stretch><a:fillRect/></a:stretch></a:blipFill>` +
            `<a:prstGeom prst="rect"/></p:spPr></p:sp>`
          : "") +
        `</p:spTree></p:cSld>` +
        `</p:sld>`,
    ),
  );
  zip.file(
    "ppt/slides/_rels/slide1.xml.rels",
    xml(
      `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">` +
        `<Relationship Id="rIdImage" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="../media/image1.png"/>` +
        (options.includeUnusedImageRelationship === true
          ? `<Relationship Id="rIdUnusedImage" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="../media/image1.png"/>`
          : "") +
        `</Relationships>`,
    ),
  );
  if (options.includeSecondSlide === true) {
    const firstSlide = zip.file("ppt/slides/slide1.xml");
    if (firstSlide === null) throw new Error("first slide fixture was not created");
    zip.file("ppt/slides/slide2.xml", firstSlide.async("uint8array"));
    zip.file(
      "ppt/slides/_rels/slide2.xml.rels",
      xml(
        `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">` +
          `<Relationship Id="rIdImage" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="../media/image1.png"/>` +
          `</Relationships>`,
      ),
    );
  }
  zip.file("ppt/media/image1.png", RED_PNG);

  return zip.generateAsync({ type: "uint8array" });
}

export async function buildTwoSlideFixture(): Promise<Uint8Array> {
  const zip = new JSZip();
  zip.file(
    "[Content_Types].xml",
    xml(
      `<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">` +
        `<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>` +
        `<Default Extension="xml" ContentType="application/xml"/>` +
        `<Override PartName="/ppt/presentation.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml"/>` +
        `<Override PartName="/ppt/slides/slide1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/>` +
        `<Override PartName="/ppt/slides/slide2.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/>` +
        `<Override PartName="/ppt/slideLayouts/slideLayout1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slideLayout+xml"/>` +
        `</Types>`,
    ),
  );
  zip.file(
    "_rels/.rels",
    xml(
      `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">` +
        `<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="ppt/presentation.xml"/>` +
        `</Relationships>`,
    ),
  );
  zip.file(
    "ppt/presentation.xml",
    xml(
      `<p:presentation xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">` +
        `<p:sldIdLst><p:sldId id="256" r:id="rIdSlide1"/><p:sldId id="257" r:id="rIdSlide2"/></p:sldIdLst>` +
        `<p:sldSz cx="9144000" cy="5143500"/>` +
        `</p:presentation>`,
    ),
  );
  zip.file(
    "ppt/_rels/presentation.xml.rels",
    xml(
      `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">` +
        `<Relationship Id="rIdSlide1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" Target="slides/slide1.xml"/>` +
        `<Relationship Id="rIdSlide2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" Target="slides/slide2.xml"/>` +
        `</Relationships>`,
    ),
  );
  zip.file("ppt/slides/slide1.xml", textSlideXml(10, "First"));
  zip.file("ppt/slides/slide2.xml", textSlideXml(20, "Second"));
  const slideLayoutRelationship = xml(
    `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">` +
      `<Relationship Id="rIdLayout" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout" Target="../slideLayouts/slideLayout1.xml"/>` +
      `</Relationships>`,
  );
  zip.file("ppt/slides/_rels/slide1.xml.rels", slideLayoutRelationship);
  zip.file("ppt/slides/_rels/slide2.xml.rels", slideLayoutRelationship);
  zip.file(
    "ppt/slideLayouts/slideLayout1.xml",
    xml(
      `<p:sldLayout xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">` +
        `<p:cSld><p:spTree/></p:cSld>` +
        `</p:sldLayout>`,
    ),
  );

  return zip.generateAsync({ type: "uint8array" });
}

function textSlideXml(shapeId: number, text: string): Uint8Array {
  return xml(
    `<p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">` +
      `<p:cSld><p:spTree>` +
      `<p:sp><p:nvSpPr><p:cNvPr id="${String(shapeId)}" name="${text}"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>` +
      `<p:spPr><a:xfrm><a:off x="914400" y="914400"/><a:ext cx="2743200" cy="914400"/></a:xfrm><a:prstGeom prst="rect"/></p:spPr>` +
      `<p:txBody><a:bodyPr/><a:lstStyle/><a:p><a:r><a:t>${text}</a:t></a:r></a:p></p:txBody>` +
      `</p:sp>` +
      `</p:spTree></p:cSld>` +
      `</p:sld>`,
  );
}

export function firstShape(source: ReturnType<typeof readPptx>): SourceShape {
  const shape = source.slides[0]?.shapes.find((node): node is SourceShape => node.kind === "shape");
  if (shape === undefined) throw new Error("fixture shape not found");
  return shape;
}

export function firstText(source: ReturnType<typeof readPptx>): string {
  const run = firstShape(source).textBody?.paragraphs[0]?.runs[0];
  if (run === undefined) throw new Error("fixture text run not found");
  return run.text;
}

export function shapeByText(source: ReturnType<typeof readPptx>, text: string): SourceShape {
  const shape = source.slides[0]?.shapes.find(
    (node): node is SourceShape =>
      node.kind === "shape" &&
      node.textBody?.paragraphs.some((paragraph) =>
        paragraph.runs.some((run) => run.text === text),
      ) === true,
  );
  if (shape === undefined) throw new Error(`shape text not found: ${text}`);
  return shape;
}

export function connectorByName(
  source: ReturnType<typeof readPptx>,
  name: string,
): SourceConnector {
  const connector = source.slides[0]?.shapes.find(
    (node): node is SourceConnector => node.kind === "connector" && node.name === name,
  );
  if (connector === undefined) throw new Error(`connector not found: ${name}`);
  return connector;
}

export function handleKey(handle: unknown): string {
  if (handle === undefined || handle === null || typeof handle !== "object") return "";
  const value = handle as {
    partPath?: string;
    nodeId?: string;
    relationshipId?: string;
    orderingSlot?: number;
  };
  return [
    value.partPath ?? "",
    value.nodeId ?? "",
    value.relationshipId ?? "",
    value.orderingSlot ?? "",
  ].join("\u0000");
}

export function mediaBytes(source: ReturnType<typeof readPptx>, partPath: string): Uint8Array {
  const media = source.packageGraph.media.find((part) => part.partPath === partPath);
  if (media === undefined) throw new Error(`media not found: ${partPath}`);
  return media.bytes;
}
