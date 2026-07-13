import { Buffer } from "node:buffer";
import { spawnSync } from "node:child_process";
import { copyFileSync, existsSync, mkdtempSync, readFileSync, writeFileSync } from "node:fs";
import { tmpdir } from "node:os";
import { basename, dirname, join } from "node:path";
import { fileURLToPath } from "node:url";

import sharp from "sharp";

import {
  addChart,
  addConnector,
  addEmptySlideFromLayout,
  addPicture,
  addShape,
  addSlideNumber,
  addTable,
  addTextBox,
  asEmu,
  asHundredthPt,
  asOoxmlAngle,
  asOoxmlPercent,
  asPt,
  createPptx,
  deleteShape,
  deleteSlide,
  duplicateSlide,
  moveSlide,
  readPptx,
  replaceImageBytes,
  setSlideBackground,
  type SourceConnector,
  type SourceHandle,
  type SourceImage,
  type SourceParagraph,
  type SourceShape,
  type SourceShapeNode,
  type SourceTextRun,
  writePptx,
} from "../../packages/document/src/index.js";
import { createEditorSession } from "../../packages/editor-core/src/index.js";
import { compareImageBuffers } from "../compare-utils.js";

const __dirname = dirname(fileURLToPath(import.meta.url));
const FIXTURE_DIR = join(__dirname, "fixtures");
const DIFF_DIR = join(__dirname, "diffs");
const LIBREOFFICE_IMAGE_CANDIDATES = [
  process.env.PPTX_GLIMPSE_LO_VRT_IMAGE,
  "ghcr.io/hirokisakabe/pptx-glimpse-vrt:latest",
  "pptx-glimpse-vrt",
].filter((image): image is string => image !== undefined && image.length > 0);

const TEXT_EDITED_VALUE = "Edited LibreOffice text";
const TRANSFORM_EDIT = {
  offsetX: asEmu(2743200),
  offsetY: asEmu(1920240),
  width: asEmu(2926080),
  height: asEmu(1463040),
} as const;
const FORMATTING_EDITED_VALUE = "Editable formatting target";
const PARAGRAPH_EDITED_VALUE = "Paragraph properties target";
const BLUE_PNG = new Uint8Array(
  Buffer.from(
    "iVBORw0KGgoAAAANSUhEUgAAAAQAAAAECAIAAAAmkwkpAAAAE0lEQVR4nGNkYPjPAANMcBZeDgAx0wEH1s7nlgAAAABJRU5ErkJggg==",
    "base64",
  ),
);

const LO_EDITOR_VALIDITY_CASES = [
  {
    name: "text replacement",
    sourceFixture: "editor-validity-text-source.pptx",
    expectedFixture: "editor-validity-text-expected.pptx",
    createEditedPptx: (input: Uint8Array) => {
      const source = readPptx(input);
      const session = createEditorSession(source);
      const result = session.apply({
        kind: "replaceTextRunPlainText",
        handle: requireHandle(findTextRun(source, "Original LibreOffice text").handle),
        text: TEXT_EDITED_VALUE,
      });

      if (!result.ok) throw new Error(result.message);
      return writePptx(result.document);
    },
  },
  {
    name: "shape move and resize",
    sourceFixture: "editor-validity-transform-source.pptx",
    expectedFixture: "editor-validity-transform-expected.pptx",
    createEditedPptx: (input: Uint8Array) => {
      const source = readPptx(input);
      const session = createEditorSession(source);
      const handle = requireHandle(
        findShapeByName(source.slides[0].shapes, "Move Resize Target").handle,
      );
      const moveResult = session.apply({
        kind: "moveShape",
        handle,
        offsetX: TRANSFORM_EDIT.offsetX,
        offsetY: TRANSFORM_EDIT.offsetY,
      });

      if (!moveResult.ok) throw new Error(moveResult.message);

      const resizeResult = session.apply({
        kind: "resizeShape",
        handle,
        width: TRANSFORM_EDIT.width,
        height: TRANSFORM_EDIT.height,
      });

      if (!resizeResult.ok) throw new Error(resizeResult.message);
      return writePptx(resizeResult.document);
    },
  },
  {
    name: "text run formatting",
    sourceFixture: "editor-validity-formatting-source.pptx",
    expectedFixture: "editor-validity-formatting-expected.pptx",
    createEditedPptx: (input: Uint8Array) => {
      const source = readPptx(input);
      const session = createEditorSession(source);
      const handle = requireHandle(findTextRun(source, FORMATTING_EDITED_VALUE).handle);
      const setResult = session.apply({
        kind: "setTextRunProperties",
        handle,
        properties: {
          bold: false,
          fontSize: asPt(30),
          color: { kind: "srgb", hex: "9C0000" },
          typeface: "Liberation Serif",
        },
      });

      if (!setResult.ok) throw new Error(setResult.message);

      const clearResult = session.apply({
        kind: "clearTextRunProperties",
        handle,
        properties: ["italic", "underline"],
      });

      if (!clearResult.ok) throw new Error(clearResult.message);
      return writePptx(clearResult.document);
    },
  },
  {
    name: "paragraph properties",
    sourceFixture: "editor-validity-paragraph-source.pptx",
    expectedFixture: "editor-validity-paragraph-expected.pptx",
    createEditedPptx: (input: Uint8Array) => {
      const source = readPptx(input);
      const session = createEditorSession(source);
      const paragraph = findParagraph(source, PARAGRAPH_EDITED_VALUE);
      const setResult = session.apply({
        kind: "setParagraphProperties",
        handle: requireHandle(paragraph.handle),
        properties: {
          align: "right",
          level: 1,
          bullet: { type: "char", char: "\u2022" },
        },
      });

      if (!setResult.ok) throw new Error(setResult.message);
      return writePptx(setResult.document);
    },
  },
  {
    name: "image replacement",
    sourceFixture: "editor-validity-image-source.pptx",
    expectedFixture: "editor-validity-image-expected.pptx",
    createEditedPptx: (input: Uint8Array) => {
      const source = readPptx(input);
      const image = findFirstImage(source);
      return writePptx(replaceImageBytes(source, requireHandle(image.handle), BLUE_PNG));
    },
  },
] as const;

const libreOfficeImage = findLibreOfficeDockerImage();
const hasFixtures =
  existsSync(join(FIXTURE_DIR, "basic-shapes.pptx")) &&
  LO_EDITOR_VALIDITY_CASES.every(
    (testCase) =>
      existsSync(join(FIXTURE_DIR, testCase.sourceFixture)) &&
      existsSync(join(FIXTURE_DIR, testCase.expectedFixture)),
  );
const describeOrSkip = libreOfficeImage !== undefined && hasFixtures ? describe : describe.skip;
const describeFromScratchOrSkip = libreOfficeImage !== undefined ? describe : describe.skip;

describeOrSkip("LibreOffice edited PPTX validity", { timeout: 120000 }, () => {
  for (const testCase of LO_EDITOR_VALIDITY_CASES) {
    it(`${testCase.name}: opens edited PPTX and matches expected LibreOffice render`, async () => {
      const sourcePptx = readFileSync(join(FIXTURE_DIR, testCase.sourceFixture));
      const expectedFixturePath = join(FIXTURE_DIR, testCase.expectedFixture);
      const editedPptx = testCase.createEditedPptx(sourcePptx);
      const rendered = renderWithLibreOffice(
        libreOfficeImage,
        `${slugify(testCase.name)}-edited.pptx`,
        editedPptx,
        expectedFixturePath,
      );

      const comparison = await compareImageBuffers(
        readFileSync(rendered.editedPngPath),
        readFileSync(rendered.expectedPngPath),
        join(DIFF_DIR, `editor-validity-${slugify(testCase.name)}-diff.png`),
        {
          pixelThreshold: 0,
          mismatchTolerance: 0,
          includeAntiAliased: true,
        },
      );

      expect(
        comparison.passed,
        `${testCase.name}: ${(comparison.mismatchPercentage * 100).toFixed(3)}% pixels differ ` +
          `(${comparison.mismatchedPixels}/${comparison.totalPixels})`,
      ).toBe(true);
    });
  }
});

describeOrSkip("LibreOffice slide topology validity", { timeout: 120000 }, () => {
  it("opens PPTX after adding an empty slide from a layout", () => {
    const sourcePptx = readFileSync(join(FIXTURE_DIR, "basic-shapes.pptx"));
    const source = readPptx(sourcePptx);
    const layout = source.slideLayouts[0];
    if (layout === undefined) throw new Error("basic-shapes fixture has no slide layout");
    const edited = addEmptySlideFromLayout(source, { layoutPartPath: layout.partPath });

    renderSingleWithLibreOffice(
      libreOfficeImage,
      "editor-validity-empty-slide-from-layout-edited.pptx",
      writePptx(edited),
    );
  });

  it("opens PPTX after slide duplicate and delete edits", () => {
    const sourcePptx = readFileSync(join(FIXTURE_DIR, "basic-shapes.pptx"));
    const source = readPptx(sourcePptx);
    const duplicated = duplicateSlide(source, requireHandle(source.slides[0]?.handle));
    const deleted = deleteSlide(duplicated, requireHandle(duplicated.slides[1]?.handle));

    renderSingleWithLibreOffice(
      libreOfficeImage,
      "editor-validity-slide-topology-edited.pptx",
      writePptx(deleted),
    );
  });

  it("opens PPTX after moving a slide", () => {
    const sourcePptx = readFileSync(join(FIXTURE_DIR, "basic-shapes.pptx"));
    const source = readPptx(sourcePptx);
    const duplicated = duplicateSlide(source, requireHandle(source.slides[0]?.handle));
    const moved = moveSlide(duplicated, requireHandle(duplicated.slides[0]?.handle), {
      toIndex: 1,
    });

    renderSingleWithLibreOffice(
      libreOfficeImage,
      "editor-validity-slide-move-edited.pptx",
      writePptx(moved),
    );
  });
});

describeFromScratchOrSkip("LibreOffice from-scratch PPTX validity", { timeout: 120000 }, () => {
  it("opens from-scratch PPTX with native charts and embedded workbooks", () => {
    const source = createPptx();
    const handle = requireHandle(source.slides[0]?.handle);
    const types = ["bar", "line", "pie", "area", "doughnut", "radar"] as const;
    const edited = types.reduce(
      (current, chartType, index) =>
        addChart(current, handle, {
          chartType,
          offsetX: asEmu(index * 1400000),
          offsetY: asEmu(400000),
          width: asEmu(1300000),
          height: asEmu(1800000),
          title: chartType,
          titleStyle: {
            fontFace: "Liberation Sans",
            fontSize: asPt(12),
            color: { kind: "srgb", hex: "203864" },
            bold: true,
            italic: index % 2 === 0,
          },
          displayBlanksAs: (["gap", "zero", "span"] as const)[index % 3],
          roundedCorners: index === 0,
          chartArea: {
            fill: { kind: "solid", color: { kind: "srgb", hex: "F2F2F2" } },
            outline: {
              width: asEmu(12700),
              fill: { kind: "solid", color: { kind: "srgb", hex: "BFBFBF" } },
            },
          },
          plotArea: {
            fill: { kind: "solid", color: { kind: "srgb", hex: "FFFFFF" } },
          },
          ...(chartType === "bar" ||
          chartType === "line" ||
          chartType === "area" ||
          chartType === "radar"
            ? {
                categoryAxis: {
                  hidden: chartType === "bar",
                  majorTickMark: "outside" as const,
                  labelPosition: "nextTo" as const,
                  numberFormat: { formatCode: "General", sourceLinked: true },
                  line: {
                    fill: {
                      kind: "solid" as const,
                      color: { kind: "srgb" as const, hex: "808080" },
                    },
                  },
                  textStyle: {
                    fontFace: "Liberation Sans",
                    fontSize: asPt(8),
                    color: { kind: "srgb" as const, hex: "404040" },
                  },
                  showMultiLevelLabels: false,
                },
                valueAxis: {
                  hidden: chartType === "bar",
                  majorTickMark: "none" as const,
                  labelPosition: "low" as const,
                  numberFormat: { formatCode: "0.0", sourceLinked: false },
                  gridLinesVisible: true,
                  majorGridline: {
                    fill: {
                      kind: "solid" as const,
                      color: { kind: "srgb" as const, hex: "D9D9D9" },
                    },
                    dash: "dot" as const,
                  },
                },
              }
            : {}),
          ...(chartType === "bar"
            ? { plotLayout: { coordinateMode: "edge" as const, x: 0, y: 0, width: 1, height: 1 } }
            : {}),
          series: [
            {
              name: "Plan",
              categories: ["Q1", "Q2", "Q3"],
              values: [10, 20, 15],
              fill: { kind: "solid", color: { kind: "srgb", hex: "4472C4" } },
              outline: {
                width: asEmu(12700),
                fill: { kind: "solid", color: { kind: "srgb", hex: "203864" } },
              },
              ...(chartType === "line" || chartType === "radar"
                ? {
                    marker: {
                      symbol: "diamond" as const,
                      size: 7,
                      fill: {
                        kind: "solid" as const,
                        color: { kind: "srgb" as const, hex: "FFC000" },
                      },
                      outline: {
                        fill: {
                          kind: "solid" as const,
                          color: { kind: "srgb" as const, hex: "7F6000" },
                        },
                      },
                    },
                  }
                : {}),
              ...(chartType === "bar" || chartType === "pie" || chartType === "doughnut"
                ? {
                    dataPoints: [
                      {
                        index: 0,
                        fill: {
                          kind: "solid" as const,
                          color: { kind: "srgb" as const, hex: "ED7D31" },
                        },
                      },
                      {
                        index: 1,
                        fill: {
                          kind: "solid" as const,
                          color: { kind: "srgb" as const, hex: "70AD47" },
                        },
                      },
                    ],
                  }
                : {}),
            },
          ],
        }),
      source,
    );
    renderSingleWithLibreOffice(
      libreOfficeImage,
      "editor-validity-from-scratch-native-charts.pptx",
      writePptx(edited),
    );
  });

  it("opens from-scratch PPTX after adding a text box", () => {
    const source = createPptx();
    const edited = addTextBox(source, requireHandle(source.slides[0]?.handle), {
      offsetX: asEmu(914400),
      offsetY: asEmu(914400),
      width: asEmu(3657600),
      height: asEmu(914400),
      text: "LibreOffice from-scratch text box",
    });

    renderSingleWithLibreOffice(
      libreOfficeImage,
      "editor-validity-from-scratch-text-box.pptx",
      writePptx(edited),
    );
  });

  it("opens from-scratch PPTX with plain, bulleted, and numbered text", () => {
    const source = createPptx();
    const slideHandle = requireHandle(source.slides[0]?.handle);
    const withPlainText = addTextBox(source, slideHandle, {
      offsetX: asEmu(457200),
      offsetY: asEmu(228600),
      width: asEmu(8229600),
      height: asEmu(685800),
      body: { autoFit: "shape" },
      paragraphs: [
        {
          runs: [
            {
              text: "Plain text",
              properties: {
                fontFace: "Liberation Sans",
                fontSize: asPt(24),
                baseline: { type: "percent", value: asOoxmlPercent(0) },
              },
            },
          ],
        },
      ],
    });
    const withBullets = addTextBox(withPlainText, slideHandle, {
      offsetX: asEmu(685800),
      offsetY: asEmu(1371600),
      width: asEmu(3657600),
      height: asEmu(2286000),
      body: { autoFit: "shape" },
      paragraphs: ["First bullet", "Second bullet"].map((text) => ({
        properties: {
          marginLeft: asEmu(342900),
          indent: asEmu(-285750),
          lineSpacing: { type: "percent" as const, value: asOoxmlPercent(110000) },
          bullet: {
            type: "character" as const,
            character: "•",
            fontFace: "Liberation Sans",
            size: asOoxmlPercent(100000),
          },
        },
        runs: [{ text, properties: { fontFace: "Liberation Sans", fontSize: asPt(20) } }],
      })),
    });
    const edited = addShape(withBullets, slideHandle, {
      geometry: { kind: "preset", preset: "roundRect" },
      offsetX: asEmu(4800600),
      offsetY: asEmu(1371600),
      width: asEmu(3657600),
      height: asEmu(2286000),
      body: { autoFit: "shape" },
      paragraphs: [
        {
          properties: {
            marginLeft: asEmu(457200),
            indent: asEmu(-228600),
            lineSpacing: { type: "points", value: asHundredthPt(2400) },
            bullet: {
              type: "auto-number",
              scheme: "arabicPeriod",
              startAt: 2,
              fontFace: "Liberation Sans",
              size: asOoxmlPercent(100000),
            },
          },
          runs: [{ text: "Numbered item", properties: { fontSize: asPt(20) } }],
        },
        {
          properties: { bullet: { type: "none" } },
          runs: [{ text: "Explicitly unbulleted", properties: { fontSize: asPt(18) } }],
        },
      ],
    });

    renderSingleWithLibreOffice(
      libreOfficeImage,
      "editor-validity-from-scratch-text-lists.pptx",
      writePptx(edited),
    );
  });

  it("opens from-scratch PPTX after adding a formatted text box", () => {
    const source = createPptx();
    const edited = addTextBox(source, requireHandle(source.slides[0]?.handle), {
      offsetX: asEmu(914400),
      offsetY: asEmu(914400),
      width: asEmu(5486400),
      height: asEmu(1828800),
      rotation: asOoxmlAngle(900000),
      body: {
        anchor: "middle",
        marginLeft: asEmu(91440),
        marginRight: asEmu(91440),
        marginTop: asEmu(45720),
        marginBottom: asEmu(45720),
      },
      paragraphs: [
        {
          properties: {
            align: "center",
            lineSpacing: { type: "points", value: asHundredthPt(1800) },
          },
          runs: [
            {
              text: "Formatted ",
              properties: {
                fontFace: "Liberation Sans",
                fontSize: asPt(24),
                color: {
                  kind: "srgb",
                  hex: "112233",
                  transforms: [{ kind: "alpha", value: asOoxmlPercent(70000) }],
                },
                bold: true,
                italic: true,
                underline: { style: "sng", color: { kind: "srgb", hex: "445566" } },
                strike: true,
                highlight: { kind: "srgb", hex: "ffff00" },
                glow: {
                  radius: asEmu(12700),
                  color: {
                    kind: "srgb",
                    hex: "00aaff",
                    transforms: [{ kind: "alpha", value: asOoxmlPercent(50000) }],
                  },
                },
                outline: { width: asEmu(6350), color: { kind: "srgb", hex: "aa00aa" } },
                charSpacing: 80,
              },
            },
            {
              text: "gradient",
              properties: {
                gradientFill: {
                  gradientType: "linear",
                  angle: asOoxmlAngle(2700000),
                  stops: [
                    {
                      position: asOoxmlPercent(0),
                      color: {
                        kind: "srgb",
                        hex: "ff0000",
                        transforms: [{ kind: "alpha", value: asOoxmlPercent(60000) }],
                      },
                    },
                    {
                      position: asOoxmlPercent(100000),
                      color: { kind: "srgb", hex: "0000ff" },
                    },
                  ],
                },
                baseline: "superscript",
              },
            },
          ],
        },
        {
          properties: {
            align: "right",
            lineSpacing: { type: "points", value: asHundredthPt(2200) },
          },
          runs: [{ text: "Subscript", properties: { baseline: "subscript" } }],
        },
      ],
    });

    renderSingleWithLibreOffice(
      libreOfficeImage,
      "editor-validity-from-scratch-formatted-text-box.pptx",
      writePptx(edited),
    );
  });

  it("opens from-scratch PPTX with run hyperlinks in a text box and shape", () => {
    const source = createPptx();
    const handle = requireHandle(source.slides[0]?.handle);
    const withTextBox = addTextBox(source, handle, {
      offsetX: asEmu(914400),
      offsetY: asEmu(914400),
      width: asEmu(5486400),
      height: asEmu(914400),
      paragraphs: [
        {
          runs: [
            { text: "Plain text and " },
            { text: "linked text", hyperlink: "https://example.com/text-box" },
          ],
        },
      ],
    });
    const edited = addShape(withTextBox, handle, {
      geometry: { kind: "preset", preset: "roundRect" },
      offsetX: asEmu(914400),
      offsetY: asEmu(2286000),
      width: asEmu(5486400),
      height: asEmu(914400),
      paragraphs: [
        {
          runs: [{ text: "Shape " }, { text: "hyperlink", hyperlink: "https://example.com/shape" }],
        },
      ],
    });

    renderSingleWithLibreOffice(
      libreOfficeImage,
      "editor-validity-from-scratch-run-hyperlinks.pptx",
      writePptx(edited),
    );
  });

  it("opens from-scratch PPTX after adding a native table", () => {
    const source = createPptx();
    const edited = addTable(source, requireHandle(source.slides[0]?.handle), {
      offsetX: asEmu(914400),
      offsetY: asEmu(914400),
      width: asEmu(5486400),
      height: asEmu(1828800),
      columnWidths: [asEmu(2743200), asEmu(2743200)],
      rows: [
        {
          height: asEmu(914400),
          cells: [
            {
              runs: [
                {
                  text: "Native\ntable",
                  hyperlink: "https://example.com/table-validity",
                  properties: {
                    bold: true,
                    strike: true,
                    underline: {
                      style: "dashLongHeavy",
                      color: { kind: "srgb", hex: "FFFF00" },
                    },
                    highlight: { kind: "srgb", hex: "1F4E78" },
                    color: { kind: "srgb", hex: "FFFFFF" },
                  },
                },
              ],
              fill: "4472C4",
              colspan: 2,
              marginLeft: asEmu(100000),
              marginRight: asEmu(110000),
              marginTop: asEmu(50000),
              marginBottom: asEmu(60000),
              borders: {
                left: { width: asEmu(12700), color: "FFFFFF", dash: "lgDashDotDot" },
                right: { width: asEmu(12700), color: "FFFFFF", dash: "sysDash" },
                top: { width: asEmu(12700), color: "FFFFFF", dash: "dashDot" },
                bottom: { width: asEmu(12700), color: "FFFFFF", dash: "sysDot" },
              },
            },
            {},
          ],
        },
        {
          height: asEmu(914400),
          cells: [{ text: "LibreOffice" }, { text: "validity" }],
        },
      ],
    });

    renderSingleWithLibreOffice(
      libreOfficeImage,
      "editor-validity-from-scratch-native-table.pptx",
      writePptx(edited),
    );
  });

  it("opens a from-scratch PPTX with authored master objects and multiple slides", () => {
    let source = createPptx({
      slideMaster: {
        name: "LibreOffice Authored Master",
        background: { kind: "solid", color: { kind: "srgb", hex: "F1F5F9" } },
      },
      slideLayout: {
        name: "LibreOffice Authored Layout",
        margin: {
          left: asEmu(120000),
          right: asEmu(120000),
          top: asEmu(80000),
          bottom: asEmu(80000),
        },
      },
    });
    const masterHandle = requireHandle(source.slideMasters[0]?.handle);
    const layout = source.slideLayouts[0];
    if (layout === undefined) throw new Error("createPptx should create a layout");
    source = addTextBox(source, masterHandle, {
      offsetX: asEmu(300000),
      offsetY: asEmu(180000),
      width: asEmu(3000000),
      height: asEmu(500000),
      text: "Inherited master object",
    });
    source = addShape(source, masterHandle, {
      geometry: { kind: "preset", preset: "rect" },
      offsetX: asEmu(0),
      offsetY: asEmu(5000000),
      width: asEmu(9144000),
      height: asEmu(143500),
      fill: { kind: "solid", color: { kind: "srgb", hex: "4472C4" } },
      effects: {
        outerShadow: {
          blurRadius: asEmu(40000),
          distance: asEmu(20000),
          direction: asOoxmlAngle(2700000),
          color: {
            kind: "srgb",
            hex: "000000",
            transforms: [{ kind: "alpha", value: asOoxmlPercent(35000) }],
          },
          alignment: "b",
          rotateWithShape: false,
        },
      },
    });
    source = addConnector(source, masterHandle, {
      preset: "straightConnector1",
      offsetX: asEmu(300000),
      offsetY: asEmu(800000),
      width: asEmu(2500000),
      height: asEmu(1),
    });
    source = addPicture(source, masterHandle, {
      bytes: BLUE_PNG,
      offsetX: asEmu(8200000),
      offsetY: asEmu(180000),
      width: asEmu(500000),
      height: asEmu(500000),
      effects: {
        innerShadow: {
          blurRadius: asEmu(30000),
          distance: asEmu(15000),
          direction: asOoxmlAngle(8100000),
          color: { kind: "srgb", hex: "1E293B" },
        },
        outerShadow: {
          blurRadius: asEmu(45000),
          distance: asEmu(25000),
          direction: asOoxmlAngle(13500000),
          color: {
            kind: "srgb",
            hex: "000000",
            transforms: [{ kind: "alpha", value: asOoxmlPercent(25000) }],
          },
          alignment: "br",
          rotateWithShape: true,
        },
      },
    });
    source = addSlideNumber(source, masterHandle, {
      offsetX: asEmu(8200000),
      offsetY: asEmu(4650000),
      width: asEmu(500000),
      height: asEmu(300000),
      align: "right",
      properties: { fontSize: asPt(12), color: { kind: "srgb", hex: "334155" } },
    });
    source = addEmptySlideFromLayout(source, { layoutPartPath: layout.partPath });
    source = addEmptySlideFromLayout(source, { layoutPartPath: layout.partPath });

    renderSingleWithLibreOffice(
      libreOfficeImage,
      "editor-validity-from-scratch-authored-master.pptx",
      writePptx(source),
    );
  });

  it("opens a from-scratch PPTX with individual slide backgrounds", async () => {
    let source = createPptx();
    const masterHandle = requireHandle(source.slideMasters[0]?.handle);
    const layout = source.slideLayouts[0];
    if (layout === undefined) throw new Error("createPptx should create a layout");
    source = addTextBox(source, masterHandle, {
      offsetX: asEmu(300000),
      offsetY: asEmu(180000),
      width: asEmu(3000000),
      height: asEmu(500000),
      text: "Inherited above each slide background",
    });
    for (let index = 0; index < 4; index += 1) {
      source = addEmptySlideFromLayout(source, { layoutPartPath: layout.partPath });
    }
    const handles = source.slides.map((slide) => requireHandle(slide.handle));
    source = setSlideBackground(source, handles[0], {
      kind: "solid",
      color: { kind: "srgb", hex: "E2E8F0" },
    });
    source = setSlideBackground(source, handles[1], {
      kind: "gradient",
      gradientType: "linear",
      angle: asOoxmlAngle(2700000),
      stops: [
        { position: asOoxmlPercent(0), color: { kind: "srgb", hex: "DBEAFE" } },
        { position: asOoxmlPercent(100000), color: { kind: "srgb", hex: "BFDBFE" } },
      ],
    });
    source = setSlideBackground(source, handles[2], {
      kind: "gradient",
      gradientType: "radial",
      centerX: asOoxmlPercent(35000),
      centerY: asOoxmlPercent(40000),
      stops: [
        { position: asOoxmlPercent(0), color: { kind: "srgb", hex: "FEF3C7" } },
        { position: asOoxmlPercent(100000), color: { kind: "srgb", hex: "FDE68A" } },
      ],
    });
    source = setSlideBackground(source, handles[3], { kind: "image", bytes: BLUE_PNG });
    source = setSlideBackground(source, handles[4], {
      kind: "image",
      bytes: new Uint8Array(await sharp(BLUE_PNG).jpeg().toBuffer()),
    });

    renderSingleWithLibreOffice(
      libreOfficeImage,
      "editor-validity-from-scratch-slide-backgrounds.pptx",
      writePptx(source),
    );
  });
});

describeOrSkip("LibreOffice shape add/delete validity", { timeout: 120000 }, () => {
  it("opens PPTX after adding a text box", () => {
    const sourcePptx = readFileSync(join(FIXTURE_DIR, "basic-shapes.pptx"));
    const source = readPptx(sourcePptx);
    const edited = addTextBox(source, requireHandle(source.slides[0]?.handle), {
      offsetX: asEmu(914400),
      offsetY: asEmu(914400),
      width: asEmu(3657600),
      height: asEmu(914400),
      text: "LibreOffice added text box",
    });

    renderSingleWithLibreOffice(
      libreOfficeImage,
      "editor-validity-shape-add-edited.pptx",
      writePptx(edited),
    );
  });

  it("opens PPTX after adding a preset geometry shape", () => {
    const sourcePptx = readFileSync(join(FIXTURE_DIR, "basic-shapes.pptx"));
    const source = readPptx(sourcePptx);
    const edited = addShape(source, requireHandle(source.slides[0]?.handle), {
      geometry: { kind: "preset", preset: "roundRect" },
      offsetX: asEmu(914400),
      offsetY: asEmu(914400),
      width: asEmu(3657600),
      height: asEmu(914400),
      rotation: asOoxmlAngle(900000),
      fill: {
        kind: "gradient",
        gradientType: "radial",
        centerX: asOoxmlPercent(35000),
        centerY: asOoxmlPercent(65000),
        stops: [
          {
            position: asOoxmlPercent(0),
            color: {
              kind: "srgb",
              hex: "FF0000",
              transforms: [{ kind: "alpha", value: asOoxmlPercent(70000) }],
            },
          },
          { position: asOoxmlPercent(100000), color: { kind: "srgb", hex: "0000FF" } },
        ],
      },
      outline: {
        width: asEmu(12700),
        fill: {
          kind: "gradient",
          gradientType: "linear",
          angle: asOoxmlAngle(2700000),
          stops: [
            {
              position: asOoxmlPercent(0),
              color: {
                kind: "srgb",
                hex: "00AA44",
                transforms: [{ kind: "alpha", value: asOoxmlPercent(50000) }],
              },
            },
            {
              position: asOoxmlPercent(100000),
              color: { kind: "srgb", hex: "0044AA" },
            },
          ],
        },
        dash: "dash",
        tailEnd: { type: "triangle", width: "med", length: "lg" },
      },
      effects: {
        glow: {
          radius: asEmu(25400),
          color: {
            kind: "srgb",
            hex: "AA00AA",
            transforms: [{ kind: "alpha", value: asOoxmlPercent(60000) }],
          },
        },
      },
      paragraphs: [{ runs: [{ text: "LibreOffice added shape" }] }],
    });

    renderSingleWithLibreOffice(
      libreOfficeImage,
      "editor-validity-shape-add-preset-edited.pptx",
      writePptx(edited),
    );
  });

  it("opens PPTX after adding adjusted, custom, flipped, and zero-extent shapes", () => {
    const sourcePptx = readFileSync(join(FIXTURE_DIR, "basic-shapes.pptx"));
    let edited = readPptx(sourcePptx);
    const slideHandle = requireHandle(edited.slides[0]?.handle);
    edited = addShape(edited, slideHandle, {
      geometry: { kind: "preset", preset: "roundRect", adjustValues: { adj: 20000 } },
      offsetX: asEmu(457200),
      offsetY: asEmu(457200),
      width: asEmu(1828800),
      height: asEmu(914400),
      flipHorizontal: true,
    });
    edited = addShape(edited, slideHandle, {
      geometry: {
        kind: "custom",
        paths: [
          {
            width: 100,
            height: 100,
            commands: [
              { kind: "moveTo", x: 0, y: 100 },
              { kind: "lineTo", x: 50, y: 0 },
              { kind: "lineTo", x: 100, y: 100 },
              { kind: "close" },
            ],
          },
        ],
      },
      offsetX: asEmu(3200400),
      offsetY: asEmu(457200),
      width: asEmu(1828800),
      height: asEmu(914400),
      flipVertical: true,
    });
    edited = addShape(edited, slideHandle, {
      geometry: { kind: "preset", preset: "line" },
      offsetX: asEmu(5943600),
      offsetY: asEmu(457200),
      width: asEmu(0),
      height: asEmu(914400),
      flipHorizontal: true,
      flipVertical: true,
    });

    renderSingleWithLibreOffice(
      libreOfficeImage,
      "editor-validity-shape-add-geometry-transform-edited.pptx",
      writePptx(edited),
    );
  });

  it("opens PPTX after adding a connector", () => {
    const sourcePptx = readFileSync(join(FIXTURE_DIR, "basic-shapes.pptx"));
    const source = readPptx(sourcePptx);
    const shapes = source.slides[0]?.shapes.filter(
      (shape): shape is SourceShape => shape.kind === "shape",
    );
    const edited = addConnector(source, requireHandle(source.slides[0]?.handle), {
      preset: "straightConnector1",
      offsetX: asEmu(914400),
      offsetY: asEmu(1371600),
      width: asEmu(3657600),
      height: asEmu(914400),
      start: {
        shapeHandle: requireHandle(shapes?.[0]?.handle),
        connectionSiteIndex: 1,
      },
      end: {
        shapeHandle: requireHandle(shapes?.[1]?.handle),
        connectionSiteIndex: 3,
      },
      outline: {
        tailEnd: { type: "triangle", width: "med", length: "med" },
      },
    });

    renderSingleWithLibreOffice(
      libreOfficeImage,
      "editor-validity-shape-connector-edited.pptx",
      writePptx(edited),
    );
  });

  it("opens PPTX after editing shape fill and outline", () => {
    const sourcePptx = readFileSync(join(FIXTURE_DIR, "basic-shapes.pptx"));
    const source = readPptx(sourcePptx);
    const shapes = source.slides[0]?.shapes.filter(
      (shape): shape is SourceShape => shape.kind === "shape",
    );
    const firstShape = shapes?.[0];
    const secondShape = shapes?.[1];
    const session = createEditorSession(source);
    const setFirstFill = session.apply({
      kind: "setShapeFill",
      handle: requireHandle(firstShape?.handle),
      fill: { kind: "solid", color: { kind: "srgb", hex: "00AA44" } },
    });

    if (!setFirstFill.ok) throw new Error(setFirstFill.message);

    const setFirstOutline = session.apply({
      kind: "setShapeOutline",
      handle: requireHandle(firstShape?.handle),
      outline: {
        width: asEmu(25400),
        fill: { kind: "solid", color: { kind: "srgb", hex: "336699" } },
      },
    });

    if (!setFirstOutline.ok) throw new Error(setFirstOutline.message);

    const setSecondNoFill = session.apply({
      kind: "setShapeFill",
      handle: requireHandle(secondShape?.handle),
      fill: { kind: "none" },
    });

    if (!setSecondNoFill.ok) throw new Error(setSecondNoFill.message);

    const setSecondNoOutline = session.apply({
      kind: "setShapeOutline",
      handle: requireHandle(secondShape?.handle),
      outline: { fill: { kind: "none" } },
    });

    if (!setSecondNoOutline.ok) throw new Error(setSecondNoOutline.message);

    renderSingleWithLibreOffice(
      libreOfficeImage,
      "editor-validity-shape-style-edited.pptx",
      writePptx(setSecondNoOutline.document),
    );
  });

  it("opens PPTX after adding a free connector", () => {
    const sourcePptx = readFileSync(join(FIXTURE_DIR, "basic-shapes.pptx"));
    const source = readPptx(sourcePptx);
    const edited = addConnector(source, requireHandle(source.slides[0]?.handle), {
      preset: "straightConnector1",
      offsetX: asEmu(914400),
      offsetY: asEmu(1371600),
      width: asEmu(3657600),
      height: asEmu(914400),
      outline: {
        tailEnd: { type: "triangle", width: "med", length: "med" },
      },
    });

    renderSingleWithLibreOffice(
      libreOfficeImage,
      "editor-validity-shape-free-connector-edited.pptx",
      writePptx(edited),
    );
  });

  it("opens PPTX after deleting a connector", () => {
    const sourcePptx = readFileSync(join(FIXTURE_DIR, "basic-shapes.pptx"));
    const source = readPptx(sourcePptx);
    const withConnector = addConnector(source, requireHandle(source.slides[0]?.handle), {
      preset: "straightConnector1",
      offsetX: asEmu(914400),
      offsetY: asEmu(1371600),
      width: asEmu(3657600),
      height: asEmu(914400),
      outline: {
        tailEnd: { type: "triangle", width: "med", length: "med" },
      },
    });
    const persisted = readPptx(writePptx(withConnector));
    const connector = persisted.slides[0]?.shapes.find(
      (shape): shape is SourceConnector => shape.kind === "connector",
    );
    const edited = deleteShape(persisted, requireHandle(connector?.handle));

    renderSingleWithLibreOffice(
      libreOfficeImage,
      "editor-validity-shape-connector-delete-edited.pptx",
      writePptx(edited),
    );
  });

  it("opens PPTX after deleting a shape", () => {
    const sourcePptx = readFileSync(join(FIXTURE_DIR, "basic-shapes.pptx"));
    const source = readPptx(sourcePptx);
    const shape = source.slides[0]?.shapes.find((candidate) => candidate.kind === "shape");
    const edited = deleteShape(source, requireHandle(shape?.handle));

    renderSingleWithLibreOffice(
      libreOfficeImage,
      "editor-validity-shape-delete-edited.pptx",
      writePptx(edited),
    );
  });
});

function findLibreOfficeDockerImage(): string | undefined {
  for (const image of LIBREOFFICE_IMAGE_CANDIDATES) {
    const result = spawnSync("docker", ["image", "inspect", image], { encoding: "utf8" });
    if (result.status === 0) return image;
  }
  return undefined;
}

function renderWithLibreOffice(
  image: string,
  editedFilename: string,
  editedPptx: Uint8Array,
  expectedFixturePath: string,
): { readonly editedPngPath: string; readonly expectedPngPath: string } {
  const workDir = mkdtempSync(join(tmpdir(), "pptx-glimpse-lo-editor-validity-"));
  const outputDir = join(workDir, "out");
  const editedPath = join(workDir, editedFilename);
  const expectedFilename = basename(expectedFixturePath);
  const expectedPath = join(workDir, expectedFilename);

  writeFileSync(editedPath, editedPptx);
  copyFileSync(expectedFixturePath, expectedPath);

  const result = spawnSync(
    "docker",
    [
      "run",
      "--rm",
      "-v",
      `${workDir}:/work`,
      image,
      "bash",
      "-lc",
      "mkdir -p /work/out && libreoffice --headless --convert-to png --outdir /work/out /work/*.pptx",
    ],
    { encoding: "utf8" },
  );

  if (result.status !== 0) {
    throw new Error(
      `LibreOffice conversion failed for '${editedFilename}'.\n` +
        `stdout:\n${result.stdout}\n\nstderr:\n${result.stderr}`,
    );
  }

  const editedPngPath = join(outputDir, `${basename(editedFilename, ".pptx")}.png`);
  const expectedPngPath = join(outputDir, `${basename(expectedFilename, ".pptx")}.png`);
  if (!existsSync(editedPngPath) || !existsSync(expectedPngPath)) {
    throw new Error(
      `LibreOffice conversion did not produce expected PNGs for '${editedFilename}' and ` +
        `'${expectedFilename}'.`,
    );
  }

  return { editedPngPath, expectedPngPath };
}

function renderSingleWithLibreOffice(
  image: string,
  editedFilename: string,
  editedPptx: Uint8Array,
): string {
  const workDir = mkdtempSync(join(tmpdir(), "pptx-glimpse-lo-editor-validity-"));
  const outputDir = join(workDir, "out");
  const editedPath = join(workDir, editedFilename);

  writeFileSync(editedPath, editedPptx);

  const result = spawnSync(
    "docker",
    [
      "run",
      "--rm",
      "-v",
      `${workDir}:/work`,
      image,
      "bash",
      "-lc",
      `mkdir -p /work/out && libreoffice --headless --convert-to png --outdir /work/out /work/${editedFilename}`,
    ],
    { encoding: "utf8" },
  );

  if (result.status !== 0) {
    throw new Error(
      `LibreOffice conversion failed for '${editedFilename}'.\n` +
        `stdout:\n${result.stdout}\n\nstderr:\n${result.stderr}`,
    );
  }

  const editedPngPath = join(outputDir, `${basename(editedFilename, ".pptx")}.png`);
  if (!existsSync(editedPngPath)) {
    throw new Error(`LibreOffice conversion did not produce expected PNG for '${editedFilename}'.`);
  }
  return editedPngPath;
}

function findTextRun(source: ReturnType<typeof readPptx>, text: string): SourceTextRun {
  for (const slide of source.slides) {
    for (const shape of flattenShapes(slide.shapes)) {
      for (const paragraph of shape.textBody?.paragraphs ?? []) {
        const run = paragraph.runs.find((candidate) => candidate.text === text);
        if (run !== undefined) return run;
      }
    }
  }
  throw new Error(`Text run not found: ${text}`);
}

function findParagraph(source: ReturnType<typeof readPptx>, text: string): SourceParagraph {
  for (const slide of source.slides) {
    for (const shape of flattenShapes(slide.shapes)) {
      const paragraph = shape.textBody?.paragraphs.find(
        (candidate) => candidate.runs.map((run) => run.text).join("") === text,
      );
      if (paragraph !== undefined) return paragraph;
    }
  }
  throw new Error(`Paragraph not found: ${text}`);
}

function findShapeByName(shapes: readonly SourceShapeNode[], name: string): SourceShape {
  const shape = flattenShapes(shapes).find((candidate) => candidate.name === name);
  if (shape === undefined) throw new Error(`Shape not found: ${name}`);
  return shape;
}

function findFirstImage(source: ReturnType<typeof readPptx>): SourceImage {
  for (const slide of source.slides) {
    const image = slide.shapes.find((shape): shape is SourceImage => shape.kind === "image");
    if (image !== undefined) return image;
  }
  throw new Error("Image shape not found");
}

function flattenShapes(shapes: readonly SourceShapeNode[]): SourceShape[] {
  return shapes.flatMap((shape): SourceShape[] => {
    if (shape.kind === "shape") return [shape];
    if (shape.kind === "group") return flattenShapes(shape.children);
    return [];
  });
}

function requireHandle(handle: SourceHandle | undefined): SourceHandle {
  if (handle === undefined) throw new Error("Source handle not found");
  return handle;
}

function slugify(value: string): string {
  return value.replace(/[^A-Za-z0-9_-]+/g, "-").replace(/^-+|-+$/g, "") || "case";
}
