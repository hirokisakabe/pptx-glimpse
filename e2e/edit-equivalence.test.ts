import { existsSync } from "node:fs";
import { mkdir, writeFile } from "node:fs/promises";
import { tmpdir } from "node:os";
import { join } from "node:path";

import { describe, expect, it } from "vitest";

import type { ConvertOptions } from "../packages/core/src/converter.js";
import {
  addTextBox,
  asEmu,
  asPt,
  replaceTextRunPlainText,
} from "../packages/document/src/index.js";
import { createEditorSession } from "../packages/editor-core/src/index.js";
import {
  buildPptx,
  shapeXml,
  slideRelsXml,
  textBodyXmlHelper,
  wrapSlideXml,
} from "../vrt/snapshot/create-fixtures.js";
import { unsafeVrtInteropAssertion } from "../vrt/unsafe-type-assertion.js";
import {
  assertEditEquivalence,
  defineEditEquivalenceTests,
  type EditEquivalenceFixture,
  type EditEquivalenceOperation,
} from "./edit-equivalence.js";

const TEST_FONT_FAMILY = "Pptx Glimpse Edit Equivalence";
const TEST_FONT_DIR = join(tmpdir(), "pptx-glimpse-edit-equivalence-font-v2");
const TEST_FONT_PATH = join(TEST_FONT_DIR, "PptxGlimpseEditEquivalence-Regular.ttf");

const originalTextRunFixture = textRunFixture("source text-run fixture", "Original text");
const editedTextRunFixture = textRunFixture("expected edited text-run fixture", "Edited text");
const originalDecoratedRunFixture = textRunFixture(
  "source text-run decoration fixture",
  "Edited text",
);
const expectedDecoratedRunFixture = textRunFixture(
  "expected decorated text-run fixture",
  "Edited text",
  {
    bold: true,
    italic: true,
    underline: true,
    fontSize: 32,
    color: "9C0000",
    typeface: TEST_FONT_FAMILY,
  },
);
const mismatchedTextRunFixture = textRunFixture(
  "mismatched expected text-run fixture",
  "Alterd text",
);
const emptySlideFixture = {
  name: "source empty slide fixture",
  create: async () =>
    await buildPptx({
      slides: [
        {
          xml: wrapSlideXml(""),
          rels: slideRelsXml(),
        },
      ],
    }),
} as const satisfies EditEquivalenceFixture;
const expectedAddedTextBoxFixture = {
  name: "expected added text box fixture",
  create: async () =>
    await buildPptx({
      slides: [
        {
          xml: wrapSlideXml(
            textBoxShapeXml(1, "TextBox 1", {
              x: 914400,
              y: 914400,
              cx: 3657600,
              cy: 914400,
              text: "Edited added box",
            }),
          ),
          rels: slideRelsXml(),
        },
      ],
    }),
} as const satisfies EditEquivalenceFixture;

const replaceFirstRunWithEditedText = {
  name: "replace first text run with edited text",
  apply: (source) => {
    const run =
      source.slides[0]?.shapes[0]?.kind === "shape"
        ? source.slides[0].shapes[0].textBody?.paragraphs[0]?.runs[0]
        : undefined;
    if (run?.handle === undefined) throw new Error("editable text run not found");
    return replaceTextRunPlainText(source, run.handle, "Edited text");
  },
} as const satisfies EditEquivalenceOperation;

const decorateFirstRun = {
  name: "set first text run decoration",
  apply: (source) => {
    const run =
      source.slides[0]?.shapes[0]?.kind === "shape"
        ? source.slides[0].shapes[0].textBody?.paragraphs[0]?.runs[0]
        : undefined;
    if (run?.handle === undefined) throw new Error("editable text run not found");
    const session = createEditorSession(source);
    const result = session.apply({
      kind: "setTextRunProperties",
      handle: run.handle,
      properties: {
        bold: true,
        italic: true,
        underline: true,
        fontSize: asPt(32),
        color: { kind: "srgb", hex: "9C0000" },
        typeface: TEST_FONT_FAMILY,
      },
    });
    if (!result.ok) throw new Error(result.message);
    return result.document;
  },
} as const satisfies EditEquivalenceOperation;

const addThenEditTextBox = {
  name: "add then edit text box",
  apply: (source) => {
    const withTextBox = addTextBox(source, source.slides[0].handle!, {
      offsetX: asEmu(457200),
      offsetY: asEmu(457200),
      width: asEmu(2743200),
      height: asEmu(685800),
      text: "Initial added box",
    });
    const shape = withTextBox.slides[0]?.shapes.find(
      (candidate) => candidate.kind === "shape" && candidate.name === "TextBox 1",
    );
    const runHandle =
      shape?.kind === "shape" ? shape.textBody?.paragraphs[0]?.runs[0]?.handle : undefined;
    if (shape?.handle === undefined || runHandle === undefined) {
      throw new Error("added text box handles not found");
    }
    const session = createEditorSession(withTextBox);
    const textResult = session.apply({
      kind: "replaceTextRunPlainText",
      handle: runHandle,
      text: "Edited added box",
    });
    if (!textResult.ok) throw new Error(textResult.message);
    const xfrmResult = session.apply({
      kind: "setShapeTransform",
      handle: shape.handle,
      offsetX: asEmu(914400),
      offsetY: asEmu(914400),
      width: asEmu(3657600),
      height: asEmu(914400),
    });
    if (!xfrmResult.ok) throw new Error(xfrmResult.message);
    return xfrmResult.document;
  },
} as const satisfies EditEquivalenceOperation;

defineEditEquivalenceTests([
  {
    name: "text-run replacement",
    sourceFixture: originalTextRunFixture,
    operations: [replaceFirstRunWithEditedText],
    expectedFixture: editedTextRunFixture,
    renderOptionsProvider: getTinyTestFontRenderOptions,
    renderOptions: { width: 480 },
  },
  {
    name: "text-run decoration",
    sourceFixture: originalDecoratedRunFixture,
    operations: [decorateFirstRun],
    expectedFixture: expectedDecoratedRunFixture,
    renderOptionsProvider: getTinyTestFontRenderOptions,
    renderOptions: { width: 480 },
  },
  {
    name: "add edit text box",
    sourceFixture: emptySlideFixture,
    operations: [addThenEditTextBox],
    expectedFixture: expectedAddedTextBoxFixture,
    renderOptionsProvider: getTinyTestFontRenderOptions,
    renderOptions: { width: 480 },
  },
]);

describe("edit equivalence rendering oracle", { timeout: 60000 }, () => {
  it("fails when edit operations and expected fixture intentionally disagree", async () => {
    await expect(
      assertEditEquivalence({
        name: "text-run replacement mismatch proof",
        sourceFixture: originalTextRunFixture,
        operations: [replaceFirstRunWithEditedText],
        expectedFixture: mismatchedTextRunFixture,
        renderOptionsProvider: getTinyTestFontRenderOptions,
        renderOptions: { width: 480 },
      }),
    ).rejects.toThrow(/pixels differ/);
  });
});

function textRunFixture(
  name: string,
  text: string,
  opts?: Parameters<typeof textBodyXmlHelper>[1],
): EditEquivalenceFixture {
  return {
    name,
    create: async () =>
      await buildPptx({
        slides: [
          {
            xml: wrapSlideXml(textRunShapeXml(text, opts)),
            rels: slideRelsXml(),
          },
        ],
      }),
  };
}

function textRunShapeXml(text: string, opts?: Parameters<typeof textBodyXmlHelper>[1]): string {
  return decoratedTextRunShapeXml(text, {
    fontSize: 28,
    color: "FFFFFF",
    typeface: TEST_FONT_FAMILY,
    align: "ctr",
    ...opts,
  });
}

function decoratedTextRunShapeXml(
  text: string,
  opts: Parameters<typeof textBodyXmlHelper>[1],
): string {
  return shapeXml(2, "Editable Text", {
    preset: "rect",
    x: 914400,
    y: 914400,
    cx: 5486400,
    cy: 1371600,
    fillXml: `<a:solidFill><a:srgbClr val="4472C4"/></a:solidFill>`,
    outlineXml: `<a:ln w="12700"><a:solidFill><a:srgbClr val="2F528F"/></a:solidFill></a:ln>`,
    textBodyXml: textBodyXmlHelper(text, opts),
  });
}

function textBoxShapeXml(
  id: number,
  name: string,
  opts: { x: number; y: number; cx: number; cy: number; text: string },
): string {
  return `<p:sp>
  <p:nvSpPr><p:cNvPr id="${id}" name="${name}"/><p:cNvSpPr txBox="1"/><p:nvPr/></p:nvSpPr>
  <p:spPr>
    <a:xfrm><a:off x="${opts.x}" y="${opts.y}"/><a:ext cx="${opts.cx}" cy="${opts.cy}"/></a:xfrm>
    <a:prstGeom prst="rect"><a:avLst/></a:prstGeom>
    <a:noFill/>
    <a:ln><a:noFill/></a:ln>
  </p:spPr>
  <p:txBody><a:bodyPr wrap="square"/><a:lstStyle/><a:p><a:r><a:t>${opts.text}</a:t></a:r><a:endParaRPr/></a:p></p:txBody>
</p:sp>`;
}

let tinyTestFontDirPromise: Promise<string> | null = null;

async function getTinyTestFontRenderOptions(): Promise<ConvertOptions> {
  return {
    fontDirs: [await ensureTinyTestFontDir()],
    skipSystemFonts: true,
  };
}

async function ensureTinyTestFontDir(): Promise<string> {
  tinyTestFontDirPromise ??= createTinyTestFontDir();
  return tinyTestFontDirPromise;
}

async function createTinyTestFontDir(): Promise<string> {
  await mkdir(TEST_FONT_DIR, { recursive: true });
  if (existsSync(TEST_FONT_PATH)) return TEST_FONT_DIR;

  const opentype = unsafeVrtInteropAssertion<OpentypeModule>(await import("opentype.js"));
  const glyphs = [
    new opentype.Glyph({
      name: ".notdef",
      advanceWidth: 600,
      path: createGlyphPath(opentype),
    }),
    ...Array.from(new Set("EditedOriginalAlterd added box text")).map(
      (char) =>
        new opentype.Glyph({
          name: char === " " ? "space" : char,
          unicode: char.codePointAt(0),
          advanceWidth: glyphAdvanceWidth(char),
          path: char === " " ? new opentype.Path() : createGlyphPath(opentype, char),
        }),
    ),
  ];
  const font = new opentype.Font({
    familyName: TEST_FONT_FAMILY,
    styleName: "Regular",
    unitsPerEm: 1000,
    ascender: 800,
    descender: -200,
    glyphs,
  });

  await writeFile(TEST_FONT_PATH, Buffer.from(font.toArrayBuffer()));
  return TEST_FONT_DIR;
}

function createGlyphPath(opentype: OpentypeModule, char = ".notdef"): OpentypePath {
  const codePoint = char.codePointAt(0) ?? 0;
  const width = 220 + (codePoint % 7) * 45;
  const top = 520 + (codePoint % 5) * 35;
  const skew = (codePoint % 3) * 35;
  const path = new opentype.Path();
  path.moveTo(100, 0);
  path.lineTo(100 + width, 0);
  path.lineTo(100 + width + skew, top);
  path.lineTo(100 + skew, top);
  path.close();
  return path;
}

function glyphAdvanceWidth(char: string): number {
  if (char === " ") return 300;
  return 440 + ((char.codePointAt(0) ?? 0) % 6) * 45;
}

interface OpentypeModule {
  readonly Font: new (options: {
    familyName: string;
    styleName: string;
    unitsPerEm: number;
    ascender: number;
    descender: number;
    glyphs: readonly unknown[];
  }) => {
    toArrayBuffer(): ArrayBuffer;
  };
  readonly Glyph: new (options: {
    name: string;
    unicode?: number;
    advanceWidth: number;
    path: OpentypePath;
  }) => unknown;
  readonly Path: new () => OpentypePath;
}

interface OpentypePath {
  moveTo(x: number, y: number): void;
  lineTo(x: number, y: number): void;
  close(): void;
}
