import { spawnSync } from "node:child_process";
import { copyFileSync, existsSync, mkdtempSync, readFileSync, writeFileSync } from "node:fs";
import { tmpdir } from "node:os";
import { basename, dirname, join } from "node:path";
import { fileURLToPath } from "node:url";

import {
  asEmu,
  asPt,
  deleteSlide,
  duplicateSlide,
  readPptx,
  type SourceHandle,
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
] as const;

const libreOfficeImage = findLibreOfficeDockerImage();
const hasFixtures = LO_EDITOR_VALIDITY_CASES.every(
  (testCase) =>
    existsSync(join(FIXTURE_DIR, testCase.sourceFixture)) &&
    existsSync(join(FIXTURE_DIR, testCase.expectedFixture)),
);
const describeOrSkip = libreOfficeImage !== undefined && hasFixtures ? describe : describe.skip;

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

function findShapeByName(shapes: readonly SourceShapeNode[], name: string): SourceShape {
  const shape = flattenShapes(shapes).find((candidate) => candidate.name === name);
  if (shape === undefined) throw new Error(`Shape not found: ${name}`);
  return shape;
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
