import { readFileSync } from "node:fs";
import { fileURLToPath } from "node:url";

import { describe, expect, it } from "vitest";

import { convertPptxToSvg } from "../packages/core/src/converter.js";
import {
  asEmu,
  findShapeNodeBySourceHandle,
  type PptxSourceModel,
  readPptx,
  type SourceHandle,
  type SourceShapeNode,
  type SourceTransform,
  writePptx,
} from "../packages/document/src/index.js";
import {
  createEditorSession,
  type EditorApplyCommandResult,
} from "../packages/editor-core/src/index.js";

describe("editor-core xfrm rendering integration", () => {
  it("renders edited shape position and size after move and resize commands", async () => {
    const input = readSharedFixture("real-product-page.pptx");
    const source = readPptx(input);
    const editable = firstTransformShape(source);
    const session = createEditorSession(source);

    expectApplied(
      session.apply({
        kind: "moveShape",
        handle: editable.shape.handle,
        offsetX: asEmu(914400),
        offsetY: asEmu(1828800),
      }),
    );
    const edited = expectApplied(
      session.apply({
        kind: "resizeShape",
        handle: editable.shape.handle,
        width: asEmu(2743200),
        height: asEmu(914400),
      }),
    );
    const output = writePptx(edited);
    const reread = readPptx(output);
    const rereadShape = findShapeNodeBySourceHandle(reread, editable.shape.handle);

    expect(rereadShape?.transform).toMatchObject({
      offsetX: 914400,
      offsetY: 1828800,
      width: 2743200,
      height: 914400,
    });

    const { slides } = await convertPptxToSvg(output, {
      slides: [editable.slideNumber],
      textOutput: "text",
      skipSystemFonts: true,
    });

    expect(slides).toHaveLength(1);
    expect(slides[0].svg).toContain('transform="translate(96, 192)"');
    expect(slides[0].svg).toContain('width="288" height="96"');
  });
});

interface EditableShape {
  readonly slideNumber: number;
  readonly shape: SourceShapeNode & {
    readonly handle: SourceHandle;
    readonly transform: SourceTransform;
  };
}

function readSharedFixture(name: string): Buffer {
  return readFileSync(fileURLToPath(new URL(`../shared-fixtures/${name}`, import.meta.url)));
}

function firstTransformShape(source: PptxSourceModel): EditableShape {
  for (const [slideIndex, slide] of source.slides.entries()) {
    for (const shape of slide.shapes) {
      if (shape.kind === "raw" || shape.handle === undefined || shape.transform === undefined) {
        continue;
      }
      return { slideNumber: slideIndex + 1, shape };
    }
  }
  throw new Error("editable shape not found");
}

function expectApplied(result: EditorApplyCommandResult): PptxSourceModel {
  if (!result.ok) throw new Error(result.message);
  return result.document;
}
