#!/usr/bin/env bash
set -euo pipefail

REPO_DIR=$(pwd)

# Create temporary directory and clean up on exit
WORK_DIR=$(mktemp -d)
trap 'rm -rf "$WORK_DIR"' EXIT

echo "=== Package publish verification ==="
echo "Working directory: $WORK_DIR"

# Generate tarballs for publishing target packages with pnpm pack.
# pnpm rewrites workspace: dependency ranges the same way publish does.
pnpm --dir packages/core pack --pack-destination "$WORK_DIR" > /dev/null
TARBALL_PATH=$(find "$WORK_DIR" -maxdepth 1 -type f -name "pptx-glimpse-[0-9]*.tgz" -print -quit)
TARBALL=$(basename "$TARBALL_PATH")
pnpm --dir packages/document pack --pack-destination "$WORK_DIR" > /dev/null
DOCUMENT_TARBALL_PATH=$(find "$WORK_DIR" -maxdepth 1 -type f -name "pptx-glimpse-document-*.tgz" -print -quit)
DOCUMENT_TARBALL=$(basename "$DOCUMENT_TARBALL_PATH")
pnpm --dir packages/editor pack --pack-destination "$WORK_DIR" > /dev/null
EDITOR_TARBALL_PATH=$(find "$WORK_DIR" -maxdepth 1 -type f -name "pptx-glimpse-editor-*.tgz" -print -quit)
EDITOR_TARBALL=$(basename "$EDITOR_TARBALL_PATH")
echo "Packed: $TARBALL"
echo "Packed: $DOCUMENT_TARBALL"
echo "Packed: $EDITOR_TARBALL"

EDITOR_PACKAGE_DIR="$WORK_DIR/editor-package"
mkdir -p "$EDITOR_PACKAGE_DIR"
tar -xzf "$EDITOR_TARBALL_PATH" -C "$EDITOR_PACKAGE_DIR"
node --input-type=module - "$EDITOR_PACKAGE_DIR/package/package.json" << 'TESTEOF'
import { readFileSync } from "node:fs";

const packageJson = JSON.parse(readFileSync(process.argv[2], "utf8"));
const runtimeDependencies = [
  ...Object.keys(packageJson.dependencies ?? {}),
  ...Object.keys(packageJson.optionalDependencies ?? {}),
  ...Object.keys(packageJson.peerDependencies ?? {}),
];
if (
  runtimeDependencies.length !== 1 ||
  runtimeDependencies[0] !== "@pptx-glimpse/document"
) {
  throw new Error(
    `@pptx-glimpse/editor runtime dependencies must contain only @pptx-glimpse/document: ${runtimeDependencies.join(", ")}`,
  );
}
if (packageJson.private === true) {
  throw new Error("@pptx-glimpse/editor must be publishable");
}
TESTEOF
if grep -R -i "prosemirror" "$EDITOR_PACKAGE_DIR/package/dist"; then
  echo "FAIL: @pptx-glimpse/editor dist must not expose ProseMirror"
  exit 1
fi
echo "Editor package boundary verification passed!"

# Install in core package test directory
TEST_DIR="$WORK_DIR/core-test-project"
mkdir -p "$TEST_DIR"
cd "$TEST_DIR"
npm init -y > /dev/null 2>&1
npm install "$DOCUMENT_TARBALL_PATH" "$EDITOR_TARBALL_PATH" "$TARBALL_PATH" > /dev/null 2>&1

echo ""

# --- core CJS test ---
echo "--- Test: pptx-glimpse CJS (require) ---"
cat > test-cjs.cjs << 'TESTEOF'
const pkg = require("pptx-glimpse");

const assert = (condition, message) => {
  if (!condition) {
    console.error("FAIL:", message);
    process.exit(1);
  }
};

assert(typeof pkg.convertPptxToSvg === "function", "convertPptxToSvg should be a function");
assert(typeof pkg.convertPptxToPng === "function", "convertPptxToPng should be a function");
assert(
  typeof pkg.renderPptxSourceModelToSvg === "function",
  "renderPptxSourceModelToSvg should be a function",
);

console.log("  convertPptxToSvg: function OK");
console.log("  convertPptxToPng: function OK");
console.log("  renderPptxSourceModelToSvg: function OK");
console.log("CJS test passed!");
TESTEOF
node test-cjs.cjs

echo ""

# --- core ESM test ---
echo "--- Test: pptx-glimpse ESM (import) ---"
cat > test-esm.mjs << 'TESTEOF'
import { convertPptxToSvg, convertPptxToPng, renderPptxSourceModelToSvg } from "pptx-glimpse";

const assert = (condition, message) => {
  if (!condition) {
    console.error("FAIL:", message);
    process.exit(1);
  }
};

assert(typeof convertPptxToSvg === "function", "convertPptxToSvg should be a function");
assert(typeof convertPptxToPng === "function", "convertPptxToPng should be a function");
assert(
  typeof renderPptxSourceModelToSvg === "function",
  "renderPptxSourceModelToSvg should be a function",
);

console.log("  convertPptxToSvg: function OK");
console.log("  convertPptxToPng: function OK");
console.log("  renderPptxSourceModelToSvg: function OK");
console.log("ESM test passed!");
TESTEOF
node test-esm.mjs

echo ""

# --- core TypeScript type resolution test ---
echo "--- Test: pptx-glimpse TypeScript type resolution ---"
npm install typescript@latest @types/node > /dev/null 2>&1

# Set test project to ESM (pptx-glimpse is "type": "module")
npm pkg set type=module > /dev/null 2>&1

cat > tsconfig.json << 'TESTEOF'
{
  "compilerOptions": {
    "target": "ES2022",
    "module": "Node16",
    "moduleResolution": "Node16",
    "strict": true,
    "noEmit": true,
    "skipLibCheck": true,
    "types": ["node"]
  },
  "include": ["test-types.ts"]
}
TESTEOF

cat > test-types.ts << 'TESTEOF'
import { readPptx } from "@pptx-glimpse/document";
import { collectUsedFonts, convertPptxToSvg, convertPptxToPng, renderPptxSourceModelToSvg } from "pptx-glimpse";
import type {
  ConvertOptions,
  PngConversionReport,
  PptxSourceModel,
  SvgConversionReport,
  UsedFonts,
} from "pptx-glimpse";

// Verify function signatures
const _svgFn: (input: Uint8Array, options?: ConvertOptions) => Promise<SvgConversionReport> =
  convertPptxToSvg;
const _pngFn: (input: Uint8Array, options?: ConvertOptions) => Promise<PngConversionReport> =
  convertPptxToPng;
const _sourceModelSvgFn: (
  source: ReturnType<typeof readPptx>,
  options?: ConvertOptions,
) => Promise<SvgConversionReport> = renderPptxSourceModelToSvg;
const _fontFn: (input: Uint8Array) => UsedFonts = collectUsedFonts;

// Verify SlideImage.png is Uint8Array
async function _verifyPngType(input: Uint8Array) {
  const { slides } = await convertPptxToPng(input);
  const _png: Uint8Array = slides[0].png;
  void _png;
}

// Verify Node Buffer remains accepted as a Uint8Array subclass.
function _verifyBufferInput(input: Buffer) {
  void convertPptxToSvg(input);
  void convertPptxToPng(input);
  void collectUsedFonts(input);
}

// Verify @pptx-glimpse/document readPptx() output is accepted by the source-model render API.
async function _verifyDocumentSourceModel(input: Uint8Array) {
  const source = readPptx(input);
  const _source: PptxSourceModel = source;
  await renderPptxSourceModelToSvg(source);
  void _source;
}

// Verify ConvertOptions includes fontDirs
const _options: ConvertOptions = { slides: [1], width: 960, fontDirs: ["/custom/fonts"] };
void _svgFn;
void _pngFn;
void _sourceModelSvgFn;
void _fontFn;
void _options;
void _verifyPngType;
void _verifyBufferInput;
void _verifyDocumentSourceModel;
TESTEOF
npx tsc --noEmit
echo "TypeScript type resolution test passed!"

echo ""

# Install in the test directory of the editor package.
EDITOR_TEST_DIR="$WORK_DIR/editor-test-project"
mkdir -p "$EDITOR_TEST_DIR"
cp "$REPO_DIR/shared-fixtures/real-basic-theme.pptx" "$EDITOR_TEST_DIR/fixture.pptx"
cd "$EDITOR_TEST_DIR"
npm init -y > /dev/null 2>&1
npm install "$DOCUMENT_TARBALL_PATH" "$EDITOR_TARBALL_PATH" > /dev/null 2>&1

# --- editor Node CJS test ---
echo "--- Test: @pptx-glimpse/editor Node CJS consumer ---"
cat > test-editor-node.cjs << 'TESTEOF'
const { readFileSync } = require("node:fs");

const { findTextRunBySourceHandle, readPptx } = require("@pptx-glimpse/document");
const { createEditorSession } = require("@pptx-glimpse/editor");

const source = readPptx(readFileSync("fixture.pptx"));
const run = source.slides
  .flatMap((slide) => slide.shapes)
  .find((shape) => shape.kind === "shape" && shape.textBody?.paragraphs[0]?.runs[0]?.handle)
  ?.textBody?.paragraphs[0]?.runs[0];
if (!run?.handle) throw new Error("fixture text run not found");

const session = createEditorSession(source);
const result = session.apply({
  kind: "replaceTextRunPlainText",
  handle: run.handle,
  text: "CJS consumer edited",
});
if (!result.ok) throw new Error(result.message);
if (findTextRunBySourceHandle(session.document, run.handle)?.text !== "CJS consumer edited") {
  throw new Error("CJS consumer edit was not applied");
}

console.log("Editor Node CJS consumer test passed!");
TESTEOF
node test-editor-node.cjs

echo ""

# --- editor Node ESM test ---
echo "--- Test: @pptx-glimpse/editor Node ESM consumer ---"
cat > test-editor-node.mjs << 'TESTEOF'
import { readFileSync } from "node:fs";

import { findTextRunBySourceHandle, readPptx } from "@pptx-glimpse/document";
import { createEditorSession } from "@pptx-glimpse/editor";

const source = readPptx(readFileSync("fixture.pptx"));
const run = source.slides
  .flatMap((slide) => slide.shapes)
  .find((shape) => shape.kind === "shape" && shape.textBody?.paragraphs[0]?.runs[0]?.handle)
  ?.textBody?.paragraphs[0]?.runs[0];
if (!run?.handle) throw new Error("fixture text run not found");

const session = createEditorSession(source);
const result = session.apply({
  kind: "replaceTextRunPlainText",
  handle: run.handle,
  text: "Package consumer edited",
});
if (!result.ok) throw new Error(result.message);
if (findTextRunBySourceHandle(session.document, run.handle)?.text !== "Package consumer edited") {
  throw new Error("ESM consumer edit was not applied");
}

console.log("Editor Node ESM consumer test passed!");
TESTEOF
node test-editor-node.mjs

echo ""

# --- editor browser bundle test ---
echo "--- Test: @pptx-glimpse/editor browser consumer bundle ---"
npm install --save-dev esbuild > /dev/null 2>&1
cat > test-editor-browser.mjs << 'TESTEOF'
import { build } from "esbuild";

const result = await build({
  stdin: {
    contents: `
      import { readPptx } from "@pptx-glimpse/document";
      import { createEditorSession } from "@pptx-glimpse/editor";
      export function edit(input, handle) {
        const session = createEditorSession(readPptx(input));
        return session.apply({ kind: "replaceTextRunPlainText", handle, text: "Browser edited" });
      }
    `,
    loader: "js",
    resolveDir: process.cwd(),
  },
  bundle: true,
  write: false,
  format: "esm",
  platform: "browser",
});
const bundled = result.outputFiles[0]?.text ?? "";
if (!bundled.includes("replaceTextRunPlainText")) {
  throw new Error("browser bundle does not contain the editor command implementation");
}
if (/(?:node:fs|node:path|node:buffer|from "fs"|from "path")/.test(bundled)) {
  throw new Error("browser bundle contains a Node built-in");
}

console.log("Editor browser consumer bundle test passed!");
TESTEOF
node test-editor-browser.mjs

echo ""

# --- editor TypeScript public API test ---
echo "--- Test: @pptx-glimpse/editor TypeScript public API ---"
npm install --save-dev typescript@latest > /dev/null 2>&1
npm pkg set type=module > /dev/null 2>&1
cat > tsconfig.json << 'TESTEOF'
{
  "compilerOptions": {
    "target": "ES2022",
    "module": "Node16",
    "moduleResolution": "Node16",
    "strict": true,
    "noEmit": true,
    "skipLibCheck": true
  },
  "include": ["test-editor-types.ts", "test-editor-cjs.cts"]
}
TESTEOF
cat > test-editor-types.ts << 'TESTEOF'
import { createEditorSession, EditorSession } from "@pptx-glimpse/editor";
import type {
  AddConnectorCommand,
  AddEmptySlideFromLayoutCommand,
  AddTextBoxCommand,
  ClearParagraphPropertiesCommand,
  ClearTextRunPropertiesCommand,
  DeleteShapeCommand,
  DeleteSlideCommand,
  DuplicateSlideCommand,
  EditorApplyCommandResult,
  EditorCommand,
  EditorCommandWarning,
  EditorHistoryResult,
  EditorSelection,
  EditorSelectShapeResult,
  MoveShapeCommand,
  MoveSlideCommand,
  ReplaceImageCommand,
  ReplaceParagraphPlainTextCommand,
  ReplaceTextRunPlainTextCommand,
  ResizeShapeCommand,
  SetParagraphPropertiesCommand,
  SetShapeFillCommand,
  SetShapeOutlineCommand,
  SetShapeTransformCommand,
  SetTextRunPropertiesCommand,
} from "@pptx-glimpse/editor";
import type { PptxSourceModel } from "@pptx-glimpse/document";

const _create: (document: PptxSourceModel) => EditorSession = createEditorSession;
type _AllCommands =
  | AddConnectorCommand
  | AddEmptySlideFromLayoutCommand
  | AddTextBoxCommand
  | ClearParagraphPropertiesCommand
  | ClearTextRunPropertiesCommand
  | DeleteShapeCommand
  | DeleteSlideCommand
  | DuplicateSlideCommand
  | MoveShapeCommand
  | MoveSlideCommand
  | ReplaceImageCommand
  | ReplaceParagraphPlainTextCommand
  | ReplaceTextRunPlainTextCommand
  | ResizeShapeCommand
  | SetParagraphPropertiesCommand
  | SetShapeFillCommand
  | SetShapeOutlineCommand
  | SetShapeTransformCommand
  | SetTextRunPropertiesCommand;
declare const _command: _AllCommands;
const _editorCommand: EditorCommand = _command;
declare const _apply: EditorApplyCommandResult;
declare const _history: EditorHistoryResult;
declare const _selection: EditorSelection;
declare const _selectResult: EditorSelectShapeResult;
declare const _warning: EditorCommandWarning;
void _create;
void _editorCommand;
void _apply;
void _history;
void _selection;
void _selectResult;
void _warning;
TESTEOF
cat > test-editor-cjs.cts << 'TESTEOF'
import editor = require("@pptx-glimpse/editor");

declare const source: Parameters<typeof editor.createEditorSession>[0];
const _session: editor.EditorSession = editor.createEditorSession(source);
void _session;
TESTEOF
npx tsc --noEmit
echo "Editor TypeScript public API test passed!"

echo ""

# Install in the test directory of the document package
DOCUMENT_TEST_DIR="$WORK_DIR/document-test-project"
mkdir -p "$DOCUMENT_TEST_DIR"
cd "$DOCUMENT_TEST_DIR"
npm init -y > /dev/null 2>&1
npm install "$DOCUMENT_TARBALL_PATH" > /dev/null 2>&1

# --- document CJS test ---
echo "--- Test: @pptx-glimpse/document CJS (require) ---"
cat > test-document-cjs.cjs << 'TESTEOF'
const documentPkg = require("@pptx-glimpse/document");

const assert = (condition, message) => {
  if (!condition) {
    console.error("FAIL:", message);
    process.exit(1);
  }
};

assert(typeof documentPkg.readPptx === "function", "readPptx should be a function");
assert(typeof documentPkg.createComputedView === "function", "createComputedView should be a function");
assert(typeof documentPkg.writePptx === "function", "writePptx should be a function");
assert(
  typeof documentPkg.replaceTextRunPlainText === "function",
  "replaceTextRunPlainText should be a function",
);

console.log("  @pptx-glimpse/document root CJS: function OK");
console.log("Document CJS test passed!");
TESTEOF
node test-document-cjs.cjs

echo ""

# --- document ESM test ---
echo "--- Test: @pptx-glimpse/document ESM (import) ---"
cat > test-document-esm.mjs << 'TESTEOF'
import {
  createComputedView,
  readPptx,
  replaceTextRunPlainText,
  writePptx,
} from "@pptx-glimpse/document";

const assert = (condition, message) => {
  if (!condition) {
    console.error("FAIL:", message);
    process.exit(1);
  }
};

assert(typeof readPptx === "function", "readPptx should be a function");
assert(typeof createComputedView === "function", "createComputedView should be a function");
assert(typeof writePptx === "function", "writePptx should be a function");
assert(
  typeof replaceTextRunPlainText === "function",
  "replaceTextRunPlainText should be a function",
);

console.log("  @pptx-glimpse/document root ESM: function OK");
console.log("Document ESM test passed!");
TESTEOF
node test-document-esm.mjs

echo ""

# --- document TypeScript type resolution test ---
echo "--- Test: @pptx-glimpse/document TypeScript type resolution ---"
npm install typescript@latest @types/node > /dev/null 2>&1
npm pkg set type=module > /dev/null 2>&1

cat > tsconfig.json << 'TESTEOF'
{
  "compilerOptions": {
    "target": "ES2022",
    "module": "Node16",
    "moduleResolution": "Node16",
    "strict": true,
    "noEmit": true,
    "skipLibCheck": true,
    "types": ["node"]
  },
  "include": ["test-document-types.ts"]
}
TESTEOF

cat > test-document-types.ts << 'TESTEOF'
import {
  asPartPath,
  asRawSidecarId,
  createComputedView,
  findTextRunBySourceHandle,
  readPptx,
  replaceTextRunPlainText,
  writePptx,
} from "@pptx-glimpse/document";
import type {
  PptxComputedView,
  PptxSourceModel,
  RawOoxmlNode,
  RawPackagePart,
  RawSidecar,
  ReadPptxInput,
  SourceHandle,
  SourceTextRun,
  WritePptxOutput,
} from "@pptx-glimpse/document";

const _readPptxFn: (input: ReadPptxInput) => PptxSourceModel = readPptx;
const _createComputedViewFn: (source: PptxSourceModel) => PptxComputedView = createComputedView;
const _writePptxFn: (source: PptxSourceModel) => WritePptxOutput = writePptx;
const _replaceTextRunPlainTextFn: (
  source: PptxSourceModel,
  handle: SourceHandle,
  text: string,
) => PptxSourceModel = replaceTextRunPlainText;
const _findTextRunBySourceHandleFn: (
  source: PptxSourceModel,
  handle: SourceHandle,
) => SourceTextRun | undefined = findTextRunBySourceHandle;

const _rawNode: RawOoxmlNode = { name: "p:extLst" };
const _rawSidecar: RawSidecar = { id: asRawSidecarId("raw-1"), node: _rawNode };
const _rawPart: RawPackagePart = {
  kind: "xml",
  partPath: asPartPath("ppt/customXml/item1.xml"),
  contentType: "application/xml",
  xml: _rawNode,
};

void _readPptxFn;
void _createComputedViewFn;
void _writePptxFn;
void _replaceTextRunPlainTextFn;
void _findTextRunBySourceHandleFn;
void _rawSidecar;
void _rawPart;
TESTEOF
npx tsc --noEmit
echo "Document TypeScript type resolution test passed!"

echo ""
echo "=== All package verification tests passed! ==="
