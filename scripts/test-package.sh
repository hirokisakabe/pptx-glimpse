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

CORE_PACKAGE_DIR="$WORK_DIR/core-package"
mkdir -p "$CORE_PACKAGE_DIR"
tar -xzf "$TARBALL_PATH" -C "$CORE_PACKAGE_DIR"
test -f "$CORE_PACKAGE_DIR/package/README.md"
echo "Core package README verification passed!"
node --input-type=module - "$CORE_PACKAGE_DIR/package/package.json" << 'TESTEOF'
import { readFileSync } from "node:fs";

const packageJson = JSON.parse(readFileSync(process.argv[2], "utf8"));
const expectedDependencies = [
  "@pptx-glimpse/document",
  "@pptx-glimpse/editor",
  "@resvg/resvg-wasm",
  "fast-xml-parser",
  "opentype.js",
];
const runtimeDependencies = [
  ...Object.keys(packageJson.dependencies ?? {}),
  ...Object.keys(packageJson.optionalDependencies ?? {}),
  ...Object.keys(packageJson.peerDependencies ?? {}),
].sort();
if (runtimeDependencies.join("\n") !== expectedDependencies.sort().join("\n")) {
  throw new Error(
    `pptx-glimpse runtime dependencies do not match generated JavaScript: ${runtimeDependencies.join(", ")}`,
  );
}
TESTEOF
node --input-type=module - "$CORE_PACKAGE_DIR/package/dist" << 'TESTEOF'
import { readdirSync, readFileSync } from "node:fs";
import { join } from "node:path";

const distDir = process.argv[2];
const files = readdirSync(distDir).filter((file) => /\.(?:[cm]?js|d\.[cm]?ts)$/.test(file));
const contents = new Map(
  files.map((file) => [file, readFileSync(join(distDir, file), "utf8")]),
);
const javascript = [...contents]
  .filter(([file]) => /\.[cm]?js$/.test(file))
  .map(([, content]) => content)
  .join("\n");
const declarations = [...contents]
  .filter(([file]) => /\.d\.[cm]?ts$/.test(file))
  .map(([, content]) => content)
  .join("\n");

for (const packageName of ["@pptx-glimpse/document", "@pptx-glimpse/editor"]) {
  if (!javascript.includes(packageName)) {
    throw new Error(`generated JavaScript must keep ${packageName} as a package import`);
  }
}
for (const privatePackageName of [
  "@pptx-glimpse/renderer",
  "@pptx-glimpse/editor-core",
]) {
  if (javascript.includes(privatePackageName) || declarations.includes(privatePackageName)) {
    throw new Error(`published output contains private package import: ${privatePackageName}`);
  }
}
for (const removedEditorApi of [
  "BrowserPptxEditorSession",
  "createBrowserPptxEditorSession",
  "BrowserEditor",
  "PptxEditorTextBodyInfo",
  "editableTextBody",
  "applyTextBodyDocJson",
]) {
  if (declarations.includes(removedEditorApi)) {
    throw new Error(`published declarations contain removed editor API: ${removedEditorApi}`);
  }
}
TESTEOF
echo "Core package boundary verification passed!"

DOCUMENT_PACKAGE_DIR="$WORK_DIR/document-package"
mkdir -p "$DOCUMENT_PACKAGE_DIR"
tar -xzf "$DOCUMENT_TARBALL_PATH" -C "$DOCUMENT_PACKAGE_DIR"
test -f "$DOCUMENT_PACKAGE_DIR/package/README.md"
echo "Document package README verification passed!"
node --input-type=module - "$DOCUMENT_PACKAGE_DIR/package/package.json" << 'TESTEOF'
import { readFileSync } from "node:fs";

const packageJson = JSON.parse(readFileSync(process.argv[2], "utf8"));
const runtimeDependencies = [
  ...Object.keys(packageJson.dependencies ?? {}),
  ...Object.keys(packageJson.optionalDependencies ?? {}),
  ...Object.keys(packageJson.peerDependencies ?? {}),
].sort();
if (runtimeDependencies.join("\n") !== ["fast-xml-parser", "fflate"].join("\n")) {
  throw new Error(
    `@pptx-glimpse/document runtime dependencies must contain only fast-xml-parser and fflate: ${runtimeDependencies.join(", ")}`,
  );
}
TESTEOF
echo "Document package boundary verification passed!"

EDITOR_PACKAGE_DIR="$WORK_DIR/editor-package"
mkdir -p "$EDITOR_PACKAGE_DIR"
tar -xzf "$EDITOR_TARBALL_PATH" -C "$EDITOR_PACKAGE_DIR"
test -f "$EDITOR_PACKAGE_DIR/package/README.md"
echo "Editor package README verification passed!"
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
cp "$REPO_DIR/shared-fixtures/real-basic-theme.pptx" "$TEST_DIR/fixture.pptx"
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
assert(
  typeof pkg.createPptxEditorSession === "function",
  "createPptxEditorSession should be a function",
);
assert(typeof pkg.PptxEditorSession === "function", "PptxEditorSession should be a class");
assert(typeof pkg.PptxEditorError === "function", "PptxEditorError should be a class");
assert(typeof pkg.isPptxEditorError === "function", "isPptxEditorError should be a function");

console.log("  convertPptxToSvg: function OK");
console.log("  convertPptxToPng: function OK");
console.log("  renderPptxSourceModelToSvg: function OK");
console.log("  createPptxEditorSession: function OK");
console.log("  PptxEditorSession: class OK");
console.log("  PptxEditorError: class OK");
console.log("  isPptxEditorError: function OK");
console.log("CJS test passed!");
TESTEOF
node test-cjs.cjs

echo ""

# --- core ESM test ---
echo "--- Test: pptx-glimpse ESM (import) ---"
cat > test-esm.mjs << 'TESTEOF'
import {
  convertPptxToSvg,
  convertPptxToPng,
  createPptxEditorSession,
  isPptxEditorError,
  PptxEditorError,
  PptxEditorSession,
  renderPptxSourceModelToSvg,
} from "pptx-glimpse";
import { readFile } from "node:fs/promises";

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
assert(
  typeof createPptxEditorSession === "function",
  "createPptxEditorSession should be a function",
);
assert(typeof PptxEditorSession === "function", "PptxEditorSession should be a class");
assert(typeof PptxEditorError === "function", "PptxEditorError should be a class");
assert(typeof isPptxEditorError === "function", "isPptxEditorError should be a function");

console.log("  convertPptxToSvg: function OK");
console.log("  convertPptxToPng: function OK");
console.log("  renderPptxSourceModelToSvg: function OK");
console.log("  createPptxEditorSession: function OK");
console.log("  PptxEditorSession: class OK");
console.log("  PptxEditorError: class OK");
console.log("  isPptxEditorError: function OK");
const input = new Uint8Array(await readFile("fixture.pptx"));
const result = await convertPptxToSvg(input, { skipSystemFonts: true });
assert(result.slides[0]?.svg.startsWith("<svg"), "Node SVG conversion should produce SVG");
console.log("  Node SVG conversion: OK");
const editor = await createPptxEditorSession(input, { skipSystemFonts: true });
const run = editor
  .shapes(1)
  .flatMap((shape) => shape.textBody?.paragraphs ?? [])
  .flatMap((paragraph) => paragraph.runs)
  .find((candidate) => candidate.handle !== undefined);
assert(run?.handle !== undefined, "Node editor fixture text run should exist");
const edited = await editor.apply({
  kind: "replaceTextRunPlainText",
  handle: run.handle,
  text: "Node package consumer edited",
});
assert(edited.slides[0]?.svg.startsWith("<svg"), "Node editor should rerender SVG");
assert(editor.save().pptx instanceof Uint8Array, "Node editor should save Uint8Array");
console.log("  Node high-level editor: OK");
console.log("ESM test passed!");
TESTEOF
node test-esm.mjs

echo ""

# --- core browser consumer bundle and execution test ---
echo "--- Test: pptx-glimpse browser consumer bundle ---"
npm install --save-dev esbuild > /dev/null 2>&1
cat > browser-entry.mjs << 'TESTEOF'
import {
  convertPptxToSvg,
  createPptxEditorSession,
  isPptxEditorError,
  PptxEditorError,
} from "pptx-glimpse";

export async function verifyBrowserApis(input) {
  const sampleError = new PptxEditorError("invalid-command", "sample");
  if (!isPptxEditorError(sampleError) || sampleError.code !== "invalid-command") {
    throw new Error("browser editor error API is unavailable");
  }
  const converted = await convertPptxToSvg(input, { skipSystemFonts: true });
  if (!converted.slides[0]?.svg.startsWith("<svg")) {
    throw new Error("browser SVG conversion did not produce SVG");
  }

  const session = await createPptxEditorSession(input, { skipSystemFonts: true });
  const run = session
    .shapes(1)
    .flatMap((shape) => shape.textBody?.paragraphs ?? [])
    .flatMap((paragraph) => paragraph.runs)
    .find((candidate) => candidate.handle !== undefined);
  if (run?.handle === undefined) {
    throw new Error("browser editor fixture text run not found");
  }
  const response = await session.apply({
    kind: "replaceTextRunPlainText",
    handle: run.handle,
    text: "Browser package consumer edited",
  });
  if (!response.slides[0]?.svg.startsWith("<svg")) {
    throw new Error("browser editor did not rerender SVG");
  }
  if (!(session.save().pptx instanceof Uint8Array)) {
    throw new Error("browser editor save did not return Uint8Array");
  }
}
TESTEOF
cat > build-browser.mjs << 'TESTEOF'
import { build } from "esbuild";

await build({
  entryPoints: ["browser-entry.mjs"],
  bundle: true,
  outfile: "browser-bundle.mjs",
  format: "esm",
  platform: "browser",
});
TESTEOF
node build-browser.mjs
if grep -E 'node:(fs|path|buffer)|from "(fs|path)"' browser-bundle.mjs; then
  echo "FAIL: pptx-glimpse browser bundle contains a Node built-in"
  exit 1
fi
cat > run-browser-bundle.mjs << 'TESTEOF'
import { readFile } from "node:fs/promises";

import { verifyBrowserApis } from "./browser-bundle.mjs";

await verifyBrowserApis(new Uint8Array(await readFile("fixture.pptx")));
console.log("Browser consumer bundle test passed!");
TESTEOF
node run-browser-bundle.mjs

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
    "skipLibCheck": false,
    "types": ["node"]
  },
  "include": ["test-types.ts"]
}
TESTEOF

cat > test-types.ts << 'TESTEOF'
import {
  collectUsedFonts,
  convertPptxToSvg,
  convertPptxToPng,
  createPptxEditorSession,
  isPptxEditorError,
  PptxEditorError,
  renderPptxSourceModelToSvg,
} from "pptx-glimpse";
import type {
  ConvertOptions,
  EditorCommand,
  FontBuffer,
  FontMapping,
  OpentypeSetup,
  PngConversionReport,
  PptxEditorErrorCode,
  PptxEditorRenderOptions,
  PptxEditorSession,
  PptxEditorShapeInfo,
  PptxSourceModel,
  ResvgWasmInput,
  SvgConversionReport,
  UsedFonts,
} from "pptx-glimpse";

// Verify function signatures
const _svgFn: (input: Uint8Array, options?: ConvertOptions) => Promise<SvgConversionReport> =
  convertPptxToSvg;
const _pngFn: (input: Uint8Array, options?: ConvertOptions) => Promise<PngConversionReport> =
  convertPptxToPng;
const _sourceModelSvgFn: (
  source: PptxSourceModel,
  options?: ConvertOptions,
) => Promise<SvgConversionReport> = renderPptxSourceModelToSvg;
const _fontFn: (input: Uint8Array) => UsedFonts = collectUsedFonts;
const _editorFn: (
  input: Uint8Array,
  options?: PptxEditorRenderOptions,
) => Promise<PptxEditorSession> = createPptxEditorSession;
declare const _shapeInfo: PptxEditorShapeInfo;

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

// Verify ConvertOptions includes fontDirs
const _options: ConvertOptions = { slides: [1], width: 960, fontDirs: ["/custom/fonts"] };
const _fontBuffer: FontBuffer = { name: "Test", data: new Uint8Array() };
const _fontMapping: FontMapping = { Arial: "Inter" };
declare const _opentypeSetup: OpentypeSetup;
const _wasm: ResvgWasmInput = new Uint8Array();
declare const _editorCommand: EditorCommand;
const _editorError = new PptxEditorError("invalid-command", "typed");
const _editorErrorCode: PptxEditorErrorCode = _editorError.code;
const _isEditorError: boolean = isPptxEditorError(_editorError);
void _svgFn;
void _pngFn;
void _sourceModelSvgFn;
void _fontFn;
void _editorFn;
void _shapeInfo;
void _options;
void _fontBuffer;
void _fontMapping;
void _opentypeSetup;
void _wasm;
void _editorCommand;
void _editorErrorCode;
void _isEditorError;
void _verifyPngType;
void _verifyBufferInput;
TESTEOF
npx tsc --noEmit

cat > tsconfig-browser.json << 'TESTEOF'
{
  "compilerOptions": {
    "target": "ES2022",
    "module": "ESNext",
    "moduleResolution": "Bundler",
    "customConditions": ["browser"],
    "strict": true,
    "noEmit": true,
    "skipLibCheck": false,
    "lib": ["ES2022", "DOM"]
  },
  "include": ["test-browser-types.ts"]
}
TESTEOF
cat > test-browser-types.ts << 'TESTEOF'
import {
  createPptxEditorSession,
  initResvgWasm,
  isPptxEditorError,
  PptxEditorError,
  type EditorCommand,
  type PptxEditorErrorCode,
  type PptxEditorSession,
  type PptxEditorShapeInfo,
  type PptxSourceModel,
  type ResvgWasmInput,
} from "pptx-glimpse";

const _create: (input: Uint8Array) => Promise<PptxEditorSession> = createPptxEditorSession;
const _init: (wasm: ResvgWasmInput) => Promise<void> = initResvgWasm;
const _editorError = new PptxEditorError("render-failed", "typed browser error");
const _editorErrorCode: PptxEditorErrorCode = _editorError.code;
const _isEditorError: boolean = isPptxEditorError(_editorError);
declare const _source: PptxSourceModel;
declare const _shape: PptxEditorShapeInfo;
const _command: EditorCommand = {
  kind: "replaceTextRunPlainText",
  handle: _source.slides[0].handle!,
  text: "typed",
};
void _create;
void _init;
void _editorErrorCode;
void _isEditorError;
void _command;
void _shape;
TESTEOF
npx tsc -p tsconfig-browser.json
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
  EditorOperationErrorCode,
  EditorOperationFailure,
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
const _operationCode: EditorOperationErrorCode = "invalid-command";
declare const _operationFailure: EditorOperationFailure;
void _create;
void _editorCommand;
void _apply;
void _history;
void _selection;
void _selectResult;
void _warning;
void _operationCode;
void _operationFailure;
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
