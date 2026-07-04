#!/usr/bin/env bash
set -euo pipefail

# Create temporary directory and clean up on exit
WORK_DIR=$(mktemp -d)
trap 'rm -rf "$WORK_DIR"' EXIT

echo "=== Package publish verification ==="
echo "Working directory: $WORK_DIR"

# Generate tarballs for publishing target package and document workspace package with npm pack
TARBALL=$(cd packages/core && npm pack --pack-destination "$WORK_DIR" 2>/dev/null)
TARBALL_PATH="$WORK_DIR/$TARBALL"
DOCUMENT_TARBALL=$(cd packages/document && npm pack --pack-destination "$WORK_DIR" 2>/dev/null)
DOCUMENT_TARBALL_PATH="$WORK_DIR/$DOCUMENT_TARBALL"
echo "Packed: $TARBALL"
echo "Packed: $DOCUMENT_TARBALL"

# Install in core package test directory
TEST_DIR="$WORK_DIR/core-test-project"
mkdir -p "$TEST_DIR"
cd "$TEST_DIR"
npm init -y > /dev/null 2>&1
npm install "$TARBALL_PATH" > /dev/null 2>&1

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

console.log("  convertPptxToSvg: function OK");
console.log("  convertPptxToPng: function OK");
console.log("CJS test passed!");
TESTEOF
node test-cjs.cjs

echo ""

# --- core ESM test ---
echo "--- Test: pptx-glimpse ESM (import) ---"
cat > test-esm.mjs << 'TESTEOF'
import { convertPptxToSvg, convertPptxToPng } from "pptx-glimpse";

const assert = (condition, message) => {
  if (!condition) {
    console.error("FAIL:", message);
    process.exit(1);
  }
};

assert(typeof convertPptxToSvg === "function", "convertPptxToSvg should be a function");
assert(typeof convertPptxToPng === "function", "convertPptxToPng should be a function");

console.log("  convertPptxToSvg: function OK");
console.log("  convertPptxToPng: function OK");
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
import { collectUsedFonts, convertPptxToSvg, convertPptxToPng } from "pptx-glimpse";
import type { ConvertOptions, PngConversionReport, SvgConversionReport, UsedFonts } from "pptx-glimpse";

// Verify function signatures
const _svgFn: (input: Uint8Array, options?: ConvertOptions) => Promise<SvgConversionReport> =
  convertPptxToSvg;
const _pngFn: (input: Uint8Array, options?: ConvertOptions) => Promise<PngConversionReport> =
  convertPptxToPng;
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

// Verify ConvertOptions includes fontDirs
const _options: ConvertOptions = { slides: [1], width: 960, fontDirs: ["/custom/fonts"] };
void _svgFn;
void _pngFn;
void _fontFn;
void _options;
void _verifyPngType;
void _verifyBufferInput;
TESTEOF
npx tsc --noEmit
echo "TypeScript type resolution test passed!"

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
