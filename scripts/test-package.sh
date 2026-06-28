#!/usr/bin/env bash
set -euo pipefail

# 一時ディレクトリを作成し、終了時にクリーンアップ
WORK_DIR=$(mktemp -d)
trap 'rm -rf "$WORK_DIR"' EXIT

echo "=== Package publish verification ==="
echo "Working directory: $WORK_DIR"

# npm pack で公開対象 package と document workspace package のタールボールを生成
TARBALL=$(cd packages/core && npm pack --pack-destination "$WORK_DIR" 2>/dev/null)
TARBALL_PATH="$WORK_DIR/$TARBALL"
DOCUMENT_TARBALL=$(cd packages/document && npm pack --pack-destination "$WORK_DIR" 2>/dev/null)
DOCUMENT_TARBALL_PATH="$WORK_DIR/$DOCUMENT_TARBALL"
echo "Packed: $TARBALL"
echo "Packed: $DOCUMENT_TARBALL"

# テスト用ディレクトリにインストール
TEST_DIR="$WORK_DIR/test-project"
mkdir -p "$TEST_DIR"
cd "$TEST_DIR"
npm init -y > /dev/null 2>&1
npm install "$TARBALL_PATH" "$DOCUMENT_TARBALL_PATH" > /dev/null 2>&1

echo ""

# --- CJS テスト ---
echo "--- Test: CJS (require) ---"
cat > test-cjs.cjs << 'TESTEOF'
const pkg = require("pptx-glimpse");
const documentPkg = require("@pptx-glimpse/document");

const assert = (condition, message) => {
  if (!condition) {
    console.error("FAIL:", message);
    process.exit(1);
  }
};

assert(typeof pkg.convertPptxToSvg === "function", "convertPptxToSvg should be a function");
assert(typeof pkg.convertPptxToPng === "function", "convertPptxToPng should be a function");
assert(typeof documentPkg.readPptx === "function", "readPptx should be a function");
assert(typeof documentPkg.createComputedView === "function", "createComputedView should be a function");
assert(typeof documentPkg.writePptx === "function", "writePptx should be a function");
assert(
  typeof documentPkg.replaceTextRunPlainText === "function",
  "replaceTextRunPlainText should be a function",
);

console.log("  convertPptxToSvg: function OK");
console.log("  convertPptxToPng: function OK");
console.log("  @pptx-glimpse/document root CJS: function OK");
console.log("CJS test passed!");
TESTEOF
node test-cjs.cjs

echo ""

# --- ESM テスト ---
echo "--- Test: ESM (import) ---"
cat > test-esm.mjs << 'TESTEOF'
import { convertPptxToSvg, convertPptxToPng } from "pptx-glimpse";
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

assert(typeof convertPptxToSvg === "function", "convertPptxToSvg should be a function");
assert(typeof convertPptxToPng === "function", "convertPptxToPng should be a function");
assert(typeof readPptx === "function", "readPptx should be a function");
assert(typeof createComputedView === "function", "createComputedView should be a function");
assert(typeof writePptx === "function", "writePptx should be a function");
assert(
  typeof replaceTextRunPlainText === "function",
  "replaceTextRunPlainText should be a function",
);

console.log("  convertPptxToSvg: function OK");
console.log("  convertPptxToPng: function OK");
console.log("  @pptx-glimpse/document root ESM: function OK");
console.log("ESM test passed!");
TESTEOF
node test-esm.mjs

echo ""

# --- TypeScript 型解決テスト ---
echo "--- Test: TypeScript type resolution ---"
npm install typescript@latest @types/node > /dev/null 2>&1

# テスト用プロジェクトを ESM に設定 (pptx-glimpse は "type": "module")
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
import { convertPptxToSvg, convertPptxToPng } from "pptx-glimpse";
import type { ConvertOptions, SlideImage, SlideSvg } from "pptx-glimpse";
import {
  createComputedView,
  findTextRunBySourceHandle,
  readPptx,
  replaceTextRunPlainText,
  writePptx,
} from "@pptx-glimpse/document";
import type {
  PptxComputedView,
  PptxSourceModel,
  ReadPptxInput,
  SourceHandle,
  SourceTextRun,
  WritePptxOutput,
} from "@pptx-glimpse/document";

// Verify function signatures
const _svgFn: (input: Buffer | Uint8Array, options?: ConvertOptions) => Promise<SlideSvg[]> =
  convertPptxToSvg;
const _pngFn: (input: Buffer | Uint8Array, options?: ConvertOptions) => Promise<SlideImage[]> =
  convertPptxToPng;
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

// Verify SlideImage.png is Buffer
async function _verifyPngType(input: Uint8Array) {
  const results = await convertPptxToPng(input);
  const _png: Buffer = results[0].png;
  void _png;
}

// Verify ConvertOptions includes fontDirs
const _options: ConvertOptions = { slides: [1], width: 960, fontDirs: ["/custom/fonts"] };
void _svgFn;
void _pngFn;
void _readPptxFn;
void _createComputedViewFn;
void _writePptxFn;
void _replaceTextRunPlainTextFn;
void _findTextRunBySourceHandleFn;
void _options;
void _verifyPngType;
TESTEOF
npx tsc --noEmit
echo "TypeScript type resolution test passed!"

echo ""
echo "=== All package verification tests passed! ==="
