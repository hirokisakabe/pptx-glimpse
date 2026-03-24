#!/usr/bin/env bash
set -euo pipefail

# 一時ディレクトリを作成し、終了時にクリーンアップ
WORK_DIR=$(mktemp -d)
trap 'rm -rf "$WORK_DIR"' EXIT

REPO_ROOT="$(cd "$(dirname "$0")/.." && pwd)"

echo "=== Package publish verification ==="
echo "Working directory: $WORK_DIR"

# renderer パッケージをパック
RENDERER_TARBALL=$(cd "$REPO_ROOT/packages/pptx-glimpse-renderer" && npm pack --pack-destination "$WORK_DIR" 2>/dev/null)
RENDERER_TARBALL_PATH="$WORK_DIR/$RENDERER_TARBALL"
echo "Packed renderer: $RENDERER_TARBALL"

# main パッケージをパック
MAIN_TARBALL=$(cd "$REPO_ROOT/packages/pptx-glimpse" && npm pack --pack-destination "$WORK_DIR" 2>/dev/null)
MAIN_TARBALL_PATH="$WORK_DIR/$MAIN_TARBALL"
echo "Packed main: $MAIN_TARBALL"

# テスト用ディレクトリにインストール
TEST_DIR="$WORK_DIR/test-project"
mkdir -p "$TEST_DIR"
cd "$TEST_DIR"
npm init -y > /dev/null 2>&1
npm install "$RENDERER_TARBALL_PATH" "$MAIN_TARBALL_PATH" > /dev/null 2>&1

echo ""

# --- CJS テスト ---
echo "--- Test: CJS (require) ---"
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

# --- ESM テスト ---
echo "--- Test: ESM (import) ---"
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
    "skipLibCheck": true
  },
  "include": ["test-types.ts"]
}
TESTEOF

cat > test-types.ts << 'TESTEOF'
import { convertPptxToSvg, convertPptxToPng } from "pptx-glimpse";
import type { ConvertOptions, SlideImage, SlideSvg } from "pptx-glimpse";

// Verify function signatures
const _svgFn: (input: Buffer | Uint8Array, options?: ConvertOptions) => Promise<SlideSvg[]> =
  convertPptxToSvg;
const _pngFn: (input: Buffer | Uint8Array, options?: ConvertOptions) => Promise<SlideImage[]> =
  convertPptxToPng;

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
void _options;
void _verifyPngType;
TESTEOF
npx tsc --noEmit
echo "TypeScript type resolution test passed!"

echo ""
echo "=== All package verification tests passed! ==="
