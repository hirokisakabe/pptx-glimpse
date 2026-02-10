# pptx-glimpse

PPTX スライドを SVG / PNG に変換する TypeScript ライブラリ。

## インストール

```bash
npm install pptx-glimpse
```

## 使い方

```typescript
import { readFileSync } from "fs";
import { convertPptxToSvg, convertPptxToPng } from "pptx-glimpse";

const pptx = readFileSync("presentation.pptx");

// SVG に変換
const svgResults = await convertPptxToSvg(pptx);
// [{ slideNumber: 1, svg: "<svg>...</svg>" }, ...]

// PNG に変換
const pngResults = await convertPptxToPng(pptx);
// [{ slideNumber: 1, png: Buffer, width: 960, height: 540 }, ...]
```

## テストレンダリング

PPTX ファイルを指定して、SVG と PNG の変換結果を確認できます。

```bash
npm run render -- <pptx-file> [output-dir]
```

- `output-dir` を省略すると `./output` に出力されます

```bash
# 例
npm run render -- presentation.pptx
npm run render -- presentation.pptx ./my-output
```

## 開発

```bash
npm run build          # ビルド (CJS + ESM + .d.ts)
npm run test           # テスト実行
npm run lint           # ESLint チェック
npm run format         # Prettier 整形
npm run typecheck      # 型チェック
```

## ライセンス

MIT
