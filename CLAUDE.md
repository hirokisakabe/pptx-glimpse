# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## プロジェクト概要

pptx-glimpse は PPTX スライドを SVG / PNG に変換する TypeScript ライブラリ。
入力は `Buffer | Uint8Array`、出力は SVG 文字列または PNG Buffer。

## コマンド

```bash
npm run build          # tsup でビルド (CJS + ESM + .d.ts)
npm run test           # vitest で全テスト実行
npm run test -- tests/utils/emu.test.ts  # 単一テストファイル実行
npm run test:watch     # テストのウォッチモード
npm run lint           # ESLint チェック
npm run lint:fix       # ESLint 自動修正
npm run format         # Prettier 整形
npm run format:check   # Prettier チェック
npm run typecheck      # tsc --noEmit で型チェック
npm run render         # tsx scripts/test-render.ts でテストレンダリング
```

CI は `lint` → `format:check` → `typecheck` → `test` → `build` の順で実行される。

## アーキテクチャ

データフロー: **PPTX バイナリ → Parser (ZIP解凍+XMLパース) → 中間モデル → Renderer (SVG生成) → PNG変換 (optional)**

- `src/parser/` — PPTX の ZIP 解凍 (`jszip`) と XML パース (`fast-xml-parser`) で中間モデルを構築
- `src/model/` — 中間モデルの TypeScript インターフェース (Slide, Shape, Fill, Text, Theme 等)
- `src/renderer/` — 中間モデルから SVG 文字列を生成。`geometry/` にプリセット図形定義
- `src/color/` — テーマカラー解決 (schemeClr → colorMap → colorScheme) と色変換 (lumMod/tint/shade)
- `src/png/` — sharp による SVG → PNG 変換
- `src/utils/` — EMU↔ピクセル変換 (1 inch = 914400 EMU, 96 DPI)

エントリポイント: `src/index.ts` が `convertPptxToSvg` と `convertPptxToPng` をエクスポート。

## 技術制約

- **SVG はインライン属性のみ使用** — CSS クラスは使わない。sharp (librsvg) が CSS を正しく解釈しないため
- **fast-xml-parser の `isArray` 設定が必須** — `sp`, `pic`, `p`, `r` 等のタグは単一要素でも配列として返す必要がある (`xml-parser.ts` の `ARRAY_TAGS`)
- **EMU 単位** — PPTX 内部座標は EMU (English Metric Units)。`emuToPixels()` で変換。16:9 スライドは 9144000×5143500 EMU = 960×540 px
- **背景フォールバック** — スライド → スライドレイアウト → スライドマスターの順で背景を探索

## コーディング規約

- Prettier: ダブルクォート、セミコロンあり、末尾カンマ、printWidth 100
- ESLint: `_` プレフィックスの未使用変数は許可
- ESM (`"type": "module"`) — インポートには `.js` 拡張子が必要
