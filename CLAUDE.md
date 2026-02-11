# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## プロジェクト概要

pptx-glimpse は PPTX スライドを SVG / PNG に変換する TypeScript ライブラリ。
入力は `Buffer | Uint8Array`、出力は SVG 文字列または PNG Buffer。

## コマンド

```bash
npm run build          # tsup でビルド (CJS + ESM + .d.ts)
npm run test           # vitest で全テスト実行
npm run test -- src/utils/emu.test.ts  # 単一テストファイル実行
npm run test:watch     # テストのウォッチモード
npm run lint           # ESLint チェック
npm run lint:fix       # ESLint 自動修正
npm run format         # Prettier 整形
npm run format:check   # Prettier チェック
npm run typecheck      # tsc --noEmit で型チェック
npm run render         # tsx scripts/test-render.ts でテストレンダリング
```

CI は 3 ジョブ構成:

- **ci**: `lint` → `format:check` → `typecheck` → `test` (VRT 除外) → `build` (Node 20/22/24)
- **vrt**: 通常 VRT (自己比較)
- **libreoffice-vrt**: LibreOffice VRT (Docker でフィクスチャ・参照画像を生成して実行)

## アーキテクチャ

データフロー: **PPTX バイナリ → Parser (ZIP解凍+XMLパース) → 中間モデル → Renderer (SVG生成) → PNG変換 (optional)**

- `src/parser/` — PPTX の ZIP 解凍 (`jszip`) と XML パース (`fast-xml-parser`) で中間モデルを構築
- `src/model/` — 中間モデルの TypeScript インターフェース (Slide, Shape, Fill, Text, Theme, Table, Chart, Image, Line, Effect, Presentation 等)
- `src/renderer/` — 中間モデルから SVG 文字列を生成。`geometry/` にプリセット図形定義、テーブル・チャート・画像の個別レンダラーも含む
- `src/color/` — テーマカラー解決 (schemeClr → colorMap → colorScheme) と色変換 (lumMod/tint/shade)
- `src/png/` — sharp による SVG → PNG 変換
- `src/data/` — フォントメトリクスデータ (OSS 互換フォントから抽出した文字幅情報)
- `src/utils/` — EMU↔ピクセル変換 (1 inch = 914400 EMU, 96 DPI)、テキスト幅計測・テキスト折り返し

エントリポイント: `src/index.ts` が `convertPptxToSvg` と `convertPptxToPng` をエクスポート。

## 技術制約

- **SVG はインライン属性のみ使用** — CSS クラスは使わない。sharp (librsvg) が CSS を正しく解釈しないため
- **fast-xml-parser の `isArray` 設定が必須** — `sp`, `pic`, `p`, `r` 等のタグは単一要素でも配列として返す必要がある (`xml-parser.ts` の `ARRAY_TAGS`)
- **EMU 単位** — PPTX 内部座標は EMU (English Metric Units)。`emuToPixels()` で変換。16:9 スライドは 9144000×5143500 EMU = 960×540 px
- **背景フォールバック** — スライド → スライドレイアウト → スライドマスターの順で背景を探索

## VRT (Visual Regression Testing)

描画結果の視覚的回帰テスト。パーサーやレンダラーを変更した場合、**必ず VRT の更新が必要か確認すること**。

### ディレクトリ構成

```
vrt/
├── compare-utils.ts                          # 共通画像比較ユーティリティ
├── internal/                                 # 通常 VRT (自己比較)
│   ├── regression.test.ts                    # テスト本体
│   ├── create-fixtures.ts                    # フィクスチャ生成スクリプト
│   ├── update-snapshots.ts                   # スナップショット更新スクリプト
│   ├── fixtures/                             # VRT 用 PPTX フィクスチャ (CI で動的生成)
│   └── snapshots/                            # 参照スナップショット画像
└── libreoffice/                              # LibreOffice VRT
    ├── regression.test.ts                    # テスト本体
    ├── create_fixtures.py                    # フィクスチャ生成 (Python, Docker)
    ├── update_snapshots.sh                   # スナップショット更新 (Docker)
    ├── fixtures/                             # CI で動的生成
    └── snapshots/                            # CI で動的生成
```

### VRT 更新手順

パーサー・レンダラー・モデルの変更で描画結果が変わる場合:

1. **フィクスチャ更新** (新機能追加や既存フィクスチャの修正が必要な場合): `vrt/internal/create-fixtures.ts` を編集し `npm run vrt:internal:fixtures` を実行
2. **スナップショット更新**: `npm run vrt:internal:update` で参照画像を再生成
3. **テスト確認**: `npm run test` で VRT テストが通ることを確認

### 同期が必要な 3 箇所

新しい描画機能を追加した場合、以下の **3 箇所すべて** を更新する必要がある:

1. **`vrt/internal/create-fixtures.ts`** — 新機能をカバーするフィクスチャ (PPTX) を追加
2. **`vrt/internal/regression.test.ts`** — `VRT_CASES` 配列に新しいテストケースを追加
3. **`vrt/internal/snapshots/`** — `npm run vrt:internal:update` でスナップショットを再生成

**よくあるミス**: パーサーやレンダラーを修正したのにスナップショットを更新し忘れて VRT テストが失敗する。描画に影響する変更を行ったら、必ず `npm run vrt:internal:update` を実行すること。

### LibreOffice VRT (Docker ベース)

python-pptx で生成した PPTX を LibreOffice でレンダリングし、pptx-glimpse の出力と比較する。Docker で環境を統一。

#### セットアップ

```bash
npm run vrt:lo:docker-build   # Docker イメージのビルド
npm run vrt:lo:update          # フィクスチャ生成 + 参照画像生成 (Docker 必須)
npm run test                   # テスト実行 (LibreOffice VRT 含む)
```

#### 許容度

- `PIXEL_THRESHOLD = 0.3` (ピクセル単位の色差許容度)
- `MISMATCH_TOLERANCE = 0.05` (5% のピクセル不一致を許容)

LibreOffice ≠ PowerPoint のため、フォントレンダリングやアンチエイリアスの差異は許容する。明らかな描画漏れ・構造的な間違いを検知する目的。

#### Docker なしの場合

Docker がない環境では LibreOffice VRT テストは自動的にスキップされる。`npm run test` は問題なく通る。

## コーディング規約

- Prettier: ダブルクォート、セミコロンあり、末尾カンマ、printWidth 100
- ESLint: `_` プレフィックスの未使用変数は許可
- ESM (`"type": "module"`) — インポートには `.js` 拡張子が必要
