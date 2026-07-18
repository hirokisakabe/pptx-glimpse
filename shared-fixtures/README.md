# shared-fixtures

実アプリケーション（PowerPoint / Google Slides / md-pptx など）が生成した実 PPTX ファイルと、公開 authoring API の固定 integration baseline を格納するディレクトリ。

## 用途

このディレクトリのファイルは複数のテストスイートから共有される:

| テスト               | ファイル                          | 目的                                              |
| -------------------- | --------------------------------- | ------------------------------------------------- |
| E2E スモークテスト   | `e2e/smoke.test.ts`               | エラーなく変換できること・主要要素の SVG 出力確認 |
| スナップショット VRT | `vrt/snapshot/regression.test.ts` | レンダリング結果のリグレッション検出              |

## プログラム合成フィクスチャとの違い

`vrt/snapshot/fixtures/` にあるフィクスチャは `create-fixtures.ts` でプログラム的に生成される。
このディレクトリのファイルは **手動で作成・管理される実 PPTX** であり、プログラム合成では再現しにくい以下の構造を含む:

- テーマフォント参照 (`+mj-lt`, `+mn-lt`)
- `presentation.xml` の `defaultTextStyle`
- スライドマスターの `txStyles`
- スタイル参照 (`sp.style` の `lnRef` / `fillRef` / `effectRef`)
- PowerPoint バージョン固有の XML 構造

## ファイル一覧

| ファイル                     | 作成元        | スライド数 | 主な内容                                                   |
| ---------------------------- | ------------- | ---------- | ---------------------------------------------------------- |
| `real-basic-theme.pptx`      | Google Slides | 2          | タイトル・コンテンツ・テーブル・画像・テーマフォント参照   |
| `real-product-page.pptx`     | 手作成        | 1          | 角丸矩形・楕円・テキストボックス                           |
| `real-financial-report.pptx` | 手作成        | 4          | チャート（棒グラフ・円グラフ等）・テキスト                 |
| `sample.pptx`                | md-pptx 生成  | 6          | 日本語テキスト・箇条書き・テキスト装飾                     |
| `sample-issue-387.pptx`      | 手作成        | 1          | インラインテキスト装飾（太字・斜体・太字斜体）             |
| `authoring-integration.pptx` | document API  | 1          | from-scratch authoring API の package/render 統合 contract |

## Authoring integration fixture の再生成

`authoring-integration.pptx` は `@pptx-glimpse/document` の package root から公開された
`createPptx` と authoring API だけで生成する。次のコマンドで再生成できる。

```bash
npm run fixture:authoring
npm run test -- e2e/authoring-integration.test.ts
```

baseline は、公開 authoring contract または生成される汎用 PPTX / OOXML package structure を
意図的に変更した場合だけ更新する。ZIP metadata や byte-for-byte equality は baseline contract に
含めない。更新時は integration test で package relationship、content type、part path、ID、reader / writer
round-trip、core document-path renderability が引き続き成立することを確認する。

## ファイルを追加する場合

1. `shared-fixtures/` に PPTX ファイルを配置する
2. `e2e/smoke.test.ts` にスモークテストを追加する
3. `vrt/snapshot/vrt-cases.ts` の `SHARED_FIXTURE_CASES` にエントリを追加する
4. `npm run vrt:snapshot:update` でスナップショットを生成する

ファイルサイズはリポジトリサイズへの影響を抑えるため **1ファイルあたり 500KB 以下** を目安とする。
ライセンス上問題のない PPTX のみ追加すること（自作または OSSテンプレート）。
