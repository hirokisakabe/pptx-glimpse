import { mkdirSync, writeFileSync } from "node:fs";
import { tmpdir } from "node:os";
import { join } from "node:path";

import { describe, expect, it } from "vitest";

import { collectFontFilePaths } from "./system-font-loader.js";

function makeTempFontDir(): { dir: string; fontPath: string } {
  const dir = join(tmpdir(), `pptx-glimpse-loader-test-${Date.now()}-${Math.random()}`);
  mkdirSync(dir, { recursive: true });
  const fontPath = join(dir, "bundled.ttf");
  writeFileSync(fontPath, "");
  return { dir, fontPath };
}

describe("collectFontFilePaths skipSystemFonts", () => {
  it("skipSystemFonts=true のとき additionalDirs のファイルのみ返す", () => {
    const { dir, fontPath } = makeTempFontDir();
    const result = collectFontFilePaths([dir], true);
    expect(result).toHaveLength(1);
    expect(result[0]).toBe(fontPath);
  });

  it("skipSystemFonts=true かつ fontDirs 未指定のとき空配列を返す", () => {
    // ユニークキーにするため空配列でも別インスタンス的に扱われる
    // (キャッシュは dirsKey+skipSystemFonts で管理されるため問題なし)
    const result = collectFontFilePaths([], true);
    expect(result).toEqual([]);
  });

  it("skipSystemFonts=false のとき additionalDirs のファイルが含まれる", () => {
    const { dir, fontPath } = makeTempFontDir();
    const result = collectFontFilePaths([dir], false);
    expect(result).toContain(fontPath);
  });

  it("skipSystemFonts=false は skipSystemFonts=true と同数以上のファイルを返す", () => {
    const { dir } = makeTempFontDir();
    const skipSystem = collectFontFilePaths([dir], true);
    // 別の dir を使わないと skipSystem のキャッシュが当たってしまうため
    // 同じ dir で false を呼ぶ (dirsKey は同じだが skipSystemFonts が異なるためキャッシュミス)
    const withSystem = collectFontFilePaths([dir], false);
    expect(withSystem.length).toBeGreaterThanOrEqual(skipSystem.length);
  });
});
