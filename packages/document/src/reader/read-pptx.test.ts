import { readFileSync } from "node:fs";
import { fileURLToPath } from "node:url";

import { zipSync } from "fflate";
import { describe, expect, it } from "vitest";

// 実際の公開面 (`@pptx-glimpse/document/experimental`) 経由で import する。
import { readPptx } from "../experimental.js";

const encoder = new TextEncoder();

function xml(content: string): Uint8Array {
  return encoder.encode(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n${content}`);
}

/**
 * 受け入れ条件を精密に検証するための合成 PPTX。slide 2 枚 (順序が rId と
 * 逆になるよう意図的に並べる)、media 1 つ、未対応 part (docProps/custom.xml)、
 * external relationship、`../` を含む相対 target を含む。
 */
function buildSyntheticPptx(): Uint8Array {
  const files: Record<string, Uint8Array> = {
    "[Content_Types].xml": xml(
      `<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">` +
        `<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>` +
        `<Default Extension="xml" ContentType="application/xml"/>` +
        `<Default Extension="png" ContentType="image/png"/>` +
        `<Override PartName="/ppt/presentation.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml"/>` +
        `<Override PartName="/ppt/slides/slide1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/>` +
        `<Override PartName="/ppt/slides/slide2.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/>` +
        `</Types>`,
    ),
    "_rels/.rels": xml(
      `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">` +
        `<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="ppt/presentation.xml"/>` +
        `</Relationships>`,
    ),
    "ppt/presentation.xml": xml(
      `<p:presentation xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">` +
        `<p:sldIdLst>` +
        // sldIdLst の並び順 (slide1 → slide2) が slide order の真実。
        `<p:sldId id="256" r:id="rIdSlide1"/>` +
        `<p:sldId id="257" r:id="rIdSlide2"/>` +
        `</p:sldIdLst>` +
        `<p:sldSz cx="9144000" cy="5143500"/>` +
        `</p:presentation>`,
    ),
    "ppt/_rels/presentation.xml.rels": xml(
      `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">` +
        `<Relationship Id="rIdSlide1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" Target="slides/slide1.xml"/>` +
        `<Relationship Id="rIdSlide2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" Target="slides/slide2.xml"/>` +
        `</Relationships>`,
    ),
    "ppt/slides/slide1.xml": xml(`<p:sld xmlns:p="x"><p:cSld/></p:sld>`),
    "ppt/slides/slide2.xml": xml(`<p:sld xmlns:p="x"><p:cSld/></p:sld>`),
    "ppt/slides/_rels/slide1.xml.rels": xml(
      `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">` +
        // `../` を含む相対 target と external relationship。
        `<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="../media/image1.png"/>` +
        `<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink" Target="https://example.com/" TargetMode="External"/>` +
        `</Relationships>`,
    ),
    "ppt/media/image1.png": new Uint8Array([0x89, 0x50, 0x4e, 0x47, 1, 2, 3]),
    "docProps/custom.xml": xml(`<Properties><custom/></Properties>`),
  };
  return zipSync(files);
}

describe("readPptx — package graph と presentation metadata", () => {
  const source = readPptx(buildSyntheticPptx());

  it("presentation metadata を含む CleanDoc source を返す", () => {
    expect(source.presentation.partPath).toBe("ppt/presentation.xml");
    expect(source.presentation.handle?.partPath).toBe("ppt/presentation.xml");
    // slide は presentation order どおり typed に読まれる (cSld が空なので shapes は空)。
    expect(source.slides.map((slide) => slide.partPath)).toEqual([
      "ppt/slides/slide1.xml",
      "ppt/slides/slide2.xml",
    ]);
    expect(source.slides.every((slide) => slide.shapes.length === 0)).toBe(true);
    // この合成 fixture は slideLayout 関係を持たないため chain は辿れない。
    expect(source.slideLayouts).toEqual([]);
    expect(source.slideMasters).toEqual([]);
    expect(source.themes).toEqual([]);
    expect(source.diagnostics).toContainEqual(
      expect.objectContaining({ code: "slide-layout-unresolved" }),
    );
  });

  it("slide count / slide order / slide size を取得できる", () => {
    expect(source.presentation.slidePartPaths).toEqual([
      "ppt/slides/slide1.xml",
      "ppt/slides/slide2.xml",
    ]);
    expect(source.presentation.slideSize).toEqual({ width: 9144000, height: 5143500 });
  });

  it("relationship IDs / targets / target modes を保持できる", () => {
    const slideRels = source.packageGraph.relationships.find(
      (rel) => rel.sourcePartPath === "ppt/slides/slide1.xml",
    );
    expect(slideRels?.relationships).toEqual([
      {
        id: "rId1",
        type: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image",
        target: "../media/image1.png",
      },
      {
        id: "rId2",
        type: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink",
        target: "https://example.com/",
        targetMode: "External",
      },
    ]);

    // package root の rels は sourcePartPath を "" として保持する。
    const rootRels = source.packageGraph.relationships.find((rel) => rel.sourcePartPath === "");
    expect(rootRels?.relationships[0]?.id).toBe("rId1");
  });

  it("content type defaults / overrides を保持できる", () => {
    expect(source.packageGraph.contentTypes.defaults).toContainEqual({
      extension: "png",
      contentType: "image/png",
    });
    // override の PartName は先頭スラッシュを除去して PartPath に正規化する。
    expect(source.packageGraph.contentTypes.overrides).toContainEqual({
      partName: "ppt/slides/slide1.xml",
      contentType: "application/vnd.openxmlformats-officedocument.presentationml.slide+xml",
    });
  });

  it("media bytes と part paths を保持できる", () => {
    expect(source.packageGraph.media).toEqual([
      {
        partPath: "ppt/media/image1.png",
        contentType: "image/png",
        bytes: new Uint8Array([0x89, 0x50, 0x4e, 0x47, 1, 2, 3]),
      },
    ]);
  });

  it("unsupported package parts を raw material として保持できる", () => {
    const rawPaths = source.packageGraph.rawParts?.map((part) => part.partPath) ?? [];
    // 未対応 part / typed 化しない slide / presentation は raw として保持する。
    expect(rawPaths).toContain("docProps/custom.xml");
    expect(rawPaths).toContain("ppt/slides/slide1.xml");
    expect(rawPaths).toContain("ppt/presentation.xml");
    // content types / rels / media は structural data / media として別管理し raw に含めない。
    expect(rawPaths).not.toContain("[Content_Types].xml");
    expect(rawPaths).not.toContain("ppt/media/image1.png");
    expect(rawPaths.some((path) => path.endsWith(".rels"))).toBe(false);

    const customPart = source.packageGraph.rawParts?.find(
      (part) => part.partPath === "docProps/custom.xml",
    );
    expect(customPart?.kind).toBe("binary");
    expect(customPart?.contentType).toBe("application/xml");
  });

  it("part manifest が全 part を content type 付きで列挙する", () => {
    const partMap = new Map(
      source.packageGraph.parts.map((part) => [part.partPath, part.contentType]),
    );
    expect(partMap.get("ppt/media/image1.png")).toBe("image/png");
    expect(partMap.get("ppt/slides/slide1.xml")).toBe(
      "application/vnd.openxmlformats-officedocument.presentationml.slide+xml",
    );
    // rels part は Default extension で content type を解決する。
    expect(partMap.get("ppt/_rels/presentation.xml.rels")).toBe(
      "application/vnd.openxmlformats-package.relationships+xml",
    );
    // [Content_Types].xml 自体は part manifest に含めない。
    expect(partMap.has("[Content_Types].xml")).toBe(false);
  });

  it("presentation part が無い場合はエラーを投げる", () => {
    const bogus = zipSync({
      "[Content_Types].xml": xml(
        `<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"/>`,
      ),
    });
    expect(() => readPptx(bogus)).toThrow(/presentation part not found/);
  });

  it("officeDocument relationship が presentation 以外を指す場合はエラーを投げる", () => {
    const bogus = zipSync({
      "[Content_Types].xml": xml(
        `<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">` +
          `<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>` +
          `<Default Extension="xml" ContentType="application/xml"/>` +
          `</Types>`,
      ),
      "_rels/.rels": xml(
        `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">` +
          // officeDocument が誤って別 XML part を指している。
          `<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="docProps/app.xml"/>` +
          `</Relationships>`,
      ),
      "docProps/app.xml": xml(`<Properties/>`),
    });
    expect(() => readPptx(bogus)).toThrow(/not a presentation part/);
  });

  it("p:sldId が slide 以外の relationship を指す場合は除外し diagnostic を残す", () => {
    const bogus = zipSync({
      "[Content_Types].xml": xml(
        `<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">` +
          `<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>` +
          `<Default Extension="xml" ContentType="application/xml"/>` +
          `</Types>`,
      ),
      "_rels/.rels": xml(
        `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">` +
          `<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="ppt/presentation.xml"/>` +
          `</Relationships>`,
      ),
      "ppt/presentation.xml": xml(
        `<p:presentation xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">` +
          `<p:sldIdLst><p:sldId id="256" r:id="rIdBogus"/></p:sldIdLst>` +
          `</p:presentation>`,
      ),
      "ppt/_rels/presentation.xml.rels": xml(
        `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">` +
          // slide ではなく notesMaster を指す relationship。
          `<Relationship Id="rIdBogus" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/notesMaster" Target="notesMasters/notesMaster1.xml"/>` +
          `</Relationships>`,
      ),
    });
    const result = readPptx(bogus);
    expect(result.presentation.slidePartPaths).toEqual([]);
    expect(result.diagnostics).toContainEqual(
      expect.objectContaining({ code: "slide-relationship-invalid" }),
    );
  });
});

describe("readPptx — real fixture smoke test", () => {
  const fixturePath = fileURLToPath(
    new URL("../../../../shared-fixtures/real-basic-theme.pptx", import.meta.url),
  );
  const source = readPptx(readFileSync(fixturePath));

  it("real-basic-theme から slide order / size / media を読める", () => {
    expect(source.presentation.slidePartPaths).toEqual([
      "ppt/slides/slide1.xml",
      "ppt/slides/slide2.xml",
    ]);
    expect(source.presentation.slideSize?.width).toBeGreaterThan(0);
    expect(source.presentation.slideSize?.height).toBeGreaterThan(0);

    const image = source.packageGraph.media.find(
      (part) => part.partPath === "ppt/media/image1.png",
    );
    expect(image?.contentType).toBe("image/png");
    expect(image && image.bytes.length).toBeGreaterThan(0);

    // 全 slide part が raw として round-trip 用に保持されている。
    const rawPaths = source.packageGraph.rawParts?.map((part) => part.partPath) ?? [];
    expect(rawPaths).toContain("ppt/slides/slide1.xml");
    expect(rawPaths).toContain("ppt/slides/slide2.xml");
  });
});
