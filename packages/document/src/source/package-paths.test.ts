import { describe, expect, it } from "vitest";

import { asPartPath } from "./handles.js";
import {
  isRelationshipPart,
  parseRelationshipTargetMode,
  relationshipsPartPath,
  relationshipsSourcePartPath,
  resolveInternalRelationshipTarget,
  resolveRelationshipTarget,
} from "./package-paths.js";

describe("package path helpers", () => {
  it("relationship part path と owner part path を相互に解決する", () => {
    expect(relationshipsPartPath(asPartPath(""))).toBe("_rels/.rels");
    expect(relationshipsPartPath(asPartPath("ppt/presentation.xml"))).toBe(
      "ppt/_rels/presentation.xml.rels",
    );
    expect(relationshipsPartPath(asPartPath("ppt/slides/slide1.xml"))).toBe(
      "ppt/slides/_rels/slide1.xml.rels",
    );

    expect(relationshipsSourcePartPath("_rels/.rels")).toBe("");
    expect(relationshipsSourcePartPath("ppt/_rels/presentation.xml.rels")).toBe(
      "ppt/presentation.xml",
    );
    expect(relationshipsSourcePartPath("ppt/slides/_rels/slide1.xml.rels")).toBe(
      "ppt/slides/slide1.xml",
    );
  });

  it("relationship part を判定する", () => {
    expect(isRelationshipPart("_rels/.rels")).toBe(true);
    expect(isRelationshipPart("ppt/slides/_rels/slide1.xml.rels")).toBe(true);
    expect(isRelationshipPart("ppt/slides/slide1.xml")).toBe(false);
  });

  it("relationship target を source part 基準で正規化する", () => {
    expect(resolveRelationshipTarget("ppt/slides/slide1.xml", "../media/image1.png")).toBe(
      "ppt/media/image1.png",
    );
    expect(resolveRelationshipTarget("ppt/presentation.xml", "/ppt/slides/slide1.xml")).toBe(
      "ppt/slides/slide1.xml",
    );
    expect(resolveRelationshipTarget("", "ppt/presentation.xml")).toBe("ppt/presentation.xml");
    expect(resolveRelationshipTarget("ppt/slides/slide1.xml", "https://example.com/a")).toBe(
      "https://example.com/a",
    );
  });

  it("internal relationship target と target mode を parse する", () => {
    expect(
      resolveInternalRelationshipTarget(asPartPath("ppt/slides/slide1.xml"), {
        id: "rId1",
        type: "image",
        target: "../media/image1.png",
      }),
    ).toBe("ppt/media/image1.png");
    expect(
      resolveInternalRelationshipTarget(asPartPath("ppt/slides/slide1.xml"), {
        id: "rId2",
        type: "hyperlink",
        target: "https://example.com/",
        targetMode: "External",
      }),
    ).toBeUndefined();

    expect(parseRelationshipTargetMode("External")).toBe("External");
    expect(parseRelationshipTargetMode("Internal")).toBe("Internal");
    expect(parseRelationshipTargetMode("Unexpected")).toBeUndefined();
    expect(parseRelationshipTargetMode(undefined)).toBeUndefined();
  });
});
