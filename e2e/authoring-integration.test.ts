import { readFileSync } from "node:fs";
import { posix } from "node:path";
import { fileURLToPath } from "node:url";

import {
  createComputedView,
  type PackageGraph,
  type PptxSourceModel,
  readPptx,
  writePptx,
} from "@pptx-glimpse/document";
import JSZip from "jszip";
import { describe, expect, it } from "vitest";

import { convertPptxToSvg } from "../packages/core/src/index.js";
import { createAuthoringIntegrationFixture } from "./fixtures/authoring-integration.js";

const fixturePath = fileURLToPath(
  new URL("../shared-fixtures/authoring-integration.pptx", import.meta.url),
);

describe("from-scratch authoring integration fixture", () => {
  it.each([
    ["freshly generated", () => createAuthoringIntegrationFixture()],
    ["committed baseline", () => new Uint8Array(readFileSync(fixturePath))],
  ] as const)("validates the %s package and source-model contract", async (_, loadFixture) => {
    const bytes = loadFixture();
    const archive = await JSZip.loadAsync(bytes);
    const entryPaths = new Set(Object.keys(archive.files));
    const source = readPptx(bytes);

    expect(source.diagnostics.filter((diagnostic) => diagnostic.severity === "error")).toEqual([]);
    expect([...entryPaths]).toEqual(expect.arrayContaining(requiredPackagePaths()));
    expectPackageGraphToBeConsistent(source.packageGraph, entryPaths);
    expectSourceContract(source);

    const reread = readPptx(writePptx(source));
    expect(reread.diagnostics.filter((diagnostic) => diagnostic.severity === "error")).toEqual([]);
    expectSourceContract(reread);
  });

  it("renders every authored element through core's document path", async () => {
    const bytes = createAuthoringIntegrationFixture();
    const source = readPptx(bytes);
    const computed = createComputedView(source);
    const report = await convertPptxToSvg(bytes, {
      textOutput: "text",
      skipSystemFonts: true,
    });

    expect(computed.slides).toHaveLength(1);
    expect(computed.slides[0]?.elements.map((element) => element.kind)).toEqual([
      "shape",
      "shape",
      "shape",
      "shape",
      "image",
      "connector",
      "table",
      "chart",
    ]);
    expect(report.slides).toHaveLength(1);
    expect(report.diagnostics.filter((diagnostic) => diagnostic.severity === "error")).toEqual([]);
    expect(report.supportCoverage.overall).toMatchObject({
      inputElements: 8,
      outputElements: 8,
      skippedElements: 0,
      unresolvedElements: 0,
    });
    expect(report.slides[0]?.svg).toContain("MASTER CONTRACT");
    expect(report.slides[0]?.svg).toContain("LAYOUT CONTRACT");
    expect(report.slides[0]?.svg).toContain("Authoring integration fixture");
    expect(report.slides[0]?.svg).toContain("Shape contract");
    expect(report.slides[0]?.svg).toContain(">Table</tspan>");
    expect(report.slides[0]?.svg).toContain(">contract</tspan>");
    expect(report.slides[0]?.svg).toContain("Chart contract");
    expect(report.slides[0]?.svg).toContain("data:image/png;base64,");
  });
});

function expectSourceContract(source: PptxSourceModel): void {
  expect(source.presentation.slidePartPaths).toEqual(["ppt/slides/slide1.xml"]);
  expect(source.slides).toHaveLength(1);
  expect(source.slideLayouts).toHaveLength(1);
  expect(source.slideMasters).toHaveLength(1);
  expect(source.themes).toHaveLength(1);
  expect(source.slideMasters[0]).toMatchObject({ name: "Authoring Contract Master" });
  expect(source.slideLayouts[0]).toMatchObject({ name: "Authoring Contract Layout" });
  expect(source.slideMasters[0]?.background).toMatchObject({ kind: "fill" });
  expect(source.slides[0]?.background).toMatchObject({ kind: "fill" });
  expect(source.slides[0]?.shapes.map((shape) => shape.kind)).toEqual([
    "shape",
    "shape",
    "image",
    "connector",
    "table",
    "chart",
  ]);
  expect(source.slideMasters[0]?.shapes.map((shape) => shape.name)).toContain(
    "Master contract text",
  );
  expect(source.slideLayouts[0]?.shapes.map((shape) => shape.name)).toContain(
    "Layout contract text",
  );

  for (const shapeOwner of [source.slides[0], source.slideLayouts[0], source.slideMasters[0]]) {
    const ids =
      shapeOwner?.shapes.map((shape) => shape.nodeId).filter((id) => id !== undefined) ?? [];
    expect(new Set(ids).size).toBe(ids.length);
  }
}

function expectPackageGraphToBeConsistent(
  graph: PackageGraph,
  entryPaths: ReadonlySet<string>,
): void {
  const partPaths = new Set(graph.parts.map((part) => part.partPath));
  for (const part of graph.parts) {
    expect(entryPaths.has(part.partPath), `missing ZIP entry ${part.partPath}`).toBe(true);
    expect(contentTypeForPart(graph, part.partPath)).toBe(part.contentType);
  }

  for (const group of graph.relationships) {
    const relationshipIds = group.relationships.map((relationship) => relationship.id);
    expect(
      new Set(relationshipIds).size,
      `duplicate relationship ID in ${group.sourcePartPath}`,
    ).toBe(relationshipIds.length);
    for (const relationship of group.relationships) {
      if (relationship.targetMode === "External") continue;
      const target = posix.normalize(
        posix.join(posix.dirname(group.sourcePartPath), relationship.target),
      );
      expect(partPaths.has(target), `${group.sourcePartPath} -> ${target}`).toBe(true);
    }
  }
}

function contentTypeForPart(graph: PackageGraph, partPath: string): string | undefined {
  const override = graph.contentTypes.overrides.find((item) => item.partName === partPath);
  if (override !== undefined) return override.contentType;
  const extension = posix.basename(partPath).split(".").at(-1) ?? "";
  return graph.contentTypes.defaults.find((item) => item.extension === extension)?.contentType;
}

function requiredPackagePaths(): string[] {
  return [
    "[Content_Types].xml",
    "_rels/.rels",
    "ppt/presentation.xml",
    "ppt/_rels/presentation.xml.rels",
    "ppt/slides/slide1.xml",
    "ppt/slides/_rels/slide1.xml.rels",
    "ppt/slideLayouts/slideLayout1.xml",
    "ppt/slideLayouts/_rels/slideLayout1.xml.rels",
    "ppt/slideMasters/slideMaster1.xml",
    "ppt/slideMasters/_rels/slideMaster1.xml.rels",
    "ppt/theme/theme1.xml",
    "ppt/media/image1.png",
    "ppt/charts/chart1.xml",
    "ppt/charts/_rels/chart1.xml.rels",
    "ppt/embeddings/Microsoft_Excel_Worksheet1.xlsx",
  ];
}
