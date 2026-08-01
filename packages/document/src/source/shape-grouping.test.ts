import { unzipSync, zipSync } from "fflate";
import { describe, expect, it } from "vitest";

import { createPptx } from "../builder/create-pptx.js";
import {
  addConnector,
  addShape,
  asEmu,
  asOoxmlAngle,
  asSourceNodeId,
  groupShapes,
  type PptxSourceModel,
  type SourceConnector,
  type SourceGroup,
  ungroupShape,
} from "../index.js";
import { readPptx } from "../reader/read-pptx.js";
import { writePptx } from "../writer/write-pptx.js";

describe("groupShapes and ungroupShape", () => {
  it("replaces one z-order range, preserves child handles, and round-trips topology", () => {
    const existing = existingShapes(3);
    const slide = requireValue(existing.slides[0]);
    const childHandles = slide.shapes.slice(0, 2).map((shape) => requireValue(shape.handle));

    const grouped = groupShapes(existing, childHandles);
    const group = requireGroup(grouped.slides[0]?.shapes[0]);
    expect(group.children.map((child) => child.nodeId)).toEqual(
      childHandles.map((handle) => handle.nodeId),
    );
    expect(group.children.map((child) => child.handle)).toEqual(childHandles);
    expect(grouped.slides[0]?.shapes.map((shape) => shape.kind)).toEqual(["group", "shape"]);
    expect(group.transform).toEqual(group.childTransform);

    const reread = readPptx(writePptx(grouped));
    const rereadGroup = requireGroup(reread.slides[0]?.shapes[0]);
    expect(rereadGroup.nodeId).toBe(group.nodeId);
    expect(rereadGroup.transform).toEqual(group.transform);
    expect(rereadGroup.childTransform).toEqual(group.childTransform);
    expect(rereadGroup.children.map((child) => child.nodeId)).toEqual(
      childHandles.map((handle) => handle.nodeId),
    );

    const ungrouped = ungroupShape(reread, requireValue(rereadGroup.handle));
    expect(ungrouped.slides[0]?.shapes.map((shape) => shape.nodeId)).toEqual(
      slide.shapes.map((shape) => shape.nodeId),
    );
    expect(ungrouped.slides[0]?.shapes.slice(0, 2).map((shape) => shape.handle)).toEqual(
      childHandles,
    );
    const writtenUngrouped = readPptx(writePptx(ungrouped));
    expect(writtenUngrouped.slides[0]?.shapes.map((shape) => shape.nodeId)).toEqual(
      slide.shapes.map((shape) => shape.nodeId),
    );
  });

  it("supports nested sibling topology edits through the recursive writer locator", () => {
    const existing = existingShapes(3);
    const handles = requireValue(existing.slides[0]).shapes.map((shape) =>
      requireValue(shape.handle),
    );
    const withInner = groupShapes(existing, handles.slice(0, 2));
    const inner = requireGroup(withInner.slides[0]?.shapes[0]);
    const withOuter = groupShapes(withInner, [requireValue(inner.handle), handles[2]]);
    const outer = requireGroup(withOuter.slides[0]?.shapes[0]);
    const expandedInner = ungroupShape(withOuter, requireValue(inner.handle));

    const expandedOuter = requireGroup(expandedInner.slides[0]?.shapes[0]);
    expect(expandedOuter.nodeId).toBe(outer.nodeId);
    expect(expandedOuter.children.map((child) => child.nodeId)).toEqual(
      handles.map((handle) => handle.nodeId),
    );
    const rereadOuter = requireGroup(readPptx(writePptx(expandedInner)).slides[0]?.shapes[0]);
    expect(rereadOuter.children.map((child) => child.nodeId)).toEqual(
      handles.map((handle) => handle.nodeId),
    );
  });

  it("rejects a forged group edit whose finalized shell already has children", () => {
    const source = existingShapes(2);
    const handles = requireValue(source.slides[0]).shapes.map((shape) =>
      requireValue(shape.handle),
    );
    const grouped = groupShapes(source, handles);
    const forged: PptxSourceModel = {
      ...grouped,
      edits: grouped.edits?.map((edit) =>
        edit.kind === "groupShapes"
          ? { ...edit, xml: edit.xml.replace("</p:grpSp>", "<p:grpSp/></p:grpSp>") }
          : edit,
      ),
    };

    expect(() => writePptx(forged)).toThrow("grouped XML must be an empty group shell");
  });

  it("rejects nested edits when the immediate parent group has no node id", () => {
    const existing = existingShapes(3);
    const handles = requireValue(existing.slides[0]).shapes.map((shape) =>
      requireValue(shape.handle),
    );
    const withInner = groupShapes(existing, handles.slice(0, 2));
    const inner = requireGroup(withInner.slides[0]?.shapes[0]);
    const idlessInnerParent = replaceFirstGroup(withInner, withoutGroupId(inner));
    const idlessInnerParentBefore = structuredClone(idlessInnerParent);

    expect(() => groupShapes(idlessInnerParent, handles.slice(0, 2))).toThrow(
      "the immediate parent group requires a node id",
    );
    expect(idlessInnerParent).toEqual(idlessInnerParentBefore);

    const withOuter = groupShapes(withInner, [requireValue(inner.handle), handles[2]]);
    const outer = requireGroup(withOuter.slides[0]?.shapes[0]);
    const idlessOuterParent = replaceFirstGroup(withOuter, withoutGroupId(outer));
    const idlessOuterParentBefore = structuredClone(idlessOuterParent);

    expect(() => ungroupShape(idlessOuterParent, requireValue(inner.handle))).toThrow(
      "the immediate parent group requires a node id",
    );
    expect(idlessOuterParent).toEqual(idlessOuterParentBefore);
  });

  it("keeps internal connector endpoint ids through group and reread", () => {
    let source = createPptx();
    const slideHandle = requireValue(source.slides[0]?.handle);
    source = addShape(source, slideHandle, shapeInput(0));
    source = addShape(source, slideHandle, shapeInput(2000));
    const shapes = requireValue(source.slides[0]).shapes;
    source = addConnector(source, slideHandle, {
      preset: "straightConnector1",
      offsetX: asEmu(1000),
      offsetY: asEmu(500),
      width: asEmu(1000),
      height: asEmu(1),
      start: { shapeHandle: requireValue(shapes[0]?.handle), connectionSiteIndex: 3 },
      end: { shapeHandle: requireValue(shapes[1]?.handle), connectionSiteIndex: 1 },
    });
    const existing = readPptx(writePptx(source));
    const existingShapes = requireValue(existing.slides[0]).shapes;
    const endpointIds = existingShapes.slice(0, 2).map((shape) => shape.nodeId);

    const grouped = groupShapes(
      existing,
      existingShapes.map((shape) => requireValue(shape.handle)),
    );
    const rereadGroup = requireGroup(readPptx(writePptx(grouped)).slides[0]?.shapes[0]);
    const connector = rereadGroup.children.find((shape) => shape.kind === "connector");
    expect(connector?.kind).toBe("connector");
    if (connector?.kind !== "connector") throw new Error("connector was not reread");
    expect(connector.connection?.start?.shapeId).toBe(endpointIds[0]);
    expect(connector.connection?.end?.shapeId).toBe(endpointIds[1]);

    const ungrouped = ungroupShape(
      grouped,
      requireValue(requireGroup(grouped.slides[0]?.shapes[0]).handle),
    );
    const rereadUngrouped = readPptx(writePptx(ungrouped));
    const ungroupedConnector = rereadUngrouped.slides[0]?.shapes.find(
      (shape) => shape.kind === "connector",
    );
    expect(ungroupedConnector?.kind).toBe("connector");
    if (ungroupedConnector?.kind !== "connector") {
      throw new Error("ungrouped connector was not reread");
    }
    expect(ungroupedConnector.connection?.start?.shapeId).toBe(endpointIds[0]);
    expect(ungroupedConnector.connection?.end?.shapeId).toBe(endpointIds[1]);
  });

  it("rejects non-contiguous siblings and selection-crossing connectors atomically", () => {
    const existing = existingShapes(3);
    const existingBefore = structuredClone(existing);
    const handles = requireValue(existing.slides[0]).shapes.map((shape) =>
      requireValue(shape.handle),
    );
    expect(() => groupShapes(existing, [handles[0], handles[2]])).toThrow(
      "selected shapes must be consecutive siblings",
    );
    expect(existing.edits).toBeUndefined();
    expect(existing).toEqual(existingBefore);

    let connected = createPptx();
    const slideHandle = requireValue(connected.slides[0]?.handle);
    connected = addShape(connected, slideHandle, shapeInput(0));
    connected = addShape(connected, slideHandle, shapeInput(2000));
    const targets = requireValue(connected.slides[0]).shapes;
    connected = addConnector(connected, slideHandle, {
      preset: "straightConnector1",
      offsetX: asEmu(1000),
      offsetY: asEmu(500),
      width: asEmu(1000),
      height: asEmu(1),
      start: { shapeHandle: requireValue(targets[0]?.handle), connectionSiteIndex: 3 },
      end: { shapeHandle: requireValue(targets[1]?.handle), connectionSiteIndex: 1 },
    });
    const beforeEdits = connected.edits;
    const connectedBefore = structuredClone(connected);
    expect(() =>
      groupShapes(
        connected,
        targets.map((shape) => requireValue(shape.handle)),
      ),
    ).toThrow("connector endpoint crosses the selection boundary");
    expect(connected.edits).toBe(beforeEdits);
    expect(connected).toEqual(connectedBefore);
  });

  it("treats preserved unsupported shape-tree nodes as z-order siblings", () => {
    const archive = unzipSync(writePptx(existingShapes(2)));
    const slidePath = "ppt/slides/slide1.xml";
    const slideXml = new TextDecoder().decode(requireValue(archive[slidePath]));
    archive[slidePath] = new TextEncoder().encode(
      slideXml.replace(
        "</p:sp><p:sp>",
        '</p:sp><p:contentPart/><p14:grpSpPr xmlns:p14="urn:extension"/><p:sp>',
      ),
    );
    const source = readPptx(zipSync(archive));
    const sourceBefore = structuredClone(source);
    const handles = requireValue(source.slides[0])
      .shapes.filter((shape) => shape.nodeId !== undefined)
      .map((shape) => requireValue(shape.handle));

    expect(source.slides[0]?.shapes.map((shape) => shape.kind)).toEqual([
      "shape",
      "raw",
      "raw",
      "shape",
    ]);
    expect(source.slides[0]?.shapes[2]?.kind).toBe("raw");
    if (source.slides[0]?.shapes[2]?.kind === "raw") {
      expect(source.slides[0].shapes[2].raw.node.name).toBe("p14:grpSpPr");
    }
    expect(() => groupShapes(source, handles)).toThrow(
      "selected shapes must be consecutive siblings",
    );
    expect(source).toEqual(sourceBefore);
  });

  it("rejects an mc:AlternateContent selection before creating an edit", () => {
    const existing = existingShapes(2);
    const archive = unzipSync(writePptx(existing));
    const slidePath = "ppt/slides/slide1.xml";
    const slideXml = new TextDecoder().decode(requireValue(archive[slidePath]));
    const firstShape = requireValue(slideXml.match(/<p:sp>[\s\S]*?<\/p:sp>/)?.[0]);
    archive[slidePath] = new TextEncoder().encode(
      slideXml
        .replace(
          "<p:sld ",
          '<p:sld xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" ',
        )
        .replace(
          firstShape,
          `<mc:AlternateContent><mc:Choice Requires="p14">${firstShape}</mc:Choice></mc:AlternateContent>`,
        ),
    );
    const source = readPptx(zipSync(archive));
    const sourceBefore = structuredClone(source);
    const handles = requireValue(source.slides[0]).shapes.map((shape) =>
      requireValue(shape.handle),
    );

    expect(() => groupShapes(source, handles)).toThrow(
      "mc:AlternateContent nodes are not supported",
    );
    expect(source.edits).toBeUndefined();
    expect(source).toEqual(sourceBefore);
  });

  it("rejects referenced or non-lossless groups atomically", () => {
    const existing = existingShapes(3);
    const handles = requireValue(existing.slides[0]).shapes.map((shape) =>
      requireValue(shape.handle),
    );
    const grouped = groupShapes(existing, handles.slice(0, 2));
    const group = requireGroup(grouped.slides[0]?.shapes[0]);
    const rotated = replaceFirstGroup(grouped, {
      ...group,
      transform: { ...requireValue(group.transform), rotation: asOoxmlAngle(60000) },
    });
    const rotatedBefore = structuredClone(rotated);
    expect(() => ungroupShape(rotated, requireValue(group.handle))).toThrow(
      "not a lossless identity child mapping",
    );
    expect(rotated.edits).toBe(grouped.edits);
    expect(rotated).toEqual(rotatedBefore);

    const sibling = requireValue(grouped.slides[0]?.shapes[1]);
    const referencingConnector: SourceConnector = {
      kind: "connector",
      nodeId: sibling.nodeId,
      name: "Group reference",
      transform: sibling.kind === "raw" ? undefined : sibling.transform,
      connection: {
        start: { shapeId: requireValue(group.nodeId), connectionSiteIndex: 0 },
      },
      handle: sibling.handle,
    };
    const referenced: PptxSourceModel = {
      ...grouped,
      slides: grouped.slides.map((slide, index) =>
        index === 0 ? { ...slide, shapes: [group, referencingConnector] } : slide,
      ),
    };
    const beforeEdits = referenced.edits;
    const referencedBefore = structuredClone(referenced);
    expect(() => ungroupShape(referenced, requireValue(group.handle))).toThrow(
      "group is referenced by connector",
    );
    expect(referenced.edits).toBe(beforeEdits);
    expect(referenced).toEqual(referencedBefore);
  });

  it("rejects group non-visual metadata that cannot be losslessly discarded", () => {
    const source = existingShapes(2);
    const handles = requireValue(source.slides[0]).shapes.map((shape) =>
      requireValue(shape.handle),
    );
    const grouped = groupShapes(source, handles);
    const slidePath = "ppt/slides/slide1.xml";
    const groupId = requireValue(requireGroup(grouped.slides[0]?.shapes[0]).nodeId);
    const mutations = [
      (xml: string) =>
        xml.replaceAll(
          "<p:cNvGrpSpPr/>",
          '<p:cNvGrpSpPr><a:grpSpLocks noUngrp="1"/></p:cNvGrpSpPr>',
        ),
      (xml: string) =>
        xml.replaceAll("<p:nvGrpSpPr>", '<p:nvGrpSpPr xmlns:vendor="urn:vendor" vendor:flag="1">'),
      (xml: string) =>
        xml.replace(
          `id="${groupId}" name="Group ${groupId}"`,
          `id="${groupId}" name="Group ${groupId}" xmlns:vendor="urn:vendor" vendor:id="${groupId}"`,
        ),
    ];

    for (const mutate of mutations) {
      const archive = unzipSync(writePptx(grouped));
      const slideXml = new TextDecoder().decode(requireValue(archive[slidePath]));
      archive[slidePath] = new TextEncoder().encode(mutate(slideXml));
      const reread = readPptx(zipSync(archive));
      const group = requireGroup(reread.slides[0]?.shapes[0]);
      const sourceBefore = structuredClone(reread);

      expect(() => ungroupShape(reread, requireValue(group.handle))).toThrow(
        "group appearance or unknown XML cannot be losslessly expanded",
      );
      expect(reread).toEqual(sourceBefore);
    }
  });

  it("preserves group-local namespace declarations when children are expanded", () => {
    const source = existingShapes(2);
    const handles = requireValue(source.slides[0]).shapes.map((shape) =>
      requireValue(shape.handle),
    );
    const grouped = groupShapes(source, handles);
    const archive = unzipSync(writePptx(grouped));
    const slidePath = "ppt/slides/slide1.xml";
    const slideXml = new TextDecoder().decode(requireValue(archive[slidePath]));
    archive[slidePath] = new TextEncoder().encode(
      slideXml
        .replace("<p:grpSp xmlns:a=", '<p:grpSp xmlns:v="urn:group-local" xmlns:a=')
        .replace("</p:nvSpPr><p:spPr>", "</p:nvSpPr><v:before/><p:spPr>")
        .replace("</p:sp>", "<v:metadata/></p:sp>"),
    );
    const reread = readPptx(zipSync(archive));
    const group = requireGroup(reread.slides[0]?.shapes[0]);
    const output = writePptx(ungroupShape(reread, requireValue(group.handle)));
    const outputXml = new TextDecoder().decode(requireValue(unzipSync(output)[slidePath]));
    const expandedShapeXml = requireValue(outputXml.match(/<p:sp\b[^>]*>.*?<\/p:sp>/s)?.[0]);

    expect(expandedShapeXml).toContain('xmlns:v="urn:group-local"');
    expect(expandedShapeXml).toContain("<v:metadata/>");
    expect(expandedShapeXml.indexOf("</p:nvSpPr>")).toBeLessThan(
      expandedShapeXml.indexOf("<v:before/>"),
    );
    expect(expandedShapeXml.indexOf("<v:before/>")).toBeLessThan(
      expandedShapeXml.indexOf("<p:spPr>"),
    );
    expect(() => readPptx(output)).not.toThrow();
  });

  it("never reuses group ids after ungroup and includes descendant and pending ids", () => {
    const existing = existingShapes(3);
    const handles = requireValue(existing.slides[0]).shapes.map((shape) =>
      requireValue(shape.handle),
    );
    const firstGrouped = groupShapes(existing, handles.slice(0, 2));
    const firstGroup = requireGroup(firstGrouped.slides[0]?.shapes[0]);
    const firstChild = requireValue(firstGroup.children[0]);
    const descendantId = asSourceNodeId("99");
    const withHighDescendant = replaceFirstGroup(firstGrouped, {
      ...firstGroup,
      children: [
        {
          ...firstChild,
          nodeId: descendantId,
          handle: { ...requireValue(firstChild.handle), nodeId: descendantId },
        },
        ...firstGroup.children.slice(1),
      ],
    });
    const outerGrouped = groupShapes(withHighDescendant, [
      requireValue(firstGroup.handle),
      handles[2],
    ]);
    expect(requireGroup(outerGrouped.slides[0]?.shapes[0]).nodeId).toBe("100");

    const ungrouped = ungroupShape(firstGrouped, requireValue(firstGroup.handle));
    const regrouped = groupShapes(ungrouped, handles.slice(0, 2));
    const secondGroup = requireGroup(regrouped.slides[0]?.shapes[0]);

    expect(secondGroup.nodeId).not.toBe(firstGroup.nodeId);
    expect(Number(secondGroup.nodeId)).toBeGreaterThan(Number(firstGroup.nodeId));
  });
});

function existingShapes(count: number): PptxSourceModel {
  let source = createPptx();
  const slideHandle = requireValue(source.slides[0]?.handle);
  for (let index = 0; index < count; index += 1) {
    source = addShape(source, slideHandle, shapeInput(index * 2000));
  }
  return readPptx(writePptx(source));
}

function shapeInput(offsetX: number) {
  return {
    geometry: { kind: "preset", preset: "rect" } as const,
    offsetX: asEmu(offsetX),
    offsetY: asEmu(0),
    width: asEmu(1000),
    height: asEmu(1000),
  };
}

function replaceFirstGroup(source: PptxSourceModel, group: SourceGroup): PptxSourceModel {
  return {
    ...source,
    slides: source.slides.map((slide, index) =>
      index === 0 ? { ...slide, shapes: [group, ...slide.shapes.slice(1)] } : slide,
    ),
  };
}

function withoutGroupId(group: SourceGroup): SourceGroup {
  const { nodeId: _nodeId, ...withoutNodeId } = group;
  const { nodeId: _handleNodeId, ...handle } = requireValue(group.handle);
  void _nodeId;
  void _handleNodeId;
  return { ...withoutNodeId, handle };
}

function requireGroup(node: PptxSourceModel["slides"][number]["shapes"][number] | undefined) {
  if (node?.kind !== "group") throw new Error("test fixture group is missing");
  return node;
}

function requireValue<T>(value: T | undefined): T {
  if (value === undefined) throw new Error("test fixture value is missing");
  return value;
}
