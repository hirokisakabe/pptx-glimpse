import { describe, expect, it } from "vitest";

import { createPptx } from "../builder/create-pptx.js";
import { addChart } from "./chart-authoring.js";
import { nextDrawingOrderingSlot, nextDrawingShapeId } from "./drawing-authoring-allocation.js";
import { asPartPath, asRawSidecarId, asSourceNodeId } from "./handles.js";
import { addPicture } from "./picture-authoring.js";
import type { PptxSourceModel, PptxSourceModelEdit } from "./pptx-source-model.js";
import { addConnector, addShape, addSlideNumber, addTextBox } from "./shape-authoring.js";
import type { SourceGroup, SourceShape } from "./shapes.js";
import { addTable } from "./table-authoring.js";
import { asEmu } from "./units.js";

const PNG_BYTES = new Uint8Array([0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a]);

describe("drawing authoring allocation", () => {
  it("reserves root id 1 and includes group descendants and root ordering slots", () => {
    const created = createPptx();
    const partPath = created.slides[0].partPath;
    const source = withShapeTreeRootId(created, partPath, 1);
    const group: SourceGroup = {
      kind: "group",
      nodeId: asSourceNodeId("4"),
      children: [shape(partPath, "12", 0)],
      handle: { partPath, nodeId: asSourceNodeId("4"), orderingSlot: 7 },
    };

    expect(nextDrawingShapeId(source, [], partPath)).toBe("2");
    expect(nextDrawingShapeId(source, [group], partPath)).toBe("13");
    expect(nextDrawingOrderingSlot([group, shape(partPath, "5", 3)])).toBe(8);
    expect(nextDrawingOrderingSlot([])).toBe(0);
  });

  it("keeps pending add and delete ids reserved within each slide, layout, or master part", () => {
    const slidePath = asPartPath("ppt/slides/slide1.xml");
    const layoutPath = asPartPath("ppt/slideLayouts/slideLayout1.xml");
    const masterPath = asPartPath("ppt/slideMasters/slideMaster1.xml");
    const edits: readonly PptxSourceModelEdit[] = [
      { kind: "addShape", slidePartPath: slidePath, shapeId: "20", xml: "<p:sp/>" },
      {
        kind: "deleteShape",
        handle: { partPath: slidePath, nodeId: asSourceNodeId("25"), orderingSlot: 0 },
      },
      { kind: "addShape", slidePartPath: layoutPath, shapeId: "30", xml: "<p:sp/>" },
      {
        kind: "deleteShape",
        handle: { partPath: masterPath, nodeId: asSourceNodeId("40"), orderingSlot: 0 },
      },
    ];

    const source = { ...createPptx(), edits };
    expect(nextDrawingShapeId(source, [shape(slidePath, "10", 0)], slidePath)).toBe("26");
    expect(nextDrawingShapeId(source, [shape(layoutPath, "3", 0)], layoutPath)).toBe("31");
    expect(nextDrawingShapeId(source, [shape(masterPath, "4", 0)], masterPath)).toBe("41");
  });

  it("keeps finalized XML and typed nodes on the same edit-time id and ordering", () => {
    let source = createPptx();
    const slideHandle = source.slides[0].handle!;
    source = addTextBox(source, slideHandle, {
      offsetX: asEmu(0),
      offsetY: asEmu(0),
      width: asEmu(1000),
      height: asEmu(1000),
      text: "A",
    });
    source = addShape(source, slideHandle, {
      geometry: { kind: "preset", preset: "rect" },
      offsetX: asEmu(0),
      offsetY: asEmu(0),
      width: asEmu(1000),
      height: asEmu(1000),
    });
    source = addConnector(source, slideHandle, {
      preset: "straightConnector1",
      offsetX: asEmu(0),
      offsetY: asEmu(0),
      width: asEmu(1000),
      height: asEmu(1),
    });
    source = addPicture(source, slideHandle, {
      bytes: PNG_BYTES,
      offsetX: asEmu(0),
      offsetY: asEmu(0),
      width: asEmu(1000),
      height: asEmu(1000),
    });
    source = addChart(source, slideHandle, {
      chartType: "bar",
      offsetX: asEmu(0),
      offsetY: asEmu(0),
      width: asEmu(1000),
      height: asEmu(1000),
      series: [{ categories: ["A"], values: [1] }],
    });
    source = addTable(source, slideHandle, {
      offsetX: asEmu(0),
      offsetY: asEmu(0),
      width: asEmu(1000),
      height: asEmu(1000),
      columnWidths: [asEmu(1000)],
      rows: [{ height: asEmu(1000), cells: [{ text: "A" }] }],
    });

    const nodes = source.slides[0].shapes;
    const edits = (source.edits ?? []).filter((edit) =>
      ["addTextBox", "addShape", "addConnector", "addPicture", "addChart", "addTable"].includes(
        edit.kind,
      ),
    );

    expect(nodes.map((node) => node.nodeId)).toEqual(["1", "2", "3", "4", "5", "6"]);
    expect(nodes.map((node) => node.handle?.orderingSlot)).toEqual([0, 1, 2, 3, 4, 5]);
    expect(edits).toHaveLength(nodes.length);
    edits.forEach((edit, index) => {
      if (!("shapeId" in edit) || !("xml" in edit)) throw new Error("expected drawing add edit");
      expect(edit.shapeId).toBe(nodes[index].nodeId);
      expect(edit.xml).toContain(`id="${edit.shapeId}"`);
    });

    const layoutHandle = source.slideLayouts[0].handle!;
    source = addSlideNumber(source, layoutHandle, {
      offsetX: asEmu(0),
      offsetY: asEmu(0),
      width: asEmu(1000),
      height: asEmu(1000),
    });
    const slideNumber = source.slideLayouts[0].shapes.at(-1);
    const slideNumberEdit = source.edits?.at(-1);
    if (slideNumberEdit?.kind !== "addTextBox") throw new Error("expected slide number edit");
    expect(slideNumber?.nodeId).toBe(slideNumberEdit.shapeId);
    expect(slideNumber?.handle?.orderingSlot).toBe(0);
    expect(slideNumberEdit.xml).toContain(`id="${slideNumberEdit.shapeId}"`);
  });
});

function shape(
  partPath: ReturnType<typeof asPartPath>,
  id: string,
  orderingSlot: number,
): SourceShape {
  return {
    kind: "shape",
    nodeId: asSourceNodeId(id),
    handle: {
      partPath,
      nodeId: asSourceNodeId(id),
      orderingSlot,
      rawSidecarIds: [asRawSidecarId(`shape-${id}`)],
    },
  };
}

function withShapeTreeRootId(
  source: PptxSourceModel,
  partPath: ReturnType<typeof asPartPath>,
  id: number,
): PptxSourceModel {
  const bytes = new TextEncoder().encode(
    `<p:sld><p:cSld><p:spTree><p:nvGrpSpPr><p:cNvPr id="${id}"/></p:nvGrpSpPr></p:spTree></p:cSld></p:sld>`,
  );
  return {
    ...source,
    packageGraph: {
      ...source.packageGraph,
      rawParts: (source.packageGraph.rawParts ?? []).map((part) =>
        part.partPath === partPath ? { ...part, kind: "binary", bytes } : part,
      ),
    },
  };
}
