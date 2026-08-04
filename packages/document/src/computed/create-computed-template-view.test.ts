import { describe, expect, it } from "vitest";

import {
  asEmu,
  createComputedTemplateView,
  createPptx,
  createPptxAuthoringSession,
} from "../index.js";

describe("createComputedTemplateView", () => {
  it("projects a layout/master cascade without reading slide content or mutating source", () => {
    const initial = createPptx({
      slideMaster: {
        background: { kind: "solid", color: { kind: "srgb", hex: "AA0000" } },
      },
    });
    const master = requireValue(initial.slideMasters[0]);
    const layout = requireValue(initial.slideLayouts[0]);
    const slide = requireValue(initial.slides[0]);
    const session = createPptxAuthoringSession(initial);
    session.target(requireValue(master.handle)).addPlaceholder({
      type: "title",
      index: 4,
      name: "Master title",
      transform: {
        offsetX: asEmu(10),
        offsetY: asEmu(20),
        width: asEmu(300),
        height: asEmu(400),
      },
      promptText: "Master prompt",
    });
    session.target(requireValue(layout.handle)).addPlaceholder({
      type: "ctrTitle",
      index: 4,
      name: "Layout title",
      promptText: "Layout prompt",
    });
    session.target(requireValue(slide.handle)).addTextBox({
      offsetX: asEmu(1),
      offsetY: asEmu(2),
      width: asEmu(3),
      height: asEmu(4),
      text: "Slide user content",
    });
    const source = session.source;
    const before = structuredClone(source);

    const computed = createComputedTemplateView(source, {
      kind: "slideLayout",
      partPath: layout.partPath,
    });

    expect(source).toEqual(before);
    expect(computed.slides).toHaveLength(1);
    expect(computed.slides[0]).toMatchObject({
      partPath: layout.partPath,
      layoutPartPath: layout.partPath,
      masterPartPath: master.partPath,
      background: { sourceLayer: "master", fill: { color: { hex: "#aa0000" } } },
    });
    const title = computed.slides[0]?.elements.find(
      (element) => element.kind === "shape" && element.sourceNode.name === "Layout title",
    );
    expect(title).toMatchObject({
      sourceLayer: "layout",
      transform: { offsetX: 10, offsetY: 20, width: 300, height: 400 },
      placeholderMatch: {
        layout: { name: "Layout title" },
        master: { name: "Master title" },
      },
    });
    expect(
      computed.slides[0]?.elements.some(
        (element) => element.sourceNode.name === "Slide user content",
      ),
    ).toBe(false);
  });
});

function requireValue<T>(value: T | undefined): T {
  if (value === undefined) throw new Error("required fixture value is missing");
  return value;
}
