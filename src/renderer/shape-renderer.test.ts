import { describe, it, expect, vi, beforeEach } from "vitest";
import { renderShape, renderConnector } from "./shape-renderer.js";
import type { ShapeElement, ConnectorElement, Transform } from "../model/shape.js";

beforeEach(() => {
  let counter = 0;
  vi.spyOn(crypto, "randomUUID").mockImplementation(() => {
    return `test-uuid-${counter++}` as ReturnType<typeof crypto.randomUUID>;
  });
});

function makeTransform(overrides: Partial<Transform> = {}): Transform {
  return {
    offsetX: 914400,
    offsetY: 914400,
    extentWidth: 1828800,
    extentHeight: 914400,
    rotation: 0,
    flipH: false,
    flipV: false,
    ...overrides,
  };
}

function makeShape(overrides: Partial<ShapeElement> = {}): ShapeElement {
  return {
    type: "shape",
    transform: makeTransform(),
    geometry: { type: "preset", preset: "rect", adjustValues: {} },
    fill: { type: "solid", color: { hex: "#FF0000", alpha: 1 } },
    outline: {
      width: 12700,
      fill: { type: "solid", color: { hex: "#000000", alpha: 1 } },
      dashStyle: "solid",
      headEnd: null,
      tailEnd: null,
    },
    textBody: null,
    effects: null,
    ...overrides,
  };
}

function makeConnector(overrides: Partial<ConnectorElement> = {}): ConnectorElement {
  return {
    type: "connector",
    transform: makeTransform(),
    geometry: { type: "preset", preset: "line", adjustValues: {} },
    outline: {
      width: 12700,
      fill: { type: "solid", color: { hex: "#000000", alpha: 1 } },
      dashStyle: "solid",
      headEnd: null,
      tailEnd: null,
    },
    effects: null,
    ...overrides,
  };
}

describe("renderShape", () => {
  it("renders basic shape with fill and outline", () => {
    const result = renderShape(makeShape());

    expect(result).toContain('<g transform="translate(96, 96)">');
    expect(result).toContain('fill="#FF0000"');
    expect(result).toContain('stroke="#000000"');
    expect(result).toContain("</g>");
  });

  it("renders shape with no fill", () => {
    const result = renderShape(makeShape({ fill: { type: "none" } }));
    expect(result).toContain('fill="none"');
  });

  it("renders shape with null fill", () => {
    const result = renderShape(makeShape({ fill: null }));
    expect(result).toContain('fill="none"');
  });

  it("renders shape with null outline", () => {
    const result = renderShape(makeShape({ outline: null }));
    expect(result).toContain('stroke="none"');
  });

  it("renders shape with rotation", () => {
    const result = renderShape(makeShape({ transform: makeTransform({ rotation: 90 }) }));
    expect(result).toContain("rotate(90, 96, 48)");
  });

  it("renders shape with effects", () => {
    const result = renderShape(
      makeShape({
        effects: {
          outerShadow: {
            blurRadius: 50800,
            distance: 38100,
            direction: 45,
            color: { hex: "#000000", alpha: 0.5 },
            alignment: "br",
            rotateWithShape: false,
          },
          innerShadow: null,
          glow: null,
          softEdge: null,
        },
      }),
    );

    expect(result).toContain('<filter id="effect-test-uuid-0"');
    expect(result).toContain('filter="url(#effect-test-uuid-0)"');
  });

  it("renders shape with gradient fill", () => {
    const result = renderShape(
      makeShape({
        fill: {
          type: "gradient",
          angle: 90,
          gradientType: "linear",
          stops: [
            { position: 0, color: { hex: "#FF0000", alpha: 1 } },
            { position: 1, color: { hex: "#0000FF", alpha: 1 } },
          ],
        },
      }),
    );

    expect(result).toContain("<linearGradient");
    expect(result).toContain("url(#grad-");
  });
});

describe("renderConnector", () => {
  it("renders basic connector with line element", () => {
    const result = renderConnector(makeConnector());

    expect(result).toContain('<g transform="translate(96, 96)">');
    expect(result).toContain('<line x1="0" y1="0" x2="192" y2="96"');
    expect(result).toContain('stroke="#000000"');
    expect(result).toContain('fill="none"');
    expect(result).toContain("</g>");
  });

  it("renders connector with null outline", () => {
    const result = renderConnector(makeConnector({ outline: null }));
    expect(result).toContain('stroke="none"');
  });

  it("renders connector with rotation", () => {
    const result = renderConnector(makeConnector({ transform: makeTransform({ rotation: 45 }) }));
    expect(result).toContain("rotate(45, 96, 48)");
  });

  it("renders connector with flipH", () => {
    const result = renderConnector(makeConnector({ transform: makeTransform({ flipH: true }) }));
    expect(result).toContain("translate(192, 0) scale(-1, 1)");
  });

  it("renders connector with effects", () => {
    const result = renderConnector(
      makeConnector({
        effects: {
          outerShadow: null,
          innerShadow: null,
          glow: { radius: 127000, color: { hex: "#00FF00", alpha: 1 } },
          softEdge: null,
        },
      }),
    );

    expect(result).toContain("<filter");
    expect(result).toContain("filter=");
  });

  it("renders connector with dash style", () => {
    const result = renderConnector(
      makeConnector({
        outline: {
          width: 12700,
          fill: { type: "solid", color: { hex: "#FF0000", alpha: 1 } },
          dashStyle: "dash",
          headEnd: null,
          tailEnd: null,
        },
      }),
    );

    expect(result).toContain('stroke="#FF0000"');
    expect(result).toContain("stroke-dasharray=");
  });
});
