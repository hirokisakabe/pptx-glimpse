import { describe, it, expect } from "vitest";
import { renderFillAttrs, renderOutlineAttrs } from "../../src/renderer/fill-renderer.js";

describe("renderFillAttrs", () => {
  it("renders null fill as none", () => {
    const result = renderFillAttrs(null);
    expect(result.attrs).toBe('fill="none"');
    expect(result.defs).toBe("");
  });

  it("renders noFill", () => {
    const result = renderFillAttrs({ type: "none" });
    expect(result.attrs).toBe('fill="none"');
  });

  it("renders solid fill", () => {
    const result = renderFillAttrs({
      type: "solid",
      color: { hex: "#FF0000", alpha: 1 },
    });
    expect(result.attrs).toBe('fill="#FF0000"');
    expect(result.defs).toBe("");
  });

  it("renders solid fill with alpha", () => {
    const result = renderFillAttrs({
      type: "solid",
      color: { hex: "#FF0000", alpha: 0.5 },
    });
    expect(result.attrs).toContain('fill="#FF0000"');
    expect(result.attrs).toContain('fill-opacity="0.5"');
  });

  it("renders gradient fill with defs", () => {
    const result = renderFillAttrs({
      type: "gradient",
      angle: 90,
      stops: [
        { position: 0, color: { hex: "#FF0000", alpha: 1 } },
        { position: 1, color: { hex: "#0000FF", alpha: 1 } },
      ],
    });
    expect(result.attrs).toContain("url(#grad-");
    expect(result.defs).toContain("<linearGradient");
    expect(result.defs).toContain("#FF0000");
    expect(result.defs).toContain("#0000FF");
  });
});

describe("renderOutlineAttrs", () => {
  it("renders null outline as stroke none", () => {
    expect(renderOutlineAttrs(null)).toBe('stroke="none"');
  });

  it("renders outline with color", () => {
    const result = renderOutlineAttrs({
      width: 12700,
      fill: { type: "solid", color: { hex: "#000000", alpha: 1 } },
      dashStyle: "solid",
    });
    expect(result).toContain('stroke="#000000"');
    expect(result).toContain("stroke-width=");
  });

  it("renders dash style", () => {
    const result = renderOutlineAttrs({
      width: 12700,
      fill: { type: "solid", color: { hex: "#000000", alpha: 1 } },
      dashStyle: "dash",
    });
    expect(result).toContain("stroke-dasharray=");
  });
});
