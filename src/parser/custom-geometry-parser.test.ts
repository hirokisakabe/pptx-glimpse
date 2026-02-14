import { describe, it, expect } from "vitest";
import { parseCustomGeometry } from "./custom-geometry-parser.js";

describe("parseCustomGeometry", () => {
  it("parses moveTo + lnTo + close", () => {
    const custGeom = {
      avLst: {},
      gdLst: {},
      pathLst: {
        path: {
          "@_w": "1000",
          "@_h": "1000",
          moveTo: { pt: [{ "@_x": "0", "@_y": "0" }] },
          lnTo: [
            { pt: [{ "@_x": "1000", "@_y": "0" }] },
            { pt: [{ "@_x": "1000", "@_y": "1000" }] },
          ],
          close: "",
        },
      },
    };
    const result = parseCustomGeometry(custGeom);
    expect(result).not.toBeNull();
    expect(result).toHaveLength(1);
    expect(result![0].width).toBe(1000);
    expect(result![0].height).toBe(1000);
    expect(result![0].commands).toBe("M 0 0 L 1000 0 L 1000 1000 Z");
  });

  it("parses cubicBezTo", () => {
    const custGeom = {
      pathLst: {
        path: {
          "@_w": "1000",
          "@_h": "1000",
          moveTo: { pt: [{ "@_x": "0", "@_y": "500" }] },
          cubicBezTo: {
            pt: [
              { "@_x": "250", "@_y": "0" },
              { "@_x": "750", "@_y": "1000" },
              { "@_x": "1000", "@_y": "500" },
            ],
          },
          close: "",
        },
      },
    };
    const result = parseCustomGeometry(custGeom);
    expect(result).not.toBeNull();
    expect(result![0].commands).toBe("M 0 500 C 250 0, 750 1000, 1000 500 Z");
  });

  it("parses quadBezTo", () => {
    const custGeom = {
      pathLst: {
        path: {
          "@_w": "1000",
          "@_h": "1000",
          moveTo: { pt: [{ "@_x": "0", "@_y": "1000" }] },
          quadBezTo: {
            pt: [
              { "@_x": "500", "@_y": "0" },
              { "@_x": "1000", "@_y": "1000" },
            ],
          },
          close: "",
        },
      },
    };
    const result = parseCustomGeometry(custGeom);
    expect(result).not.toBeNull();
    expect(result![0].commands).toBe("M 0 1000 Q 500 0, 1000 1000 Z");
  });

  it("parses arcTo and calculates correct endpoint", () => {
    // 半円: start at (500, 0), arcTo with wR=500, hR=500,
    // stAng=16200000 (270°), swAng=10800000 (180°)
    // 270° → start at top of circle, sweep 180° → end at bottom
    const custGeom = {
      pathLst: {
        path: {
          "@_w": "1000",
          "@_h": "1000",
          moveTo: { pt: [{ "@_x": "500", "@_y": "0" }] },
          arcTo: {
            "@_wR": "500",
            "@_hR": "500",
            "@_stAng": "16200000",
            "@_swAng": "10800000",
          },
        },
      },
    };
    const result = parseCustomGeometry(custGeom);
    expect(result).not.toBeNull();
    const cmd = result![0].commands;
    expect(cmd).toMatch(/^M 500 0 A 500 500 0 0 1/);
    // endpoint should be approximately (500, 1000) - bottom of the circle
    const match = cmd.match(/A .+ (\S+) (\S+)$/);
    expect(match).not.toBeNull();
    expect(Number(match![1])).toBeCloseTo(500, 0);
    expect(Number(match![2])).toBeCloseTo(1000, 0);
  });

  it("resolves guide values in coordinates", () => {
    const custGeom = {
      avLst: { gd: [{ "@_name": "adj", "@_fmla": "val 50000" }] },
      gdLst: {
        gd: [
          { "@_name": "midX", "@_fmla": "*/ w adj 100000" },
          { "@_name": "midY", "@_fmla": "*/ h adj 100000" },
        ],
      },
      pathLst: {
        path: {
          "@_w": "1000",
          "@_h": "1000",
          moveTo: { pt: [{ "@_x": "0", "@_y": "0" }] },
          lnTo: [
            { pt: [{ "@_x": "midX", "@_y": "0" }] },
            { pt: [{ "@_x": "midX", "@_y": "midY" }] },
            { pt: [{ "@_x": "0", "@_y": "midY" }] },
          ],
          close: "",
        },
      },
    };
    const result = parseCustomGeometry(custGeom);
    expect(result).not.toBeNull();
    // midX = 1000 * 50000 / 100000 = 500
    // midY = 1000 * 50000 / 100000 = 500
    expect(result![0].commands).toBe("M 0 0 L 500 0 L 500 500 L 0 500 Z");
  });

  it("returns path dimensions for scaling", () => {
    const custGeom = {
      pathLst: {
        path: {
          "@_w": "2000",
          "@_h": "1500",
          moveTo: { pt: [{ "@_x": "0", "@_y": "0" }] },
          lnTo: { pt: [{ "@_x": "2000", "@_y": "1500" }] },
        },
      },
    };
    const result = parseCustomGeometry(custGeom);
    expect(result).not.toBeNull();
    expect(result![0].width).toBe(2000);
    expect(result![0].height).toBe(1500);
  });

  it("returns null for empty pathLst", () => {
    expect(parseCustomGeometry({})).toBeNull();
    expect(parseCustomGeometry({ pathLst: {} })).toBeNull();
  });

  it("handles multiple paths", () => {
    const custGeom = {
      pathLst: {
        path: [
          {
            "@_w": "1000",
            "@_h": "1000",
            moveTo: { pt: [{ "@_x": "0", "@_y": "0" }] },
            lnTo: { pt: [{ "@_x": "1000", "@_y": "1000" }] },
          },
          {
            "@_w": "500",
            "@_h": "500",
            moveTo: { pt: [{ "@_x": "0", "@_y": "0" }] },
            lnTo: { pt: [{ "@_x": "500", "@_y": "500" }] },
          },
        ],
      },
    };
    const result = parseCustomGeometry(custGeom);
    expect(result).not.toBeNull();
    expect(result).toHaveLength(2);
    expect(result![0].width).toBe(1000);
    expect(result![1].width).toBe(500);
  });

  it("skips paths with w=0 and h=0", () => {
    const custGeom = {
      pathLst: {
        path: {
          "@_w": "0",
          "@_h": "0",
          moveTo: { pt: [{ "@_x": "0", "@_y": "0" }] },
        },
      },
    };
    expect(parseCustomGeometry(custGeom)).toBeNull();
  });
});
