import { describe, it, expect } from "vitest";
import { emuToPixels, emuToPoints, rotationToDegrees, hundredthPointToPoint } from "./emu.js";
import { asEmu, asHundredthPt } from "./unit-types.js";

describe("emuToPixels", () => {
  it("converts 914400 EMU (1 inch) to 96 pixels at 96 DPI", () => {
    expect(emuToPixels(asEmu(914400))).toBe(96);
  });

  it("converts 0 EMU to 0 pixels", () => {
    expect(emuToPixels(asEmu(0))).toBe(0);
  });

  it("converts standard 16:9 slide width to 960 pixels", () => {
    expect(emuToPixels(asEmu(9144000))).toBe(960);
  });

  it("converts standard 16:9 slide height to 540 pixels", () => {
    expect(emuToPixels(asEmu(5143500))).toBeCloseTo(540, 0);
  });

  it("supports custom DPI", () => {
    expect(emuToPixels(asEmu(914400), 72)).toBe(72);
  });
});

describe("emuToPoints", () => {
  it("converts 12700 EMU to 1 point", () => {
    expect(emuToPoints(asEmu(12700))).toBe(1);
  });
});

describe("rotationToDegrees", () => {
  it("converts 5400000 to 90 degrees", () => {
    expect(rotationToDegrees(5400000)).toBe(90);
  });

  it("converts 0 to 0 degrees", () => {
    expect(rotationToDegrees(0)).toBe(0);
  });
});

describe("hundredthPointToPoint", () => {
  it("converts 2800 to 28 points", () => {
    expect(hundredthPointToPoint(asHundredthPt(2800))).toBe(28);
  });
});
