import { describe, it, expect } from "vitest";
import { readFileSync } from "fs";
import { join } from "path";
import { convertPptxToSvg, convertPptxToPng } from "../../src/converter.js";

const FIXTURE_PATH = join(__dirname, "../fixtures/basic-shapes.pptx");

describe("convertPptxToSvg", () => {
  it("converts a PPTX file to SVG", async () => {
    const input = readFileSync(FIXTURE_PATH);
    const results = await convertPptxToSvg(input);

    expect(results).toHaveLength(1);
    expect(results[0].slideNumber).toBe(1);
    expect(results[0].svg).toContain("<svg");
    expect(results[0].svg).toContain("</svg>");
  });

  it("renders basic shapes", async () => {
    const input = readFileSync(FIXTURE_PATH);
    const results = await convertPptxToSvg(input);
    const svg = results[0].svg;

    // Should contain rect (blue rectangle)
    expect(svg).toContain("<rect");
    // Should contain ellipse (orange)
    expect(svg).toContain("<ellipse");
    // Should contain text
    expect(svg).toContain("Hello World");
    expect(svg).toContain("Rounded Rectangle");
  });

  it("has correct viewBox dimensions for 16:9", async () => {
    const input = readFileSync(FIXTURE_PATH);
    const results = await convertPptxToSvg(input);
    const svg = results[0].svg;

    // 9144000 EMU = 960px, 5143500 EMU ≈ 540px
    expect(svg).toContain('viewBox="0 0 960');
  });

  it("applies fill colors", async () => {
    const input = readFileSync(FIXTURE_PATH);
    const results = await convertPptxToSvg(input);
    const svg = results[0].svg;

    // Blue rectangle fill
    expect(svg).toContain("#4472C4");
    // Orange ellipse fill
    expect(svg).toContain("#ED7D31");
  });

  it("supports slide number filtering", async () => {
    const input = readFileSync(FIXTURE_PATH);
    const results = await convertPptxToSvg(input, { slides: [1] });

    expect(results).toHaveLength(1);
    expect(results[0].slideNumber).toBe(1);
  });

  it("returns empty for non-existent slide numbers", async () => {
    const input = readFileSync(FIXTURE_PATH);
    const results = await convertPptxToSvg(input, { slides: [99] });

    expect(results).toHaveLength(0);
  });
});

describe("convertPptxToPng", () => {
  it("converts a PPTX file to PNG", async () => {
    const input = readFileSync(FIXTURE_PATH);
    const results = await convertPptxToPng(input);

    expect(results).toHaveLength(1);
    expect(results[0].slideNumber).toBe(1);
    expect(results[0].png).toBeInstanceOf(Buffer);
    expect(results[0].png.length).toBeGreaterThan(0);
    // PNG magic bytes
    expect(results[0].png[0]).toBe(0x89);
    expect(results[0].png[1]).toBe(0x50); // P
    expect(results[0].png[2]).toBe(0x4e); // N
    expect(results[0].png[3]).toBe(0x47); // G
  });

  it("respects width option", async () => {
    const input = readFileSync(FIXTURE_PATH);
    const results = await convertPptxToPng(input, { width: 480 });

    expect(results[0].width).toBe(480);
    expect(results[0].height).toBeGreaterThan(0);
  });
});
