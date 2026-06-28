import { describe, expect, it, vi } from "vitest";

const mocks = vi.hoisted(() => {
  const mockRender = vi.fn(() => ({
    asPng: () => Buffer.from([0x89, 0x50, 0x4e, 0x47]),
    width: 960,
    height: 540,
  }));
  const MockResvg = vi.fn().mockImplementation(function (_svg: string, _opts?: unknown) {
    return {
      render: mockRender,
    };
  });
  return { MockResvg, mockRender };
});

vi.mock("@resvg/resvg-wasm", () => ({
  initWasm: vi.fn().mockResolvedValue(undefined),
  Resvg: mocks.MockResvg,
}));

vi.mock("fs/promises", () => ({
  readFile: vi.fn().mockResolvedValue(Buffer.alloc(0)),
}));

import { unsafeFixtureAssertion } from "../unsafe-type-assertion.js";
import { svgToPng } from "./png-converter.js";

const MINIMAL_SVG = '<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 10 10"></svg>';

describe("svgToPng", () => {
  it("Do not pass font option to Resvg when fontBuffers is not specified", async () => {
    mocks.MockResvg.mockClear();
    await svgToPng(MINIMAL_SVG);
    const opts = unsafeFixtureAssertion<Record<string, unknown> | undefined>(
      mocks.MockResvg.mock.calls[0]?.[1],
    );
    expect(opts?.font).toBeUndefined();
  });

  it("Do not pass font option to Resvg when fontBuffers is empty array", async () => {
    mocks.MockResvg.mockClear();
    await svgToPng(MINIMAL_SVG, { fontBuffers: [] });
    const opts = unsafeFixtureAssertion<Record<string, unknown> | undefined>(
      mocks.MockResvg.mock.calls[0]?.[1],
    );
    expect(opts?.font).toBeUndefined();
  });

  it("If fontBuffers is specified, it will be passed to Resvg as font: { fontBuffers }", async () => {
    const fontBuffers = [new Uint8Array([1, 2, 3])];
    mocks.MockResvg.mockClear();
    await svgToPng(MINIMAL_SVG, { fontBuffers });
    const opts = unsafeFixtureAssertion<Record<string, unknown> | undefined>(
      mocks.MockResvg.mock.calls[0]?.[1],
    );
    expect(opts?.font).toEqual({ fontBuffers });
  });
});
