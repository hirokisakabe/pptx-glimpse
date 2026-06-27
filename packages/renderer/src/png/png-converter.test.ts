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

import { unsafeTypeAssertion } from "../unsafe-type-assertion.js";
import { svgToPng } from "./png-converter.js";

const MINIMAL_SVG = '<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 10 10"></svg>';

describe("svgToPng", () => {
  it("fontBuffers 未指定のとき font オプションを Resvg に渡さない", async () => {
    mocks.MockResvg.mockClear();
    await svgToPng(MINIMAL_SVG);
    const opts = unsafeTypeAssertion<Record<string, unknown> | undefined>(
      mocks.MockResvg.mock.calls[0]?.[1],
    );
    expect(opts?.font).toBeUndefined();
  });

  it("fontBuffers が空配列のとき font オプションを Resvg に渡さない", async () => {
    mocks.MockResvg.mockClear();
    await svgToPng(MINIMAL_SVG, { fontBuffers: [] });
    const opts = unsafeTypeAssertion<Record<string, unknown> | undefined>(
      mocks.MockResvg.mock.calls[0]?.[1],
    );
    expect(opts?.font).toBeUndefined();
  });

  it("fontBuffers を指定すると font: { fontBuffers } として Resvg に渡る", async () => {
    const fontBuffers = [new Uint8Array([1, 2, 3])];
    mocks.MockResvg.mockClear();
    await svgToPng(MINIMAL_SVG, { fontBuffers });
    const opts = unsafeTypeAssertion<Record<string, unknown> | undefined>(
      mocks.MockResvg.mock.calls[0]?.[1],
    );
    expect(opts?.font).toEqual({ fontBuffers });
  });
});
