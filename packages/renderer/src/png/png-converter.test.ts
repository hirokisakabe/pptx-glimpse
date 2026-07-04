import { describe, expect, it, vi } from "vitest";

const mocks = vi.hoisted(() => {
  const initWasm = vi.fn().mockResolvedValue(undefined);
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
  const readFile = vi.fn().mockResolvedValue(new Uint8Array([0]));
  const requireResolve = vi.fn().mockReturnValue("/mock/resvg.wasm");
  const createRequire = vi.fn().mockReturnValue({ resolve: requireResolve });
  return { createRequire, initWasm, MockResvg, mockRender, readFile, requireResolve };
});

vi.mock("@resvg/resvg-wasm", () => ({
  initWasm: mocks.initWasm,
  Resvg: mocks.MockResvg,
}));

vi.mock("node:fs/promises", () => ({
  readFile: mocks.readFile,
}));

vi.mock("node:module", () => ({
  createRequire: mocks.createRequire,
}));

import { unsafeFixtureAssertion } from "../unsafe-type-assertion.js";

const MINIMAL_SVG = '<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 10 10"></svg>';

async function loadPngConverter() {
  vi.resetModules();
  mocks.initWasm.mockClear();
  mocks.MockResvg.mockClear();
  mocks.mockRender.mockClear();
  mocks.readFile.mockClear();
  mocks.createRequire.mockClear();
  mocks.requireResolve.mockClear();
  return import("./png-converter.js");
}

describe("initResvgWasm", () => {
  it("loads the bundled WASM through Node.js APIs when no input is specified", async () => {
    const { initResvgWasm } = await loadPngConverter();

    await initResvgWasm();

    expect(mocks.createRequire).toHaveBeenCalledTimes(1);
    expect(mocks.requireResolve).toHaveBeenCalledWith("@resvg/resvg-wasm/index_bg.wasm");
    expect(mocks.readFile).toHaveBeenCalledWith("/mock/resvg.wasm");
    expect(mocks.initWasm).toHaveBeenCalledWith(new Uint8Array([0]));
  });

  it("initializes from an ArrayBuffer without executing Node.js WASM loading", async () => {
    const { initResvgWasm } = await loadPngConverter();
    const wasm = new Uint8Array([1, 2, 3]).buffer;

    await initResvgWasm(wasm);

    expect(mocks.initWasm).toHaveBeenCalledWith(wasm);
    expect(mocks.createRequire).not.toHaveBeenCalled();
    expect(mocks.readFile).not.toHaveBeenCalled();
  });

  it("initializes from a Uint8Array without executing Node.js WASM loading", async () => {
    const { initResvgWasm } = await loadPngConverter();
    const wasm = new Uint8Array([1, 2, 3]);

    await initResvgWasm(wasm);

    expect(mocks.initWasm).toHaveBeenCalledWith(wasm);
    expect(mocks.createRequire).not.toHaveBeenCalled();
    expect(mocks.readFile).not.toHaveBeenCalled();
  });

  it("initializes from a Response without executing Node.js WASM loading", async () => {
    const { initResvgWasm } = await loadPngConverter();
    const wasm = new Uint8Array([1, 2, 3]).buffer;
    const response = new Response(wasm);

    await initResvgWasm(response);

    const actual = unsafeFixtureAssertion<ArrayBuffer | Uint8Array>(
      mocks.initWasm.mock.calls[0]?.[0],
    );
    expect([...new Uint8Array(actual)]).toEqual([1, 2, 3]);
    expect(mocks.createRequire).not.toHaveBeenCalled();
    expect(mocks.readFile).not.toHaveBeenCalled();
  });
});

describe("svgToPng", () => {
  it("Do not pass font option to Resvg when fontBuffers is not specified", async () => {
    const { svgToPng } = await loadPngConverter();

    await svgToPng(MINIMAL_SVG);
    const opts = unsafeFixtureAssertion<Record<string, unknown> | undefined>(
      mocks.MockResvg.mock.calls[0]?.[1],
    );
    expect(opts?.font).toBeUndefined();
  });

  it("Do not pass font option to Resvg when fontBuffers is empty array", async () => {
    const { svgToPng } = await loadPngConverter();

    await svgToPng(MINIMAL_SVG, { fontBuffers: [] });
    const opts = unsafeFixtureAssertion<Record<string, unknown> | undefined>(
      mocks.MockResvg.mock.calls[0]?.[1],
    );
    expect(opts?.font).toBeUndefined();
  });

  it("If fontBuffers is specified, it will be passed to Resvg as font: { fontBuffers }", async () => {
    const { svgToPng } = await loadPngConverter();
    const fontBuffers = [new Uint8Array([1, 2, 3])];

    await svgToPng(MINIMAL_SVG, { fontBuffers });
    const opts = unsafeFixtureAssertion<Record<string, unknown> | undefined>(
      mocks.MockResvg.mock.calls[0]?.[1],
    );
    expect(opts?.font).toEqual({ fontBuffers });
  });
});
