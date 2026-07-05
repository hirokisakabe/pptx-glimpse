import { describe, expect, it, vi } from "vitest";

const mocks = vi.hoisted(() => {
  const initWasm = vi.fn().mockResolvedValue(undefined);
  const MockResvg = vi.fn().mockImplementation(function (_svg: string, _opts?: unknown) {
    return {
      render: () => ({
        asPng: () => new Uint8Array([0x89, 0x50, 0x4e, 0x47]),
        width: 960,
        height: 540,
      }),
    };
  });
  return { initWasm, MockResvg };
});

vi.mock("@resvg/resvg-wasm", () => ({
  initWasm: mocks.initWasm,
  Resvg: mocks.MockResvg,
}));

async function loadBrowserPngConverter() {
  vi.resetModules();
  mocks.initWasm.mockReset();
  mocks.initWasm.mockResolvedValue(undefined);
  mocks.MockResvg.mockClear();
  return import("./browser-png-converter.js");
}

describe("browser initResvgWasm", () => {
  it("allows retrying after initialization failure", async () => {
    const { initResvgWasm } = await loadBrowserPngConverter();
    mocks.initWasm.mockRejectedValueOnce(new Error("bad wasm"));

    await expect(initResvgWasm(new Uint8Array([1]))).rejects.toThrow("bad wasm");
    await initResvgWasm(new Uint8Array([2]));

    expect(mocks.initWasm).toHaveBeenCalledTimes(2);
    expect(mocks.initWasm).toHaveBeenNthCalledWith(2, new Uint8Array([2]));
  });

  it("rejects non-ok Response input before calling initWasm", async () => {
    const { initResvgWasm } = await loadBrowserPngConverter();
    const response = new Response("not found", { status: 404 });

    await expect(initResvgWasm(response)).rejects.toThrow("HTTP 404");
    expect(mocks.initWasm).not.toHaveBeenCalled();
  });
});
