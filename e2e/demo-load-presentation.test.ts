import { describe, expect, it, vi } from "vitest";

import { loadPresentationWith } from "../demo/src/components/load-presentation.js";

describe("demo presentation loading", () => {
  it("uses only SVG conversion for view mode", async () => {
    const file = memoryFile("slides.pptx", [1, 2, 3]);
    const font = memoryFile("DemoFont.ttf", [4, 5]);
    const convertToSvg = vi.fn(() => Promise.resolve([{ slideNumber: 1, svg: "<svg />" }]));
    const createEditor = vi.fn(() => Promise.resolve({ slides: [{}] }));

    const result = await loadPresentationWith(file, "view", [font], {
      convertToSvg,
      createEditor,
    });

    expect(result.mode).toBe("view");
    expect(convertToSvg).toHaveBeenCalledOnce();
    expect(createEditor).not.toHaveBeenCalled();
    expect(convertToSvg.mock.calls[0]?.[0]).toEqual(new Uint8Array([1, 2, 3]));
    expect(convertToSvg.mock.calls[0]?.[1].map(({ name }) => name)).toEqual(["DemoFont"]);
  });

  it("uses only editor session creation for edit mode", async () => {
    const editor = { slides: [{ slideNumber: 1 }], marker: "editor" };
    const convertToSvg = vi.fn(() => Promise.resolve([{ slideNumber: 1, svg: "<svg />" }]));
    const createEditor = vi.fn(() => Promise.resolve(editor));

    const result = await loadPresentationWith(memoryFile("slides.pptx", [7, 8]), "edit", [], {
      convertToSvg,
      createEditor,
    });

    expect(result).toMatchObject({ mode: "edit", fileName: "slides.pptx", editor });
    expect(createEditor).toHaveBeenCalledOnce();
    expect(convertToSvg).not.toHaveBeenCalled();
  });

  it.each(["view", "edit"] as const)("rejects an empty %s result", async (mode) => {
    await expect(
      loadPresentationWith(memoryFile("empty.pptx", []), mode, [], {
        convertToSvg: () => Promise.resolve([]),
        createEditor: () => Promise.resolve({ slides: [] }),
      }),
    ).rejects.toThrow("No slides found in the selected file");
  });
});

function memoryFile(name: string, bytes: readonly number[]) {
  return {
    name,
    arrayBuffer: () => Promise.resolve(new Uint8Array(bytes).buffer),
  };
}
