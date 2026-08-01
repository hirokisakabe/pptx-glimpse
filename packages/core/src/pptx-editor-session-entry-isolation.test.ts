import { createPptx, writePptx } from "@pptx-glimpse/document";
import { beforeEach, expect, it, vi } from "vitest";

const nodeFontMocks = vi.hoisted(() => ({
  createOpentypeSetupFromSystem: vi.fn().mockResolvedValue(null),
}));

vi.mock("@pptx-glimpse/renderer/node", () => ({
  createOpentypeSetupFromSystem: nodeFontMocks.createOpentypeSetupFromSystem,
}));

beforeEach(() => {
  vi.resetModules();
  nodeFontMocks.createOpentypeSetupFromSystem.mockClear();
});

it("keeps browser and Node editor renderers isolated when the browser entry loads first", async () => {
  const input = writePptx(createPptx());
  const renderOptions = {
    fontDirs: ["/browser-first-node-fonts"],
    skipSystemFonts: true,
  };

  const browserEntry = await import("./browser.js");
  await browserEntry.createPptxEditorSession(input, renderOptions);
  expect(nodeFontMocks.createOpentypeSetupFromSystem).not.toHaveBeenCalled();

  const nodeEntry = await import("./index.js");
  await nodeEntry.createPptxEditorSession(input, renderOptions);
  expect(nodeFontMocks.createOpentypeSetupFromSystem).toHaveBeenCalledOnce();

  nodeFontMocks.createOpentypeSetupFromSystem.mockClear();
  await browserEntry.PptxEditorSession.create(input, renderOptions);
  expect(nodeFontMocks.createOpentypeSetupFromSystem).not.toHaveBeenCalled();
});
