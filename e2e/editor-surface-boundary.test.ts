import { readFile } from "node:fs/promises";
import { resolve } from "node:path";

import { describe, expect, it } from "vitest";

const componentsRoot = resolve(import.meta.dirname, "../demo/src/components");
const reusableComponentFiles = [
  "EditorSurface.tsx",
  "EditorToolbar.tsx",
  "EditorSlideStrip.tsx",
  "editor-controller.ts",
  "use-editor-controller.ts",
];

describe("demo editor component boundary", () => {
  it("keeps reusable editor components independent from Next.js and the demo shell", async () => {
    for (const fileName of reusableComponentFiles) {
      const source = await readFile(resolve(componentsRoot, fileName), "utf8");
      expect(source, fileName).not.toMatch(/from ["']next(?:\/|["'])/);
      expect(source, fileName).not.toContain("DemoEditorShell");
    }
  });

  it("composes file-oriented host policy around the reusable surface", async () => {
    const workspace = await readFile(resolve(componentsRoot, "EditorWorkspace.tsx"), "utf8");
    const shell = await readFile(resolve(componentsRoot, "DemoEditorShell.tsx"), "utf8");
    const surface = await readFile(resolve(componentsRoot, "EditorSurface.tsx"), "utf8");

    expect(workspace).toContain("<EditorSurface");
    expect(workspace).toContain("<DemoEditorShell");
    for (const hostResponsibility of [
      "Open PPTX",
      "Open sample",
      "Add fonts",
      "Download PPTX",
      "Presentation file name",
      "beforeunload",
    ]) {
      expect(shell).toContain(hostResponsibility);
      expect(surface).not.toContain(hostResponsibility);
    }
  });
});
