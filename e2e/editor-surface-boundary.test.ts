import { readFile } from "node:fs/promises";
import { resolve } from "node:path";

import { describe, expect, it } from "vitest";

const componentsRoot = resolve(import.meta.dirname, "../demo/src/components");
const reusableComponentFiles = [
  "EditorSurface.tsx",
  "EditorToolbar.tsx",
  "EditorSlideStrip.tsx",
  "direct-text-editor-lifecycle.ts",
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

  it("resets toolbar-local run selection when the editor session identity changes", async () => {
    const toolbar = await readFile(resolve(componentsRoot, "EditorToolbar.tsx"), "utf8");
    const surface = await readFile(resolve(componentsRoot, "EditorSurface.tsx"), "utf8");

    expect(toolbar).toContain("readonly selectionScope: object");
    expect(toolbar).toContain('setFontSize("24")');
    expect(toolbar).toContain('setTypeface("")');
    expect(toolbar).toContain('setColor("#2454a6")');
    expect(toolbar).toContain("[selectedShapeKey, selectionScope]");
    expect(toolbar).toContain("[selectionScope]");
    expect(surface).toContain("selectionScope={controller}");
  });

  it("resets the demo download name from an explicit presentation identity", async () => {
    const shell = await readFile(resolve(componentsRoot, "DemoEditorShell.tsx"), "utf8");
    const surface = await readFile(resolve(componentsRoot, "EditorSurface.tsx"), "utf8");

    expect(surface).toContain("resetScope: controller");
    expect(shell).toContain("[controls.resetScope, fileName]");
  });

  it("keeps navigation guard listeners independent from volatile host-control snapshots", async () => {
    const shell = await readFile(resolve(componentsRoot, "DemoEditorShell.tsx"), "utf8");

    expect(shell).not.toContain("}, [controls]);");
    expect(shell).toContain("[controls.hasUnsavedChanges]");
  });

  it("moves session lifecycle invalidation out of render and scopes transient gestures", async () => {
    const surface = await readFile(resolve(componentsRoot, "EditorSurface.tsx"), "utf8");
    const slideStrip = await readFile(resolve(componentsRoot, "EditorSlideStrip.tsx"), "utf8");

    expect(surface).not.toContain("directTextSessionRef");
    expect(surface).toContain("useLayoutEffect(() => {");
    expect(surface).toContain("committedInteractionScopeRef.current = controller");
    expect(surface).toContain("interactionScope={controller}");
    expect(slideStrip).toContain("readonly interactionScope: object");
    expect(slideStrip).toContain("drag.interactionScope !== interactionScope");
  });
});
