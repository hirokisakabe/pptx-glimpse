import { describe, expect, it } from "vitest";

import { DirectTextEditorLifecycle } from "./direct-text-editor-lifecycle.js";

describe("DirectTextEditorLifecycle", () => {
  it("releases an IME composition waiter when the editor unmounts", async () => {
    const lifecycle = new DirectTextEditorLifecycle();
    lifecycle.beginComposition();
    const pendingHostSave = (async () => {
      const composition = lifecycle.compositionPromise();
      if (composition === undefined || !(await composition)) return undefined;
      return "saved";
    })();

    lifecycle.invalidate();

    await expect(pendingHostSave).resolves.toBeUndefined();
  });

  it("isolates an old commit completion after the editor session changes", async () => {
    const lifecycle = new DirectTextEditorLifecycle();
    const oldGeneration = lifecycle.currentGeneration();
    let resolveOldCommit: (committed: boolean) => void = () => {};
    const oldCommit = new Promise<boolean>((resolve) => {
      resolveOldCommit = resolve;
    });
    expect(lifecycle.setCommit(oldGeneration, oldCommit)).toBe(true);

    lifecycle.invalidate();
    let activeEditor: "new" | null = "new";
    const newGeneration = lifecycle.currentGeneration();
    const newCommit = new Promise<boolean>(() => {});
    expect(lifecycle.setCommit(newGeneration, newCommit)).toBe(true);

    resolveOldCommit(true);
    const committed = await oldCommit;
    if (committed && lifecycle.isCurrent(oldGeneration)) activeEditor = null;
    lifecycle.clearCommit(oldGeneration);

    expect(activeEditor).toBe("new");
    expect(lifecycle.isCurrent(oldGeneration)).toBe(false);
    expect(lifecycle.currentCommit()).toBe(newCommit);
  });
});
