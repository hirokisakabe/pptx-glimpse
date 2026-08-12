import { describe, expect, it } from "vitest";

import { LayoutPickerAddLifecycle } from "./layout-picker-add-lifecycle.js";

describe("LayoutPickerAddLifecycle", () => {
  it("isolates a deferred add completion after the interaction scope changes", async () => {
    const oldScope = {};
    const replacementScope = {};
    const lifecycle = new LayoutPickerAddLifecycle(oldScope);
    const oldCompletion = lifecycle.capture();
    let resolveOldAdd: (added: boolean) => void = () => {};
    const oldAdd = new Promise<boolean>((resolve) => {
      resolveOldAdd = resolve;
    });
    const replacementPicker = { adding: true, open: true, focused: false };

    lifecycle.invalidate();
    lifecycle.activate(replacementScope);
    resolveOldAdd(true);
    const added = await oldAdd;
    if (lifecycle.isCurrent(oldCompletion)) {
      replacementPicker.adding = false;
      replacementPicker.open = !added;
      replacementPicker.focused = added;
    }

    expect(replacementPicker).toEqual({ adding: true, open: true, focused: false });
    expect(lifecycle.isCurrent(oldCompletion)).toBe(false);
  });
});
