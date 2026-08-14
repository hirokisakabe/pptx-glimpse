import { describe, expect, it } from "vitest";

import { invalidCommandFailure } from "./command-contract.js";

describe("invalid command failure classification", () => {
  const expectedPrefixes = ["batchValidation:"] as const;

  it("converts an Error with an expected prefix", () => {
    const cause = new Error("batchValidation: commands conflict");

    expect(invalidCommandFailure(expectedPrefixes, cause)).toEqual({
      ok: false,
      code: "invalid-command",
      message: cause.message,
      cause,
    });
  });

  it("propagates an Error without an expected prefix", () => {
    const cause = new Error("EditorSession: broken batch invariant");

    expect(() => invalidCommandFailure(expectedPrefixes, cause)).toThrow(cause);
  });

  it("propagates a non-Error thrown value", () => {
    const cause = { reason: "unexpected non-Error" };

    expect(() => invalidCommandFailure(expectedPrefixes, cause)).toThrow(cause);
  });
});
