import { describe, expect, it } from "vitest";

import { resolveGeneratedVrtCases, resolveSnapshotCases } from "./vrt-cases.js";

describe("resolveSnapshotCases", () => {
  it("resolves selected generated and shared fixture case names", () => {
    expect(resolveSnapshotCases(["shapes", "real-basic-theme"]).map(({ name }) => name)).toEqual([
      "shapes",
      "real-basic-theme",
    ]);
  });

  it("deduplicates selected case names", () => {
    expect(resolveSnapshotCases(["shapes", "shapes"]).map(({ name }) => name)).toEqual(["shapes"]);
  });

  it("points unknown case names to vrt-cases.ts", () => {
    expect(() => resolveSnapshotCases(["missing-case"])).toThrow(
      /Unknown VRT snapshot case name\(s\): "missing-case".*vrt\/snapshot\/vrt-cases\.ts/,
    );
  });
});

describe("resolveGeneratedVrtCases", () => {
  it("returns all generated cases when no case names are specified", () => {
    expect(resolveGeneratedVrtCases([]).length).toBeGreaterThan(1);
  });

  it("filters shared fixture cases out of fixture generation", () => {
    expect(
      resolveGeneratedVrtCases(["shapes", "real-basic-theme"]).map(({ name }) => name),
    ).toEqual(["shapes"]);
  });
});
