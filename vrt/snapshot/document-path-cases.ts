/**
 * Document path VRT は public snapshot を更新せず、current parser path の PNG を
 * 参照画像としてその場で比較する opt-in parity harness。
 */

export const DOCUMENT_PATH_VRT_RENDER_WIDTH = 320;

export const DOCUMENT_PATH_VRT_SNAPSHOT_POLICY =
  "No committed snapshot update is required: document path VRT compares against the current parser path in-memory until the public default path changes.";

export const DOCUMENT_PATH_VRT_OPT_IN_CASES = [
  {
    name: "real-basic-theme",
    fixture: "real-basic-theme.pptx",
    slides: [1],
    tolerance: 0.001,
    reason:
      "Shared fixture already covered by focused document render tests; slide 1 stays within the current CleanDoc shape/text subset.",
    expectedDiagnosticCodes: [
      "cleandoc-adapter.raw-element-skipped",
      "document-render.cjk-font-context-unsupported",
    ],
  },
  {
    name: "real-product-page",
    fixture: "real-product-page.pptx",
    slides: [1],
    tolerance: 0.04,
    reason:
      "Shared fixture exercises the image/text render path and intentionally records the current visual parity gap before default-path migration.",
    expectedDiagnosticCodes: [],
  },
] as const;

export const DOCUMENT_PATH_VRT_EXCLUDED_CASES = [
  {
    name: "real-financial-report",
    fixture: "real-financial-report.pptx",
    reason:
      "Deferred because the fixture is broader than the initial selected shared fixture slice and would mix document path dogfood with unrelated parity gaps.",
  },
  {
    name: "sample",
    fixture: "sample.pptx",
    reason:
      "Deferred until the selected real app fixtures have stable document path parity metrics.",
  },
  {
    name: "sample-issue-387",
    fixture: "sample-issue-387.pptx",
    reason:
      "Deferred because issue-specific regression fixtures should only opt in after their document path support scope is explicitly reviewed.",
  },
  {
    name: "generated snapshot VRT cases",
    fixture: "vrt/snapshot/fixtures/*.pptx",
    reason:
      "Deferred because generated cases cover many unsupported tables, charts, groups, effects, and advanced text features outside this issue's selected fixture scope.",
  },
] as const;
