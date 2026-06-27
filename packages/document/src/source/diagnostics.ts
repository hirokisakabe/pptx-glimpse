/**
 * Diagnostics の型。document の正しさ (round-trip / preservation / 参照解決) に
 * 関する診断であり、rendering fidelity や UI 挙動の警告ではない
 * (`docs/document-boundaries.md`)。
 */

import type { SourceHandle } from "./handles.js";

export type DiagnosticSeverity = "info" | "warning" | "error";

export interface Diagnostic {
  readonly severity: DiagnosticSeverity;
  /** 機械判別用の安定コード (例: `unsupported-node-dropped`)。 */
  readonly code: string;
  readonly message: string;
  /** 診断の発生元 source node (特定できる場合)。 */
  readonly handle?: SourceHandle;
}
