/**
 * Diagnostics type. Document correctness (round-trip / preservation / reference resolution)
 * and are not warnings about rendering fidelity or UI behavior.
 */

import type { SourceHandle } from "./handles.js";

export type DiagnosticSeverity = "info" | "warning" | "error";

export interface Diagnostic {
  readonly severity: DiagnosticSeverity;
  /** Stable machine-readable code (Example: `unsupported-node-dropped`). */
  readonly code: string;
  readonly message: string;
  /** The source node from which the diagnosis originates, if it can be determined. */
  readonly handle?: SourceHandle;
}
