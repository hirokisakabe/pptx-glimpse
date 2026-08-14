import {
  type PptxSourceModel,
  type SourceHandle,
  updateBubbleChartData,
  type UpdateBubbleChartDataInput,
  updateChartData,
  type UpdateChartDataInput,
  updateScatterChartData,
  type UpdateScatterChartDataInput,
} from "@pptx-glimpse/document";

import { type ApplyCommandAttempt, attemptCommand } from "../command-contract.js";

/** Update the series data of one chart. @inline */
export interface UpdateChartDataCommand extends UpdateChartDataInput {
  readonly kind: "updateChartData";
  readonly handle: SourceHandle;
}

/** Update the XY series data of one scatter chart. @inline */
export interface UpdateScatterChartDataCommand extends UpdateScatterChartDataInput {
  readonly kind: "updateScatterChartData";
  readonly handle: SourceHandle;
}

/** Update the XYZ series data of one bubble chart. @inline */
export interface UpdateBubbleChartDataCommand extends UpdateBubbleChartDataInput {
  readonly kind: "updateBubbleChartData";
  readonly handle: SourceHandle;
}

export type ChartEditorCommand =
  | UpdateChartDataCommand
  | UpdateScatterChartDataCommand
  | UpdateBubbleChartDataCommand;

const EXPECTED_REJECTION_PREFIXES = [
  "updateChartData:",
  "updateScatterChartData:",
  "updateBubbleChartData:",
] as const;

export function applyChartCommand(
  document: PptxSourceModel,
  command: ChartEditorCommand,
): ApplyCommandAttempt {
  return attemptCommand(EXPECTED_REJECTION_PREFIXES, () => executeChartCommand(document, command));
}

function executeChartCommand(
  document: PptxSourceModel,
  command: ChartEditorCommand,
): PptxSourceModel {
  if (!Array.isArray(command.series) || command.series.length === 0) {
    throw new Error(`${command.kind}: series must be a non-empty array`);
  }
  for (const [index, series] of command.series.entries()) {
    if (!isObject(series)) {
      throw new Error(`${command.kind}: series[${index}] must be an object`);
    }
    if (typeof series.name !== "string") {
      throw new Error(`${command.kind}: series[${index}].name must be a string`);
    }
    switch (command.kind) {
      case "updateChartData":
        if (!Array.isArray(series.categories) || !Array.isArray(series.values)) {
          throw new Error(`updateChartData: series[${index}] categories and values must be arrays`);
        }
        break;
      case "updateScatterChartData":
        if (!Array.isArray(series.xValues) || !Array.isArray(series.yValues)) {
          throw new Error(`updateScatterChartData: series[${index}] X and Y values must be arrays`);
        }
        break;
      case "updateBubbleChartData":
        if (
          !Array.isArray(series.xValues) ||
          !Array.isArray(series.yValues) ||
          !Array.isArray(series.bubbleSizes)
        ) {
          throw new Error(
            `updateBubbleChartData: series[${index}] X, Y, and bubble size values must be arrays`,
          );
        }
        break;
    }
  }
  switch (command.kind) {
    case "updateChartData":
      return updateChartData(document, command.handle, command);
    case "updateScatterChartData":
      return updateScatterChartData(document, command.handle, command);
    case "updateBubbleChartData":
      return updateBubbleChartData(document, command.handle, command);
  }
}

function isObject(value: unknown): value is Record<string, unknown> {
  return typeof value === "object" && value !== null;
}
