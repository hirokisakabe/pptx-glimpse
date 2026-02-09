import type { TextBody } from "../model/text.js";
import type { Transform } from "../model/shape.js";
import { emuToPixels } from "../utils/emu.js";

export function renderTextBody(textBody: TextBody, transform: Transform): string {
  const { bodyProperties, paragraphs } = textBody;
  const width = emuToPixels(transform.extentWidth);
  const height = emuToPixels(transform.extentHeight);
  const marginLeft = emuToPixels(bodyProperties.marginLeft);
  const marginRight = emuToPixels(bodyProperties.marginRight);
  const marginTop = emuToPixels(bodyProperties.marginTop);
  const _marginBottom = emuToPixels(bodyProperties.marginBottom);

  const hasText = paragraphs.some((p) => p.runs.some((r) => r.text.length > 0));
  if (!hasText) return "";

  const textWidth = width - marginLeft - marginRight;

  let anchorAttr: string;
  let xPos: number;
  switch (getEffectiveAlignment(paragraphs)) {
    case "ctr":
      anchorAttr = `text-anchor="middle"`;
      xPos = marginLeft + textWidth / 2;
      break;
    case "r":
      anchorAttr = `text-anchor="end"`;
      xPos = width - marginRight;
      break;
    default:
      anchorAttr = `text-anchor="start"`;
      xPos = marginLeft;
      break;
  }

  const tspans: string[] = [];
  let firstLine = true;

  for (const para of paragraphs) {
    if (para.runs.length === 0) {
      tspans.push(`<tspan x="${xPos}" dy="1.2em"> </tspan>`);
      firstLine = false;
      continue;
    }

    for (let i = 0; i < para.runs.length; i++) {
      const run = para.runs[i];
      if (run.text.length === 0) continue;

      const props = run.properties;
      const styles: string[] = [];

      if (props.fontSize) {
        styles.push(`font-size="${props.fontSize}pt"`);
      }
      if (props.fontFamily) {
        styles.push(`font-family="${escapeXml(props.fontFamily)}"`);
      }
      if (props.bold) {
        styles.push(`font-weight="bold"`);
      }
      if (props.italic) {
        styles.push(`font-style="italic"`);
      }
      if (props.color) {
        styles.push(`fill="${props.color.hex}"`);
        if (props.color.alpha < 1) {
          styles.push(`fill-opacity="${props.color.alpha}"`);
        }
      }

      const decorations: string[] = [];
      if (props.underline) decorations.push("underline");
      if (props.strikethrough) decorations.push("line-through");
      if (decorations.length > 0) {
        styles.push(`text-decoration="${decorations.join(" ")}"`);
      }

      const dyAttr = i === 0 ? ` x="${xPos}" dy="${firstLine ? "0" : "1.2em"}"` : "";
      const escapedText = escapeXml(run.text);
      tspans.push(`<tspan${dyAttr} ${styles.join(" ")}>${escapedText}</tspan>`);
    }
    firstLine = false;
  }

  let yStart = marginTop;
  const defaultFontSize = getDefaultFontSize(paragraphs);
  if (bodyProperties.anchor === "ctr") {
    const totalTextHeight = estimateTextHeight(paragraphs, defaultFontSize);
    yStart = Math.max(marginTop, (height - totalTextHeight) / 2);
  } else if (bodyProperties.anchor === "b") {
    const totalTextHeight = estimateTextHeight(paragraphs, defaultFontSize);
    yStart = Math.max(marginTop, height - totalTextHeight - _marginBottom);
  }

  yStart += defaultFontSize;

  return `<text ${anchorAttr} x="${xPos}" y="${yStart}" ${tspans.length > 0 ? "" : ""}>${tspans.join("")}</text>`;
}

function getEffectiveAlignment(paragraphs: TextBody["paragraphs"]): "l" | "ctr" | "r" | "just" {
  for (const p of paragraphs) {
    if (p.runs.some((r) => r.text.length > 0)) {
      return p.properties.alignment;
    }
  }
  return "l";
}

function getDefaultFontSize(paragraphs: TextBody["paragraphs"]): number {
  for (const p of paragraphs) {
    for (const r of p.runs) {
      if (r.properties.fontSize) return r.properties.fontSize;
    }
  }
  return 18;
}

function estimateTextHeight(paragraphs: TextBody["paragraphs"], defaultFontSize: number): number {
  let lines = 0;
  for (const p of paragraphs) {
    lines += Math.max(1, p.runs.length > 0 ? 1 : 0);
  }
  return lines * defaultFontSize * 1.2;
}

function escapeXml(str: string): string {
  return str
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;")
    .replace(/'/g, "&apos;");
}
