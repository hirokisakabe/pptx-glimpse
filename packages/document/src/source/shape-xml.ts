/**
 * Edit-time XML finalization for newly added shapes.
 *
 * New-content edits (text box / preset shape / connector / picture additions) finalize their
 * `p:sp` / `p:cxnSp` / `p:pic` XML fragment here at edit time. The serialized fragment stored on the
 * edit is the single source of truth: the in-memory `SourceShapeNode` for the edited
 * model is derived by parsing the fragment with the same reader used for regular
 * PPTX input, and the writer only splices the fragment into the target `p:spTree`
 * without regenerating any shape content.
 */

import { XMLBuilder } from "fast-xml-parser";

import { createSidecarIdFactory } from "../reader/raw-node.js";
import { parseShapeTree } from "../reader/shape-tree.js";
import { parseXml, parseXmlOrdered } from "../reader/xml.js";
import type { PartPath, RelationshipId } from "./handles.js";
import type { ConnectorPresetGeometry } from "./pptx-source-model.js";
import type {
  SourceArrowEndpoint,
  SourceAutoNumScheme,
  SourceDashStyle,
  SourceImageCrop,
  SourceShapeNode,
  SourceTextAlign,
  SourceUnderlineStyle,
  SourceVerticalAnchor,
} from "./shapes.js";
import type { Emu, HundredthPt, OoxmlAngle, OoxmlPercent, Pt } from "./units.js";

export type TextBoxColorInput = { readonly kind: "srgb"; readonly hex: string };

export interface TextBoxGradientStopInput {
  readonly position: OoxmlPercent;
  readonly color: TextBoxColorInput;
}

export interface TextBoxGradientFillInput {
  readonly stops: readonly TextBoxGradientStopInput[];
  readonly angle?: OoxmlAngle;
}

export type TextBoxUnderlineStyle = SourceUnderlineStyle;

export interface TextBoxUnderlineInput {
  readonly style?: TextBoxUnderlineStyle;
  readonly color?: TextBoxColorInput;
}

export type TextBoxBaselineInput =
  | "subscript"
  | "superscript"
  | { readonly type: "percent"; readonly value: OoxmlPercent };

export type TextBoxLineSpacingInput =
  /** Legacy point input retained for compatibility. */
  | HundredthPt
  | { readonly type: "points"; readonly value: HundredthPt }
  | { readonly type: "percent"; readonly value: OoxmlPercent };

interface TextBoxBulletFormattingInput {
  readonly fontFace?: string;
  /** Bullet size as an OOXML percentage, where 100% is `asOoxmlPercent(100000)`. */
  readonly size?: OoxmlPercent;
}

export type TextBoxBulletInput =
  | { readonly type: "none" }
  | ({ readonly type: "character"; readonly character: string } & TextBoxBulletFormattingInput)
  | ({
      readonly type: "auto-number";
      readonly scheme: SourceAutoNumScheme;
      readonly startAt?: number;
    } & TextBoxBulletFormattingInput);

export interface TextBoxGlowInput {
  readonly radius: Emu;
  readonly color: TextBoxColorInput;
}

export type ShapeColorInput = TextBoxColorInput;

export type ShapeGradientFillInput = TextBoxGradientFillInput;

export type ShapeFillInput =
  | { readonly kind: "none" }
  | { readonly kind: "solid"; readonly color: ShapeColorInput }
  | ({ readonly kind: "gradient" } & ShapeGradientFillInput);

export interface ShapeGlowInput {
  readonly radius: Emu;
  readonly color: ShapeColorInput;
}

export interface ShapeEffectsInput {
  readonly glow: ShapeGlowInput;
}

export interface ShapeOutlineInput {
  readonly width?: Emu;
  readonly fill?: ShapeFillInput;
  readonly dash?: SourceDashStyle;
  readonly headEnd?: SourceArrowEndpoint;
  readonly tailEnd?: SourceArrowEndpoint;
}

export interface ShapePresetGeometryInput {
  readonly kind: "preset";
  readonly preset: string;
  readonly adjustValues?: Readonly<Record<string, number>>;
}

export type ShapeCustomGeometryPathCommandInput =
  | { readonly kind: "moveTo"; readonly x: number; readonly y: number }
  | { readonly kind: "lineTo"; readonly x: number; readonly y: number }
  | { readonly kind: "close" };

export interface ShapeCustomGeometryPathInput {
  readonly width: number;
  readonly height: number;
  readonly commands: readonly ShapeCustomGeometryPathCommandInput[];
}

export interface ShapeCustomGeometryInput {
  readonly kind: "custom";
  readonly paths: readonly ShapeCustomGeometryPathInput[];
}

export type ShapeGeometryInput = ShapePresetGeometryInput | ShapeCustomGeometryInput;

export interface TextBoxOutlineInput {
  readonly width?: Emu;
  readonly color?: TextBoxColorInput;
}

export interface TextBoxRunPropertiesInput {
  readonly fontFace?: string;
  readonly fontSize?: Pt;
  readonly color?: TextBoxColorInput;
  readonly gradientFill?: TextBoxGradientFillInput;
  readonly bold?: boolean;
  readonly italic?: boolean;
  readonly underline?: boolean | TextBoxUnderlineInput;
  readonly strike?: boolean;
  readonly baseline?: TextBoxBaselineInput;
  readonly highlight?: TextBoxColorInput;
  readonly glow?: TextBoxGlowInput;
  readonly outline?: TextBoxOutlineInput;
  /** OOXML ST_TextPoint value for `a:rPr@spc` (-400000..400000, 100 = 1 pt). */
  readonly charSpacing?: number;
}

export interface TextBoxRunInput {
  readonly text: string;
  readonly properties?: TextBoxRunPropertiesInput;
  /** External HTTP(S) hyperlink applied to this run. */
  readonly hyperlink?: string;
}

export interface TextBoxParagraphPropertiesInput {
  readonly align?: SourceTextAlign;
  readonly marginLeft?: Emu;
  readonly indent?: Emu;
  readonly lineSpacing?: TextBoxLineSpacingInput;
  readonly bullet?: TextBoxBulletInput;
}

export interface TextBoxParagraphInput {
  readonly runs: readonly TextBoxRunInput[];
  readonly properties?: TextBoxParagraphPropertiesInput;
}

export interface TextBoxBodyPropertiesInput {
  readonly anchor?: SourceVerticalAnchor;
  readonly marginLeft?: Emu;
  readonly marginRight?: Emu;
  readonly marginTop?: Emu;
  readonly marginBottom?: Emu;
  /** Lets the shape grow to fit its text (`a:spAutoFit`). */
  readonly autoFit?: "shape";
}

interface TextBoxXmlParams {
  readonly shapeId: string;
  readonly name: string;
  readonly offsetX: Emu;
  readonly offsetY: Emu;
  readonly width: Emu;
  readonly height: Emu;
  readonly rotation?: OoxmlAngle;
  readonly text?: string;
  readonly paragraphs?: readonly TextBoxParagraphInput[];
  readonly body?: TextBoxBodyPropertiesInput;
  readonly hyperlinkIds?: ReadonlyMap<string, RelationshipId>;
}

interface SlideNumberXmlParams {
  readonly partPath: PartPath;
  readonly shapeId: string;
  readonly name: string;
  readonly offsetX: Emu;
  readonly offsetY: Emu;
  readonly width: Emu;
  readonly height: Emu;
  readonly body?: TextBoxBodyPropertiesInput;
  readonly properties?: TextBoxRunPropertiesInput;
  readonly align?: SourceTextAlign;
}

interface ShapeXmlParams {
  readonly shapeId: string;
  readonly name: string;
  readonly geometry: ShapeGeometryInput;
  readonly offsetX: Emu;
  readonly offsetY: Emu;
  readonly width: Emu;
  readonly height: Emu;
  readonly rotation?: OoxmlAngle;
  readonly flipHorizontal?: boolean;
  readonly flipVertical?: boolean;
  readonly fill?: ShapeFillInput;
  readonly outline?: ShapeOutlineInput;
  readonly effects?: ShapeEffectsInput;
  readonly text?: string;
  readonly paragraphs?: readonly TextBoxParagraphInput[];
  readonly body?: TextBoxBodyPropertiesInput;
  readonly hyperlinkIds?: ReadonlyMap<string, RelationshipId>;
}

interface ConnectorXmlParams {
  readonly shapeId: string;
  readonly name: string;
  readonly preset: ConnectorPresetGeometry;
  readonly offsetX: Emu;
  readonly offsetY: Emu;
  readonly width: Emu;
  readonly height: Emu;
  readonly startShapeId?: string;
  readonly startConnectionSiteIndex?: number;
  readonly endShapeId?: string;
  readonly endConnectionSiteIndex?: number;
  readonly outline?: {
    readonly headEnd?: SourceArrowEndpoint;
    readonly tailEnd?: SourceArrowEndpoint;
  };
}

interface PictureXmlParams {
  readonly shapeId: string;
  readonly name: string;
  readonly relationshipId: RelationshipId;
  readonly offsetX: Emu;
  readonly offsetY: Emu;
  readonly width: Emu;
  readonly height: Emu;
  readonly rotation?: OoxmlAngle;
  readonly crop?: SourceImageCrop;
}

const xmlBuilder = new XMLBuilder({
  ignoreAttributes: false,
  attributeNamePrefix: "@_",
  format: false,
  suppressEmptyNode: true,
});

export function buildTextBoxXml(params: TextBoxXmlParams): string {
  const paragraphs = params.paragraphs ?? [{ runs: [{ text: params.text ?? "" }] }];
  return xmlBuilder.build({
    "p:sp": {
      "p:nvSpPr": {
        "p:cNvPr": {
          "@_id": params.shapeId,
          "@_name": params.name,
        },
        "p:cNvSpPr": {
          "@_txBox": "1",
        },
        "p:nvPr": {},
      },
      "p:spPr": {
        "a:xfrm": {
          ...(params.rotation !== undefined ? { "@_rot": String(params.rotation) } : {}),
          "a:off": {
            "@_x": String(params.offsetX),
            "@_y": String(params.offsetY),
          },
          "a:ext": {
            "@_cx": String(params.width),
            "@_cy": String(params.height),
          },
        },
        "a:prstGeom": {
          "@_prst": "rect",
          "a:avLst": {},
        },
        "a:noFill": {},
        "a:ln": {
          "a:noFill": {},
        },
      },
      "p:txBody": {
        "a:bodyPr": createTextBodyPropertiesXml(params.body),
        "a:lstStyle": {},
        "a:p": paragraphs.map((paragraph) =>
          createParagraphXml(paragraph, params.hyperlinkIds ?? new Map()),
        ),
      },
    },
  });
}

export function buildSlideNumberXml(params: SlideNumberXmlParams): string {
  return xmlBuilder.build({
    "p:sp": {
      "p:nvSpPr": {
        "p:cNvPr": { "@_id": params.shapeId, "@_name": params.name },
        "p:cNvSpPr": { "a:spLocks": { "@_noGrp": "1" } },
        "p:nvPr": {},
      },
      "p:spPr": {
        "a:xfrm": {
          "a:off": { "@_x": String(params.offsetX), "@_y": String(params.offsetY) },
          "a:ext": { "@_cx": String(params.width), "@_cy": String(params.height) },
        },
        "a:prstGeom": { "@_prst": "rect", "a:avLst": {} },
        "a:noFill": {},
        "a:ln": { "a:noFill": {} },
      },
      "p:txBody": {
        "a:bodyPr": createTextBodyPropertiesXml(params.body),
        "a:lstStyle": {},
        "a:p": {
          ...(params.align !== undefined
            ? { "a:pPr": createParagraphPropertiesXml({ align: params.align }) }
            : {}),
          "a:fld": {
            "@_id": slideNumberFieldId(params.partPath, params.shapeId),
            "@_type": "slidenum",
            "a:rPr": {
              "@_lang": "en-US",
              ...(params.properties !== undefined
                ? createTextRunPropertiesXml(params.properties)
                : {}),
            },
            "a:t": "1",
          },
          "a:endParaRPr": { "@_lang": "en-US" },
        },
      },
    },
  });
}

function slideNumberFieldId(partPath: PartPath, shapeId: string): string {
  let partHash = 2166136261;
  for (const character of partPath) {
    partHash ^= character.codePointAt(0) ?? 0;
    partHash = Math.imul(partHash, 16777619);
  }
  const prefix = (partHash >>> 0).toString(16).padStart(8, "0");
  const suffix = Number.parseInt(shapeId, 10).toString(16).slice(-12).padStart(12, "0");
  return `{${prefix}-0000-4000-8000-${suffix}}`;
}

export function buildShapeXml(params: ShapeXmlParams): string {
  const paragraphs =
    params.paragraphs ?? (params.text !== undefined ? [{ runs: [{ text: params.text }] }] : []);
  return xmlBuilder.build({
    "p:sp": {
      "p:nvSpPr": {
        "p:cNvPr": {
          "@_id": params.shapeId,
          "@_name": params.name,
        },
        "p:cNvSpPr": {},
        "p:nvPr": {},
      },
      "p:spPr": {
        "a:xfrm": {
          ...(params.rotation !== undefined ? { "@_rot": String(params.rotation) } : {}),
          ...(params.flipHorizontal ? { "@_flipH": "1" } : {}),
          ...(params.flipVertical ? { "@_flipV": "1" } : {}),
          "a:off": {
            "@_x": String(params.offsetX),
            "@_y": String(params.offsetY),
          },
          "a:ext": {
            "@_cx": String(params.width),
            "@_cy": String(params.height),
          },
        },
        ...createShapeGeometryXml(params.geometry),
        ...createShapeFillChildXml(params.fill),
        ...(params.outline !== undefined ? { "a:ln": createShapeLineXml(params.outline) } : {}),
        ...(params.effects?.glow !== undefined
          ? { "a:effectLst": { "a:glow": createGlowXml(params.effects.glow) } }
          : {}),
      },
      ...(paragraphs.length > 0
        ? {
            "p:txBody": {
              "a:bodyPr": createTextBodyPropertiesXml(params.body),
              "a:lstStyle": {},
              "a:p": paragraphs.map((paragraph) =>
                createParagraphXml(paragraph, params.hyperlinkIds ?? new Map()),
              ),
            },
          }
        : {}),
    },
  });
}

function createShapeGeometryXml(geometry: ShapeGeometryInput): Record<string, unknown> {
  if (geometry.kind === "preset") {
    return {
      "a:prstGeom": {
        "@_prst": geometry.preset,
        "a:avLst": {
          ...(geometry.adjustValues !== undefined
            ? {
                "a:gd": Object.entries(geometry.adjustValues).map(([name, value]) => ({
                  "@_name": name,
                  "@_fmla": `val ${value}`,
                })),
              }
            : {}),
        },
      },
    };
  }

  return {
    "a:custGeom": {
      "a:avLst": {},
      "a:gdLst": {},
      "a:ahLst": {},
      "a:cxnLst": {},
      "a:rect": { "@_l": "l", "@_t": "t", "@_r": "r", "@_b": "b" },
      "a:pathLst": {
        "a:path": geometry.paths.map((path) => ({
          "@_w": String(path.width),
          "@_h": String(path.height),
          "a:moveTo": createCustomGeometryPointXml(path.commands[0]),
          ...(path.commands.some((command) => command.kind === "lineTo")
            ? {
                "a:lnTo": path.commands
                  .filter((command) => command.kind === "lineTo")
                  .map(createCustomGeometryPointXml),
              }
            : {}),
          ...(path.commands.at(-1)?.kind === "close" ? { "a:close": {} } : {}),
        })),
      },
    },
  };
}

function createCustomGeometryPointXml(
  command: ShapeCustomGeometryPathCommandInput | undefined,
): Record<string, unknown> {
  if (command === undefined || command.kind === "close") return {};
  return { "a:pt": { "@_x": String(command.x), "@_y": String(command.y) } };
}

function createTextBodyPropertiesXml(
  body: TextBoxBodyPropertiesInput | undefined,
): Record<string, unknown> {
  return {
    "@_wrap": "square",
    ...(body?.anchor !== undefined ? { "@_anchor": verticalAnchorToken(body.anchor) } : {}),
    ...(body?.marginLeft !== undefined ? { "@_lIns": String(body.marginLeft) } : {}),
    ...(body?.marginRight !== undefined ? { "@_rIns": String(body.marginRight) } : {}),
    ...(body?.marginTop !== undefined ? { "@_tIns": String(body.marginTop) } : {}),
    ...(body?.marginBottom !== undefined ? { "@_bIns": String(body.marginBottom) } : {}),
    ...(body?.autoFit === "shape" ? { "a:spAutoFit": {} } : {}),
  };
}

function createParagraphXml(
  paragraph: TextBoxParagraphInput,
  hyperlinkIds: ReadonlyMap<string, RelationshipId>,
): Record<string, unknown> {
  return {
    ...(paragraph.properties !== undefined
      ? { "a:pPr": createParagraphPropertiesXml(paragraph.properties) }
      : {}),
    "a:r": paragraph.runs.map((run) => createTextRunXml(run, hyperlinkIds)),
    "a:endParaRPr": {},
  };
}

function createParagraphPropertiesXml(
  properties: TextBoxParagraphPropertiesInput,
): Record<string, unknown> {
  return {
    ...(properties.align !== undefined ? { "@_algn": textAlignToken(properties.align) } : {}),
    ...(properties.marginLeft !== undefined ? { "@_marL": String(properties.marginLeft) } : {}),
    ...(properties.indent !== undefined ? { "@_indent": String(properties.indent) } : {}),
    ...(properties.lineSpacing !== undefined
      ? { "a:lnSpc": createLineSpacingXml(properties.lineSpacing) }
      : {}),
    ...(properties.bullet !== undefined ? createBulletXml(properties.bullet) : {}),
  };
}

function createLineSpacingXml(spacing: TextBoxLineSpacingInput): Record<string, unknown> {
  if (typeof spacing === "number") return { "a:spcPts": { "@_val": String(spacing) } };
  return spacing.type === "points"
    ? { "a:spcPts": { "@_val": String(spacing.value) } }
    : { "a:spcPct": { "@_val": String(spacing.value) } };
}

function createBulletXml(bullet: TextBoxBulletInput): Record<string, unknown> {
  if (bullet.type === "none") return { "a:buNone": {} };
  return {
    ...(bullet.size !== undefined ? { "a:buSzPct": { "@_val": String(bullet.size) } } : {}),
    ...(bullet.fontFace !== undefined ? { "a:buFont": { "@_typeface": bullet.fontFace } } : {}),
    ...(bullet.type === "character"
      ? { "a:buChar": { "@_char": bullet.character } }
      : {
          "a:buAutoNum": {
            "@_type": bullet.scheme,
            ...(bullet.startAt !== undefined ? { "@_startAt": String(bullet.startAt) } : {}),
          },
        }),
  };
}

function createTextRunXml(
  run: TextBoxRunInput,
  hyperlinkIds: ReadonlyMap<string, RelationshipId>,
): Record<string, unknown> {
  const hyperlinkId = run.hyperlink === undefined ? undefined : hyperlinkIds.get(run.hyperlink);
  if (run.hyperlink !== undefined && hyperlinkId === undefined) {
    throw new Error("buildTextBoxXml: hyperlink relationship id was not allocated");
  }
  return {
    ...(run.properties !== undefined || hyperlinkId !== undefined
      ? { "a:rPr": createTextRunPropertiesXml(run.properties ?? {}, hyperlinkId) }
      : {}),
    "a:t": textElementValue(run.text),
  };
}

export function createTextRunPropertiesXml(
  properties: TextBoxRunPropertiesInput,
  hyperlinkId?: RelationshipId,
): Record<string, unknown> {
  return {
    ...(properties.bold !== undefined ? { "@_b": boolToken(properties.bold) } : {}),
    ...(properties.italic !== undefined ? { "@_i": boolToken(properties.italic) } : {}),
    ...(properties.underline !== undefined
      ? { "@_u": underlineStyleToken(properties.underline) }
      : {}),
    ...(properties.strike !== undefined
      ? { "@_strike": properties.strike ? "sngStrike" : "noStrike" }
      : {}),
    ...(properties.baseline !== undefined
      ? { "@_baseline": String(baselineToken(properties.baseline)) }
      : {}),
    ...(properties.fontSize !== undefined
      ? { "@_sz": String(Math.round(properties.fontSize * 100)) }
      : {}),
    ...(properties.charSpacing !== undefined
      ? { "@_spc": textPointToken(properties.charSpacing) }
      : {}),
    ...(properties.outline !== undefined
      ? { "a:ln": createTextOutlineXml(properties.outline) }
      : {}),
    ...(properties.color !== undefined
      ? { "a:solidFill": createSolidFillXml(properties.color) }
      : {}),
    ...(properties.gradientFill !== undefined
      ? { "a:gradFill": createGradientFillXml(properties.gradientFill) }
      : {}),
    ...(properties.glow !== undefined
      ? { "a:effectLst": { "a:glow": createGlowXml(properties.glow) } }
      : {}),
    ...(properties.highlight !== undefined
      ? { "a:highlight": createColorXml(properties.highlight) }
      : {}),
    ...(properties.underline !== undefined &&
    typeof properties.underline !== "boolean" &&
    properties.underline.color !== undefined
      ? { "a:uFill": { "a:solidFill": createSolidFillXml(properties.underline.color) } }
      : {}),
    ...(properties.fontFace !== undefined
      ? {
          "a:latin": { "@_typeface": properties.fontFace },
          "a:ea": { "@_typeface": properties.fontFace },
          "a:cs": { "@_typeface": properties.fontFace },
        }
      : {}),
    ...(hyperlinkId !== undefined ? { "a:hlinkClick": { "@_r:id": hyperlinkId } } : {}),
  };
}

function createSolidFillXml(color: TextBoxColorInput): Record<string, unknown> {
  return createColorXml(color);
}

function createColorXml(color: TextBoxColorInput): Record<string, unknown> {
  if (!/^[0-9A-Fa-f]{6}$/.test(color.hex)) {
    throw new Error("buildTextBoxXml: color hex must be a 6-digit RGB value");
  }
  return {
    "a:srgbClr": {
      "@_val": color.hex.toUpperCase(),
    },
  };
}

function createGradientFillXml(fill: TextBoxGradientFillInput): Record<string, unknown> {
  return {
    "a:gsLst": {
      "a:gs": fill.stops.map((stop) => ({
        "@_pos": String(stop.position),
        ...createColorXml(stop.color),
      })),
    },
    "a:lin": {
      "@_ang": String(fill.angle ?? 0),
      "@_scaled": "1",
    },
  };
}

function createShapeFillChildXml(fill: ShapeFillInput | undefined): Record<string, unknown> {
  if (fill === undefined) return {};
  switch (fill.kind) {
    case "none":
      return { "a:noFill": {} };
    case "solid":
      return { "a:solidFill": createSolidFillXml(fill.color) };
    case "gradient":
      return { "a:gradFill": createGradientFillXml(fill) };
  }
}

function createShapeLineXml(outline: ShapeOutlineInput): Record<string, unknown> {
  return {
    ...(outline.width !== undefined ? { "@_w": String(outline.width) } : {}),
    ...createShapeFillChildXml(outline.fill),
    ...(outline.dash !== undefined ? { "a:prstDash": { "@_val": outline.dash } } : {}),
    ...(outline.headEnd !== undefined
      ? { "a:headEnd": createArrowEndpointXml(outline.headEnd) }
      : {}),
    ...(outline.tailEnd !== undefined
      ? { "a:tailEnd": createArrowEndpointXml(outline.tailEnd) }
      : {}),
  };
}

function createTextOutlineXml(outline: TextBoxOutlineInput): Record<string, unknown> {
  return {
    ...(outline.width !== undefined ? { "@_w": String(outline.width) } : {}),
    ...(outline.color !== undefined ? { "a:solidFill": createSolidFillXml(outline.color) } : {}),
  };
}

function createGlowXml(glow: TextBoxGlowInput): Record<string, unknown> {
  return {
    "@_rad": String(glow.radius),
    ...createColorXml(glow.color),
  };
}

function underlineStyleToken(underline: boolean | TextBoxUnderlineInput): TextBoxUnderlineStyle {
  if (typeof underline === "boolean") return underline ? "sng" : "none";
  return underline.style ?? "sng";
}

function baselineToken(baseline: TextBoxBaselineInput): number {
  if (baseline === "superscript") return 30000;
  if (baseline === "subscript") return -25000;
  return baseline.value;
}

function textPointToken(value: number): string {
  if (!Number.isInteger(value) || value < -400000 || value > 400000) {
    throw new Error("buildTextBoxXml: charSpacing must be an integer between -400000 and 400000");
  }
  return String(value);
}

function boolToken(value: boolean): "1" | "0" {
  return value ? "1" : "0";
}

function textAlignToken(align: SourceTextAlign): "l" | "ctr" | "r" | "just" {
  switch (align) {
    case "left":
      return "l";
    case "center":
      return "ctr";
    case "right":
      return "r";
    case "justify":
      return "just";
  }
}

function verticalAnchorToken(anchor: SourceVerticalAnchor): "t" | "ctr" | "b" {
  switch (anchor) {
    case "top":
      return "t";
    case "middle":
      return "ctr";
    case "bottom":
      return "b";
  }
}

export function buildConnectorXml(params: ConnectorXmlParams): string {
  return xmlBuilder.build({
    "p:cxnSp": {
      "p:nvCxnSpPr": {
        "p:cNvPr": {
          "@_id": params.shapeId,
          "@_name": params.name,
        },
        "p:cNvCxnSpPr": createConnectorConnectionXml(params),
        "p:nvPr": {},
      },
      "p:spPr": {
        "a:xfrm": {
          "a:off": {
            "@_x": String(params.offsetX),
            "@_y": String(params.offsetY),
          },
          "a:ext": {
            "@_cx": String(params.width),
            "@_cy": String(params.height),
          },
        },
        "a:prstGeom": {
          "@_prst": params.preset,
          "a:avLst": {},
        },
        "a:ln": createConnectorLineXml(params),
      },
    },
  });
}

export function buildPictureXml(params: PictureXmlParams): string {
  return xmlBuilder.build({
    "p:pic": {
      "p:nvPicPr": {
        "p:cNvPr": {
          "@_id": params.shapeId,
          "@_name": params.name,
        },
        "p:cNvPicPr": {},
        "p:nvPr": {},
      },
      "p:blipFill": {
        "a:blip": {
          "@_r:embed": params.relationshipId,
        },
        ...(params.crop !== undefined
          ? { "a:srcRect": createSourceRectangleXml(params.crop) }
          : {}),
        "a:stretch": {
          "a:fillRect": {},
        },
      },
      "p:spPr": {
        "a:xfrm": {
          ...(params.rotation !== undefined ? { "@_rot": String(params.rotation) } : {}),
          "a:off": {
            "@_x": String(params.offsetX),
            "@_y": String(params.offsetY),
          },
          "a:ext": {
            "@_cx": String(params.width),
            "@_cy": String(params.height),
          },
        },
        "a:prstGeom": {
          "@_prst": "rect",
          "a:avLst": {},
        },
      },
    },
  });
}

function createSourceRectangleXml(crop: SourceImageCrop): Record<string, unknown> {
  return {
    ...(crop.left !== undefined ? { "@_l": String(crop.left) } : {}),
    ...(crop.top !== undefined ? { "@_t": String(crop.top) } : {}),
    ...(crop.right !== undefined ? { "@_r": String(crop.right) } : {}),
    ...(crop.bottom !== undefined ? { "@_b": String(crop.bottom) } : {}),
  };
}

function createConnectorConnectionXml(params: ConnectorXmlParams): Record<string, unknown> {
  return {
    ...(params.startShapeId !== undefined && params.startConnectionSiteIndex !== undefined
      ? {
          "a:stCxn": {
            "@_id": params.startShapeId,
            "@_idx": String(params.startConnectionSiteIndex),
          },
        }
      : {}),
    ...(params.endShapeId !== undefined && params.endConnectionSiteIndex !== undefined
      ? {
          "a:endCxn": {
            "@_id": params.endShapeId,
            "@_idx": String(params.endConnectionSiteIndex),
          },
        }
      : {}),
  };
}

function createConnectorLineXml(params: ConnectorXmlParams): Record<string, unknown> {
  return {
    "a:solidFill": {
      "a:srgbClr": {
        "@_val": "000000",
      },
    },
    ...(params.outline?.headEnd !== undefined
      ? { "a:headEnd": createArrowEndpointXml(params.outline.headEnd) }
      : {}),
    ...(params.outline?.tailEnd !== undefined
      ? { "a:tailEnd": createArrowEndpointXml(params.outline.tailEnd) }
      : {}),
  };
}

function createArrowEndpointXml(endpoint: SourceArrowEndpoint): Record<string, unknown> {
  return {
    "@_type": endpoint.type,
    "@_w": endpoint.width,
    "@_len": endpoint.length,
  };
}

function textElementValue(text: string): unknown {
  return text.startsWith(" ") || text.endsWith(" ")
    ? { "@_xml:space": "preserve", "#text": text }
    : text;
}

/**
 * Derive the edited-model shape node from a finalized shape XML fragment using the
 * regular reader, so the in-memory shape and the written XML cannot drift apart.
 */
export function parseShapeNodeXml(
  xml: string,
  partPath: PartPath,
  orderingSlot: number,
): SourceShapeNode {
  const nodes = parseShapeTree(
    parseXml(xml),
    partPath,
    createSidecarIdFactory(`${partPath}#added-shape-${orderingSlot}`),
    parseXmlOrdered(xml),
  );
  const node = nodes[0];
  if (node === undefined || nodes.length !== 1) {
    throw new Error("parseShapeNodeXml: shape XML fragment must contain exactly one shape node");
  }
  if (node.handle === undefined) return node;
  return { ...node, handle: { ...node.handle, orderingSlot } };
}
