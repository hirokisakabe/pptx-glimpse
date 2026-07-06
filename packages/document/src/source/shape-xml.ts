/**
 * Edit-time XML finalization for newly added shapes.
 *
 * New-content edits (text box / connector additions) finalize their `p:sp` /
 * `p:cxnSp` XML fragment here at edit time. The serialized fragment stored on the
 * edit is the single source of truth: the in-memory `SourceShapeNode` for the edited
 * model is derived by parsing the fragment with the same reader used for regular
 * PPTX input, and the writer only splices the fragment into the target `p:spTree`
 * without regenerating any shape content.
 */

import { XMLBuilder } from "fast-xml-parser";

import { createSidecarIdFactory } from "../reader/raw-node.js";
import { parseShapeTree } from "../reader/shape-tree.js";
import { parseXml } from "../reader/xml.js";
import type { PartPath } from "./handles.js";
import type { ConnectorPresetGeometry } from "./pptx-source-model.js";
import type { SourceArrowEndpoint, SourceShapeNode } from "./shapes.js";
import type { Emu } from "./units.js";

interface TextBoxXmlParams {
  readonly shapeId: string;
  readonly name: string;
  readonly offsetX: Emu;
  readonly offsetY: Emu;
  readonly width: Emu;
  readonly height: Emu;
  readonly text: string;
}

interface ConnectorXmlParams {
  readonly shapeId: string;
  readonly name: string;
  readonly preset: ConnectorPresetGeometry;
  readonly offsetX: Emu;
  readonly offsetY: Emu;
  readonly width: Emu;
  readonly height: Emu;
  readonly startShapeId: string;
  readonly startConnectionSiteIndex: number;
  readonly endShapeId: string;
  readonly endConnectionSiteIndex: number;
  readonly outline?: {
    readonly headEnd?: SourceArrowEndpoint;
    readonly tailEnd?: SourceArrowEndpoint;
  };
}

const xmlBuilder = new XMLBuilder({
  ignoreAttributes: false,
  attributeNamePrefix: "@_",
  format: false,
  suppressEmptyNode: true,
});

export function buildTextBoxXml(params: TextBoxXmlParams): string {
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
        "a:bodyPr": {
          "@_wrap": "square",
        },
        "a:lstStyle": {},
        "a:p": {
          "a:r": {
            "a:t": textElementValue(params.text),
          },
          "a:endParaRPr": {},
        },
      },
    },
  });
}

export function buildConnectorXml(params: ConnectorXmlParams): string {
  return xmlBuilder.build({
    "p:cxnSp": {
      "p:nvCxnSpPr": {
        "p:cNvPr": {
          "@_id": params.shapeId,
          "@_name": params.name,
        },
        "p:cNvCxnSpPr": {
          "a:stCxn": {
            "@_id": params.startShapeId,
            "@_idx": String(params.startConnectionSiteIndex),
          },
          "a:endCxn": {
            "@_id": params.endShapeId,
            "@_idx": String(params.endConnectionSiteIndex),
          },
        },
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
  );
  const node = nodes[0];
  if (node === undefined || nodes.length !== 1) {
    throw new Error("parseShapeNodeXml: shape XML fragment must contain exactly one shape node");
  }
  if (node.handle === undefined) return node;
  return { ...node, handle: { ...node.handle, orderingSlot } };
}
