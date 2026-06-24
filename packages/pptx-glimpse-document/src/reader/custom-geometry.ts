import type { SourceCustomGeometryPath } from "../source/index.js";
import {
  getAttr,
  getChild,
  getChildArray,
  localName,
  type XmlNode,
  type XmlOrderedNode,
} from "./xml.js";

interface GuideDefinition {
  readonly name: string;
  readonly formula: string;
}

export function parseCustomGeometry(
  custGeom: XmlNode | undefined,
  orderedCustGeom?: readonly XmlOrderedNode[],
): SourceCustomGeometryPath[] | undefined {
  const pathLst = getChild(custGeom, "pathLst");
  const paths = getChildArray(pathLst, "path");
  if (paths.length === 0) return undefined;

  const avGuides = parseGuideList(getChild(custGeom, "avLst"));
  const guides = parseGuideList(getChild(custGeom, "gdLst"));
  const orderedPaths = orderedChildEntries(
    orderedChildChildren(orderedCustGeom, "pathLst"),
    "path",
  );
  const result: SourceCustomGeometryPath[] = [];

  for (const [index, path] of paths.entries()) {
    const width = numericAttr(path, "w") ?? 0;
    const height = numericAttr(path, "h") ?? 0;
    if (width === 0 && height === 0) continue;

    const variables = evaluateGuides(avGuides, guides, width, height);
    const commands = buildPathCommands(path, variables, orderedNodeChildren(orderedPaths[index]));
    if (commands !== undefined) result.push({ width, height, commands });
  }

  return result.length > 0 ? result : undefined;
}

function parseGuideList(parent: XmlNode | undefined): GuideDefinition[] {
  return getChildArray(parent, "gd")
    .map((guide) => ({
      name: getAttr(guide, "name") ?? "",
      formula: getAttr(guide, "fmla") ?? "",
    }))
    .filter((guide) => guide.name !== "" && guide.formula !== "");
}

function buildPathCommands(
  path: XmlNode,
  variables: Record<string, number>,
  orderedCommands?: readonly XmlOrderedNode[],
): string | undefined {
  const parts: string[] = [];
  let currentX = 0;
  let currentY = 0;
  let startX = 0;
  let startY = 0;

  for (const { local, nodes } of pathCommandNodes(path, orderedCommands)) {
    for (const node of nodes) {
      if (local === "moveTo") {
        const point = firstPoint(node, variables);
        if (point !== undefined) {
          parts.push(`M ${point.x} ${point.y}`);
          currentX = point.x;
          currentY = point.y;
          startX = point.x;
          startY = point.y;
        }
      } else if (local === "lnTo") {
        const point = firstPoint(node, variables);
        if (point !== undefined) {
          parts.push(`L ${point.x} ${point.y}`);
          currentX = point.x;
          currentY = point.y;
        }
      } else if (local === "cubicBezTo") {
        const points = allPoints(node, variables);
        if (points.length >= 3) {
          parts.push(`C ${points.map((point) => `${point.x} ${point.y}`).join(", ")}`);
          currentX = points[points.length - 1].x;
          currentY = points[points.length - 1].y;
        }
      } else if (local === "quadBezTo") {
        const points = allPoints(node, variables);
        if (points.length >= 2) {
          parts.push(`Q ${points.map((point) => `${point.x} ${point.y}`).join(", ")}`);
          currentX = points[points.length - 1].x;
          currentY = points[points.length - 1].y;
        }
      } else if (local === "arcTo") {
        const arc = convertArcTo(node, currentX, currentY, variables);
        if (arc !== undefined) {
          parts.push(arc.command);
          currentX = arc.endX;
          currentY = arc.endY;
        }
      } else if (local === "close") {
        parts.push("Z");
        currentX = startX;
        currentY = startY;
      }
    }
  }

  return parts.length > 0 ? parts.join(" ") : undefined;
}

function pathCommandNodes(
  path: XmlNode,
  orderedCommands: readonly XmlOrderedNode[] | undefined,
): Array<{ readonly local: string; readonly nodes: readonly XmlNode[] }> {
  if (orderedCommands === undefined) {
    return Object.keys(path)
      .filter((key) => !key.startsWith("@_"))
      .map((key) => ({ local: localName(key), nodes: getChildArray(path, localName(key)) }));
  }

  const counters: Record<string, number> = {};
  const result: Array<{ readonly local: string; readonly nodes: readonly XmlNode[] }> = [];
  for (const command of orderedCommands) {
    const key = orderedNodeKey(command);
    if (key === undefined) continue;
    const local = localName(key);
    const index = counters[local] ?? 0;
    counters[local] = index + 1;
    const node = getChildArray(path, local)[index];
    if (node !== undefined) result.push({ local, nodes: [node] });
  }
  return result;
}

function firstPoint(
  node: XmlNode,
  variables: Record<string, number>,
): { readonly x: number; readonly y: number } | undefined {
  return allPoints(node, variables)[0];
}

function allPoints(
  node: XmlNode,
  variables: Record<string, number>,
): Array<{ readonly x: number; readonly y: number }> {
  return getChildArray(node, "pt").map((point) => ({
    x: resolveValue(getAttr(point, "x") ?? "0", variables),
    y: resolveValue(getAttr(point, "y") ?? "0", variables),
  }));
}

function convertArcTo(
  arc: XmlNode,
  currentX: number,
  currentY: number,
  variables: Record<string, number>,
):
  | {
      readonly command: string;
      readonly endX: number;
      readonly endY: number;
    }
  | undefined {
  const widthRadius = resolveValue(getAttr(arc, "wR") ?? "0", variables);
  const heightRadius = resolveValue(getAttr(arc, "hR") ?? "0", variables);
  const startAngle = resolveValue(getAttr(arc, "stAng") ?? "0", variables);
  const sweepAngle = resolveValue(getAttr(arc, "swAng") ?? "0", variables);
  if ((widthRadius === 0 && heightRadius === 0) || sweepAngle === 0) return undefined;

  const startRadians = (startAngle / 60000) * (Math.PI / 180);
  const endRadians = ((startAngle + sweepAngle) / 60000) * (Math.PI / 180);
  const centerX = currentX - widthRadius * Math.cos(startRadians);
  const centerY = currentY - heightRadius * Math.sin(startRadians);
  const endX = centerX + widthRadius * Math.cos(endRadians);
  const endY = centerY + heightRadius * Math.sin(endRadians);
  const largeArcFlag = Math.abs(sweepAngle / 60000) > 180 ? 1 : 0;
  const sweepFlag = sweepAngle > 0 ? 1 : 0;

  return {
    command: `A ${round(widthRadius)} ${round(heightRadius)} 0 ${largeArcFlag} ${sweepFlag} ${round(endX)} ${round(endY)}`,
    endX,
    endY,
  };
}

function evaluateGuides(
  avGuides: readonly GuideDefinition[],
  guides: readonly GuideDefinition[],
  width: number,
  height: number,
): Record<string, number> {
  const variables = createBuiltinVariables(width, height);
  for (const guide of avGuides) variables[guide.name] = evaluateFormula(guide.formula, variables);
  for (const guide of guides) variables[guide.name] = evaluateFormula(guide.formula, variables);
  return variables;
}

function createBuiltinVariables(width: number, height: number): Record<string, number> {
  return {
    w: width,
    h: height,
    l: 0,
    t: 0,
    r: width,
    b: height,
    wd2: width / 2,
    hd2: height / 2,
    wd4: width / 4,
    hd4: height / 4,
    ss: Math.min(width, height),
    ls: Math.max(width, height),
    cd2: 10800000,
    cd4: 5400000,
    cd8: 2700000,
    "3cd4": 16200000,
  };
}

function evaluateFormula(formula: string, variables: Record<string, number>): number {
  const tokens = formula.trim().split(/\s+/);
  const op = tokens[0];
  const resolve = (token: string | undefined): number => {
    if (token === undefined) return 0;
    const value = Number(token);
    return Number.isNaN(value) ? (variables[token] ?? 0) : value;
  };

  if (op === "val") return resolve(tokens[1]);
  if (op === "+-") return resolve(tokens[1]) + resolve(tokens[2]) - resolve(tokens[3]);
  if (op === "*/")
    return Math.round((resolve(tokens[1]) * resolve(tokens[2])) / (resolve(tokens[3]) || 1));
  if (op === "+/")
    return Math.round((resolve(tokens[1]) + resolve(tokens[2])) / (resolve(tokens[3]) || 1));
  if (op === "pin")
    return Math.max(resolve(tokens[1]), Math.min(resolve(tokens[2]), resolve(tokens[3])));
  if (op === "min") return Math.min(resolve(tokens[1]), resolve(tokens[2]));
  if (op === "max") return Math.max(resolve(tokens[1]), resolve(tokens[2]));
  if (op === "abs") return Math.abs(resolve(tokens[1]));
  if (op === "sqrt") return Math.round(Math.sqrt(resolve(tokens[1])));
  if (op === "sin") return Math.round(resolve(tokens[1]) * Math.sin(toRadians(resolve(tokens[2]))));
  if (op === "cos") return Math.round(resolve(tokens[1]) * Math.cos(toRadians(resolve(tokens[2]))));
  if (op === "tan") return Math.round(resolve(tokens[1]) * Math.tan(toRadians(resolve(tokens[2]))));
  if (op === "at2")
    return Math.round(Math.atan2(resolve(tokens[2]), resolve(tokens[1])) * (180 / Math.PI) * 60000);
  if (op === "mod") {
    const a = resolve(tokens[1]);
    const b = resolve(tokens[2]);
    const c = resolve(tokens[3]);
    return Math.round(Math.sqrt(a * a + b * b + c * c));
  }
  if (op === "cat2") {
    return Math.round(
      resolve(tokens[1]) * Math.cos(Math.atan2(resolve(tokens[3]), resolve(tokens[2]))),
    );
  }
  if (op === "sat2") {
    return Math.round(
      resolve(tokens[1]) * Math.sin(Math.atan2(resolve(tokens[3]), resolve(tokens[2]))),
    );
  }
  if (op === "?:") return resolve(tokens[1]) > 0 ? resolve(tokens[2]) : resolve(tokens[3]);
  return 0;
}

function resolveValue(value: string, variables: Record<string, number>): number {
  const numeric = Number(value);
  return Number.isNaN(numeric) ? (variables[value] ?? 0) : numeric;
}

function numericAttr(node: XmlNode | undefined, name: string): number | undefined {
  const raw = getAttr(node, name);
  if (raw === undefined) return undefined;
  const value = Number(raw);
  return Number.isFinite(value) ? value : undefined;
}

function round(value: number): number {
  return Math.round(value * 1000) / 1000;
}

function toRadians(ooxmlAngle: number): number {
  return (ooxmlAngle / 60000) * (Math.PI / 180);
}

function orderedChildChildren(
  parent: readonly XmlOrderedNode[] | undefined,
  childLocalName: string,
): readonly XmlOrderedNode[] | undefined {
  const child = orderedChildEntries(parent, childLocalName)[0];
  return orderedNodeChildren(child);
}

function orderedChildEntries(
  parent: readonly XmlOrderedNode[] | undefined,
  childLocalName: string,
): readonly XmlOrderedNode[] {
  if (parent === undefined) return [];
  return parent.filter((child) => {
    const key = orderedNodeKey(child);
    return key !== undefined && localName(key) === childLocalName;
  });
}

function orderedNodeChildren(
  node: XmlOrderedNode | undefined,
): readonly XmlOrderedNode[] | undefined {
  const key = node !== undefined ? orderedNodeKey(node) : undefined;
  const value = key !== undefined ? node?.[key] : undefined;
  return Array.isArray(value) ? (value as readonly XmlOrderedNode[]) : undefined;
}

function orderedNodeKey(node: XmlOrderedNode): string | undefined {
  return Object.keys(node).find((key) => key !== ":@");
}
