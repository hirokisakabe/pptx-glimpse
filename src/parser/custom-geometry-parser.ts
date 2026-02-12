import type { CustomGeometryPath } from "../model/shape.js";
import { evaluateGuides, resolveValue, type GuideDefinition } from "./geometry-formula.js";

/** custGeom ノードをパースして CustomGeometryPath[] を返す */
// eslint-disable-next-line @typescript-eslint/no-explicit-any
export function parseCustomGeometry(custGeom: any): CustomGeometryPath[] | null {
  const pathLst = custGeom.pathLst;
  if (!pathLst?.path) return null;

  const avLst = parseGuideList(custGeom.avLst?.gd);
  const gdLst = parseGuideList(custGeom.gdLst?.gd);

  const paths = Array.isArray(pathLst.path) ? pathLst.path : [pathLst.path];
  const result: CustomGeometryPath[] = [];

  for (const path of paths) {
    const w = Number(path["@_w"] ?? 0);
    const h = Number(path["@_h"] ?? 0);
    if (w === 0 && h === 0) continue;

    const vars = evaluateGuides(avLst, gdLst, w, h);
    const commands = buildPathCommands(path, vars);
    if (!commands) continue;

    result.push({ width: w, height: h, commands });
  }

  return result.length > 0 ? result : null;
}

// eslint-disable-next-line @typescript-eslint/no-explicit-any
function parseGuideList(gd: any): GuideDefinition[] {
  if (!gd) return [];
  const list = Array.isArray(gd) ? gd : [gd];
  return list
    .map((g: Record<string, string>) => ({
      name: g["@_name"] ?? "",
      fmla: g["@_fmla"] ?? "",
    }))
    .filter((g: GuideDefinition) => g.name && g.fmla);
}

/** パスオブジェクトからSVGパスコマンド文字列を生成 */
// eslint-disable-next-line @typescript-eslint/no-explicit-any
function buildPathCommands(path: any, vars: Record<string, number>): string | null {
  const parts: string[] = [];
  let curX = 0;
  let curY = 0;
  let startX = 0;
  let startY = 0;

  for (const key of Object.keys(path)) {
    if (key.startsWith("@_")) continue;

    const value = path[key];
    const items = Array.isArray(value) ? value : [value];

    for (const item of items) {
      switch (key) {
        case "moveTo": {
          const pt = resolveFirstPoint(item.pt, vars);
          if (pt) {
            parts.push(`M ${pt.x} ${pt.y}`);
            curX = pt.x;
            curY = pt.y;
            startX = pt.x;
            startY = pt.y;
          }
          break;
        }
        case "lnTo": {
          const pt = resolveFirstPoint(item.pt, vars);
          if (pt) {
            parts.push(`L ${pt.x} ${pt.y}`);
            curX = pt.x;
            curY = pt.y;
          }
          break;
        }
        case "cubicBezTo": {
          const pts = resolvePoints(item.pt, vars);
          if (pts.length >= 3) {
            parts.push(`C ${pts.map((p) => `${p.x} ${p.y}`).join(", ")}`);
            curX = pts[pts.length - 1].x;
            curY = pts[pts.length - 1].y;
          }
          break;
        }
        case "quadBezTo": {
          const pts = resolvePoints(item.pt, vars);
          if (pts.length >= 2) {
            parts.push(`Q ${pts.map((p) => `${p.x} ${p.y}`).join(", ")}`);
            curX = pts[pts.length - 1].x;
            curY = pts[pts.length - 1].y;
          }
          break;
        }
        case "arcTo": {
          const arcResult = convertArcTo(item, curX, curY, vars);
          if (arcResult) {
            parts.push(arcResult.svg);
            curX = arcResult.endX;
            curY = arcResult.endY;
          }
          break;
        }
        case "close": {
          parts.push("Z");
          curX = startX;
          curY = startY;
          break;
        }
      }
    }
  }

  return parts.length > 0 ? parts.join(" ") : null;
}

function resolveFirstPoint(
  // eslint-disable-next-line @typescript-eslint/no-explicit-any
  pt: any,
  vars: Record<string, number>,
): { x: number; y: number } | null {
  if (!pt) return null;
  const p = Array.isArray(pt) ? pt[0] : pt;
  if (!p) return null;
  return {
    x: resolveValue(p["@_x"], vars),
    y: resolveValue(p["@_y"], vars),
  };
}

function resolvePoints(
  // eslint-disable-next-line @typescript-eslint/no-explicit-any
  pt: any,
  vars: Record<string, number>,
): Array<{ x: number; y: number }> {
  if (!pt) return [];
  const pts = Array.isArray(pt) ? pt : [pt];
  return pts.map((p: Record<string, string>) => ({
    x: resolveValue(p["@_x"], vars),
    y: resolveValue(p["@_y"], vars),
  }));
}

/** arcTo を SVG A コマンドに変換 */
function convertArcTo(
  // eslint-disable-next-line @typescript-eslint/no-explicit-any
  arc: any,
  curX: number,
  curY: number,
  vars: Record<string, number>,
): { svg: string; endX: number; endY: number } | null {
  const wR = resolveValue(arc["@_wR"] ?? "0", vars);
  const hR = resolveValue(arc["@_hR"] ?? "0", vars);
  const stAng = resolveValue(arc["@_stAng"] ?? "0", vars);
  const swAng = resolveValue(arc["@_swAng"] ?? "0", vars);

  if (wR === 0 && hR === 0) return null;
  if (swAng === 0) return null;

  const stAngDeg = stAng / 60000;
  const swAngDeg = swAng / 60000;
  const stAngRad = (stAngDeg * Math.PI) / 180;
  const endAngRad = ((stAngDeg + swAngDeg) * Math.PI) / 180;

  // 楕円の中心を逆算 (現在位置は楕円上の stAng の点)
  const cx = curX - wR * Math.cos(stAngRad);
  const cy = curY - hR * Math.sin(stAngRad);

  // 終点を計算
  const endX = cx + wR * Math.cos(endAngRad);
  const endY = cy + hR * Math.sin(endAngRad);

  // SVG arc フラグ
  const largeArcFlag = Math.abs(swAngDeg) > 180 ? 1 : 0;
  const sweepFlag = swAngDeg > 0 ? 1 : 0;

  // 小数点以下の丸め
  const rx = Math.round(wR * 1000) / 1000;
  const ry = Math.round(hR * 1000) / 1000;
  const ex = Math.round(endX * 1000) / 1000;
  const ey = Math.round(endY * 1000) / 1000;

  return {
    svg: `A ${rx} ${ry} 0 ${largeArcFlag} ${sweepFlag} ${ex} ${ey}`,
    endX,
    endY,
  };
}
