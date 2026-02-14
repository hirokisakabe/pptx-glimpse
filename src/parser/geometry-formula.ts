/**
 * OOXML ジオメトリガイド値のフォーミュラエバリュエータ。
 * avLst / gdLst で定義されるフォーミュラを評価し、
 * パス座標値として使われるガイド名を数値に解決する。
 */

export interface GuideDefinition {
  name: string;
  fmla: string;
}

/** 組み込み変数を生成する */
function createBuiltinVariables(w: number, h: number): Record<string, number> {
  return {
    w,
    h,
    l: 0,
    t: 0,
    r: w,
    b: h,
    wd2: w / 2,
    hd2: h / 2,
    wd4: w / 4,
    hd4: h / 4,
    wd5: w / 5,
    hd5: h / 5,
    wd6: w / 6,
    hd6: h / 6,
    wd8: w / 8,
    hd8: h / 8,
    wd10: w / 10,
    hd10: h / 10,
    wd12: w / 12,
    hd12: h / 12,
    wd32: w / 32,
    hd32: h / 32,
    ss: Math.min(w, h),
    ls: Math.max(w, h),
    ssd2: Math.min(w, h) / 2,
    ssd4: Math.min(w, h) / 4,
    ssd6: Math.min(w, h) / 6,
    ssd8: Math.min(w, h) / 8,
    ssd16: Math.min(w, h) / 16,
    ssd32: Math.min(w, h) / 32,
    cd2: 10800000,
    cd4: 5400000,
    cd8: 2700000,
    "3cd4": 16200000,
    "3cd8": 8100000,
    "5cd8": 13500000,
    "7cd8": 18900000,
  };
}

/** フォーミュラ文字列を評価する */
export function evaluateFormula(fmla: string, vars: Record<string, number>): number {
  const tokens = fmla.trim().split(/\s+/);
  const op = tokens[0];

  const resolve = (token: string): number => {
    if (token === undefined) return 0;
    const num = Number(token);
    return Number.isNaN(num) ? (vars[token] ?? 0) : num;
  };

  switch (op) {
    case "val":
      return resolve(tokens[1]);
    case "+-":
      return resolve(tokens[1]) + resolve(tokens[2]) - resolve(tokens[3]);
    case "*/":
      return Math.round((resolve(tokens[1]) * resolve(tokens[2])) / (resolve(tokens[3]) || 1));
    case "+/":
      return Math.round((resolve(tokens[1]) + resolve(tokens[2])) / (resolve(tokens[3]) || 1));
    case "sin": {
      const a = resolve(tokens[1]);
      const b = resolve(tokens[2]);
      return Math.round(a * Math.sin((b / 60000) * (Math.PI / 180)));
    }
    case "cos": {
      const a = resolve(tokens[1]);
      const b = resolve(tokens[2]);
      return Math.round(a * Math.cos((b / 60000) * (Math.PI / 180)));
    }
    case "tan": {
      const a = resolve(tokens[1]);
      const b = resolve(tokens[2]);
      const tanVal = Math.tan((b / 60000) * (Math.PI / 180));
      return Math.round(a * tanVal);
    }
    case "at2": {
      const x = resolve(tokens[1]);
      const y = resolve(tokens[2]);
      return Math.round(Math.atan2(y, x) * (180 / Math.PI) * 60000);
    }
    case "sqrt":
      return Math.round(Math.sqrt(resolve(tokens[1])));
    case "min":
      return Math.min(resolve(tokens[1]), resolve(tokens[2]));
    case "max":
      return Math.max(resolve(tokens[1]), resolve(tokens[2]));
    case "abs":
      return Math.abs(resolve(tokens[1]));
    case "pin": {
      const lo = resolve(tokens[1]);
      const val = resolve(tokens[2]);
      const hi = resolve(tokens[3]);
      return Math.max(lo, Math.min(hi, val));
    }
    case "mod": {
      const a = resolve(tokens[1]);
      const b = resolve(tokens[2]);
      const c = resolve(tokens[3]);
      return Math.round(Math.sqrt(a * a + b * b + c * c));
    }
    case "cat2": {
      const a = resolve(tokens[1]);
      const b = resolve(tokens[2]);
      const c = resolve(tokens[3]);
      return Math.round(a * Math.cos(Math.atan2(c, b)));
    }
    case "sat2": {
      const a = resolve(tokens[1]);
      const b = resolve(tokens[2]);
      const c = resolve(tokens[3]);
      return Math.round(a * Math.sin(Math.atan2(c, b)));
    }
    case "?:":
      return resolve(tokens[1]) > 0 ? resolve(tokens[2]) : resolve(tokens[3]);
    default:
      return 0;
  }
}

/** avLst + gdLst を評価し、全ガイド値を解決する */
export function evaluateGuides(
  avLst: GuideDefinition[],
  gdLst: GuideDefinition[],
  w: number,
  h: number,
): Record<string, number> {
  const vars = createBuiltinVariables(w, h);

  for (const gd of avLst) {
    vars[gd.name] = evaluateFormula(gd.fmla, vars);
  }

  for (const gd of gdLst) {
    vars[gd.name] = evaluateFormula(gd.fmla, vars);
  }

  return vars;
}

/** 値を解決する（数値リテラルまたはガイド名） */
export function resolveValue(val: string | number, vars: Record<string, number>): number {
  if (typeof val === "number") return val;
  const num = Number(val);
  if (!Number.isNaN(num)) return num;
  return vars[val] ?? 0;
}
