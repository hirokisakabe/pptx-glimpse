// OOXML DrawingML プリセット図形 (ECMA-376 §20.1.10.56 prst)
// adj 値は 100,000 分率で正規化 (e.g. adj=50000 → 50%)

type GeometryGenerator = (w: number, h: number, adj: Record<string, number>) => string;

// --- Helper functions ---

function regularPolygon(w: number, h: number, sides: number): string {
  const cx = w / 2;
  const cy = h / 2;
  const points = Array.from({ length: sides }, (_, i) => {
    const angle = (Math.PI * 2 * i) / sides - Math.PI / 2;
    return `${cx + cx * Math.cos(angle)},${cy + cy * Math.sin(angle)}`;
  }).join(" ");
  return `<polygon points="${points}"/>`;
}

function starPolygon(w: number, h: number, points: number, innerRatio: number): string {
  const cx = w / 2;
  const cy = h / 2;
  const coords: string[] = [];
  for (let i = 0; i < points * 2; i++) {
    const angle = (Math.PI * 2 * i) / (points * 2) - Math.PI / 2;
    const r = i % 2 === 0 ? 1 : innerRatio;
    coords.push(`${cx + cx * r * Math.cos(angle)},${cy + cy * r * Math.sin(angle)}`);
  }
  return `<polygon points="${coords.join(" ")}"/>`;
}

// OOXML 角度 (1/60,000 度) → ラジアン
function ooxmlAngleToRadians(angle60k: number): number {
  return (angle60k / 60000) * (Math.PI / 180);
}

const presetGeometries: Record<string, GeometryGenerator> = {
  // =====================
  // Basic shapes (existing)
  // =====================

  rect: (w, h) => `<rect width="${w}" height="${h}"/>`,

  ellipse: (w, h) => `<ellipse cx="${w / 2}" cy="${h / 2}" rx="${w / 2}" ry="${h / 2}"/>`,

  roundRect: (w, h, adj) => {
    const r = ((adj["adj"] ?? 16667) / 100000) * Math.min(w, h);
    return `<rect width="${w}" height="${h}" rx="${r}" ry="${r}"/>`;
  },

  triangle: (w, h, adj) => {
    const topX = ((adj["adj"] ?? 50000) / 100000) * w;
    return `<polygon points="${topX},0 ${w},${h} 0,${h}"/>`;
  },

  rtTriangle: (w, h) => `<polygon points="0,0 ${w},${h} 0,${h}"/>`,

  diamond: (w, h) => `<polygon points="${w / 2},0 ${w},${h / 2} ${w / 2},${h} 0,${h / 2}"/>`,

  parallelogram: (w, h, adj) => {
    const offset = ((adj["adj"] ?? 25000) / 100000) * w;
    return `<polygon points="${offset},0 ${w},0 ${w - offset},${h} 0,${h}"/>`;
  },

  trapezoid: (w, h, adj) => {
    const offset = ((adj["adj"] ?? 25000) / 100000) * w;
    return `<polygon points="${offset},0 ${w - offset},0 ${w},${h} 0,${h}"/>`;
  },

  pentagon: (w, h) => regularPolygon(w, h, 5),

  hexagon: (w, h, adj) => {
    const offset = ((adj["adj"] ?? 25000) / 100000) * w;
    return `<polygon points="${offset},0 ${w - offset},0 ${w},${h / 2} ${w - offset},${h} ${offset},${h} 0,${h / 2}"/>`;
  },

  star4: (w, h) => {
    const cx = w / 2;
    const cy = h / 2;
    const ir = 0.38;
    return `<polygon points="${cx},0 ${cx + cx * ir},${cy - cy * ir} ${w},${cy} ${cx + cx * ir},${cy + cy * ir} ${cx},${h} ${cx - cx * ir},${cy + cy * ir} 0,${cy} ${cx - cx * ir},${cy - cy * ir}"/>`;
  },

  star5: (w, h) => starPolygon(w, h, 5, 0.38),

  rightArrow: (w, h, adj) => {
    const headWidth = ((adj["adj1"] ?? 50000) / 100000) * h;
    const headLength = ((adj["adj2"] ?? 50000) / 100000) * w;
    const bodyTop = (h - headWidth) / 2;
    const bodyBottom = h - bodyTop;
    const shaftEnd = w - headLength;
    return `<polygon points="0,${bodyTop} ${shaftEnd},${bodyTop} ${shaftEnd},0 ${w},${h / 2} ${shaftEnd},${h} ${shaftEnd},${bodyBottom} 0,${bodyBottom}"/>`;
  },

  leftArrow: (w, h, adj) => {
    const headWidth = ((adj["adj1"] ?? 50000) / 100000) * h;
    const headLength = ((adj["adj2"] ?? 50000) / 100000) * w;
    const bodyTop = (h - headWidth) / 2;
    const bodyBottom = h - bodyTop;
    return `<polygon points="${headLength},${bodyTop} ${headLength},0 0,${h / 2} ${headLength},${h} ${headLength},${bodyBottom} ${w},${bodyBottom} ${w},${bodyTop}"/>`;
  },

  upArrow: (w, h, adj) => {
    const headWidth = ((adj["adj1"] ?? 50000) / 100000) * w;
    const headLength = ((adj["adj2"] ?? 50000) / 100000) * h;
    const bodyLeft = (w - headWidth) / 2;
    const bodyRight = w - bodyLeft;
    return `<polygon points="${bodyLeft},${headLength} 0,${headLength} ${w / 2},0 ${w},${headLength} ${bodyRight},${headLength} ${bodyRight},${h} ${bodyLeft},${h}"/>`;
  },

  downArrow: (w, h, adj) => {
    const headWidth = ((adj["adj1"] ?? 50000) / 100000) * w;
    const headLength = ((adj["adj2"] ?? 50000) / 100000) * h;
    const bodyLeft = (w - headWidth) / 2;
    const bodyRight = w - bodyLeft;
    const shaftEnd = h - headLength;
    return `<polygon points="${bodyLeft},0 ${bodyRight},0 ${bodyRight},${shaftEnd} ${w},${shaftEnd} ${w / 2},${h} 0,${shaftEnd} ${bodyLeft},${shaftEnd}"/>`;
  },

  line: () => "",

  cloud: (w, h) => `<rect width="${w}" height="${h}" rx="${Math.min(w, h) * 0.15}"/>`,

  heart: (w, h) => {
    const cx = w / 2;
    return `<path d="M ${cx} ${h * 0.35} C ${cx} ${h * 0.1}, 0 0, 0 ${h * 0.35} C 0 ${h * 0.65}, ${cx} ${h * 0.85}, ${cx} ${h} C ${cx} ${h * 0.85}, ${w} ${h * 0.65}, ${w} ${h * 0.35} C ${w} 0, ${cx} ${h * 0.1}, ${cx} ${h * 0.35} Z"/>`;
  },

  // =====================
  // Additional polygons
  // =====================

  heptagon: (w, h) => regularPolygon(w, h, 7),

  octagon: (w, h) => regularPolygon(w, h, 8),

  decagon: (w, h) => regularPolygon(w, h, 10),

  dodecagon: (w, h) => regularPolygon(w, h, 12),

  // =====================
  // Stars
  // =====================

  star6: (w, h) => starPolygon(w, h, 6, 0.5),

  star8: (w, h) => starPolygon(w, h, 8, 0.38),

  star10: (w, h) => starPolygon(w, h, 10, 0.38),

  star12: (w, h) => starPolygon(w, h, 12, 0.38),

  star16: (w, h) => starPolygon(w, h, 16, 0.38),

  star24: (w, h) => starPolygon(w, h, 24, 0.38),

  star32: (w, h) => starPolygon(w, h, 32, 0.38),

  irregularSeal1: (w, h) =>
    `<polygon points="${w * 0.15},${h * 0.35} ${w * 0.27},${h * 0.03} ${w * 0.38},${h * 0.28} ${w * 0.5},0 ${w * 0.6},${h * 0.23} ${w * 0.73},${h * 0.08} ${w * 0.72},${h * 0.35} ${w},${h * 0.35} ${w * 0.78},${h * 0.5} ${w * 0.95},${h * 0.7} ${w * 0.73},${h * 0.65} ${w * 0.65},${h} ${w * 0.5},${h * 0.72} ${w * 0.35},${h * 0.95} ${w * 0.32},${h * 0.65} ${w * 0.05},${h * 0.7} ${w * 0.18},${h * 0.5} 0,${h * 0.35}"/>`,

  irregularSeal2: (w, h) =>
    `<polygon points="${w * 0.1},${h * 0.4} ${w * 0.18},${h * 0.08} ${w * 0.32},${h * 0.3} ${w * 0.45},0 ${w * 0.55},${h * 0.18} ${w * 0.72},${h * 0.05} ${w * 0.68},${h * 0.32} ${w},${h * 0.3} ${w * 0.82},${h * 0.5} ${w * 0.98},${h * 0.68} ${w * 0.75},${h * 0.65} ${w * 0.8},${h * 0.92} ${w * 0.55},${h * 0.75} ${w * 0.42},${h} ${w * 0.38},${h * 0.72} ${w * 0.12},${h * 0.88} ${w * 0.22},${h * 0.6} 0,${h * 0.55}"/>`,

  // =====================
  // Additional arrows
  // =====================

  leftRightArrow: (w, h, adj) => {
    const headW = ((adj["adj1"] ?? 50000) / 100000) * h;
    const headL = ((adj["adj2"] ?? 50000) / 100000) * w;
    const bodyTop = (h - headW) / 2;
    const bodyBot = h - bodyTop;
    return `<polygon points="${headL},${bodyTop} ${headL},0 0,${h / 2} ${headL},${h} ${headL},${bodyBot} ${w - headL},${bodyBot} ${w - headL},${h} ${w},${h / 2} ${w - headL},0 ${w - headL},${bodyTop}"/>`;
  },

  upDownArrow: (w, h, adj) => {
    const headW = ((adj["adj1"] ?? 50000) / 100000) * w;
    const headL = ((adj["adj2"] ?? 50000) / 100000) * h;
    const bodyL = (w - headW) / 2;
    const bodyR = w - bodyL;
    return `<polygon points="${bodyL},${headL} 0,${headL} ${w / 2},0 ${w},${headL} ${bodyR},${headL} ${bodyR},${h - headL} ${w},${h - headL} ${w / 2},${h} 0,${h - headL} ${bodyL},${h - headL}"/>`;
  },

  notchedRightArrow: (w, h, adj) => {
    const headWidth = ((adj["adj1"] ?? 50000) / 100000) * h;
    const headLength = ((adj["adj2"] ?? 50000) / 100000) * w;
    const bodyTop = (h - headWidth) / 2;
    const bodyBottom = h - bodyTop;
    const shaftEnd = w - headLength;
    const notch = headLength * 0.5;
    return `<polygon points="0,${bodyTop} ${shaftEnd},${bodyTop} ${shaftEnd},0 ${w},${h / 2} ${shaftEnd},${h} ${shaftEnd},${bodyBottom} 0,${bodyBottom} ${notch},${h / 2}"/>`;
  },

  stripedRightArrow: (w, h, adj) => {
    const headWidth = ((adj["adj1"] ?? 50000) / 100000) * h;
    const headLength = ((adj["adj2"] ?? 50000) / 100000) * w;
    const bodyTop = (h - headWidth) / 2;
    const bodyBottom = h - bodyTop;
    const shaftEnd = w - headLength;
    const sw = w * 0.05;
    return `<path d="M 0 ${bodyTop} L ${sw} ${bodyTop} L ${sw} ${bodyBottom} L 0 ${bodyBottom} Z M ${sw * 1.5} ${bodyTop} L ${sw * 2.5} ${bodyTop} L ${sw * 2.5} ${bodyBottom} L ${sw * 1.5} ${bodyBottom} Z M ${sw * 3} ${bodyTop} L ${shaftEnd} ${bodyTop} L ${shaftEnd} 0 L ${w} ${h / 2} L ${shaftEnd} ${h} L ${shaftEnd} ${bodyBottom} L ${sw * 3} ${bodyBottom} Z"/>`;
  },

  chevron: (w, h, adj) => {
    const offset = ((adj["adj"] ?? 50000) / 100000) * w;
    return `<polygon points="0,0 ${w - offset},0 ${w},${h / 2} ${w - offset},${h} 0,${h} ${offset},${h / 2}"/>`;
  },

  homePlate: (w, h, adj) => {
    const offset = ((adj["adj"] ?? 50000) / 100000) * w;
    return `<polygon points="0,0 ${w - offset},0 ${w},${h / 2} ${w - offset},${h} 0,${h}"/>`;
  },

  leftRightUpArrow: (w, h, adj) => {
    const headW = ((adj["adj1"] ?? 25000) / 100000) * Math.min(w, h);
    const headL = ((adj["adj2"] ?? 25000) / 100000) * Math.min(w, h);
    const bodyW = ((adj["adj3"] ?? 25000) / 100000) * Math.min(w, h);
    const cx = w / 2;
    const bodyHalf = bodyW / 2;
    const bodyMid = h - headL - bodyW;
    return `<path d="M ${cx} 0 L ${cx + headW / 2} ${headL} L ${cx + bodyHalf} ${headL} L ${cx + bodyHalf} ${bodyMid} L ${w - headL} ${bodyMid} L ${w - headL} ${h / 2 + bodyMid / 2 - headW / 2} L ${w} ${h / 2 + bodyMid / 2} L ${w - headL} ${h / 2 + bodyMid / 2 + headW / 2} L ${w - headL} ${bodyMid + bodyW} L ${cx + bodyHalf} ${bodyMid + bodyW} L ${cx + bodyHalf} ${bodyMid + bodyW} L ${cx - bodyHalf} ${bodyMid + bodyW} L ${cx - bodyHalf} ${bodyMid + bodyW} L ${headL} ${bodyMid + bodyW} L ${headL} ${h / 2 + bodyMid / 2 + headW / 2} L 0 ${h / 2 + bodyMid / 2} L ${headL} ${h / 2 + bodyMid / 2 - headW / 2} L ${headL} ${bodyMid} L ${cx - bodyHalf} ${bodyMid} L ${cx - bodyHalf} ${headL} L ${cx - headW / 2} ${headL} Z"/>`;
  },

  quadArrow: (w, h, adj) => {
    const headW = ((adj["adj1"] ?? 22500) / 100000) * Math.min(w, h);
    const headL = ((adj["adj2"] ?? 22500) / 100000) * Math.min(w, h);
    const bodyW = ((adj["adj3"] ?? 11250) / 100000) * Math.min(w, h);
    const cx = w / 2;
    const cy = h / 2;
    const bh = bodyW / 2;
    return `<path d="M ${cx} 0 L ${cx + headW / 2} ${headL} L ${cx + bh} ${headL} L ${cx + bh} ${cy - bh} L ${w - headL} ${cy - bh} L ${w - headL} ${cy - headW / 2} L ${w} ${cy} L ${w - headL} ${cy + headW / 2} L ${w - headL} ${cy + bh} L ${cx + bh} ${cy + bh} L ${cx + bh} ${h - headL} L ${cx + headW / 2} ${h - headL} L ${cx} ${h} L ${cx - headW / 2} ${h - headL} L ${cx - bh} ${h - headL} L ${cx - bh} ${cy + bh} L ${headL} ${cy + bh} L ${headL} ${cy + headW / 2} L 0 ${cy} L ${headL} ${cy - headW / 2} L ${headL} ${cy - bh} L ${cx - bh} ${cy - bh} L ${cx - bh} ${headL} L ${cx - headW / 2} ${headL} Z"/>`;
  },

  bentArrow: (w, h, adj) => {
    const headW = ((adj["adj1"] ?? 25000) / 100000) * h;
    const headL = ((adj["adj2"] ?? 25000) / 100000) * w;
    const bodyW = ((adj["adj3"] ?? 25000) / 100000) * h;
    const shaftEnd = w - headL;
    const arrowTop = 0;
    const bodyTop = headW / 2 - bodyW / 2;
    const bodyBot = headW / 2 + bodyW / 2;
    return `<polygon points="${shaftEnd},${bodyTop} ${shaftEnd},${arrowTop} ${w},${headW / 2} ${shaftEnd},${headW} ${shaftEnd},${bodyBot} ${bodyW},${bodyBot} ${bodyW},${h} 0,${h} 0,${h - bodyW} ${0},${h - bodyW}"/>`;
  },

  bendUpArrow: (w, h, adj) => {
    const headW = ((adj["adj1"] ?? 25000) / 100000) * w;
    const headL = ((adj["adj2"] ?? 25000) / 100000) * h;
    const bodyW = ((adj["adj3"] ?? 25000) / 100000) * w;
    const cx = w - headW / 2;
    const bodyLeft = cx - bodyW / 2;
    const bodyRight = cx + bodyW / 2;
    return `<polygon points="${cx - headW / 2},${headL} ${cx},0 ${cx + headW / 2},${headL} ${bodyRight},${headL} ${bodyRight},${h - bodyW} ${bodyW},${h - bodyW} ${bodyW},${h} 0,${h} 0,${h - bodyW} ${bodyLeft},${h - bodyW} ${bodyLeft},${headL}"/>`;
  },

  leftUpArrow: (w, h, adj) => {
    const headW = ((adj["adj1"] ?? 25000) / 100000) * Math.min(w, h);
    const headL = ((adj["adj2"] ?? 25000) / 100000) * Math.min(w, h);
    const bodyW = ((adj["adj3"] ?? 25000) / 100000) * Math.min(w, h);
    const bh = bodyW / 2;
    const topCx = w - headW / 2;
    const leftCy = h - headW / 2;
    return `<path d="M ${topCx} 0 L ${topCx + headW / 2} ${headL} L ${topCx + bh} ${headL} L ${topCx + bh} ${leftCy - bh} L ${headL} ${leftCy - bh} L ${headL} ${leftCy - headW / 2} L 0 ${leftCy} L ${headL} ${leftCy + headW / 2} L ${headL} ${leftCy + bh} L ${topCx - bh} ${leftCy + bh} L ${topCx - bh} ${headL} L ${topCx - headW / 2} ${headL} Z"/>`;
  },

  uturnArrow: (w, h, adj) => {
    const headW = ((adj["adj1"] ?? 25000) / 100000) * w;
    const headL = ((adj["adj2"] ?? 25000) / 100000) * h;
    const bodyW = ((adj["adj3"] ?? 25000) / 100000) * w;
    const arcR = w * 0.35;
    const cx = w / 2;
    const arrowBot = h;
    const arrowStart = h - headL;
    const bodyRight = w - (headW / 2 - bodyW / 2);
    const bodyLeft = w - (headW / 2 + bodyW / 2);
    return `<path d="M ${w - headW / 2 - headW / 2} ${arrowStart} L ${w - headW / 2} ${arrowBot} L ${w} ${arrowStart} L ${bodyRight} ${arrowStart} L ${bodyRight} ${arcR} A ${arcR} ${arcR} 0 0 0 ${bodyW} ${arcR} L ${bodyW} ${arrowStart} L 0 ${arrowStart} L 0 ${arcR} A ${cx} ${cx} 0 0 1 ${bodyLeft} ${arcR} L ${bodyLeft} ${arrowStart} Z"/>`;
  },

  // =====================
  // Flowchart shapes
  // =====================

  flowChartProcess: (w, h) => `<rect width="${w}" height="${h}"/>`,

  flowChartAlternateProcess: (w, h) => {
    const r = Math.min(w, h) / 6;
    return `<rect width="${w}" height="${h}" rx="${r}" ry="${r}"/>`;
  },

  flowChartDecision: (w, h) =>
    `<polygon points="${w / 2},0 ${w},${h / 2} ${w / 2},${h} 0,${h / 2}"/>`,

  flowChartInputOutput: (w, h) => {
    const offset = w / 5;
    return `<polygon points="${offset},0 ${w},0 ${w - offset},${h} 0,${h}"/>`;
  },

  flowChartPredefinedProcess: (w, h) => {
    const d = w / 8;
    return `<path d="M 0 0 L ${w} 0 L ${w} ${h} L 0 ${h} Z M ${d} 0 L ${d} ${h} M ${w - d} 0 L ${w - d} ${h}"/>`;
  },

  flowChartInternalStorage: (w, h) => {
    const dx = w / 8;
    const dy = h / 8;
    return `<path d="M 0 0 L ${w} 0 L ${w} ${h} L 0 ${h} Z M ${dx} 0 L ${dx} ${h} M 0 ${dy} L ${w} ${dy}"/>`;
  },

  flowChartDocument: (w, h) => {
    const bh = h * 0.83;
    return `<path d="M 0 0 L ${w} 0 L ${w} ${bh} C ${w * 0.75} ${h}, ${w * 0.25} ${h * 0.66}, 0 ${bh} Z"/>`;
  },

  flowChartMultidocument: (w, h) => {
    const dx = w * 0.1;
    const dy = h * 0.1;
    const bw = w - dx;
    const bh = (h - dy) * 0.83;
    return `<path d="M ${dx} ${dy} L ${w} ${dy} L ${w} ${dy + bh} C ${w - bw * 0.25} ${dy + (h - dy)}, ${dx + bw * 0.25} ${dy + (h - dy) * 0.66}, ${dx} ${dy + bh} Z M ${dx / 2} ${dy / 2} L ${dx} ${dy / 2} L ${dx} ${dy} M 0 0 L ${dx / 2} 0 L ${dx / 2} ${dy / 2}"/>`;
  },

  flowChartTerminator: (w, h) => {
    const r = h / 2;
    return `<rect width="${w}" height="${h}" rx="${r}" ry="${r}"/>`;
  },

  flowChartPreparation: (w, h) => {
    const offset = w / 5;
    return `<polygon points="${offset},0 ${w - offset},0 ${w},${h / 2} ${w - offset},${h} ${offset},${h} 0,${h / 2}"/>`;
  },

  flowChartManualInput: (w, h) => {
    const topY = h / 5;
    return `<polygon points="0,${topY} ${w},0 ${w},${h} 0,${h}"/>`;
  },

  flowChartManualOperation: (w, h) => {
    const offset = w / 5;
    return `<polygon points="0,0 ${w},0 ${w - offset},${h} ${offset},${h}"/>`;
  },

  flowChartConnector: (w, h) =>
    `<ellipse cx="${w / 2}" cy="${h / 2}" rx="${w / 2}" ry="${h / 2}"/>`,

  flowChartOffpageConnector: (w, h) => {
    const arrowH = h * 0.2;
    return `<polygon points="0,0 ${w},0 ${w},${h - arrowH} ${w / 2},${h} 0,${h - arrowH}"/>`;
  },

  flowChartPunchedCard: (w, h) => {
    const cut = Math.min(w, h) * 0.2;
    return `<polygon points="${cut},0 ${w},0 ${w},${h} 0,${h} 0,${cut}"/>`;
  },

  flowChartPunchedTape: (w, h) => {
    const wave = h * 0.1;
    return `<path d="M 0 ${wave} C ${w * 0.25} ${-wave}, ${w * 0.75} ${wave * 3}, ${w} ${wave} L ${w} ${h - wave} C ${w * 0.75} ${h + wave}, ${w * 0.25} ${h - wave * 3}, 0 ${h - wave} Z"/>`;
  },

  flowChartCollate: (w, h) =>
    `<polygon points="0,0 ${w},0 ${w / 2},${h / 2} ${w},${h} 0,${h} ${w / 2},${h / 2}"/>`,

  flowChartSort: (w, h) =>
    `<path d="M ${w / 2} 0 L ${w} ${h / 2} L 0 ${h / 2} Z M 0 ${h / 2} L ${w} ${h / 2} M ${w / 2} ${h} L ${w} ${h / 2} L 0 ${h / 2} Z"/>`,

  flowChartExtract: (w, h) => `<polygon points="${w / 2},0 ${w},${h} 0,${h}"/>`,

  flowChartMerge: (w, h) => `<polygon points="0,0 ${w},0 ${w / 2},${h}"/>`,

  flowChartOnlineStorage: (w, h) => {
    const arcW = w * 0.15;
    return `<path d="M ${arcW} 0 L ${w} 0 L ${w} ${h} L ${arcW} ${h} A ${arcW} ${h / 2} 0 0 1 ${arcW} 0 Z"/>`;
  },

  flowChartDelay: (w, h) => {
    const arcW = w * 0.35;
    return `<path d="M 0 0 L ${w - arcW} 0 A ${arcW} ${h / 2} 0 0 1 ${w - arcW} ${h} L 0 ${h} Z"/>`;
  },

  flowChartDisplay: (w, h) => {
    const leftW = w * 0.15;
    const arcW = w * 0.35;
    return `<path d="M ${leftW} 0 L ${w - arcW} 0 A ${arcW} ${h / 2} 0 0 1 ${w - arcW} ${h} L ${leftW} ${h} L 0 ${h / 2} Z"/>`;
  },

  flowChartMagneticTape: (w, h) => {
    const r = Math.min(w, h) / 2;
    const cx = w / 2;
    const cy = h / 2;
    return `<path d="M ${cx + r} ${cy} A ${r} ${r} 0 1 1 ${cx + r - 0.01} ${cy + 0.01} L ${w} ${cy} L ${w} ${h} L ${w - r * 0.3} ${h} L ${cx + r * Math.cos(Math.PI / 6)} ${cy + r * Math.sin(Math.PI / 6)}"/>`;
  },

  flowChartMagneticDisk: (w, h) => {
    const ry = h * 0.15;
    return `<path d="M 0 ${ry} A ${w / 2} ${ry} 0 0 1 ${w} ${ry} L ${w} ${h - ry} A ${w / 2} ${ry} 0 0 1 0 ${h - ry} Z M 0 ${ry} A ${w / 2} ${ry} 0 0 0 ${w} ${ry}"/>`;
  },

  flowChartMagneticDrum: (w, h) => {
    const rx = w * 0.15;
    return `<path d="M ${rx} 0 A ${rx} ${h / 2} 0 0 0 ${rx} ${h} L ${w - rx} ${h} A ${rx} ${h / 2} 0 0 0 ${w - rx} 0 Z M ${w - rx} 0 A ${rx} ${h / 2} 0 0 1 ${w - rx} ${h}"/>`;
  },

  flowChartSummingJunction: (w, h) => {
    const cx = w / 2;
    const cy = h / 2;
    const rx = w / 2;
    const ry = h / 2;
    const d = 0.707;
    return `<path d="M ${cx + rx} ${cy} A ${rx} ${ry} 0 1 1 ${cx + rx - 0.01} ${cy - 0.01} Z M ${cx - rx * d} ${cy - ry * d} L ${cx + rx * d} ${cy + ry * d} M ${cx + rx * d} ${cy - ry * d} L ${cx - rx * d} ${cy + ry * d}"/>`;
  },

  flowChartOr: (w, h) => {
    const cx = w / 2;
    const cy = h / 2;
    const rx = w / 2;
    const ry = h / 2;
    return `<path d="M ${cx + rx} ${cy} A ${rx} ${ry} 0 1 1 ${cx + rx - 0.01} ${cy - 0.01} Z M ${cx} 0 L ${cx} ${h} M 0 ${cy} L ${w} ${cy}"/>`;
  },

  // =====================
  // Callout shapes
  // =====================

  wedgeRectCallout: (w, h, adj) => {
    const tipX = w / 2 + ((adj["adj1"] ?? -20833) / 100000) * w;
    const tipY = h / 2 + ((adj["adj2"] ?? 62500) / 100000) * h;
    const bx = w / 2;
    const wedgeW = w * 0.06;
    return `<path d="M 0 0 L ${w} 0 L ${w} ${h} L ${bx + wedgeW} ${h} L ${tipX} ${tipY} L ${bx - wedgeW} ${h} L 0 ${h} Z"/>`;
  },

  wedgeRoundRectCallout: (w, h, adj) => {
    const tipX = w / 2 + ((adj["adj1"] ?? -20833) / 100000) * w;
    const tipY = h / 2 + ((adj["adj2"] ?? 62500) / 100000) * h;
    const r = ((adj["adj3"] ?? 16667) / 100000) * Math.min(w, h);
    const bx = w / 2;
    const wedgeW = w * 0.06;
    return `<path d="M ${r} 0 L ${w - r} 0 A ${r} ${r} 0 0 1 ${w} ${r} L ${w} ${h - r} A ${r} ${r} 0 0 1 ${w - r} ${h} L ${bx + wedgeW} ${h} L ${tipX} ${tipY} L ${bx - wedgeW} ${h} L ${r} ${h} A ${r} ${r} 0 0 1 0 ${h - r} L 0 ${r} A ${r} ${r} 0 0 1 ${r} 0 Z"/>`;
  },

  wedgeEllipseCallout: (w, h, adj) => {
    const tipX = w / 2 + ((adj["adj1"] ?? -20833) / 100000) * w;
    const tipY = h / 2 + ((adj["adj2"] ?? 62500) / 100000) * h;
    const cx = w / 2;
    const cy = h / 2;
    const rx = w / 2;
    const ry = h / 2;
    const angle = Math.atan2(tipY - cy, tipX - cx);
    const wedgeAngle = 0.15;
    const x1 = cx + rx * Math.cos(angle - wedgeAngle);
    const y1 = cy + ry * Math.sin(angle - wedgeAngle);
    const x2 = cx + rx * Math.cos(angle + wedgeAngle);
    const y2 = cy + ry * Math.sin(angle + wedgeAngle);
    return `<path d="M ${x1} ${y1} L ${tipX} ${tipY} L ${x2} ${y2} A ${rx} ${ry} 0 1 1 ${x1} ${y1} Z"/>`;
  },

  cloudCallout: (w, h, adj) => {
    const tipX = w / 2 + ((adj["adj1"] ?? -20833) / 100000) * w;
    const tipY = h / 2 + ((adj["adj2"] ?? 62500) / 100000) * h;
    const r = Math.min(w, h) * 0.15;
    const bx = w / 2;
    const by = h / 2;
    const dx = tipX - bx;
    const dy = tipY - by;
    const d1x = bx + dx * 0.33;
    const d1y = by + dy * 0.33;
    const d2x = bx + dx * 0.66;
    const d2y = by + dy * 0.66;
    return `<path d="M ${r} ${h} A ${r} ${r} 0 0 1 0 ${h - r} A ${r} ${r} 0 0 1 ${r} ${h - 2 * r} L ${r} ${r} A ${r} ${r} 0 0 1 ${2 * r} 0 L ${w - 2 * r} 0 A ${r} ${r} 0 0 1 ${w - r} ${r} L ${w - r} ${h - 2 * r} A ${r} ${r} 0 0 1 ${w - 2 * r} ${h - r} A ${r} ${r} 0 0 1 ${w - 2 * r} ${h} Z M ${d1x} ${d1y} m ${r * 0.25} 0 a ${r * 0.25} ${r * 0.25} 0 1 1 -${r * 0.5} 0 a ${r * 0.25} ${r * 0.25} 0 1 1 ${r * 0.5} 0 Z M ${d2x} ${d2y} m ${r * 0.15} 0 a ${r * 0.15} ${r * 0.15} 0 1 1 -${r * 0.3} 0 a ${r * 0.15} ${r * 0.15} 0 1 1 ${r * 0.3} 0 Z"/>`;
  },

  borderCallout1: (w, h, adj) => {
    const y1 = ((adj["adj1"] ?? 18750) / 100000) * h;
    const x1 = ((adj["adj2"] ?? -8333) / 100000) * w;
    const y2 = ((adj["adj3"] ?? 112500) / 100000) * h;
    const x2 = ((adj["adj4"] ?? -38333) / 100000) * w;
    return `<path d="M 0 0 L ${w} 0 L ${w} ${h} L 0 ${h} Z M ${x1} ${y1} L ${x2} ${y2}"/>`;
  },

  borderCallout2: (w, h, adj) => {
    const y1 = ((adj["adj1"] ?? 18750) / 100000) * h;
    const x1 = ((adj["adj2"] ?? -8333) / 100000) * w;
    const y2 = ((adj["adj3"] ?? 18750) / 100000) * h;
    const x2 = ((adj["adj4"] ?? -16667) / 100000) * w;
    const y3 = ((adj["adj5"] ?? 112500) / 100000) * h;
    const x3 = ((adj["adj6"] ?? -46667) / 100000) * w;
    return `<path d="M 0 0 L ${w} 0 L ${w} ${h} L 0 ${h} Z M ${x1} ${y1} L ${x2} ${y2} L ${x3} ${y3}"/>`;
  },

  borderCallout3: (w, h, adj) => {
    const y1 = ((adj["adj1"] ?? 18750) / 100000) * h;
    const x1 = ((adj["adj2"] ?? -8333) / 100000) * w;
    const y2 = ((adj["adj3"] ?? 18750) / 100000) * h;
    const x2 = ((adj["adj4"] ?? -16667) / 100000) * w;
    const y3 = ((adj["adj5"] ?? 100000) / 100000) * h;
    const x3 = ((adj["adj6"] ?? -16667) / 100000) * w;
    const y4 = ((adj["adj7"] ?? 112963) / 100000) * h;
    const x4 = ((adj["adj8"] ?? -46667) / 100000) * w;
    return `<path d="M 0 0 L ${w} 0 L ${w} ${h} L 0 ${h} Z M ${x1} ${y1} L ${x2} ${y2} L ${x3} ${y3} L ${x4} ${y4}"/>`;
  },

  // =====================
  // Arc shapes
  // =====================

  arc: (w, h, adj) => {
    const stAng = ooxmlAngleToRadians(adj["adj1"] ?? 16200000);
    const endAng = ooxmlAngleToRadians(adj["adj2"] ?? 0);
    const rx = w / 2;
    const ry = h / 2;
    const cx = rx;
    const cy = ry;
    const x1 = cx + rx * Math.cos(stAng);
    const y1 = cy - ry * Math.sin(stAng);
    const x2 = cx + rx * Math.cos(endAng);
    const y2 = cy - ry * Math.sin(endAng);
    let sweep = stAng - endAng;
    if (sweep < 0) sweep += 2 * Math.PI;
    const largeArc = sweep > Math.PI ? 1 : 0;
    return `<path d="M ${x1} ${y1} A ${rx} ${ry} 0 ${largeArc} 0 ${x2} ${y2}"/>`;
  },

  chord: (w, h, adj) => {
    const stAng = ooxmlAngleToRadians(adj["adj1"] ?? 2700000);
    const endAng = ooxmlAngleToRadians(adj["adj2"] ?? 16200000);
    const rx = w / 2;
    const ry = h / 2;
    const cx = rx;
    const cy = ry;
    const x1 = cx + rx * Math.cos(stAng);
    const y1 = cy - ry * Math.sin(stAng);
    const x2 = cx + rx * Math.cos(endAng);
    const y2 = cy - ry * Math.sin(endAng);
    let sweep = stAng - endAng;
    if (sweep < 0) sweep += 2 * Math.PI;
    const largeArc = sweep > Math.PI ? 1 : 0;
    return `<path d="M ${x1} ${y1} A ${rx} ${ry} 0 ${largeArc} 0 ${x2} ${y2} Z"/>`;
  },

  pie: (w, h, adj) => {
    const stAng = ooxmlAngleToRadians(adj["adj1"] ?? 0);
    const endAng = ooxmlAngleToRadians(adj["adj2"] ?? 16200000);
    const rx = w / 2;
    const ry = h / 2;
    const cx = rx;
    const cy = ry;
    const x1 = cx + rx * Math.cos(stAng);
    const y1 = cy - ry * Math.sin(stAng);
    const x2 = cx + rx * Math.cos(endAng);
    const y2 = cy - ry * Math.sin(endAng);
    let sweep = stAng - endAng;
    if (sweep < 0) sweep += 2 * Math.PI;
    const largeArc = sweep > Math.PI ? 1 : 0;
    return `<path d="M ${cx} ${cy} L ${x1} ${y1} A ${rx} ${ry} 0 ${largeArc} 0 ${x2} ${y2} Z"/>`;
  },

  blockArc: (w, h, adj) => {
    const stAng = ooxmlAngleToRadians(adj["adj1"] ?? 10800000);
    const endAng = ooxmlAngleToRadians(adj["adj2"] ?? 0);
    const thickness = (adj["adj3"] ?? 25000) / 100000;
    const rx = w / 2;
    const ry = h / 2;
    const cx = rx;
    const cy = ry;
    const irx = rx * (1 - thickness);
    const iry = ry * (1 - thickness);
    const ox1 = cx + rx * Math.cos(stAng);
    const oy1 = cy - ry * Math.sin(stAng);
    const ox2 = cx + rx * Math.cos(endAng);
    const oy2 = cy - ry * Math.sin(endAng);
    const ix1 = cx + irx * Math.cos(stAng);
    const iy1 = cy - iry * Math.sin(stAng);
    const ix2 = cx + irx * Math.cos(endAng);
    const iy2 = cy - iry * Math.sin(endAng);
    let sweep = stAng - endAng;
    if (sweep < 0) sweep += 2 * Math.PI;
    const largeArc = sweep > Math.PI ? 1 : 0;
    return `<path d="M ${ox1} ${oy1} A ${rx} ${ry} 0 ${largeArc} 0 ${ox2} ${oy2} L ${ix2} ${iy2} A ${irx} ${iry} 0 ${largeArc} 1 ${ix1} ${iy1} Z"/>`;
  },

  // =====================
  // Math shapes
  // =====================

  mathPlus: (w, h, adj) => {
    const t = (adj["adj1"] ?? 23520) / 100000;
    const tw = t * w;
    const th = t * h;
    const lx = (w - tw) / 2;
    const rx = lx + tw;
    const ty = (h - th) / 2;
    const by = ty + th;
    return `<polygon points="${lx},0 ${rx},0 ${rx},${ty} ${w},${ty} ${w},${by} ${rx},${by} ${rx},${h} ${lx},${h} ${lx},${by} 0,${by} 0,${ty} ${lx},${ty}"/>`;
  },

  mathMinus: (w, h, adj) => {
    const t = (adj["adj1"] ?? 23520) / 100000;
    const th = t * h;
    const ty = (h - th) / 2;
    const by = ty + th;
    return `<rect x="0" y="${ty}" width="${w}" height="${by - ty}"/>`;
  },

  mathMultiply: (w, h, adj) => {
    const t = (adj["adj1"] ?? 23520) / 100000;
    const d = t * Math.min(w, h) * 0.5;
    const cx = w / 2;
    const cy = h / 2;
    return `<path d="M ${cx} ${cy - d} L ${w - d} 0 L ${w} ${d} L ${cx + d} ${cy} L ${w} ${h - d} L ${w - d} ${h} L ${cx} ${cy + d} L ${d} ${h} L 0 ${h - d} L ${cx - d} ${cy} L 0 ${d} L ${d} 0 Z"/>`;
  },

  mathDivide: (w, h, adj) => {
    const t = (adj["adj1"] ?? 23520) / 100000;
    const th = t * h;
    const ty = (h - th) / 2;
    const by = ty + th;
    const dotR = Math.min(w, h) * t * 0.5;
    const cx = w / 2;
    const topDotY = ty / 2;
    const botDotY = h - ty / 2;
    return `<path d="M 0 ${ty} L ${w} ${ty} L ${w} ${by} L 0 ${by} Z M ${cx + dotR} ${topDotY} A ${dotR} ${dotR} 0 1 1 ${cx + dotR - 0.01} ${topDotY - 0.01} Z M ${cx + dotR} ${botDotY} A ${dotR} ${dotR} 0 1 1 ${cx + dotR - 0.01} ${botDotY - 0.01} Z"/>`;
  },

  mathEqual: (w, h, adj) => {
    const t = (adj["adj1"] ?? 23520) / 100000;
    const gap = t * h;
    const barH = t * h;
    const y1 = (h - gap) / 2 - barH;
    const y2 = (h + gap) / 2;
    return `<path d="M 0 ${y1} L ${w} ${y1} L ${w} ${y1 + barH} L 0 ${y1 + barH} Z M 0 ${y2} L ${w} ${y2} L ${w} ${y2 + barH} L 0 ${y2 + barH} Z"/>`;
  },

  mathNotEqual: (w, h, adj) => {
    const t = (adj["adj1"] ?? 23520) / 100000;
    const gap = t * h;
    const barH = t * h;
    const y1 = (h - gap) / 2 - barH;
    const y2 = (h + gap) / 2;
    const slashW = w * 0.15;
    const sx = w / 2 - slashW / 2;
    return `<path d="M 0 ${y1} L ${w} ${y1} L ${w} ${y1 + barH} L 0 ${y1 + barH} Z M 0 ${y2} L ${w} ${y2} L ${w} ${y2 + barH} L 0 ${y2 + barH} Z M ${sx + slashW} ${y1 - barH} L ${sx + slashW + slashW} ${y1 - barH} L ${sx} ${y2 + barH + barH} L ${sx - slashW} ${y2 + barH + barH} Z"/>`;
  },

  // =====================
  // Other shapes
  // =====================

  plus: (w, h, adj) => {
    const t = (adj["adj"] ?? 25000) / 100000;
    const lx = t * w;
    const rx = w - lx;
    const ty = t * h;
    const by = h - ty;
    return `<polygon points="${lx},0 ${rx},0 ${rx},${ty} ${w},${ty} ${w},${by} ${rx},${by} ${rx},${h} ${lx},${h} ${lx},${by} 0,${by} 0,${ty} ${lx},${ty}"/>`;
  },

  corner: (w, h, adj) => {
    const adjX = (adj["adj1"] ?? 50000) / 100000;
    const adjY = (adj["adj2"] ?? 50000) / 100000;
    const cx = adjX * w;
    const cy = adjY * h;
    return `<polygon points="0,0 ${cx},0 ${cx},${cy} ${w},${cy} ${w},${h} 0,${h}"/>`;
  },

  diagStripe: (w, h, adj) => {
    const d = ((adj["adj"] ?? 50000) / 100000) * Math.min(w, h);
    return `<polygon points="0,${d} ${d},0 ${w},0 0,${h}"/>`;
  },

  foldedCorner: (w, h, adj) => {
    const fold = ((adj["adj"] ?? 16667) / 100000) * Math.min(w, h);
    return `<path d="M 0 0 L ${w} 0 L ${w} ${h - fold} L ${w - fold} ${h} L 0 ${h} Z M ${w - fold} ${h} L ${w - fold} ${h - fold} L ${w} ${h - fold}"/>`;
  },

  plaque: (w, h, adj) => {
    const r = ((adj["adj"] ?? 16667) / 100000) * Math.min(w, h);
    return `<path d="M 0 ${r} A ${r} ${r} 0 0 1 ${r} 0 L ${w - r} 0 A ${r} ${r} 0 0 1 ${w} ${r} L ${w} ${h - r} A ${r} ${r} 0 0 1 ${w - r} ${h} L ${r} ${h} A ${r} ${r} 0 0 1 0 ${h - r} Z"/>`;
  },

  can: (w, h, adj) => {
    const ry = ((adj["adj"] ?? 25000) / 100000) * h * 0.5;
    return `<path d="M 0 ${ry} A ${w / 2} ${ry} 0 0 1 ${w} ${ry} L ${w} ${h - ry} A ${w / 2} ${ry} 0 0 1 0 ${h - ry} Z M 0 ${ry} A ${w / 2} ${ry} 0 0 0 ${w} ${ry}"/>`;
  },

  cube: (w, h, adj) => {
    const d = ((adj["adj"] ?? 25000) / 100000) * Math.min(w, h);
    return `<path d="M 0 ${d} L ${d} 0 L ${w} 0 L ${w} ${h - d} L ${w - d} ${h} L 0 ${h} Z M 0 ${d} L ${w - d} ${d} L ${w} 0 M ${w - d} ${d} L ${w - d} ${h}"/>`;
  },

  donut: (w, h, adj) => {
    const t = (adj["adj"] ?? 25000) / 100000;
    const rx = w / 2;
    const ry = h / 2;
    const irx = rx * (1 - t);
    const iry = ry * (1 - t);
    const cx = rx;
    const cy = ry;
    return `<path fill-rule="evenodd" d="M ${cx + rx} ${cy} A ${rx} ${ry} 0 1 1 ${cx - rx} ${cy} A ${rx} ${ry} 0 1 1 ${cx + rx} ${cy} Z M ${cx + irx} ${cy} A ${irx} ${iry} 0 1 0 ${cx - irx} ${cy} A ${irx} ${iry} 0 1 0 ${cx + irx} ${cy} Z"/>`;
  },

  noSmoking: (w, h, adj) => {
    const t = (adj["adj"] ?? 18750) / 100000;
    const rx = w / 2;
    const ry = h / 2;
    const cx = rx;
    const cy = ry;
    const irx = rx * (1 - t);
    const iry = ry * (1 - t);
    const angle = Math.PI / 4;
    const lx1 = cx + irx * Math.cos(angle);
    const ly1 = cy - iry * Math.sin(angle);
    const lx2 = cx - irx * Math.cos(angle);
    const ly2 = cy + iry * Math.sin(angle);
    return `<path fill-rule="evenodd" d="M ${cx + rx} ${cy} A ${rx} ${ry} 0 1 1 ${cx - rx} ${cy} A ${rx} ${ry} 0 1 1 ${cx + rx} ${cy} Z M ${cx + irx} ${cy} A ${irx} ${iry} 0 1 0 ${cx - irx} ${cy} A ${irx} ${iry} 0 1 0 ${cx + irx} ${cy} Z M ${lx1} ${ly1} L ${lx2} ${ly2}"/>`;
  },

  smileyFace: (w, h, adj) => {
    const smile = (adj["adj"] ?? 4653) / 100000;
    const rx = w / 2;
    const ry = h / 2;
    const cx = rx;
    const cy = ry;
    const eyeRx = w * 0.06;
    const eyeRy = h * 0.06;
    const eyeY = h * 0.35;
    const leftEyeX = w * 0.35;
    const rightEyeX = w * 0.65;
    const mouthY = h * 0.6;
    const mouthW = w * 0.3;
    const mouthCurve = smile * h;
    return `<path d="M ${cx + rx} ${cy} A ${rx} ${ry} 0 1 1 ${cx - rx} ${cy} A ${rx} ${ry} 0 1 1 ${cx + rx} ${cy} Z M ${leftEyeX + eyeRx} ${eyeY} A ${eyeRx} ${eyeRy} 0 1 1 ${leftEyeX - eyeRx} ${eyeY} A ${eyeRx} ${eyeRy} 0 1 1 ${leftEyeX + eyeRx} ${eyeY} Z M ${rightEyeX + eyeRx} ${eyeY} A ${eyeRx} ${eyeRy} 0 1 1 ${rightEyeX - eyeRx} ${eyeY} A ${eyeRx} ${eyeRy} 0 1 1 ${rightEyeX + eyeRx} ${eyeY} Z M ${cx - mouthW} ${mouthY} C ${cx - mouthW * 0.5} ${mouthY + mouthCurve}, ${cx + mouthW * 0.5} ${mouthY + mouthCurve}, ${cx + mouthW} ${mouthY}"/>`;
  },

  frame: (w, h, adj) => {
    const t = ((adj["adj1"] ?? 12500) / 100000) * Math.min(w, h);
    return `<path fill-rule="evenodd" d="M 0 0 L ${w} 0 L ${w} ${h} L 0 ${h} Z M ${t} ${t} L ${t} ${h - t} L ${w - t} ${h - t} L ${w - t} ${t} Z"/>`;
  },

  bevel: (w, h, adj) => {
    const t = ((adj["adj"] ?? 12500) / 100000) * Math.min(w, h);
    return `<path d="M 0 0 L ${w} 0 L ${w} ${h} L 0 ${h} Z M ${t} ${t} L ${w - t} ${t} L ${w - t} ${h - t} L ${t} ${h - t} Z M 0 0 L ${t} ${t} M ${w} 0 L ${w - t} ${t} M ${w} ${h} L ${w - t} ${h - t} M 0 ${h} L ${t} ${h - t}"/>`;
  },

  halfFrame: (w, h, adj) => {
    const adjX = ((adj["adj1"] ?? 33333) / 100000) * w;
    const adjY = ((adj["adj2"] ?? 33333) / 100000) * h;
    return `<polygon points="0,0 ${w},0 ${w},${adjY} ${adjX},${adjY} ${adjX},${h} 0,${h}"/>`;
  },

  // Snip/Round corner rectangles

  snip1Rect: (w, h, adj) => {
    const d = ((adj["adj"] ?? 16667) / 100000) * Math.min(w, h);
    return `<polygon points="0,0 ${w - d},0 ${w},${d} ${w},${h} 0,${h}"/>`;
  },

  snip2SameRect: (w, h, adj) => {
    const d = ((adj["adj1"] ?? 16667) / 100000) * Math.min(w, h);
    const d2 = ((adj["adj2"] ?? 0) / 100000) * Math.min(w, h);
    return `<polygon points="${d},0 ${w - d},0 ${w},${d} ${w},${h - d2} ${w - d2},${h} ${d2},${h} 0,${h - d2} 0,${d}"/>`;
  },

  snip2DiagRect: (w, h, adj) => {
    const d1 = ((adj["adj1"] ?? 16667) / 100000) * Math.min(w, h);
    const d2 = ((adj["adj2"] ?? 0) / 100000) * Math.min(w, h);
    return `<polygon points="${d1},0 ${w},0 ${w},${h - d2} ${w - d2},${h} 0,${h} 0,${d1}"/>`;
  },

  snipRoundRect: (w, h, adj) => {
    const r = ((adj["adj1"] ?? 16667) / 100000) * Math.min(w, h);
    const d = ((adj["adj2"] ?? 16667) / 100000) * Math.min(w, h);
    return `<path d="M ${r} 0 L ${w - d} 0 L ${w} ${d} L ${w} ${h} L 0 ${h} L 0 ${r} A ${r} ${r} 0 0 1 ${r} 0 Z"/>`;
  },

  round1Rect: (w, h, adj) => {
    const r = ((adj["adj"] ?? 16667) / 100000) * Math.min(w, h);
    return `<path d="M 0 0 L ${w - r} 0 A ${r} ${r} 0 0 1 ${w} ${r} L ${w} ${h} L 0 ${h} Z"/>`;
  },

  round2SameRect: (w, h, adj) => {
    const r1 = ((adj["adj1"] ?? 16667) / 100000) * Math.min(w, h);
    const r2 = ((adj["adj2"] ?? 0) / 100000) * Math.min(w, h);
    return `<path d="M ${r1} 0 L ${w - r1} 0 A ${r1} ${r1} 0 0 1 ${w} ${r1} L ${w} ${h - r2} A ${r2} ${r2} 0 0 1 ${w - r2} ${h} L ${r2} ${h} A ${r2} ${r2} 0 0 1 0 ${h - r2} L 0 ${r1} A ${r1} ${r1} 0 0 1 ${r1} 0 Z"/>`;
  },

  round2DiagRect: (w, h, adj) => {
    const r1 = ((adj["adj1"] ?? 16667) / 100000) * Math.min(w, h);
    const r2 = ((adj["adj2"] ?? 0) / 100000) * Math.min(w, h);
    return `<path d="M ${r1} 0 L ${w} 0 L ${w} ${h - r2} A ${r2} ${r2} 0 0 1 ${w - r2} ${h} L 0 ${h} L 0 ${r1} A ${r1} ${r1} 0 0 1 ${r1} 0 Z"/>`;
  },

  // Brackets and braces

  leftBracket: (w, h, adj) => {
    const r = ((adj["adj"] ?? 8333) / 100000) * h;
    return `<path d="M ${w} 0 L ${r} 0 A ${r} ${r} 0 0 0 0 ${r} L 0 ${h - r} A ${r} ${r} 0 0 0 ${r} ${h} L ${w} ${h}"/>`;
  },

  rightBracket: (w, h, adj) => {
    const r = ((adj["adj"] ?? 8333) / 100000) * h;
    return `<path d="M 0 0 L ${w - r} 0 A ${r} ${r} 0 0 1 ${w} ${r} L ${w} ${h - r} A ${r} ${r} 0 0 1 ${w - r} ${h} L 0 ${h}"/>`;
  },

  leftBrace: (w, h, adj) => {
    const r = ((adj["adj1"] ?? 8333) / 100000) * h;
    const mid = ((adj["adj2"] ?? 50000) / 100000) * h;
    return `<path d="M ${w} 0 A ${w / 2} ${r} 0 0 0 ${w / 2} ${r} L ${w / 2} ${mid - r} A ${w / 2} ${r} 0 0 1 0 ${mid} A ${w / 2} ${r} 0 0 1 ${w / 2} ${mid + r} L ${w / 2} ${h - r} A ${w / 2} ${r} 0 0 0 ${w} ${h}"/>`;
  },

  rightBrace: (w, h, adj) => {
    const r = ((adj["adj1"] ?? 8333) / 100000) * h;
    const mid = ((adj["adj2"] ?? 50000) / 100000) * h;
    return `<path d="M 0 0 A ${w / 2} ${r} 0 0 1 ${w / 2} ${r} L ${w / 2} ${mid - r} A ${w / 2} ${r} 0 0 0 ${w} ${mid} A ${w / 2} ${r} 0 0 0 ${w / 2} ${mid + r} L ${w / 2} ${h - r} A ${w / 2} ${r} 0 0 1 0 ${h}"/>`;
  },

  bracketPair: (w, h, adj) => {
    const r = ((adj["adj"] ?? 16667) / 100000) * Math.min(w, h);
    return `<path d="M ${r} 0 A ${r} ${r} 0 0 0 0 ${r} L 0 ${h - r} A ${r} ${r} 0 0 0 ${r} ${h} M ${w - r} 0 A ${r} ${r} 0 0 1 ${w} ${r} L ${w} ${h - r} A ${r} ${r} 0 0 1 ${w - r} ${h}"/>`;
  },

  bracePair: (w, h, adj) => {
    const r = ((adj["adj"] ?? 8333) / 100000) * Math.min(w, h);
    return `<path d="M ${r} 0 A ${r} ${r} 0 0 0 0 ${r} L 0 ${h / 2 - r} A ${r} ${r} 0 0 1 ${-r} ${h / 2} A ${r} ${r} 0 0 1 0 ${h / 2 + r} L 0 ${h - r} A ${r} ${r} 0 0 0 ${r} ${h} M ${w - r} 0 A ${r} ${r} 0 0 1 ${w} ${r} L ${w} ${h / 2 - r} A ${r} ${r} 0 0 0 ${w + r} ${h / 2} A ${r} ${r} 0 0 0 ${w} ${h / 2 + r} L ${w} ${h - r} A ${r} ${r} 0 0 1 ${w - r} ${h}"/>`;
  },

  // Other misc shapes

  lightningBolt: (w, h) =>
    `<polygon points="${w * 0.55},0 ${w * 0.3},${h * 0.4} ${w * 0.52},${h * 0.4} ${w * 0.25},${h} ${w * 0.75},${h * 0.5} ${w * 0.52},${h * 0.5} ${w * 0.85},0"/>`,

  moon: (w, h, adj) => {
    const t = ((adj["adj"] ?? 50000) / 100000) * w;
    const rx = w / 2;
    const ry = h / 2;
    const irx = t / 2;
    return `<path d="M ${w} 0 A ${rx} ${ry} 0 1 0 ${w} ${h} A ${irx} ${ry} 0 1 1 ${w} 0 Z"/>`;
  },

  teardrop: (w, h, adj) => {
    const d = ((adj["adj"] ?? 100000) / 100000) * Math.min(w, h) * 0.5;
    const rx = w / 2;
    const ry = h / 2;
    const cx = rx;
    const cy = ry;
    return `<path d="M ${cx} 0 L ${cx + d} 0 L ${w} ${cy - d + ry} A ${rx} ${ry} 0 1 1 ${cx} 0 Z"/>`;
  },

  sun: (w, h) => {
    const cx = w / 2;
    const cy = h / 2;
    const points: string[] = [];
    for (let i = 0; i < 16; i++) {
      const angle = (Math.PI * 2 * i) / 16 - Math.PI / 2;
      const r = i % 2 === 0 ? 1 : 0.7;
      points.push(`${cx + cx * r * Math.cos(angle)},${cy + cy * r * Math.sin(angle)}`);
    }
    const circleR = 0.4;
    return `<path d="M ${points.map((p, i) => (i === 0 ? `${p}` : ` L ${p}`)).join("")} Z M ${cx + cx * circleR} ${cy} A ${cx * circleR} ${cy * circleR} 0 1 0 ${cx - cx * circleR} ${cy} A ${cx * circleR} ${cy * circleR} 0 1 0 ${cx + cx * circleR} ${cy} Z"/>`;
  },

  wave: (w, h, adj) => {
    const dy = ((adj["adj1"] ?? 12500) / 100000) * h;
    const dx = ((adj["adj2"] ?? 0) / 100000) * w;
    return `<path d="M ${dx} ${dy} C ${dx + w * 0.25} 0, ${dx + w * 0.5} 0, ${w} ${dy} L ${w - dx} ${h - dy} C ${w - dx - w * 0.25} ${h}, ${w - dx - w * 0.5} ${h}, 0 ${h - dy} Z"/>`;
  },

  doubleWave: (w, h, adj) => {
    const dy = ((adj["adj1"] ?? 6250) / 100000) * h;
    const dx = ((adj["adj2"] ?? 0) / 100000) * w;
    return `<path d="M ${dx} ${dy} C ${dx + w * 0.167} 0, ${dx + w * 0.333} ${dy * 2}, ${w / 2} ${dy} C ${w / 2 + w * 0.167} 0, ${w / 2 + w * 0.333} ${dy * 2}, ${w} ${dy} L ${w - dx} ${h - dy} C ${w - dx - w * 0.167} ${h}, ${w - dx - w * 0.333} ${h - dy * 2}, ${w / 2} ${h - dy} C ${w / 2 - w * 0.167} ${h}, ${w / 2 - w * 0.333} ${h - dy * 2}, 0 ${h - dy} Z"/>`;
  },

  ribbon: (w, h, adj) => {
    const tabH = ((adj["adj1"] ?? 16667) / 100000) * h;
    const tabW = ((adj["adj2"] ?? 50000) / 100000) * w;
    const foldW = tabW * 0.3;
    return `<path d="M 0 ${tabH} L ${foldW} ${tabH * 1.5} L ${foldW} ${h} L ${tabW} ${h - tabH} L ${w - tabW} ${h - tabH} L ${w - foldW} ${h} L ${w - foldW} ${tabH * 1.5} L ${w} ${tabH} L ${w} 0 L ${w - tabW} 0 L ${w - tabW} ${tabH} L ${tabW} ${tabH} L ${tabW} 0 L 0 0 Z"/>`;
  },

  ribbon2: (w, h, adj) => {
    const tabH = ((adj["adj1"] ?? 16667) / 100000) * h;
    const tabW = ((adj["adj2"] ?? 50000) / 100000) * w;
    const foldW = tabW * 0.3;
    return `<path d="M 0 ${h - tabH} L ${foldW} ${h - tabH * 1.5} L ${foldW} 0 L ${tabW} ${tabH} L ${w - tabW} ${tabH} L ${w - foldW} 0 L ${w - foldW} ${h - tabH * 1.5} L ${w} ${h - tabH} L ${w} ${h} L ${w - tabW} ${h} L ${w - tabW} ${h - tabH} L ${tabW} ${h - tabH} L ${tabW} ${h} L 0 ${h} Z"/>`;
  },
};

export function getPresetGeometrySvg(
  preset: string,
  width: number,
  height: number,
  adjustValues: Record<string, number>,
): string {
  const generator = presetGeometries[preset];
  if (generator) {
    return generator(width, height, adjustValues);
  }
  // Fallback: render as rectangle
  return `<rect width="${width}" height="${height}"/>`;
}
