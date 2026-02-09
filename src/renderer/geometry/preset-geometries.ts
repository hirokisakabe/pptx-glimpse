type GeometryGenerator = (w: number, h: number, adj: Record<string, number>) => string;

const presetGeometries: Record<string, GeometryGenerator> = {
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

  pentagon: (w, h) => {
    const cx = w / 2;
    const cy = h / 2;
    const points = Array.from({ length: 5 }, (_, i) => {
      const angle = (Math.PI * 2 * i) / 5 - Math.PI / 2;
      return `${cx + cx * Math.cos(angle)},${cy + cy * Math.sin(angle)}`;
    }).join(" ");
    return `<polygon points="${points}"/>`;
  },

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

  star5: (w, h) => {
    const cx = w / 2;
    const cy = h / 2;
    const points: string[] = [];
    for (let i = 0; i < 10; i++) {
      const angle = (Math.PI * 2 * i) / 10 - Math.PI / 2;
      const r = i % 2 === 0 ? 1 : 0.38;
      points.push(`${cx + cx * r * Math.cos(angle)},${cy + cy * r * Math.sin(angle)}`);
    }
    return `<polygon points="${points.join(" ")}"/>`;
  },

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
