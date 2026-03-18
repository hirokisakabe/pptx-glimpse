import type { SlideSize } from "../model/presentation.js";
import type { SlideElement } from "../model/shape.js";
import type { Slide } from "../model/slide.js";
import { convertElement } from "./element-converter.js";
import type { PomLayerChild, PomLayerNode } from "./pom-types.js";

/** pom's fixed slide width in pixels */
const POM_SLIDE_WIDTH = 1280;
/** pom's fixed slide height in pixels */
const POM_SLIDE_HEIGHT = 720;

export interface SlideScaleContext {
  scaleX: number;
  scaleY: number;
}

function createScaleContext(slideSize: SlideSize): SlideScaleContext {
  return {
    scaleX: POM_SLIDE_WIDTH / (slideSize.width as number),
    scaleY: POM_SLIDE_HEIGHT / (slideSize.height as number),
  };
}

export function convertSlideToPom(slide: Slide, slideSize: SlideSize): PomLayerNode {
  const ctx = createScaleContext(slideSize);

  const children: PomLayerChild[] = [];
  for (const element of slide.elements) {
    const child = convertSlideElement(element, ctx);
    if (child) {
      children.push(child);
    }
  }

  const layer: PomLayerNode = {
    type: "layer",
    w: POM_SLIDE_WIDTH,
    h: POM_SLIDE_HEIGHT,
    children,
  };

  if (slide.background?.fill) {
    const bg = slide.background.fill;
    if (bg.type === "solid") {
      layer.backgroundColor = stripHash(bg.color.hex);
    }
  }

  return layer;
}

function convertSlideElement(element: SlideElement, ctx: SlideScaleContext): PomLayerChild | null {
  const pomNode = convertElement(element, ctx);
  if (!pomNode) return null;

  // Line nodes already encode absolute coordinates in x1/y1/x2/y2
  if (pomNode.type === "line") {
    return { ...pomNode, x: 0, y: 0 } as PomLayerChild;
  }

  const t = element.transform;
  const x = round((t.offsetX as number) * ctx.scaleX);
  const y = round((t.offsetY as number) * ctx.scaleY);
  const w = round((t.extentWidth as number) * ctx.scaleX);
  const h = round((t.extentHeight as number) * ctx.scaleY);

  return { ...pomNode, x, y, w, h } as PomLayerChild;
}

function round(n: number): number {
  return Math.round(n * 100) / 100;
}

export function stripHash(hex: string): string {
  return hex.startsWith("#") ? hex.slice(1) : hex;
}
