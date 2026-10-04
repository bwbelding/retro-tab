// Thin helpers around PowerPoint.run used by the ribbon actions and task pane views.

import type { Rect } from "./layout";

/** An error whose message is meant for the user (shown in the pane), not a bug report. */
export class UserError extends Error {}

export interface SlideSize {
  width: number;
  height: number;
}

/** 16:9 default, used when PowerPoint can't report the slide size (before PowerPointApi 1.10). */
const DEFAULT_SLIDE: SlideSize = { width: 960, height: 540 };

export async function slideSize(context: PowerPoint.RequestContext): Promise<SlideSize> {
  if (!Office.context.requirements.isSetSupported("PowerPointApi", "1.10")) return DEFAULT_SLIDE;
  const setup = context.presentation.pageSetup;
  setup.load("slideWidth,slideHeight");
  await context.sync();
  return { width: setup.slideWidth, height: setup.slideHeight };
}

/** The slide being edited (the first selected slide). */
export async function currentSlide(context: PowerPoint.RequestContext): Promise<PowerPoint.Slide> {
  const slides = context.presentation.getSelectedSlides();
  slides.load("items/id");
  await context.sync();
  if (slides.items.length === 0) throw new UserError("Select a slide first.");
  return slides.items[0];
}

export interface Selection {
  shapes: PowerPoint.Shape[];
  rects: Rect[];
}

export async function selectedShapes(context: PowerPoint.RequestContext): Promise<Selection> {
  const coll = context.presentation.getSelectedShapes();
  coll.load("items/id,items/left,items/top,items/width,items/height");
  await context.sync();
  const shapes = coll.items;
  return { shapes, rects: shapes.map((s) => ({ id: s.id, left: s.left, top: s.top, width: s.width, height: s.height })) };
}

export async function requireSelection(context: PowerPoint.RequestContext, min: number, what: string): Promise<Selection> {
  const sel = await selectedShapes(context);
  if (sel.shapes.length < min) throw new UserError(`Select at least ${min} shapes to ${what}.`);
  return sel;
}

/** Write rectangles back to the shapes they came from (matched by id). */
export function applyRects(shapes: PowerPoint.Shape[], rects: Rect[]): void {
  const byId = new Map(rects.map((r) => [r.id, r]));
  for (const s of shapes) {
    const r = byId.get(s.id);
    if (!r) continue;
    s.left = r.left;
    s.top = r.top;
    s.width = r.width;
    s.height = r.height;
  }
}

/** A shape by its ids from the slide's top level down (one id unless it's inside groups). */
export function shapeAt(context: PowerPoint.RequestContext, slideId: string, path: string[]): PowerPoint.Shape {
  let shape = context.presentation.slides.getItem(slideId).shapes.getItem(path[0]);
  for (const id of path.slice(1)) shape = shape.group.shapes.getItem(id);
  return shape;
}
