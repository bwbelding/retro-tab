// Sticky notes and stamps. Every shape Retro adds carries a RETRO_KIND tag, so Remove All deletes
// exactly what Retro added and nothing else. Placement and styling are pure (unit tested).

import { bounds, type Rect } from "./layout";
import { currentSlide, slideSize, type SlideSize } from "./ppt";

export const KIND_TAG = "RETRO_KIND";
export type NoteKind = "sticky" | "stamp";
const NOTE_KINDS: string[] = ["sticky", "stamp"];

export const STICKY_COLORS = { yellow: "#FFE680", pink: "#FFB3C7", blue: "#A8D8FF", green: "#B5E8B0" } as const;
export type StickyColor = keyof typeof STICKY_COLORS;

export const STAMPS = ["DRAFT", "CONFIDENTIAL", "FOR REVIEW"] as const;
export type StampLabel = (typeof STAMPS)[number];

const MARGIN = 24;
export const STICKY_SIZE = { width: 180, height: 120 };

/** Beside the selection when there's room on its right, otherwise the slide's top-right corner. */
export function stickyPosition(slide: SlideSize, selection: Rect[]): { left: number; top: number } {
  const { width, height } = STICKY_SIZE;
  const corner = { left: slide.width - width - MARGIN, top: MARGIN };
  if (selection.length === 0) return corner;
  const b = bounds(selection);
  const left = b.left + b.width + 12;
  if (left + width > slide.width) return corner;
  return { left, top: Math.min(Math.max(b.top, 0), slide.height - height) };
}

export function stampRect(slide: SlideSize, label: string): { left: number; top: number; width: number; height: number } {
  const width = Math.round(label.length * 17 + 40);
  const height = 44;
  return { left: slide.width - width - MARGIN, top: MARGIN, width, height };
}

export async function addSticky(text: string, color: StickyColor, selection: Rect[] = []): Promise<void> {
  await PowerPoint.run(async (context) => {
    const slide = await currentSlide(context);
    const pos = stickyPosition(await slideSize(context), selection);
    const s = slide.shapes.addGeometricShape(PowerPoint.GeometricShapeType.foldedCorner, { ...pos, ...STICKY_SIZE });
    s.name = "Retro sticky note";
    s.fill.setSolidColor(STICKY_COLORS[color]);
    s.lineFormat.visible = false;
    s.textFrame.textRange.text = text;
    s.textFrame.verticalAlignment = PowerPoint.TextVerticalAlignment.top;
    s.textFrame.leftMargin = s.textFrame.rightMargin = s.textFrame.topMargin = s.textFrame.bottomMargin = 8;
    s.textFrame.wordWrap = true;
    s.textFrame.textRange.paragraphFormat.horizontalAlignment = PowerPoint.ParagraphHorizontalAlignment.left;
    s.textFrame.textRange.font.color = "#242424";
    s.textFrame.textRange.font.size = 12;
    s.tags.add(KIND_TAG, "sticky");
    await context.sync();
    slide.setSelectedShapes([s.id]);
    await context.sync();
  });
}

function styleStamp(s: PowerPoint.Shape, label: string) {
  s.name = `Retro stamp ${label}`;
  s.fill.clear();
  s.lineFormat.color = "#C00000";
  s.lineFormat.weight = 2.25;
  s.textFrame.textRange.text = label;
  s.textFrame.verticalAlignment = PowerPoint.TextVerticalAlignment.middle;
  s.textFrame.textRange.paragraphFormat.horizontalAlignment = PowerPoint.ParagraphHorizontalAlignment.center;
  s.textFrame.textRange.font.color = "#C00000";
  s.textFrame.textRange.font.bold = true;
  s.textFrame.textRange.font.size = 20;
  s.textFrame.wordWrap = false;
  if (Office.context.requirements.isSetSupported("PowerPointApi", "1.10")) s.rotation = -8;
  s.tags.add(KIND_TAG, "stamp");
}

/** Stamp the current slide, or every slide. Returns how many slides were stamped. */
export async function addStamp(label: StampLabel, everySlide = false): Promise<number> {
  return PowerPoint.run(async (context) => {
    const size = await slideSize(context);
    const rect = stampRect(size, label);
    let targets: PowerPoint.Slide[];
    if (everySlide) {
      const slides = context.presentation.slides;
      slides.load("items/id");
      await context.sync();
      targets = slides.items;
    } else {
      targets = [await currentSlide(context)];
    }
    for (const slide of targets) styleStamp(slide.shapes.addGeometricShape(PowerPoint.GeometricShapeType.rectangle, rect), label);
    await context.sync();
    return targets.length;
  });
}

export interface NoteCount {
  notes: number;
  slides: number;
}

/** Count (or delete) every sticky note and stamp Retro added, across the whole presentation. */
export async function scanNotes(remove: boolean): Promise<NoteCount> {
  return PowerPoint.run(async (context) => {
    const slides = context.presentation.slides;
    slides.load("items/id");
    await context.sync();
    const perSlide = slides.items.map((slide) => {
      const shapes = slide.shapes;
      shapes.load("items/id");
      return shapes;
    });
    await context.sync();
    const tagged = perSlide.map((shapes) =>
      shapes.items.map((shape) => {
        const tag = shape.tags.getItemOrNullObject(KIND_TAG);
        tag.load("value");
        return { shape, tag };
      }),
    );
    await context.sync();
    let notes = 0;
    let slideCount = 0;
    for (const onSlide of tagged) {
      const hits = onSlide.filter(({ tag }) => !tag.isNullObject && NOTE_KINDS.includes(tag.value.toLowerCase()));
      if (hits.length) slideCount++;
      notes += hits.length;
      if (remove) for (const { shape } of hits) shape.delete();
    }
    await context.sync();
    return { notes, slides: slideCount };
  });
}
