// Reading the deck for Deck Check, and going to a problem. The rules are in deckCheck.ts.

import type { ShapeSnap, SlideSnap } from "./deckCheck";
import { KIND_TAG } from "./notes";
import { slideSize, type SlideSize } from "./ppt";

/** Retro's own additions aren't checked: notes, stamps and the tracker are meant to look different. */
const SKIP_KINDS = ["sticky", "stamp", "tracker"];
const FILLED = ["GeometricShape", "TextBox", "Callout", "Freeform"];
const PICTURES = ["Image", "Graphic"];
const TITLES = ["Title", "CenterTitle", "VerticalTitle"];
/** Mixed formatting is split in halves this many times at most (up to 256 pieces per text). */
const MAX_SPLITS = 8;

export interface DeckSnapshot {
  slides: SlideSnap[];
  size: SlideSize;
  /** Whether this PowerPoint can read alt text and placeholders' text (API 1.10). */
  full: boolean;
}

interface Pending {
  snap: ShapeSnap;
  shape: PowerPoint.Shape;
  slide: number;
}

const PROPS = "items/id,items/name,items/type,items/left,items/top,items/width,items/height";

export async function snapshotDeck(): Promise<DeckSnapshot> {
  const full = Office.context.requirements.isSetSupported("PowerPointApi", "1.10");
  const groups = Office.context.requirements.isSetSupported("PowerPointApi", "1.8");
  return PowerPoint.run(async (context) => {
    const size = await slideSize(context);
    const slides = context.presentation.slides.load("items/id");
    await context.sync();
    const perSlide = slides.items.map((slide) => slide.shapes.load(PROPS));
    await context.sync();

    // Top-level shapes, minus Retro's own.
    const kinds = perSlide.map((shapes) => shapes.items.map((shape) => shape.tags.getItemOrNullObject(KIND_TAG).load("value")));
    await context.sync();
    const all: Pending[] = [];
    const snap = (shape: PowerPoint.Shape, selectId: string, nested: boolean): ShapeSnap => ({
      selectId,
      name: shape.name,
      type: shape.type,
      left: shape.left,
      top: shape.top,
      width: shape.width,
      height: shape.height,
      nested: nested || undefined,
      fonts: [],
      textColors: [],
    });
    perSlide.forEach((shapes, i) =>
      shapes.items.forEach((shape, n) => {
        const kind = kinds[i][n];
        if (!kind.isNullObject && SKIP_KINDS.includes(kind.value.toLowerCase())) return;
        all.push({ snap: snap(shape, shape.id, false), shape, slide: i });
      }),
    );

    // Shapes inside groups (and groups in groups), selected through their top-level group.
    if (groups) {
      let frontier = all.filter((p) => p.snap.type === "Group");
      for (let depth = 0; depth < 4 && frontier.length; depth++) {
        const inner = frontier.map((p) => ({ parent: p, shapes: p.shape.group.shapes.load(PROPS) }));
        await context.sync();
        frontier = [];
        for (const { parent, shapes } of inner) {
          for (const shape of shapes.items) {
            const child = { snap: snap(shape, parent.snap.selectId, true), shape, slide: parent.slide };
            all.push(child);
            if (shape.type === "Group") frontier.push(child);
          }
        }
      }
    }

    // Fills, placeholders, alt text and text frames.
    const fills = all.map((p) => (FILLED.includes(p.snap.type) ? p.shape.fill.load("type,foregroundColor") : undefined));
    const formats = all.map((p) => (p.snap.type === "Placeholder" ? p.shape.placeholderFormat.load("type,containedType") : undefined));
    const frames = all.map((p) =>
      full ? p.shape.getTextFrameOrNullObject().load("hasText") : ["GeometricShape", "TextBox"].includes(p.snap.type) ? p.shape.textFrame.load("hasText") : undefined,
    );
    await context.sync();
    const isPicture = (i: number) => PICTURES.includes(all[i].snap.type) || PICTURES.includes(formats[i]?.containedType ?? "");
    const alts = all.map((p, i) => (full && isPicture(i) ? p.shape.load("altTextDescription,isDecorative") : undefined));
    const ranges = frames.map((f) => (f && !f.isNullObject && f.hasText ? f.textRange.load("text,font/name,font/color") : undefined));
    await context.sync();

    all.forEach((p, i) => {
      const fill = fills[i];
      if (fill?.type === "Solid" && fill.foregroundColor) p.snap.fill = fill.foregroundColor;
      const format = formats[i];
      // Without API 1.10 a placeholder's text can't be read, so it's never called empty.
      const hasText = Boolean(frames[i] && !frames[i]!.isNullObject && frames[i]!.hasText);
      if (format) p.snap.placeholder = { type: format.type, empty: full && format.containedType === null && !hasText };
      if (alts[i]) p.snap.picture = { alt: alts[i]!.altTextDescription ?? "", decorative: Boolean(alts[i]!.isDecorative) };
      const range = ranges[i];
      if (range) p.snap.text = range.text.replace(/\s+/g, " ").trim().slice(0, 80);
    });

    // Fonts and text colors. Mixed text reports no single font, so it's split until each piece has one.
    let pieces = ranges.flatMap((range, i) => (range ? [{ snap: all[i].snap, range, start: 0, length: range.text.length }] : []));
    for (let round = 0; pieces.length; round++) {
      const next: typeof pieces = [];
      for (const piece of pieces) {
        const { name, color } = piece.range.font;
        if (name) piece.snap.fonts.push(name);
        if (color) piece.snap.textColors.push(color);
        if ((name && color) || piece.length < 2 || round >= MAX_SPLITS) continue;
        const half = Math.floor(piece.length / 2);
        next.push({ snap: piece.snap, range: rangeOf(piece, piece.start, half), start: piece.start, length: half });
        next.push({ snap: piece.snap, range: rangeOf(piece, piece.start + half, piece.length - half), start: piece.start + half, length: piece.length - half });
      }
      if (!next.length) break;
      await context.sync();
      pieces = next;
    }

    const result: SlideSnap[] = slides.items.map((slide, index) => ({ id: slide.id, index, title: "", shapes: [] }));
    for (const p of all) result[p.slide].shapes.push(p.snap);
    for (const s of result) {
      const title = s.shapes.find((x) => x.placeholder && TITLES.includes(x.placeholder.type) && x.text);
      s.title = title?.text ?? "";
    }
    return { slides: result, size, full };
  });
}

/** A piece of a text range, with its font loaded. Positions are within the whole text. */
function rangeOf(piece: { range: PowerPoint.TextRange; start: number }, start: number, length: number): PowerPoint.TextRange {
  return piece.range.getSubstring(start - piece.start, length).load("font/name,font/color");
}

/** Go to a slide and select a shape on it. */
export async function goToShape(slideId: string, shapeId: string): Promise<void> {
  await PowerPoint.run(async (context) => {
    context.presentation.setSelectedSlides([slideId]);
    await context.sync();
    context.presentation.slides.getItem(slideId).setSelectedShapes([shapeId]);
    await context.sync();
  });
}
