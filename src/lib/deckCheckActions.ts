// Reading the deck for Deck Check, going to a problem and fixing it. The rules are in deckCheck.ts.

import { BUILT_KINDS } from "./built";
import type { Fix, ShapeSnap, SlideSnap } from "./deckCheck";
import { KIND_TAG } from "./notes";
import { shapeAt, slideSize, type SlideSize } from "./ppt";

/** Retro's notes and stamps are only checked for being left in; its tracker isn't checked at all. */
const NOTE_KINDS = ["sticky", "stamp"];
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
    const slides = context.presentation.slides.load("items/id,items/layout/id");
    await context.sync();
    const perSlide = slides.items.map((slide) => slide.shapes.load(PROPS));
    await context.sync();

    // Top-level shapes, minus Retro's own.
    const kinds = perSlide.map((shapes) => shapes.items.map((shape) => shape.tags.getItemOrNullObject(KIND_TAG).load("value")));
    await context.sync();
    const all: Pending[] = [];
    const snap = (shape: PowerPoint.Shape, path: string[], nested: boolean): ShapeSnap => ({
      path,
      name: shape.name,
      type: shape.type,
      left: shape.left,
      top: shape.top,
      width: shape.width,
      height: shape.height,
      nested: nested || undefined,
      runs: [],
    });
    perSlide.forEach((shapes, i) =>
      shapes.items.forEach((shape, n) => {
        const kind = kinds[i][n].isNullObject ? "" : kinds[i][n].value.toLowerCase();
        if (kind === "tracker") return;
        const p = { snap: snap(shape, [shape.id], false), shape, slide: i };
        if (NOTE_KINDS.includes(kind)) p.snap.retro = kind as "sticky" | "stamp";
        if (BUILT_KINDS.includes(kind)) p.snap.built = true;
        all.push(p);
      }),
    );

    // Shapes inside groups (and groups in groups), selected through their top-level group.
    if (groups) {
      let frontier = all.filter((p) => p.snap.type === "Group" && !p.snap.retro);
      for (let depth = 0; depth < 4 && frontier.length; depth++) {
        const inner = frontier.map((p) => ({ parent: p, shapes: p.shape.group.shapes.load(PROPS) }));
        await context.sync();
        frontier = [];
        for (const { parent, shapes } of inner) {
          for (const shape of shapes.items) {
            const child = { snap: snap(shape, [...parent.snap.path, shape.id], true), shape, slide: parent.slide };
            if (parent.snap.built) child.snap.built = true;
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
    const ranges = frames.map((f) => (f && !f.isNullObject && f.hasText ? f.textRange.load(`text,${FONT}`) : undefined));
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
      if (range) p.snap.content = range.text;
    });

    // Fonts, colors and sizes. Mixed text reports no single value, so it's split until each piece
    // has one, keeping where each piece is so a fix can change just those characters.
    let pieces = ranges.flatMap((range, i) => (range ? [{ snap: all[i].snap, range, start: 0, length: range.text.length }] : []));
    for (let round = 0; pieces.length; round++) {
      const next: typeof pieces = [];
      for (const piece of pieces) {
        const { name, color, size } = piece.range.font;
        if ((name && color && size) || piece.length < 2 || round >= MAX_SPLITS) {
          piece.snap.runs.push({ start: piece.start, length: piece.length, font: name || undefined, color: color || undefined, size: size || undefined });
          continue;
        }
        const half = Math.floor(piece.length / 2);
        next.push({ snap: piece.snap, range: rangeOf(piece, piece.start, half), start: piece.start, length: half });
        next.push({ snap: piece.snap, range: rangeOf(piece, piece.start + half, piece.length - half), start: piece.start + half, length: piece.length - half });
      }
      if (!next.length) break;
      await context.sync();
      pieces = next;
    }

    const result: SlideSnap[] = slides.items.map((slide, index) => ({ id: slide.id, index, title: "", shapes: [], layoutId: slide.layout?.id }));
    for (const p of all) result[p.slide].shapes.push(p.snap);
    for (const s of result) {
      const title = s.shapes.find((x) => x.placeholder && TITLES.includes(x.placeholder.type) && x.content?.trim());
      s.title = title?.content?.replace(/\s+/g, " ").trim() ?? "";
    }
    return { slides: result, size, full };
  });
}

/** A piece of a text range, with its font loaded. Positions are within the whole text. */
function rangeOf(piece: { range: PowerPoint.TextRange; start: number }, start: number, length: number): PowerPoint.TextRange {
  return piece.range.getSubstring(start - piece.start, length).load(FONT);
}

const FONT = "font/name,font/color,font/size";

/** Go to a slide, and select a shape on it if there is one. */
export async function goToShape(slideId: string, shapeId: string): Promise<void> {
  await PowerPoint.run(async (context) => {
    context.presentation.setSelectedSlides([slideId]);
    await context.sync();
    if (!shapeId) return;
    context.presentation.slides.getItem(slideId).setSelectedShapes([shapeId]);
    await context.sync();
  });
}

export interface FixRequest {
  slideId: string;
  fix: Fix;
  /** The alt text you typed, for an "alt" fix. */
  alt?: string;
}

/** Apply fixes in one go. Deletions run last, so nothing else refers to a deleted shape. */
export async function applyFixes(requests: FixRequest[]): Promise<number> {
  await PowerPoint.run(async (context) => {
    const ordered = [...requests].sort((a, b) => Number(a.fix.kind === "delete") - Number(b.fix.kind === "delete"));
    for (const { slideId, fix, alt } of ordered) {
      const shape = shapeAt(context, slideId, fix.path);
      switch (fix.kind) {
        case "font":
        case "textColor":
          for (const r of fix.runs) {
            const font = shape.textFrame.textRange.getSubstring(r.start, r.length).font;
            if (fix.kind === "font") font.name = fix.to;
            else font.color = fix.to;
          }
          break;
        case "fill":
          shape.fill.setSolidColor(fix.to);
          break;
        case "move":
          shape.left = fix.left;
          shape.top = fix.top;
          if (fix.width !== undefined) shape.width = fix.width;
          if (fix.height !== undefined) shape.height = fix.height;
          break;
        case "alt":
          if (alt?.trim()) shape.altTextDescription = alt.trim();
          break;
        case "delete":
          shape.delete();
          break;
      }
    }
    await context.sync();
  });
  return requests.length;
}
