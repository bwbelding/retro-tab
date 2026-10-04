// Find and update in PowerPoint: reading every text in the deck and replacing chosen matches.
// Matching is in findReplace.ts.

import { replaceIn, type Match, type TextSource } from "./findReplace";
import { KIND_TAG } from "./notes";
import { shapeAt } from "./ppt";

const PROPS = "items/id,items/name,items/type";

interface Found {
  shape: PowerPoint.Shape;
  slide: number;
  path: string[];
}

/** Every text in the deck: shapes (including inside groups) and table cells. Retro's tracker is skipped. */
export async function collectText(): Promise<TextSource[]> {
  const full = Office.context.requirements.isSetSupported("PowerPointApi", "1.10");
  return PowerPoint.run(async (context) => {
    const slides = context.presentation.slides.load("items/id");
    await context.sync();
    const perSlide = slides.items.map((s) => s.shapes.load(PROPS));
    await context.sync();
    const kinds = perSlide.map((shapes) => shapes.items.map((shape) => shape.tags.getItemOrNullObject(KIND_TAG).load("value")));
    await context.sync();
    const all: Found[] = [];
    perSlide.forEach((shapes, i) =>
      shapes.items.forEach((shape, n) => {
        if (!kinds[i][n].isNullObject && kinds[i][n].value.toLowerCase() === "tracker") return;
        all.push({ shape, slide: i, path: [shape.id] });
      }),
    );
    let frontier = all.filter((f) => f.shape.type === "Group");
    for (let depth = 0; depth < 4 && frontier.length; depth++) {
      const inner = frontier.map((f) => ({ parent: f, shapes: f.shape.group.shapes.load(PROPS) }));
      await context.sync();
      frontier = [];
      for (const { parent, shapes } of inner) {
        for (const shape of shapes.items) {
          const child = { shape, slide: parent.slide, path: [...parent.path, shape.id] };
          all.push(child);
          if (shape.type === "Group") frontier.push(child);
        }
      }
    }

    const tables = all.map((f) => (f.shape.type === "Table" ? f.shape.getTable().load("rowCount,columnCount") : undefined));
    const frames = all.map((f) =>
      f.shape.type === "Table" || f.shape.type === "Group"
        ? undefined
        : full
          ? f.shape.getTextFrameOrNullObject().load("hasText")
          : ["GeometricShape", "TextBox"].includes(f.shape.type)
            ? f.shape.textFrame.load("hasText")
            : undefined,
    );
    await context.sync();
    const texts = frames.map((f) => (f && !f.isNullObject && f.hasText ? f.textRange.load("text") : undefined));
    const cells = tables.map((t) => {
      if (!t) return [];
      const out: { row: number; column: number; cell: PowerPoint.TableCell }[] = [];
      for (let row = 0; row < t.rowCount; row++) for (let column = 0; column < t.columnCount; column++) out.push({ row, column, cell: t.getCellOrNullObject(row, column).load("text") });
      return out;
    });
    await context.sync();

    const sources: TextSource[] = [];
    all.forEach((f, i) => {
      const at = { slideId: slides.items[f.slide].id, slideIndex: f.slide, path: f.path };
      const label = f.shape.type === "TextBox" ? "Text box" : f.shape.type === "Placeholder" ? "Placeholder" : "Shape";
      if (texts[i]?.text) sources.push({ ...at, where: `${label} “${f.shape.name}”`, text: texts[i]!.text });
      for (const { row, column, cell } of cells[i]) {
        if (!cell.isNullObject && cell.text) sources.push({ ...at, cell: { row, column }, where: `Table “${f.shape.name}”, row ${row + 1}, column ${column + 1}`, text: cell.text });
      }
    });
    return sources.sort((a, b) => a.slideIndex - b.slideIndex);
  });
}

/**
 * Replace the chosen matches. Each text is read again first and left alone if it changed since
 * the search, so a match is never replaced at the wrong place. Returns what was done.
 */
export async function replaceMatches(matches: Match[], replacement: string): Promise<{ replaced: number; skipped: number }> {
  const bySource = new Map<TextSource, Match[]>();
  for (const m of matches) bySource.set(m.source, [...(bySource.get(m.source) ?? []), m]);
  return PowerPoint.run(async (context) => {
    const targets = [...bySource].map(([source, found]) => {
      const shape = shapeAt(context, source.slideId, source.path);
      const target = source.cell ? shape.getTable().getCellOrNullObject(source.cell.row, source.cell.column) : shape.textFrame.textRange;
      return { source, found, target: target.load("text") };
    });
    await context.sync();
    let replaced = 0;
    let skipped = 0;
    for (const { source, found, target } of targets) {
      if (target.text !== source.text) {
        skipped += found.length;
        continue;
      }
      if (source.cell) {
        target.text = replaceIn(source.text, found, replacement);
      } else {
        // From the end, so the earlier positions still point at the right text; formatting stays.
        for (const m of [...found].sort((a, b) => b.start - a.start)) (target as PowerPoint.TextRange).getSubstring(m.start, m.length).text = replacement;
      }
      replaced += found.length;
    }
    await context.sync();
    return { replaced, skipped };
  });
}
