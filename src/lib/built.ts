// What the roadmap and KPI builders share: where to draw on the slide, and tagging what they draw
// with its source so selecting it later reopens it for editing.

import { KIND_TAG } from "./notes";
import { slideSize } from "./ppt";
import type { Area } from "./roadmap";
import { decodeName, encodeName } from "./tracker";

/** The source data on a built shape, hex-encoded JSON (tag values can come back capitalized). */
export const SOURCE_TAG = "RETRO_SOURCE";
/** Kinds of shapes Retro builds whose status colors are meant to be off-palette (Deck Check skips their colors). */
export const BUILT_KINDS = ["smart", "kpi", "roadmap"];

const TITLES = ["Title", "CenterTitle", "VerticalTitle"];

/** Below the slide's title, with margins. */
export async function contentArea(context: PowerPoint.RequestContext, slide: PowerPoint.Slide): Promise<Area> {
  const size = await slideSize(context);
  const shapes = slide.shapes.load("items/type,items/top,items/height");
  await context.sync();
  const formats = shapes.items.map((s) => (s.type === "Placeholder" ? s.placeholderFormat.load("type") : undefined));
  await context.sync();
  const title = shapes.items.find((_, i) => formats[i] && TITLES.includes(formats[i]!.type));
  const top = Math.min(title ? title.top + title.height + 16 : 100, size.height / 2);
  return { left: 36, top, width: size.width - 72, height: size.height - top - 36 };
}

export function tagBuilt(shape: PowerPoint.Shape, kind: string, source: unknown): void {
  shape.tags.add(KIND_TAG, kind);
  shape.tags.add(SOURCE_TAG, encodeName(JSON.stringify(source)));
}

/** The selected shape's id and source, if exactly one shape Retro built of this kind is selected. */
export async function readBuiltSelection<T>(kind: string): Promise<{ id: string; source: T } | undefined> {
  return PowerPoint.run(async (context) => {
    const sel = context.presentation.getSelectedShapes().load("items/id");
    await context.sync();
    if (sel.items.length !== 1) return undefined;
    const shape = sel.items[0];
    const k = shape.tags.getItemOrNullObject(KIND_TAG).load("value");
    const data = shape.tags.getItemOrNullObject(SOURCE_TAG).load("value");
    await context.sync();
    if (k.isNullObject || k.value.toLowerCase() !== kind || data.isNullObject) return undefined;
    try {
      return { id: shape.id, source: JSON.parse(decodeName(data.value)) as T };
    } catch {
      return undefined;
    }
  });
}
