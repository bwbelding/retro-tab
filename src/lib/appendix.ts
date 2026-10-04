// Appendix mover: mark slides as appendix, then move them all to the end of the deck behind an
// "Appendix" divider. Marks are slide tags, so they travel with the slides.

import { UserError } from "./ppt";

export const APPENDIX_TAG = "RETRO_APPENDIX";
/** On the divider slide Retro made, so it's reused rather than added again. */
export const DIVIDER_TAG = "RETRO_APPENDIX_DIVIDER";

export interface AppendixState {
  /** Indexes of slides marked as appendix (not the divider), in deck order. */
  marked: number[];
  divider: number;
  /** Indexes of the selected slides. */
  selected: number[];
  /** Whether the marked slides already sit together at the end, behind the divider. */
  inPlace: boolean;
}

/** Whether marked slides already sit together at the end, right after the divider. */
export function appendixInPlace(marked: number[], divider: number, count: number): boolean {
  if (marked.length === 0 || divider < 0) return false;
  return marked.every((m, i) => m === divider + 1 + i) && marked[marked.length - 1] === count - 1;
}

async function deck(context: PowerPoint.RequestContext) {
  const slides = context.presentation.slides.load("items/id");
  const selected = context.presentation.getSelectedSlides().load("items/id");
  await context.sync();
  const tags = slides.items.map((s) => ({ appendix: s.tags.getItemOrNullObject(APPENDIX_TAG).load("value"), divider: s.tags.getItemOrNullObject(DIVIDER_TAG).load("value") }));
  await context.sync();
  const chosen = new Set(selected.items.map((s) => s.id));
  return {
    slides: slides.items,
    marked: tags.flatMap((t, i) => (!t.appendix.isNullObject && t.divider.isNullObject ? [i] : [])),
    divider: tags.findIndex((t) => !t.divider.isNullObject),
    selected: slides.items.flatMap((s, i) => (chosen.has(s.id) ? [i] : [])),
  };
}

export async function readAppendix(): Promise<AppendixState> {
  return PowerPoint.run(async (context) => {
    const d = await deck(context);
    return { marked: d.marked, divider: d.divider, selected: d.selected, inPlace: appendixInPlace(d.marked, d.divider, d.slides.length) };
  });
}

/** Mark (or unmark) the selected slides as appendix. Returns how many changed. */
export async function markAppendix(mark: boolean): Promise<number> {
  return PowerPoint.run(async (context) => {
    const d = await deck(context);
    const targets = d.selected.filter((i) => i !== d.divider);
    if (targets.length === 0) throw new UserError("Select the slides first (click them in the slide thumbnails).");
    for (const i of targets) {
      if (mark) d.slides[i].tags.add(APPENDIX_TAG, "1");
      else if (d.marked.includes(i)) d.slides[i].tags.delete(APPENDIX_TAG);
    }
    await context.sync();
    return targets.length;
  });
}

/** Move the appendix slides to the end, in their order, behind an "Appendix" divider (added once). */
export async function moveAppendix(): Promise<string> {
  return PowerPoint.run(async (context) => {
    const d = await deck(context);
    if (d.marked.length === 0) throw new UserError("Mark some slides as appendix first: select them and click Mark as appendix.");
    const marked = d.marked.map((i) => d.slides[i]);
    let divider = d.divider >= 0 ? d.slides[d.divider] : undefined;
    let added = false;
    if (!divider) {
      // The template's section header layout, from the first appendix slide's master.
      const master = marked[0].slideMaster.load("id");
      const layouts = marked[0].slideMaster.layouts.load("items/id,items/name");
      await context.sync();
      const pick = (re: RegExp) => layouts.items.find((l) => re.test(l.name));
      const layout = pick(/section header/i) ?? pick(/section/i) ?? pick(/title only/i) ?? pick(/title/i) ?? layouts.items[0];
      context.presentation.slides.add({ slideMasterId: master.id, layoutId: layout?.id });
      await context.sync();
      const all = context.presentation.slides.load("items/id");
      await context.sync();
      divider = all.items[all.items.length - 1];
      divider.tags.add(DIVIDER_TAG, "1");
      divider.tags.add(APPENDIX_TAG, "1");
      const shapes = divider.shapes.load("items/id,items/type");
      await context.sync();
      const formats = shapes.items.map((s) => (s.type === "Placeholder" ? s.placeholderFormat.load("type") : undefined));
      await context.sync();
      const title = shapes.items.find((_, i) => formats[i] && ["Title", "CenterTitle"].includes(formats[i]!.type));
      if (title) title.textFrame.textRange.text = "Appendix";
      added = true;
    }
    const count = d.slides.length + (added ? 1 : 0);
    // Each goes to the end in turn, so the divider leads and the slides keep their order.
    divider.moveTo(count - 1);
    for (const slide of marked) slide.moveTo(count - 1);
    await context.sync();
    const n = marked.length;
    return `Moved ${n} appendix slide${n === 1 ? "" : "s"} to the end${added ? ", behind a new Appendix divider" : ", behind the Appendix divider"}.`;
  });
}
