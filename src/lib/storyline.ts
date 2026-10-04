// Storyline view: every slide's title in order, so you can read the deck's argument from the
// titles alone. A title under five words usually names a topic ("Market size") rather than
// stating the point ("Market grew 30% in Q3"), so it's flagged.

/** Titles shorter than this many words are flagged as topic labels. */
export const MIN_TITLE_WORDS = 5;

export interface TitleRow {
  slideId: string;
  index: number;
  /** The title placeholder's id, if the slide has one. */
  shapeId?: string;
  title: string;
}

export type TitleState = "ok" | "short" | "missing";

export const wordCount = (text: string) => text.split(/\s+/).filter((w) => /[\p{L}\p{N}]/u.test(w)).length;

export function titleState(row: TitleRow): TitleState {
  const title = row.title.trim();
  if (!title) return "missing";
  return wordCount(title) < MIN_TITLE_WORDS ? "short" : "ok";
}

const TITLE_TYPES = ["Title", "CenterTitle", "VerticalTitle"];

/** Every slide's title, in order. */
export async function readTitles(): Promise<TitleRow[]> {
  return PowerPoint.run(async (context) => {
    const slides = context.presentation.slides.load("items/id");
    await context.sync();
    const perSlide = slides.items.map((s) => s.shapes.load("items/id,items/type"));
    await context.sync();
    const formats = perSlide.map((shapes) => shapes.items.map((shape) => (shape.type === "Placeholder" ? shape.placeholderFormat.load("type") : undefined)));
    await context.sync();
    const titles = perSlide.map((shapes, i) => {
      const n = formats[i].findIndex((f) => f && TITLE_TYPES.includes(f.type));
      return n < 0 ? undefined : { shape: shapes.items[n], text: shapes.items[n].textFrame.textRange.load("text") };
    });
    await context.sync();
    return slides.items.map((slide, index) => ({
      slideId: slide.id,
      index,
      shapeId: titles[index]?.shape.id,
      title: (titles[index]?.text.text ?? "").replace(/\s+/g, " ").trim(),
    }));
  });
}

/** Change a slide's title (its formatting stays as the title's first character's). */
export async function setTitle(slideId: string, shapeId: string, title: string): Promise<void> {
  await PowerPoint.run(async (context) => {
    context.presentation.slides.getItem(slideId).shapes.getItem(shapeId).textFrame.textRange.text = title;
    await context.sync();
  });
}
