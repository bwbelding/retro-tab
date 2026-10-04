// The section tracker in PowerPoint: marking sections, building the tracker bars and the agenda
// slide. The model and geometry are in tracker.ts.

import { APPENDIX_TAG } from "./appendix";
import { activeKit } from "./brand";
import { brandStore } from "./brandActions";
import { KIND_TAG } from "./notes";
import { slideSize, UserError } from "./ppt";
import {
  AGENDA_TAG,
  agendaText,
  decodeSettings,
  encodeName,
  encodeSettings,
  SECTION_TAG,
  sectionOf,
  sectionsFrom,
  textOn,
  TRACKER_TAG,
  trackerParts,
  type Section,
  type TrackerPart,
  type TrackerSettings,
} from "./tracker";

const TRACKER_KIND = "tracker";
const AGENDA_KIND = "agenda";
const INACTIVE_FILL = "#E3E3E3";
const INACTIVE_TEXT = "#7A7A7A";
const FALLBACK_COLOR = "#404040";

export interface TrackerState {
  sections: Section[];
  settings: TrackerSettings;
  slideCount: number;
  /** The selected slide's index, or -1. */
  current: number;
  /** The agenda slide's index, or -1. */
  agenda: number;
}

interface DeckSlides {
  slides: PowerPoint.Slide[];
  sections: Section[];
  agenda: number;
  /** Appendix slides (and their divider), which get no tracker. */
  appendix: Set<number>;
}

/** Every slide with its section and agenda tags. */
async function deckSlides(context: PowerPoint.RequestContext): Promise<DeckSlides> {
  const slides = context.presentation.slides;
  slides.load("items/id");
  await context.sync();
  const tags = slides.items.map((s) => ({
    section: s.tags.getItemOrNullObject(SECTION_TAG).load("value"),
    agenda: s.tags.getItemOrNullObject(AGENDA_TAG).load("value"),
    appendix: s.tags.getItemOrNullObject(APPENDIX_TAG).load("value"),
  }));
  await context.sync();
  return {
    slides: slides.items,
    sections: sectionsFrom(tags.map((t) => (t.section.isNullObject ? undefined : t.section.value))),
    agenda: tags.findIndex((t) => !t.agenda.isNullObject),
    appendix: new Set(tags.flatMap((t, i) => (t.appendix.isNullObject ? [] : [i]))),
  };
}

export async function readTracker(): Promise<TrackerState> {
  return PowerPoint.run(async (context) => {
    const deck = await deckSlides(context);
    const setting = context.presentation.tags.getItemOrNullObject(TRACKER_TAG).load("value");
    const selected = context.presentation.getSelectedSlides();
    selected.load("items/id");
    await context.sync();
    const id = selected.items[0]?.id;
    return {
      sections: deck.sections,
      settings: decodeSettings(setting.isNullObject ? undefined : setting.value),
      slideCount: deck.slides.length,
      current: deck.slides.findIndex((s) => s.id === id),
      agenda: deck.agenda,
    };
  });
}

/** Start a section (or rename the one starting) at the selected slide. */
export async function markSection(name: string): Promise<void> {
  const clean = name.trim();
  if (!clean) throw new UserError("Type a name for the section.");
  await PowerPoint.run(async (context) => {
    const deck = await deckSlides(context);
    const selected = context.presentation.getSelectedSlides();
    selected.load("items/id");
    await context.sync();
    const index = deck.slides.findIndex((s) => s.id === selected.items[0]?.id);
    if (index < 0) throw new UserError("Select the slide where the section starts.");
    if (index === deck.agenda) throw new UserError("That's the agenda slide. Select the first slide of the section.");
    deck.slides[index].tags.add(SECTION_TAG, encodeName(clean));
    await context.sync();
  });
}

/** Rename or remove the section starting at a slide. Removing merges its slides into the one before. */
export async function changeSection(start: number, name: string | null): Promise<void> {
  await PowerPoint.run(async (context) => {
    const { slides } = await deckSlides(context);
    const slide = slides[start];
    if (!slide) return;
    if (name === null) slide.tags.delete(SECTION_TAG);
    else if (name.trim()) slide.tags.add(SECTION_TAG, encodeName(name.trim()));
    await context.sync();
  });
}

export async function goToSlide(index: number): Promise<void> {
  await PowerPoint.run(async (context) => {
    const slides = context.presentation.slides;
    slides.load("items/id");
    await context.sync();
    const slide = slides.items[index];
    if (slide) context.presentation.setSelectedSlides([slide.id]);
    await context.sync();
  });
}

/** Delete Retro's tracker shapes from every slide. Returns how many slides had one. */
async function deleteTrackers(context: PowerPoint.RequestContext, slides: PowerPoint.Slide[]): Promise<number> {
  const perSlide = slides.map((slide) => slide.shapes.load("items/id"));
  await context.sync();
  const tagged = perSlide.map((shapes) => shapes.items.map((shape) => ({ shape, tag: shape.tags.getItemOrNullObject(KIND_TAG).load("value") })));
  await context.sync();
  let count = 0;
  for (const onSlide of tagged) {
    const hits = onSlide.filter(({ tag }) => !tag.isNullObject && tag.value.toLowerCase() === TRACKER_KIND);
    if (hits.length) count++;
    for (const { shape } of hits) shape.delete();
  }
  await context.sync();
  return count;
}

/** Add one tracker part to a slide. */
function addPart(slide: PowerPoint.Slide, part: TrackerPart, color: string, font: string): PowerPoint.Shape {
  const rect = { left: part.left, top: part.top, width: part.width, height: part.height };
  let s: PowerPoint.Shape;
  if (part.kind === "label") {
    s = slide.shapes.addTextBox(part.text ?? "", rect);
    s.textFrame.textRange.font.color = part.active ? color : INACTIVE_TEXT;
    s.textFrame.textRange.font.bold = part.active;
    s.textFrame.textRange.font.size = 8;
  } else {
    s = slide.shapes.addGeometricShape(part.kind === "dot" ? PowerPoint.GeometricShapeType.ellipse : PowerPoint.GeometricShapeType.rectangle, rect);
    if (part.kind === "dot" && !part.active) {
      s.fill.clear();
      s.lineFormat.color = INACTIVE_TEXT;
      s.lineFormat.weight = 0.75;
    } else {
      s.fill.setSolidColor(part.active ? color : INACTIVE_FILL);
      s.lineFormat.visible = false;
    }
    if (part.text) {
      s.textFrame.textRange.text = part.text;
      s.textFrame.textRange.font.color = part.active ? textOn(color) : INACTIVE_TEXT;
      s.textFrame.textRange.font.bold = part.active;
      s.textFrame.textRange.font.size = 9;
      s.textFrame.textRange.paragraphFormat.horizontalAlignment = PowerPoint.ParagraphHorizontalAlignment.center;
      s.textFrame.verticalAlignment = PowerPoint.TextVerticalAlignment.middle;
    }
  }
  if (part.text) {
    s.textFrame.wordWrap = false;
    s.textFrame.autoSizeSetting = PowerPoint.ShapeAutoSize.autoSizeNone;
    s.textFrame.leftMargin = s.textFrame.rightMargin = s.textFrame.topMargin = s.textFrame.bottomMargin = part.kind === "label" ? 0 : 2;
    if (font) s.textFrame.textRange.font.name = font;
    if (part.kind === "label" && part.active && part.width >= 200) s.textFrame.textRange.paragraphFormat.horizontalAlignment = PowerPoint.ParagraphHorizontalAlignment.right;
  }
  s.name = "Retro tracker part";
  s.tags.add(KIND_TAG, TRACKER_KIND);
  return s;
}

/** Make the agenda slide (before the first section) or update its list. */
async function buildAgenda(context: PowerPoint.RequestContext, deck: DeckSlides, font: string): Promise<"added" | "updated"> {
  const names = deck.sections.map((s) => s.name);
  if (deck.agenda >= 0) {
    const slide = deck.slides[deck.agenda];
    const shapes = slide.shapes.load("items/id");
    await context.sync();
    const tags = shapes.items.map((shape) => ({ shape, tag: shape.tags.getItemOrNullObject(KIND_TAG).load("value") }));
    await context.sync();
    const list = tags.find(({ tag }) => !tag.isNullObject && tag.value.toLowerCase() === AGENDA_KIND)?.shape;
    if (list) list.textFrame.textRange.text = agendaText(names);
    else await addAgendaList(context, slide, names, font);
    await context.sync();
    return "updated";
  }

  // Use the template's "Title Only" (or "Title and Content") layout from the first section's master.
  const first = deck.slides[deck.sections[0].start];
  const master = first.slideMaster.load("id");
  const layouts = first.slideMaster.layouts.load("items/id,items/name");
  await context.sync();
  const pick = (re: RegExp) => layouts.items.find((l) => re.test(l.name));
  const layout = pick(/title only/i) ?? pick(/title and content/i) ?? pick(/title/i) ?? layouts.items[0];
  context.presentation.slides.add({ slideMasterId: master.id, layoutId: layout?.id });
  await context.sync();
  const all = context.presentation.slides.load("items/id");
  await context.sync();
  const slide = all.items[all.items.length - 1];
  slide.moveTo(deck.sections[0].start);
  slide.tags.add(AGENDA_TAG, "1");
  const shapes = slide.shapes.load("items/id,items/type");
  await context.sync();
  const placeholders = shapes.items.filter((s) => s.type === "Placeholder").map((shape) => ({ shape, format: shape.placeholderFormat.load("type") }));
  await context.sync();
  const title = placeholders.find((p) => p.format.type === "Title" || p.format.type === "CenterTitle")?.shape;
  if (title) title.textFrame.textRange.text = "Agenda";
  const body = placeholders.find((p) => p.format.type === "Body" || p.format.type === "Content")?.shape;
  if (body) {
    body.textFrame.textRange.text = agendaText(names);
    body.tags.add(KIND_TAG, AGENDA_KIND);
  } else {
    await addAgendaList(context, slide, names, font, !title);
  }
  await context.sync();
  return "added";
}

async function addAgendaList(context: PowerPoint.RequestContext, slide: PowerPoint.Slide, names: string[], font: string, withHeading = false): Promise<void> {
  const size = await slideSize(context);
  const text = withHeading ? `Agenda\n${agendaText(names)}` : agendaText(names);
  const box = slide.shapes.addTextBox(text, { left: 72, top: 140, width: size.width - 144, height: size.height - 200 });
  box.name = "Retro agenda";
  box.textFrame.textRange.font.size = 24;
  if (font) box.textFrame.textRange.font.name = font;
  box.tags.add(KIND_TAG, AGENDA_KIND);
}

/** Build (or rebuild) the tracker on every slide in a section, and the agenda slide if asked. */
export async function applyTracker(settings: TrackerSettings): Promise<string> {
  const kit = activeKit(await brandStore.load());
  const color = kit?.colors[0] ?? FALLBACK_COLOR;
  const font = kit?.bodyFont ?? "";
  return PowerPoint.run(async (context) => {
    let deck = await deckSlides(context);
    if (deck.sections.length === 0) throw new UserError("Add a section first: select the slide where it starts and click Add section.");
    context.presentation.tags.add(TRACKER_TAG, encodeSettings(settings));
    let agendaNote = "";
    if (settings.agenda) {
      agendaNote = (await buildAgenda(context, deck, font)) === "added" ? " Added an agenda slide before the first section." : " Updated the agenda slide.";
      deck = await deckSlides(context);
    }
    await deleteTrackers(context, deck.slides);
    const size = await slideSize(context);
    const names = deck.sections.map((s) => s.name);
    const built: { slide: PowerPoint.Slide; parts: PowerPoint.Shape[] }[] = [];
    deck.slides.forEach((slide, i) => {
      const section = sectionOf(deck.sections, i);
      if (section < 0 || i === deck.agenda || deck.appendix.has(i)) return;
      built.push({ slide, parts: trackerParts(settings.style, settings.position, size, names, section).map((part) => addPart(slide, part, color, font)) });
    });
    await context.sync();
    // One group per slide, so the tracker moves and selects as one piece.
    for (const { slide, parts } of built) {
      if (parts.length < 2) continue;
      const group = slide.shapes.addGroup(parts);
      group.name = "Retro tracker";
      group.tags.add(KIND_TAG, TRACKER_KIND);
    }
    await context.sync();
    const n = built.length;
    return `Tracker on ${n} slide${n === 1 ? "" : "s"}, ${names.length} section${names.length === 1 ? "" : "s"}.${agendaNote}`;
  });
}

/** Rebuild with the deck's saved settings (the ribbon's Refresh tracker). */
export async function refreshTracker(): Promise<string> {
  const state = await readTracker();
  return applyTracker(state.settings);
}

/** Remove the tracker from every slide. The agenda slide stays: it may have your edits. */
export async function removeTracker(): Promise<string> {
  return PowerPoint.run(async (context) => {
    const { slides } = await deckSlides(context);
    const n = await deleteTrackers(context, slides);
    return n ? `Removed the tracker from ${n} slide${n === 1 ? "" : "s"}. The agenda slide stays; delete it yourself if you don't need it.` : "There's no tracker to remove.";
  });
}
