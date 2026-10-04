// The shape and slide libraries' PowerPoint side: saving the selection or slides, inserting items
// and tidying up helper slides. Storage is in library.ts.

import { kv } from "./idb";
import { createLibrary, placement, type LibraryItem } from "./library";
import { KIND_TAG } from "./notes";
import { currentSlide, slideSize, UserError } from "./ppt";
import { captureShapes, HELPER_TAG, insertHelperSlide, rebuildIssues, rebuildShapes, type Recipe } from "./shapeCapture";
import { cleanSlide, stripToCaptured } from "./slideStrip";
import { AGENDA_TAG, SECTION_TAG } from "./tracker";
import { base64ToBytes, bytesToBase64 } from "./zip";

export const library = createLibrary(kv);

const PREVIEW = { width: 240, height: 160 };

function loadImage(src: string): Promise<HTMLImageElement> {
  return new Promise((resolve, reject) => {
    const img = new Image();
    img.onload = () => resolve(img);
    img.onerror = () => reject(new Error("Couldn't draw the preview."));
    img.src = src;
  });
}

/** A preview picture: each selected shape rendered by PowerPoint, placed as on the slide. */
async function makePreview(context: PowerPoint.RequestContext, shapes: PowerPoint.Shape[], recipe: Recipe): Promise<string | undefined> {
  try {
    const box = recipe.box;
    const scale = Math.min(PREVIEW.width / Math.max(box.width, 1), PREVIEW.height / Math.max(box.height, 1), 2);
    const images = shapes.map((s, i) => {
      const f = recipe.nodes[i]?.frame;
      const width = Math.max(1, Math.round((f?.width ?? box.width) * scale));
      return s.getImageAsBase64({ format: PowerPoint.ShapeGetImageFormatType.png, width });
    });
    await context.sync();
    if (shapes.length === 1) return `data:image/png;base64,${images[0].value}`;
    const canvas = document.createElement("canvas");
    canvas.width = Math.max(1, Math.round(box.width * scale));
    canvas.height = Math.max(1, Math.round(box.height * scale));
    const g = canvas.getContext("2d")!;
    for (const [i, img] of images.entries()) {
      const f = recipe.nodes[i]?.frame;
      if (!f) continue;
      const pic = await loadImage(`data:image/png;base64,${img.value}`);
      // Center each picture on its shape, since rotated shapes render to their turned bounds.
      const cx = (f.left + f.width / 2 - box.left) * scale;
      const cy = (f.top + f.height / 2 - box.top) * scale;
      g.drawImage(pic, cx - pic.width / 2, cy - pic.height / 2);
    }
    return canvas.toDataURL("image/png");
  } catch {
    return undefined; // A missing preview shouldn't stop a save.
  }
}

/** Save the selected shapes to the library. */
export async function saveSelection(name: string, category: string): Promise<LibraryItem> {
  if (!Office.context.requirements.isSetSupported("PowerPointApi", "1.10")) {
    throw new UserError("The shape library needs PowerPoint 16.105 or later.");
  }
  const captured = await PowerPoint.run(async (context) => {
    const slide = await currentSlide(context);
    const selected = context.presentation.getSelectedShapes();
    selected.load("items/id");
    await context.sync();
    if (selected.items.length === 0) throw new UserError("Select the shapes to save first.");
    const capture = await captureShapes(context, slide, selected.items);
    if (capture.found !== capture.selected) {
      throw new UserError("Retro couldn't find every selected shape in the slide. Select whole shapes or groups (not shapes inside a group) and try again.");
    }
    const preview = await makePreview(context, selected.items, capture.recipe);
    return { capture, preview };
  });
  const { recipe, pptx } = captured.capture;
  const issues = [...new Set(recipe.nodes.flatMap(rebuildIssues))];
  // Keep only the saved shapes, not the rest of the slide (other shapes, notes, the deck's properties).
  const stripped = await stripToCaptured(base64ToBytes(pptx));
  return library.save({ name, category, recipe, pptx: stripped, issues, preview: captured.preview });
}

let stripping: Promise<void> | undefined;

/**
 * Strip the stored slides of items saved before Retro did that on save (or restored from a
 * backup): shapes keep only their saved shapes, saved slides lose comments and the deck's
 * properties. An item that fails is left as it was and tried again next time.
 */
export function stripOldItems(): Promise<void> {
  stripping ??= (async () => {
    for (const item of await library.list()) {
      if (item.stripped) continue;
      try {
        if (item.kind === "slides") {
          const [first, ...more] = await Promise.all((await library.slides(item.id)).map(cleanSlide));
          await library.replacePptx(item.id, first, more);
        } else {
          await library.replacePptx(item.id, await stripToCaptured(await library.pptx(item.id)));
        }
      } catch {
        // Leave it; inserting still works.
      }
    }
  })().finally(() => (stripping = undefined));
  return stripping;
}

// Whether a helper slide might be in the deck. True at first, to tidy any left from last time.
let helpersMayExist = true;
let tidying: Promise<void> | undefined;

/**
 * Delete helper slides whose shapes have been cut away. If the empty helper is the slide being
 * shown, go back to the slide it was made for, ready to paste. Cheap when there are none.
 */
export function cleanupHelperSlides(): Promise<void> {
  if (!helpersMayExist) return Promise.resolve();
  tidying ??= tidyHelpers().finally(() => (tidying = undefined));
  return tidying;
}

async function tidyHelpers(): Promise<void> {
  await PowerPoint.run(async (context) => {
    const slides = context.presentation.slides;
    slides.load("items/id");
    await context.sync();
    const tags = slides.items.map((s) => s.tags.getItemOrNullObject(HELPER_TAG).load("value"));
    await context.sync();
    const helpers = slides.items.filter((_, i) => !tags[i].isNullObject);
    if (helpers.length === 0) {
      helpersMayExist = false;
      return;
    }
    const counts = helpers.map((s) => s.shapes.getCount());
    const shown = context.presentation.getSelectedSlides();
    shown.load("items/id");
    await context.sync();
    const shownId = shown.items[0]?.id;
    let back: string | undefined;
    helpers.forEach((s, i) => {
      if (counts[i].value > 0) return;
      if (s.id === shownId) {
        const origin = tags[slides.items.indexOf(s)].value.toLowerCase();
        back = slides.items.find((x) => x.id.toLowerCase() === origin)?.id;
      }
      s.delete();
    });
    await context.sync();
    if (back) {
      context.presentation.setSelectedSlides([back]);
      await context.sync();
    }
  });
}

/** Insert an item where it was saved. Returns which route it took. */
export async function insertItem(item: LibraryItem): Promise<LibraryItem["route"]> {
  await cleanupHelperSlides();
  const recipe = item.route === "click" ? await library.recipe(item.id) : undefined;
  const pptx = item.route === "helper" ? bytesToBase64(await library.pptx(item.id)) : undefined;
  await PowerPoint.run(async (context) => {
    const slide = await currentSlide(context);
    if (recipe) {
      const { dx, dy } = placement(item.box, await slideSize(context));
      await rebuildShapes(context, slide, recipe, dx, dy);
    } else {
      helpersMayExist = true;
      await insertHelperSlide(context, pptx!, slide.id);
    }
  });
  await library.markUsed(item.id);
  return item.route;
}

/** Save the selected slides (one or more, in deck order) to the slide library as one item. */
export async function saveSlides(name: string, category: string): Promise<LibraryItem> {
  if (!Office.context.requirements.isSetSupported("PowerPointApi", "1.8")) {
    throw new UserError("The slide library needs PowerPoint 16.96 or later.");
  }
  const exported = await PowerPoint.run(async (context) => {
    const all = context.presentation.slides.load("items/id");
    const selected = context.presentation.getSelectedSlides().load("items/id");
    await context.sync();
    const chosen = new Set(selected.items.map((s) => s.id));
    const slides = all.items.filter((s) => chosen.has(s.id));
    if (slides.length === 0) throw new UserError("Select the slides to save first (click them in the slide thumbnails).");
    const files = slides.map((s) => s.exportAsBase64());
    const image = slides[0].getImageAsBase64({ width: 480 });
    await context.sync();
    return { files: files.map((f) => f.value), image: image.value };
  });
  // Keep everything on the slides and their speaker notes; drop comments and the deck's properties.
  const [first, ...more] = await Promise.all(exported.files.map((f) => cleanSlide(base64ToBytes(f))));
  const size = await PowerPoint.run(async (context) => slideSize(context));
  const recipe: Recipe = { version: 1, box: { left: 0, top: 0, ...size }, nodes: [] };
  return library.save({ name, category, recipe, pptx: first, more, issues: [], preview: `data:image/png;base64,${exported.image}`, kind: "slides" });
}

/**
 * Insert a saved slide item after the current slide, in this deck's theme. Retro's own section
 * marks and tracker don't come along: they belong to the deck the slides were saved from.
 */
export async function insertSlides(item: LibraryItem): Promise<number> {
  const files = (await library.slides(item.id)).map(bytesToBase64);
  const added = await PowerPoint.run(async (context) => {
    const current = await currentSlide(context);
    const before = context.presentation.slides.load("items/id");
    await context.sync();
    const existing = new Set(before.items.map((s) => s.id));
    // Each goes right after the current slide, so inserting the last first keeps their order.
    for (const file of [...files].reverse()) {
      context.presentation.insertSlidesFromBase64(file, { formatting: PowerPoint.InsertSlideFormatting.useDestinationTheme, targetSlideId: current.id });
      await context.sync();
    }
    const after = context.presentation.slides.load("items/id");
    await context.sync();
    const fresh = after.items.filter((s) => !existing.has(s.id));
    const tags = fresh.map((s) => ({ slide: s, section: s.tags.getItemOrNullObject(SECTION_TAG).load("value"), agenda: s.tags.getItemOrNullObject(AGENDA_TAG).load("value"), shapes: s.shapes.load("items/id") }));
    await context.sync();
    const kinds = tags.map((t) => t.shapes.items.map((shape) => ({ shape, kind: shape.tags.getItemOrNullObject(KIND_TAG).load("value") })));
    await context.sync();
    tags.forEach((t, i) => {
      if (!t.section.isNullObject) t.slide.tags.delete(SECTION_TAG);
      if (!t.agenda.isNullObject) t.slide.tags.delete(AGENDA_TAG);
      for (const { shape, kind } of kinds[i]) if (!kind.isNullObject && kind.value.toLowerCase() === "tracker") shape.delete();
    });
    await context.sync();
    if (fresh[0]) context.presentation.setSelectedSlides([fresh[0].id]);
    await context.sync();
    return fresh.length;
  });
  await library.markUsed(item.id);
  return added;
}
