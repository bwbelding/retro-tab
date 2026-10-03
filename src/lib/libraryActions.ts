// The shape library's PowerPoint side: saving the selection, inserting items and tidying up
// helper slides. Storage is in library.ts.

import { kv } from "./idb";
import { createLibrary, placement, type LibraryItem } from "./library";
import { currentSlide, slideSize, UserError } from "./ppt";
import { captureShapes, HELPER_TAG, insertHelperSlide, rebuildIssues, rebuildShapes, type Recipe } from "./shapeCapture";
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
  return library.save({ name, category, recipe, pptx: base64ToBytes(pptx), issues, preview: captured.preview });
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
