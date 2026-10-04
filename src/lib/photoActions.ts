// The photo picker's file reading and PowerPoint inserts. The model and geometry are in photos.ts.

import { kv } from "./idb";
import { isVector, asIsRect, coversSlide, createPhotoStore, scanStats, svgSize, fitRect, type ScanStats, cropFor, outputSize, outputType, planScan, splitPath, folderOf, type Anchor, type PhotoEntry, type PhotoIndex, type Placement } from "./photos";
import { currentSlide, slideSize, UserError } from "./ppt";

export const photoStore = createPhotoStore(kv);

/** Files from the folder picked this session, by path. Empty after PowerPoint restarts. */
const connected = new Map<string, File>();

export const isConnected = () => connected.size > 0;

const THUMB = 240;

function loadImage(file: Blob): Promise<HTMLImageElement> {
  const url = URL.createObjectURL(file);
  return new Promise<HTMLImageElement>((resolve, reject) => {
    const img = new Image();
    img.onload = () => resolve(img);
    img.onerror = () => reject(new UserError("PowerPoint can't read this photo's format."));
    img.src = url;
  }).finally(() => URL.revokeObjectURL(url));
}

function canvasBlob(canvas: HTMLCanvasElement, type: string, quality?: number): Promise<Blob> {
  return new Promise((resolve, reject) => canvas.toBlob((b) => (b ? resolve(b) : reject(new Error("Couldn't encode the photo."))), type, quality));
}

/** Draw part of an image into a new canvas of the given size. */
function draw(img: HTMLImageElement, crop: { sx: number; sy: number; sw: number; sh: number }, size: { width: number; height: number }) {
  const canvas = document.createElement("canvas");
  canvas.width = size.width;
  canvas.height = size.height;
  const g = canvas.getContext("2d")!;
  g.imageSmoothingQuality = "high";
  g.drawImage(img, crop.sx, crop.sy, crop.sw, crop.sh, 0, 0, size.width, size.height);
  return canvas;
}

/** SVGs larger than this aren't previewed (they're rarely icons and slow to draw). */
const MAX_SVG = 2 * 1024 * 1024;

/** An SVG's size in pixels, from its own width, height and viewBox. */
function svgDimensions(text: string): { width: number; height: number } {
  const root = new DOMParser().parseFromString(text, "image/svg+xml").documentElement;
  if (root.localName !== "svg") throw new Error("not an SVG");
  return svgSize(root.getAttribute("width"), root.getAttribute("height"), root.getAttribute("viewBox"));
}

export interface ScanProgress {
  done: number;
  total: number;
}

export interface ScanResult {
  index: PhotoIndex;
  added: number;
  removed: number;
  unreadable: number;
  stats: ScanStats;
}

/**
 * Read a folder picked in Finder: make previews for new or changed photos, forget removed ones,
 * and remember the files for inserting this session. Picking the same folder again is quick.
 */
export async function scanFolder(files: File[], progress: (p: ScanProgress) => void): Promise<ScanResult> {
  const previous = await photoStore.index();
  const root = splitPath(files[0]?.webkitRelativePath ?? "").root;
  const sameFolder = previous?.root === root;
  const plan = planScan(files, sameFolder ? previous!.photos : []);
  if (plan.unchanged.length + plan.fresh.length === 0) throw new UserError(`There are no photos in “${root || "that folder"}”.`);

  connected.clear();
  for (const f of files) connected.set(splitPath(f.webkitRelativePath || f.name).path, f);

  const photos = [...plan.unchanged];
  let unreadable = 0;
  progress({ done: 0, total: plan.fresh.length });
  for (const [i, file] of plan.fresh.entries()) {
    const id = splitPath(file.webkitRelativePath || file.name).path;
    try {
      if (isVector(file)) {
        if (file.size > MAX_SVG) throw new Error("too large to preview");
        const size = svgDimensions(await file.text());
        await photoStore.saveThumb(id, await file.arrayBuffer()); // an SVG previews itself
        photos.push({ id, name: file.name, folder: folderOf(id), size: file.size, modified: file.lastModified, ...size, vector: true });
        continue;
      }
      const img = await loadImage(file);
      const scale = Math.min(1, THUMB / Math.max(img.naturalWidth, img.naturalHeight));
      const canvas = draw(img, { sx: 0, sy: 0, sw: img.naturalWidth, sh: img.naturalHeight }, { width: Math.max(1, Math.round(img.naturalWidth * scale)), height: Math.max(1, Math.round(img.naturalHeight * scale)) });
      await photoStore.saveThumb(id, await (await canvasBlob(canvas, "image/jpeg", 0.8)).arrayBuffer());
      photos.push({ id, name: file.name, folder: folderOf(id), size: file.size, modified: file.lastModified, width: img.naturalWidth, height: img.naturalHeight });
    } catch {
      unreadable++; // e.g. a format WebKit can't decode, or a cloud file that isn't downloaded
    } finally {
      if (i % 10 === 9 || i === plan.fresh.length - 1) progress({ done: i + 1, total: plan.fresh.length });
    }
  }

  const gone = sameFolder ? plan.removed : (previous?.photos ?? []).map((p) => p.id);
  for (const id of gone) await photoStore.deleteThumb(id);
  const index: PhotoIndex = { root, scanned: new Date().toISOString(), photos };
  await photoStore.saveIndex(index);
  return { index, added: plan.fresh.length - unreadable, removed: gone.length, unreadable, stats: scanStats(files) };
}

/** The photo cropped and sized for a target (points), as base64 without the data: prefix. */
async function prepare(
  photo: PhotoEntry,
  target: { width: number; height: number } | undefined,
  anchor: Anchor,
  slide: { width: number; height: number },
): Promise<{ base64: string; width: number; height: number }> {
  const file = connected.get(photo.id);
  if (!file) throw new UserError("Reconnect your photo folder to insert photos: click Reconnect and choose the same folder.");
  const img = await loadImage(file);
  const w = img.naturalWidth;
  const h = img.naturalHeight;
  const crop = target ? cropFor(w, h, target.width / target.height, anchor) : { sx: 0, sy: 0, sw: w, sh: h };
  const size = outputSize(crop, target ?? asIsRect(w, h, slide));
  const type = outputType(photo.name);
  const blob = await canvasBlob(draw(img, crop, size), type, type === "image/jpeg" ? 0.88 : undefined);
  const bytes = new Uint8Array(await blob.arrayBuffer());
  let bin = "";
  for (let i = 0; i < bytes.length; i += 0x8000) bin += String.fromCharCode(...bytes.subarray(i, i + 0x8000));
  return { base64: btoa(bin), width: w, height: h };
}

/**
 * Insert a picture (base64) or SVG graphic (text) through Office's image coercion, which makes a
 * real picture or graphic rather than a picture-filled shape.
 */
function insertPicture(data: string, rect: { left: number; top: number; width: number; height: number }, svg = false): Promise<void> {
  return new Promise((resolve, reject) =>
    Office.context.document.setSelectedDataAsync(
      data,
      { coercionType: svg ? Office.CoercionType.XmlSvg : Office.CoercionType.Image, imageLeft: rect.left, imageTop: rect.top, imageWidth: rect.width, imageHeight: rect.height },
      (r) => (r.status === Office.AsyncResultStatus.Succeeded ? resolve() : reject(new Error(r.error?.message ?? "PowerPoint couldn't insert the photo."))),
    ),
  );
}

/** Clear the shape selection so a new picture doesn't replace or land in a selected shape. */
async function deselect(): Promise<void> {
  await PowerPoint.run(async (context) => {
    const slide = await currentSlide(context);
    slide.setSelectedShapes([]);
    await context.sync();
  }).catch(() => undefined);
}

/** Name the just-inserted (selected) picture, and send it to the back for a full bleed. */
async function finishPicture(name: string, toBack: boolean): Promise<void> {
  await PowerPoint.run(async (context) => {
    const shapes = context.presentation.getSelectedShapes();
    shapes.load("items/id");
    await context.sync();
    const pic = shapes.items[0];
    if (!pic) return;
    pic.name = name;
    if (toBack) pic.setZOrder(PowerPoint.ShapeZOrder.sendToBack);
    await context.sync();
  });
}

/** Place a photo. Returns a short note for the pane. */
export async function insertPhoto(photo: PhotoEntry, placement: Placement, anchor: Anchor): Promise<string> {
  const name = `Retro photo ${photo.name}`;
  const slide = await PowerPoint.run(async (context) => slideSize(context));
  if (photo.vector) return insertGraphic(photo, slide);

  if (placement === "box") {
    const target = await PowerPoint.run(async (context) => {
      const sel = context.presentation.getSelectedShapes();
      sel.load("items/id,items/type,items/width,items/height");
      await context.sync();
      if (sel.items.length !== 1) throw new UserError("Select one shape to fill with the photo (a rectangle, circle or any other shape).");
      const s = sel.items[0];
      if (s.type === "Image" || s.type === "Group" || s.type === "Line" || s.type === "Table") throw new UserError("Fill box works on shapes like rectangles and circles. Select one of those.");
      return { width: s.width, height: s.height };
    });
    const pic = await prepare(photo, target, anchor, slide);
    await PowerPoint.run(async (context) => {
      const s = context.presentation.getSelectedShapes().getItemAt(0);
      s.fill.setImage(pic.base64);
      await context.sync();
    });
    return `Filled the shape with “${photo.name}”.`;
  }

  if (placement === "background") {
    if (!Office.context.requirements.isSetSupported("PowerPointApi", "1.10")) throw new UserError("Backgrounds need PowerPoint 16.105 or later. Use Full bleed instead.");
    const pic = await prepare(photo, slide, anchor, slide);
    const removed = await PowerPoint.run(async (context) => {
      const s = await currentSlide(context);
      s.background.isMasterBackgroundFollowed = false;
      // Replace the template's background completely: its layout graphics (patterns, logos)
      // would otherwise still draw on top, like PowerPoint's "Hide background graphics".
      s.background.areBackgroundGraphicsHidden = true;
      s.background.fill.setPictureOrTextureFill({ imageBase64: pic.base64, transparency: 0 });
      // A full-slide picture on the slide itself (common on title slides) would hide the new
      // background, so it goes too: pictures, filled picture placeholders and picture-filled shapes.
      s.shapes.load("items/name,items/type,items/left,items/top,items/width,items/height");
      await context.sync();
      const covering = s.shapes.items.filter((sh) => coversSlide(sh, slide));
      const fills = covering.map((sh) => (sh.type === "GeometricShape" ? sh.fill.load("type") : undefined));
      const placeholders = covering.map((sh) => (sh.type === "Placeholder" ? sh.placeholderFormat.load("containedType") : undefined));
      await context.sync();
      const pictures = covering.filter((sh, i) => sh.type === "Image" || fills[i]?.type === "PictureAndTexture" || placeholders[i]?.containedType === "Image");
      for (const sh of pictures) sh.delete();
      await context.sync();
      // Deleting a filled picture placeholder leaves it empty ("Click to insert picture") over
      // the background, so empty full-slide picture placeholders go as well.
      const after = context.presentation.slides.getItem(s.id).shapes;
      after.load("items/type,items/left,items/top,items/width,items/height");
      await context.sync();
      const empty = after.items.filter((sh) => sh.type === "Placeholder" && coversSlide(sh, slide));
      const formats = empty.map((sh) => sh.placeholderFormat.load("type,containedType"));
      await context.sync();
      empty.forEach((sh, i) => formats[i].type === "Picture" && formats[i].containedType === null && sh.delete());
      await context.sync();
      return pictures.map((sh) => sh.name);
    });
    const names = removed.map((n) => `“${n}”`).join(", ");
    const note = removed.length ? ` Removed ${removed.length === 1 ? "the full-slide picture" : "the full-slide pictures"} that covered it (${names}); ⌘Z undoes this.` : "";
    return `Set “${photo.name}” as this slide's background, with the template's background graphics hidden on this slide.${note}`;
  }

  const pic = await prepare(photo, placement === "full" ? slide : undefined, anchor, slide);
  const rect = placement === "full" ? { left: 0, top: 0, ...slide } : asIsRect(pic.width, pic.height, slide);
  await deselect();
  await insertPicture(pic.base64, rect);
  await finishPicture(name, placement === "full");
  return placement === "full" ? `Placed “${photo.name}” full bleed, behind everything on the slide.` : `Inserted “${photo.name}”.`;
}

/**
 * Place an SVG graphic: inside the selected shape when one is selected (fitted, centered),
 * otherwise centered on the slide at 40% of its size. Placement and crop don't apply to graphics.
 */
async function insertGraphic(photo: PhotoEntry, slide: { width: number; height: number }): Promise<string> {
  const file = connected.get(photo.id);
  if (!file) throw new UserError("Reconnect your photo folder to insert graphics: click Reconnect and choose the same folder.");
  const box = await PowerPoint.run(async (context) => {
    const sel = context.presentation.getSelectedShapes();
    sel.load("items/left,items/top,items/width,items/height");
    await context.sync();
    return sel.items.length === 1 ? sel.items[0] : undefined;
  });
  // Icons come in at a usable size: 80% of the selected shape, or 40% of the slide.
  const rect = box ? fitRect(photo.width, photo.height, box, 0.8) : fitRect(photo.width, photo.height, { left: 0, top: 0, ...slide }, 0.4);
  await deselect();
  await insertPicture(await file.text(), rect, true);
  await finishPicture(`Retro graphic ${photo.name}`, false);
  return box ? `Placed “${photo.name}” inside the selected shape.` : `Inserted “${photo.name}”.`;
}
