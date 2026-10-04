// The photo picker's model: which photos are in the chosen folder, their previews, and the
// geometry of placing them. Pure functions here are unit tested; reading files and inserting
// into PowerPoint is in photoActions.ts.
//
// The folder is picked in Finder each session (browsers can't keep access to a folder), so the
// list and previews are kept in IndexedDB: browsing works straight away after a restart, and
// only inserting a full-size photo needs the folder picked again ("Reconnect").

import type { Store } from "./library";

export type Placement = "full" | "box" | "background" | "asis";
export type Anchor = "top" | "center" | "bottom";

export interface PhotoEntry {
  /** Path inside the chosen folder, e.g. "Cities/Paris 01.jpg". Unique. */
  id: string;
  name: string;
  /** Subfolder path inside the chosen folder ("" for photos at the top level). */
  folder: string;
  size: number;
  modified: number;
  width: number;
  height: number;
  /** An SVG graphic rather than a photo: inserted as a vector, previewed from the file itself. */
  vector?: boolean;
}

export interface PhotoIndex {
  /** Name of the chosen folder. */
  root: string;
  scanned: string;
  photos: PhotoEntry[];
}

/** What the folder picker returns for each file. */
export interface PickedFile {
  name: string;
  type: string;
  size: number;
  lastModified: number;
  webkitRelativePath: string;
}

const IMAGE = /\.(jpe?g|jpe|jfif|png|gif|webp|avif|heic|heif|tiff?|bmp)$/i;

export function isPhoto(f: { name: string; type: string }): boolean {
  if (f.name.startsWith(".")) return false; // macOS ._ files and other hidden files
  return IMAGE.test(f.name) || (f.type.startsWith("image/") && f.type !== "image/svg+xml");
}

export function isVector(f: { name: string; type: string }): boolean {
  if (f.name.startsWith(".")) return false;
  return /\.svg$/i.test(f.name) || f.type === "image/svg+xml";
}

/** Photos and SVG graphics: everything the picker shows. */
export const isPickable = (f: { name: string; type: string }) => isPhoto(f) || isVector(f);

export interface ScanStats {
  /** Every file the folder picker returned. */
  files: number;
  /** Folders those files are in, counting the chosen folder itself. */
  folders: number;
  /** How many levels of subfolders below the chosen folder held files. */
  depth: number;
  /** Photos and SVG graphics. */
  photos: number;
  /** Files that aren't photos or graphics, by extension, most common first. */
  skipped: { ext: string; count: number }[];
}

/** What the folder picker returned, so a scan can say where photos went missing. */
export function scanStats(files: { name: string; type: string; webkitRelativePath: string }[]): ScanStats {
  const folders = new Set<string>();
  const skipped = new Map<string, number>();
  let depth = 0;
  let photos = 0;
  for (const f of files) {
    const parts = (f.webkitRelativePath || f.name).split("/");
    folders.add(parts.slice(0, -1).join("/"));
    depth = Math.max(depth, parts.length - 2);
    if (isPickable(f)) photos++;
    else if (!f.name.startsWith(".")) {
      const ext = f.name.includes(".") ? f.name.slice(f.name.lastIndexOf(".")).toLowerCase() : "(no extension)";
      skipped.set(ext, (skipped.get(ext) ?? 0) + 1);
    }
  }
  return {
    files: files.length,
    folders: folders.size,
    depth,
    photos,
    skipped: [...skipped].map(([ext, count]) => ({ ext, count })).sort((a, b) => b.count - a.count || a.ext.localeCompare(b.ext)),
  };
}

/** Split "Root/Cities/Paris.jpg" into the root folder name and the path inside it. */
export function splitPath(relative: string): { root: string; path: string } {
  const i = relative.indexOf("/");
  return i < 0 ? { root: "", path: relative } : { root: relative.slice(0, i), path: relative.slice(i + 1) };
}

export function folderOf(path: string): string {
  const i = path.lastIndexOf("/");
  return i < 0 ? "" : path.slice(0, i);
}

/** Which picked files are new or changed, and which known photos are gone. */
export function planScan<F extends PickedFile>(files: F[], known: PhotoEntry[]): { root: string; unchanged: PhotoEntry[]; fresh: F[]; removed: string[] } {
  const photos = files.filter(isPickable);
  const root = splitPath(photos[0]?.webkitRelativePath ?? files[0]?.webkitRelativePath ?? "").root;
  const byId = new Map(known.map((p) => [p.id, p]));
  const seen = new Set<string>();
  const unchanged: PhotoEntry[] = [];
  const fresh: F[] = [];
  for (const f of photos) {
    const id = splitPath(f.webkitRelativePath || f.name).path;
    seen.add(id);
    const old = byId.get(id);
    if (old && old.size === f.size && old.modified === f.lastModified) unchanged.push(old);
    else fresh.push(f);
  }
  return { root, unchanged, fresh, removed: known.filter((p) => !seen.has(p.id)).map((p) => p.id) };
}

export interface FolderInfo {
  name: string;
  /** Path inside the chosen folder. */
  path: string;
  /** Photos and graphics in it and all its subfolders. */
  count: number;
}

/** The subfolders directly inside a folder ("" for the chosen folder), with what's in each. */
export function childFolders(photos: PhotoEntry[], parent: string): FolderInfo[] {
  const prefix = parent ? `${parent}/` : "";
  const counts = new Map<string, number>();
  for (const p of photos) {
    if (parent && !p.folder.startsWith(prefix)) continue;
    const rest = p.folder.slice(prefix.length);
    if (!rest) continue;
    const name = rest.split("/")[0];
    counts.set(name, (counts.get(name) ?? 0) + 1);
  }
  return [...counts]
    .map(([name, count]) => ({ name, path: prefix + name, count }))
    .sort((a, b) => a.name.localeCompare(b.name, undefined, { numeric: true }));
}

export type Kind = "all" | "photos" | "graphics";

/** Photos and graphics in a folder and its subfolders ("" for all), matching a search, sorted by path. */
export function filterPhotos(photos: PhotoEntry[], query: string, folder: string, kind: Kind = "all"): PhotoEntry[] {
  const q = query.trim().toLowerCase();
  return photos
    .filter((p) => !folder || p.folder === folder || p.folder.startsWith(`${folder}/`))
    .filter((p) => kind === "all" || (kind === "graphics") === Boolean(p.vector))
    .filter((p) => !q || p.id.toLowerCase().includes(q))
    .sort((a, b) => a.id.localeCompare(b.id, undefined, { numeric: true }));
}

// --- Geometry ---

export interface Crop {
  sx: number;
  sy: number;
  sw: number;
  sh: number;
}

/**
 * The part of a w×h photo to keep so it fills a target of the given aspect (width / height)
 * without stretching. Height is trimmed from the top, middle or bottom as anchored; width is
 * always trimmed evenly from both sides.
 */
export function cropFor(w: number, h: number, aspect: number, anchor: Anchor): Crop {
  if (w / h > aspect) {
    const sw = h * aspect;
    return { sx: (w - sw) / 2, sy: 0, sw, sh: h };
  }
  const sh = w / aspect;
  const sy = anchor === "top" ? 0 : anchor === "bottom" ? h - sh : (h - sh) / 2;
  return { sx: 0, sy, sw: w, sh };
}

/** Pixel size to save a crop at for a target in points: 2 px per point, never enlarged, at most 3000 px. */
export function outputSize(crop: { sw: number; sh: number }, target: { width: number; height: number }): { width: number; height: number } {
  const scale = Math.min(1, (target.width * 2) / crop.sw, (target.height * 2) / crop.sh, 3000 / Math.max(crop.sw, crop.sh));
  return { width: Math.max(1, Math.round(crop.sw * scale)), height: Math.max(1, Math.round(crop.sh * scale)) };
}

/** Where an uncropped photo goes: as large as fits in 80% of the slide, centered. */
export function asIsRect(w: number, h: number, slide: { width: number; height: number }) {
  const scale = Math.min((slide.width * 0.8) / w, (slide.height * 0.8) / h);
  const round = (v: number) => Math.round(v * 100) / 100;
  const width = w * scale;
  const height = h * scale;
  return { left: round((slide.width - width) / 2), top: round((slide.height - height) / 2), width: round(width), height: round(height) };
}

const UNITS: Record<string, number> = { "": 1, px: 1, pt: 4 / 3, pc: 16, mm: 96 / 25.4, cm: 96 / 2.54, in: 96 };

/** An SVG's size in pixels from its width, height and viewBox attributes (100×100 if none). */
export function svgSize(width: string | null, height: string | null, viewBox: string | null): { width: number; height: number } {
  const length = (v: string | null) => {
    const m = v?.trim().match(/^([\d.]+)\s*(px|pt|pc|mm|cm|in)?$/i);
    return m ? Number(m[1]) * UNITS[(m[2] ?? "").toLowerCase()] : undefined;
  };
  const box = viewBox?.trim().split(/[\s,]+/).map(Number);
  const vb = box?.length === 4 && box[2] > 0 && box[3] > 0 ? { width: box[2], height: box[3] } : undefined;
  let w = length(width);
  let h = length(height);
  if (w && !h && vb) h = (w * vb.height) / vb.width;
  if (h && !w && vb) w = (h * vb.width) / vb.height;
  if (w && h) return { width: w, height: h };
  return vb ?? { width: 100, height: 100 };
}

/** A w×h graphic scaled to fit `fraction` of an area's width and height, centered in it. */
export function fitRect(w: number, h: number, area: { left: number; top: number; width: number; height: number }, fraction: number) {
  const round = (v: number) => Math.round(v * 100) / 100;
  const scale = Math.min((area.width * fraction) / w, (area.height * fraction) / h);
  const width = w * scale;
  const height = h * scale;
  return { left: round(area.left + (area.width - width) / 2), top: round(area.top + (area.height - height) / 2), width: round(width), height: round(height) };
}

/** Whether a shape covers at least `fraction` of the slide (the part of it on the slide counts). */
export function coversSlide(shape: { left: number; top: number; width: number; height: number }, slide: { width: number; height: number }, fraction = 0.9): boolean {
  const w = Math.min(shape.left + shape.width, slide.width) - Math.max(shape.left, 0);
  const h = Math.min(shape.top + shape.height, slide.height) - Math.max(shape.top, 0);
  return w > 0 && h > 0 && w * h >= fraction * slide.width * slide.height;
}

/** Pictures keep a format that can hold transparency; everything else becomes JPEG. */
export function outputType(name: string): "image/png" | "image/jpeg" {
  return /\.(png|gif|webp)$/i.test(name) ? "image/png" : "image/jpeg";
}

// --- Storage ---

const INDEX = "photo.index";
const thumbKey = (id: string) => `photo.thumb.${id}`;

export function createPhotoStore(store: Store) {
  return {
    index: () => store.get<PhotoIndex>(INDEX),
    saveIndex: (index: PhotoIndex) => store.set(INDEX, index),
    thumb: (id: string) => store.get<ArrayBuffer>(thumbKey(id)),
    saveThumb: (id: string, data: ArrayBuffer) => store.set(thumbKey(id), data),
    deleteThumb: (id: string) => store.del(thumbKey(id)),
  };
}
