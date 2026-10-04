// The shape library's storage. Items live in IndexedDB, shared by both ribbon tabs.
//
// Every item keeps the exported slide it was saved from, so it can always be inserted as an
// exact copy. Those files are stored part by part, keyed by content: the template's masters,
// layouts and pictures are the same for every shape saved from one deck, so they're kept once.

import type { Recipe } from "./shapeCapture";
import { base64ToBytes, bytesToBase64, openZip, writeZip, type RawEntry } from "./zip";

export interface Box {
  left: number;
  top: number;
  width: number;
  height: number;
}

export interface LibraryItem {
  id: string;
  name: string;
  category: string;
  favorite: boolean;
  created: string;
  /** When it was last inserted. */
  used?: string;
  uses: number;
  /** "click" when Retro can rebuild it exactly; otherwise it uses the helper slide. */
  route: "click" | "helper";
  /** Why it needs the helper slide. */
  issues: string[];
  /** Where it sat on the slide it was saved from, in points. */
  box: Box;
  /** PNG data URL. */
  preview?: string;
  /** Its stored slide holds only the saved shapes (see slideStrip.ts). Older items get stripped on load. */
  stripped?: boolean;
  /** "slides" for whole slides saved to the slide library; shapes otherwise. */
  kind?: "slides";
  /** Slide library items: how many slides. */
  slides?: number;
}

interface StoredPart {
  method: number;
  crc: number;
  size: number;
  data: Uint8Array;
}

interface ItemData {
  recipe: Recipe;
  /** The exported slide's zip entries, in order, by content hash. */
  parts: { name: string; hash: string }[];
  /** Slide library items with more than one slide: the other slides, the same way. */
  more?: { name: string; hash: string }[][];
}

/** Every slide's parts in an item. */
const allParts = (data: ItemData) => [data.parts, ...(data.more ?? [])];

export interface Store {
  get<T>(key: string): Promise<T | undefined>;
  set(key: string, value: unknown): Promise<void>;
  del(key: string): Promise<void>;
}

const INDEX = "lib.index";
const REFS = "lib.refs";
const dataKey = (id: string) => `lib.data.${id}`;
const partKey = (hash: string) => `lib.part.${hash}`;

async function sha256(bytes: Uint8Array): Promise<string> {
  const digest = await crypto.subtle.digest("SHA-256", bytes as BufferSource);
  return Array.from(new Uint8Array(digest), (b) => b.toString(16).padStart(2, "0")).join("");
}

// --- Pure helpers (unit tested) ---

export type Filter = "all" | "favorites" | `category:${string}`;

export function categoriesOf(items: LibraryItem[]): string[] {
  return [...new Set(items.map((i) => i.category).filter(Boolean))].sort((a, b) => a.localeCompare(b));
}

/** Items matching a search and filter, most recently used (or saved) first. */
export function filterItems(items: LibraryItem[], query: string, filter: Filter): LibraryItem[] {
  const q = query.trim().toLowerCase();
  return items
    .filter((i) => filter === "all" || (filter === "favorites" ? i.favorite : i.category === filter.slice("category:".length)))
    .filter((i) => !q || i.name.toLowerCase().includes(q) || i.category.toLowerCase().includes(q))
    .sort((a, b) => (b.used ?? b.created).localeCompare(a.used ?? a.created));
}

/** How far to move a saved box so it lands where it was saved, kept on a slide of this size. */
export function placement(box: Box, slide: { width: number; height: number }): { dx: number; dy: number } {
  const fit = (start: number, size: number, max: number) => (size >= max ? -start : Math.min(Math.max(start, 0), max - size) - start);
  return { dx: fit(box.left, box.width, slide.width), dy: fit(box.top, box.height, slide.height) };
}

// --- Storage ---

export interface NewItem {
  name: string;
  category: string;
  recipe: Recipe;
  /** The exported slide (.pptx) the shapes were captured from. */
  pptx: Uint8Array;
  issues: string[];
  preview?: string;
  /** Slide library items: "slides", with any further slides after the first in `more`. */
  kind?: "slides";
  more?: Uint8Array[];
}

/** A backup file's contents: the library and Retro's settings, as JSON. */
export interface Backup {
  retro: "backup";
  version: 1;
  exported: string;
  items: LibraryItem[];
  data: Record<string, ItemData>;
  parts: Record<string, Omit<StoredPart, "data"> & { data: string }>;
  settings: Record<string, unknown>;
}

/** Settings (kv keys) that go into backups. */
export const SETTING_KEYS = ["pref.matchReference", "pref.photo", "pref.checkRulesOff", "brand.kits"];

/** Version 1.5.0 stored the star as "favourite"; read it as "favorite". */
function upgrade(item: LibraryItem & { favourite?: boolean }): LibraryItem {
  if (item.favourite === undefined) return item;
  const { favourite, ...rest } = item;
  return { ...rest, favorite: rest.favorite ?? favourite };
}

export function createLibrary(store: Store) {
  const list = async () => ((await store.get<LibraryItem[]>(INDEX)) ?? []).map(upgrade);
  const writeIndex = (items: LibraryItem[]) => store.set(INDEX, items);

  // Changes run one at a time: they read and rewrite the index and the part counts, so two at once
  // (stripping old items while you save one) could lose an update and delete a part still in use.
  let queue: Promise<unknown> = Promise.resolve();
  const exclusive = <A extends unknown[], R>(fn: (...args: A) => Promise<R>) => (...args: A): Promise<R> => {
    const run = queue.then(() => fn(...args));
    queue = run.catch(() => undefined);
    return run;
  };

  /** Store a .pptx's entries by content, counting references. */
  async function storeParts(pptx: Uint8Array, refs: Record<string, number>): Promise<ItemData["parts"]> {
    const zip = openZip(pptx);
    const parts: ItemData["parts"] = [];
    for (const name of zip.names) {
      const raw = zip.raw(name)!;
      const hash = await sha256((await zip.bytes(name))!);
      if (!refs[hash] || !(await store.get(partKey(hash)))) {
        const part: StoredPart = { method: raw.method, crc: raw.crc, size: raw.size, data: raw.data.slice() };
        await store.set(partKey(hash), part);
      }
      refs[hash] = (refs[hash] ?? 0) + 1;
      parts.push({ name, hash });
    }
    return parts;
  }

  /** Drop references to parts, deleting those nothing else uses. */
  async function releaseParts(parts: ItemData["parts"], refs: Record<string, number>): Promise<void> {
    for (const { hash } of parts) {
      refs[hash] = (refs[hash] ?? 1) - 1;
      if (refs[hash] <= 0) {
        delete refs[hash];
        await store.del(partKey(hash));
      }
    }
  }

  async function slidesOf(id: string): Promise<Uint8Array[]> {
    const data = await store.get<ItemData>(dataKey(id));
    if (!data) throw new Error("This library item's data is missing.");
    const files: Uint8Array[] = [];
    for (const parts of allParts(data)) {
      const entries: RawEntry[] = [];
      for (const { name, hash } of parts) {
        const part = await store.get<StoredPart>(partKey(hash));
        if (!part) throw new Error(`Part of this library item is missing (${name}).`);
        entries.push({ name, ...part });
      }
      files.push(writeZip(entries));
    }
    return files;
  }

  return {
    list,

    save: exclusive(async (item: NewItem): Promise<LibraryItem> => {
      const refs = (await store.get<Record<string, number>>(REFS)) ?? {};
      const parts = await storeParts(item.pptx, refs);
      const more: ItemData["parts"][] = [];
      for (const pptx of item.more ?? []) more.push(await storeParts(pptx, refs));
      const saved: LibraryItem = {
        id: crypto.randomUUID(),
        name: item.name.trim() || (item.kind ? "Slide" : "Shape"),
        category: item.category.trim(),
        favorite: false,
        created: new Date().toISOString(),
        uses: 0,
        route: item.issues.length ? "helper" : "click",
        issues: item.issues,
        box: item.recipe.box,
        preview: item.preview,
        stripped: true,
        ...(item.kind ? { kind: item.kind, slides: 1 + more.length } : {}),
      };
      await store.set(dataKey(saved.id), { recipe: item.recipe, parts, ...(more.length ? { more } : {}) } satisfies ItemData);
      await store.set(REFS, refs);
      await writeIndex([saved, ...(await list())]);
      return saved;
    }),

    update: exclusive(async (id: string, patch: Partial<Pick<LibraryItem, "name" | "category" | "favorite">>): Promise<void> => {
      await writeIndex((await list()).map((i) => (i.id === id ? { ...i, ...patch } : i)));
    }),

    markUsed: exclusive(async (id: string): Promise<void> => {
      const now = new Date().toISOString();
      await writeIndex((await list()).map((i) => (i.id === id ? { ...i, used: now, uses: i.uses + 1 } : i)));
    }),

    remove: exclusive(async (id: string): Promise<void> => {
      const data = await store.get<ItemData>(dataKey(id));
      const refs = (await store.get<Record<string, number>>(REFS)) ?? {};
      for (const parts of data ? allParts(data) : []) await releaseParts(parts, refs);
      await store.set(REFS, refs);
      await store.del(dataKey(id));
      await writeIndex((await list()).filter((i) => i.id !== id));
    }),

    /** Swap an item's stored slides for new ones (used to strip older items), and mark it stripped. */
    replacePptx: exclusive(async (id: string, pptx: Uint8Array, morePptx: Uint8Array[] = []): Promise<void> => {
      const data = await store.get<ItemData>(dataKey(id));
      if (!data) return;
      const refs = (await store.get<Record<string, number>>(REFS)) ?? {};
      const parts = await storeParts(pptx, refs);
      const more: ItemData["parts"][] = [];
      for (const extra of morePptx) more.push(await storeParts(extra, refs));
      for (const old of allParts(data)) await releaseParts(old, refs);
      await store.set(REFS, refs);
      await store.set(dataKey(id), { recipe: data.recipe, parts, ...(more.length ? { more } : {}) } satisfies ItemData);
      await writeIndex((await list()).map((i) => (i.id === id ? { ...i, stripped: true } : i)));
    }),

    async recipe(id: string): Promise<Recipe> {
      const data = await store.get<ItemData>(dataKey(id));
      if (!data) throw new Error("This library item's data is missing.");
      return data.recipe;
    },

    async backup(): Promise<Backup> {
      const items = await list();
      const backup: Backup = { retro: "backup", version: 1, exported: new Date().toISOString(), items, data: {}, parts: {}, settings: {} };
      for (const item of items) {
        const data = await store.get<ItemData>(dataKey(item.id));
        if (!data) continue;
        backup.data[item.id] = data;
        for (const { hash } of allParts(data).flat()) {
          if (backup.parts[hash]) continue;
          const part = await store.get<StoredPart>(partKey(hash));
          if (part) backup.parts[hash] = { ...part, data: bytesToBase64(part.data) };
        }
      }
      for (const key of SETTING_KEYS) {
        const value = await store.get(key);
        if (value !== undefined) backup.settings[key] = value;
      }
      return backup;
    },

    /** Add a backup's items (skipping ones already here) and settings. Returns how many were added. */
    restore: exclusive(async (backup: Backup): Promise<number> => {
      if (backup?.retro !== "backup" || backup.version !== 1) throw new Error("That file isn't a Retro backup.");
      const items = await list();
      const have = new Set(items.map((i) => i.id));
      const refs = (await store.get<Record<string, number>>(REFS)) ?? {};
      const added: LibraryItem[] = [];
      for (const item of backup.items.map(upgrade)) {
        const data = backup.data[item.id];
        if (have.has(item.id) || !data) continue;
        for (const { hash } of allParts(data).flat()) {
          const part = backup.parts[hash];
          if (!part) throw new Error("The backup file is incomplete.");
          if (!refs[hash] || !(await store.get(partKey(hash)))) await store.set(partKey(hash), { ...part, data: base64ToBytes(part.data) });
          refs[hash] = (refs[hash] ?? 0) + 1;
        }
        await store.set(dataKey(item.id), data);
        // Previews must be embedded pictures (never addresses the pane would fetch), and restored slides are stripped again on load
        // whatever the file says (it may come from an older version, or have been edited).
        added.push({ ...item, preview: item.preview?.startsWith("data:image/") ? item.preview : undefined, stripped: undefined });
      }
      await store.set(REFS, refs);
      await writeIndex([...added, ...items]);
      for (const [key, value] of Object.entries(backup.settings ?? {})) if (SETTING_KEYS.includes(key)) await store.set(key, value);
      return added.length;
    }),

    /** The exported slide the item was saved from, reassembled. */
    async pptx(id: string): Promise<Uint8Array> {
      return (await slidesOf(id))[0];
    },

    /** Every slide in the item (one for shapes), each as its own .pptx, in order. */
    slides: slidesOf,
  };
}

export type Library = ReturnType<typeof createLibrary>;
