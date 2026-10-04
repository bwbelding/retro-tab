import { describe, expect, it } from "vitest";
import { categoriesOf, createLibrary, filterItems, placement, type Backup, type LibraryItem, type Store } from "../src/lib/library";
import type { Recipe } from "../src/lib/shapeCapture";
import { openZip } from "../src/lib/zip";
import { pptxBytes } from "./helpers/pptx";

function memoryStore(): Store & { map: Map<string, unknown> } {
  const map = new Map<string, unknown>();
  return {
    map,
    get: async <T>(key: string) => map.get(key) as T | undefined,
    set: async (key, value) => void map.set(key, structuredClone(value)),
    del: async (key) => void map.delete(key),
  };
}

const recipe: Recipe = { version: 1, box: { left: 10, top: 20, width: 100, height: 50 }, nodes: [] };
const parts = (store: ReturnType<typeof memoryStore>) => [...store.map.keys()].filter((k) => k.startsWith("lib.part.")).length;

async function saveTwo(store: ReturnType<typeof memoryStore>) {
  const lib = createLibrary(store);
  const a = await lib.save({ name: "Badge", category: "Callouts", recipe, pptx: await pptxBytes(), issues: [] });
  const b = await lib.save({ name: "Icon", category: "Icons", recipe, pptx: await pptxBytes(["9", "2", "3", "4", "5", "6", "7", "8"]), issues: ["SVG icon"] });
  return { lib, a, b };
}

describe("shape library storage", () => {
  it("saves items newest first, with the route their issues imply", async () => {
    const { lib, a, b } = await saveTwo(memoryStore());
    expect((await lib.list()).map((i) => i.name)).toEqual(["Icon", "Badge"]);
    expect([a.route, b.route]).toEqual(["click", "helper"]);
    expect(a.box).toEqual(recipe.box);
  });

  it("keeps parts shared between items once, and drops them when nothing uses them", async () => {
    const store = memoryStore();
    const { lib, a, b } = await saveTwo(store);
    const files = openZip(await pptxBytes()).names.length;
    // The two exports differ only in one tag part.
    expect(parts(store)).toBe(files + 1);
    await lib.remove(a.id);
    expect(parts(store)).toBe(files);
    await lib.remove(b.id);
    expect(parts(store)).toBe(0);
    expect(await lib.list()).toEqual([]);
  });

  it("reassembles the exported slide exactly", async () => {
    const { lib, a } = await saveTwo(memoryStore());
    const original = openZip(await pptxBytes());
    const copy = openZip(await lib.pptx(a.id));
    expect(copy.names).toEqual(original.names);
    for (const name of original.names) expect(await copy.text(name)).toBe(await original.text(name));
  });

  it("renames, files, stars and records use", async () => {
    const { lib, a } = await saveTwo(memoryStore());
    await lib.update(a.id, { name: "Big badge", category: "Badges", favorite: true });
    await lib.markUsed(a.id);
    const item = (await lib.list()).find((i) => i.id === a.id)!;
    expect(item).toMatchObject({ name: "Big badge", category: "Badges", favorite: true, uses: 1 });
    expect(item.used).toBeDefined();
  });

  it("reads stars saved by version 1.5.0 under the British spelling", async () => {
    const store = memoryStore();
    store.map.set("lib.index", [{ id: "x", name: "Old", category: "", favourite: true, created: "2026-10-03", uses: 0, route: "click", issues: [], box: recipe.box }]);
    expect((await createLibrary(store).list())[0]).toMatchObject({ name: "Old", favorite: true });
    expect((await createLibrary(store).list())[0]).not.toHaveProperty("favourite");
  });

  it("backs up and restores into another library, skipping items already there", async () => {
    const source = memoryStore();
    source.map.set("pref.matchReference", "last");
    const { lib } = await saveTwo(source);
    const backup: Backup = JSON.parse(JSON.stringify(await lib.backup()));
    expect(backup.settings).toEqual({ "pref.matchReference": "last" });

    const target = memoryStore();
    const restored = createLibrary(target);
    expect(await restored.restore(backup)).toBe(2);
    expect(await restored.restore(backup)).toBe(0);
    expect((await restored.list()).map((i) => i.name)).toEqual(["Icon", "Badge"]);
    expect(target.map.get("pref.matchReference")).toBe("last");
    const id = (await restored.list())[1].id;
    expect(await openZip(await restored.pptx(id)).text("ppt/slides/slide1.xml")).toBe(await openZip(await pptxBytes()).text("ppt/slides/slide1.xml"));
    await expect(restored.restore({ retro: "nope" } as unknown as Backup)).rejects.toThrow(/isn't a Retro backup/);

    // A preview pointing at a web address is dropped rather than loaded.
    const tampered: Backup = JSON.parse(JSON.stringify(backup));
    tampered.items.forEach((i) => (i.preview = "https://example.com/pixel.png"));
    const third = createLibrary(memoryStore());
    await third.restore(tampered);
    expect((await third.list()).every((i) => i.preview === undefined)).toBe(true);
  });

  it("marks new saves stripped, and swaps an older item's slide without leaking parts", async () => {
    const store = memoryStore();
    const { lib, a } = await saveTwo(store);
    expect(a.stripped).toBe(true);
    // Restored items are stripped again, since a backup can come from an older version.
    const target = memoryStore();
    const restored = createLibrary(target);
    await restored.restore(JSON.parse(JSON.stringify(await lib.backup())));
    expect((await restored.list()).every((i) => i.stripped === undefined)).toBe(true);

    const smaller = await pptxBytes(["a", "b", "c", "d", "e", "f", "g", "h"]);
    await lib.replacePptx(a.id, smaller);
    expect((await lib.list()).find((i) => i.id === a.id)?.stripped).toBe(true);
    const copy = openZip(await lib.pptx(a.id));
    const expected = openZip(smaller);
    for (const name of expected.names) expect(await copy.text(name)).toBe(await expected.text(name));
    await lib.remove(a.id);
    const b = (await lib.list())[0];
    await lib.remove(b.id);
    expect(parts(store)).toBe(0);
  });

  it("keeps part counts right when a save runs while an old item is being stripped", async () => {
    // IndexedDB answers asynchronously, so two changes at once can interleave.
    const store = memoryStore();
    const tick = () => new Promise((r) => setTimeout(r, 1));
    const { get, set } = store;
    store.get = async (key) => (await tick(), get(key));
    store.set = async (key, value) => (await tick(), set(key, value));
    const { lib, a, b } = await saveTwo(store);
    const stripped = await pptxBytes(["a", "b", "c", "d", "e", "f", "g", "h"]);
    const third = await pptxBytes(["z", "2", "3", "4", "5", "6", "7", "8"]);
    const [, c] = await Promise.all([lib.replacePptx(a.id, stripped), lib.save({ name: "Third", category: "", recipe, pptx: third, issues: [] })]);
    expect((await lib.list()).map((i) => i.name).sort()).toEqual(["Badge", "Icon", "Third"]);
    // Every stored count matches the items that use the part.
    const used: Record<string, number> = {};
    for (const [key, value] of store.map) if (key.startsWith("lib.data.")) for (const { hash } of (value as { parts: { hash: string }[] }).parts) used[hash] = (used[hash] ?? 0) + 1;
    expect(store.map.get("lib.refs")).toEqual(used);
    for (const id of [a.id, b.id]) await lib.remove(id);
    expect(openZip(await lib.pptx(c.id)).names.length).toBeGreaterThan(0);
    await lib.remove(c.id);
    expect(parts(store)).toBe(0);
  });
});

describe("slide library items", () => {
  it("keeps several slides in one item, backs them up, and frees their parts when deleted", async () => {
    const store = memoryStore();
    const lib = createLibrary(store);
    const slides = [await pptxBytes(), await pptxBytes(["s", "2", "3", "4", "5", "6", "7", "8"]), await pptxBytes(["t", "2", "3", "4", "5", "6", "7", "8"])];
    const item = await lib.save({ name: "", category: "Company", recipe, pptx: slides[0], more: slides.slice(1), issues: [], kind: "slides" });
    expect(item).toMatchObject({ name: "Slide", kind: "slides", slides: 3, route: "click" });
    const back = await lib.slides(item.id);
    expect(back).toHaveLength(3);
    for (const [i, file] of back.entries()) expect(await openZip(file).text("ppt/tags/tag1.xml")).toBe(await openZip(slides[i]).text("ppt/tags/tag1.xml"));

    const copy = createLibrary(memoryStore());
    await copy.restore(JSON.parse(JSON.stringify(await lib.backup())));
    const restored = (await copy.list())[0];
    expect(restored).toMatchObject({ kind: "slides", slides: 3 });
    expect(await copy.slides(restored.id)).toHaveLength(3);

    await lib.remove(item.id);
    expect(parts(store)).toBe(0);
  });
});

describe("finding and placing items", () => {
  const item = (name: string, category: string, favorite: boolean, created: string, used?: string): LibraryItem => ({
    id: name,
    name,
    category,
    favorite,
    created,
    used,
    uses: used ? 1 : 0,
    route: "click",
    issues: [],
    box: recipe.box,
  });
  const items = [
    item("Chevron", "Process", false, "2026-01-01"),
    item("Callout", "Callouts", true, "2026-01-02"),
    item("Step arrow", "Process", false, "2026-01-03", "2026-02-01"),
  ];

  it("sorts most recently used first and filters by search, favorites and category", () => {
    expect(filterItems(items, "", "all").map((i) => i.name)).toEqual(["Step arrow", "Callout", "Chevron"]);
    expect(filterItems(items, "", "favorites").map((i) => i.name)).toEqual(["Callout"]);
    expect(filterItems(items, "", "category:Process").map((i) => i.name)).toEqual(["Step arrow", "Chevron"]);
    expect(filterItems(items, "CALL", "all").map((i) => i.name)).toEqual(["Callout"]);
    expect(filterItems(items, "process", "all")).toHaveLength(2);
    expect(categoriesOf(items)).toEqual(["Callouts", "Process"]);
  });

  it("puts items where they were saved, kept on the slide", () => {
    const slide = { width: 960, height: 540 };
    expect(placement({ left: 10, top: 20, width: 100, height: 50 }, slide)).toEqual({ dx: 0, dy: 0 });
    expect(placement({ left: 900, top: 520, width: 100, height: 50 }, slide)).toEqual({ dx: -40, dy: -30 });
    expect(placement({ left: -20, top: 0, width: 100, height: 50 }, slide)).toEqual({ dx: 20, dy: 0 });
    expect(placement({ left: 100, top: 0, width: 1200, height: 50 }, slide)).toEqual({ dx: -100, dy: 0 });
  });
});
