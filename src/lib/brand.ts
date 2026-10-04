// Brand kits: saved palettes and heading/body fonts. Pure parts and storage are here (unit
// tested); applying them in PowerPoint is in brandActions.ts.

import type { Store } from "./library";

export interface BrandKit {
  id: string;
  name: string;
  /** "#RRGGBB", in order: Fill 1, Fill 2, Fill 3 on the ribbon are the first three. */
  colors: string[];
  headingFont: string;
  bodyFont: string;
}

export interface BrandState {
  active: string;
  kits: BrandKit[];
}

export type ColorTarget = "fill" | "line" | "text";

/** "#abc", "abc", "aabbcc" or "#AABBCC" → "#AABBCC"; anything else → undefined. */
export function normalizeHex(input: string): string | undefined {
  const s = input.trim().replace(/^#/, "");
  if (/^[0-9a-f]{3}$/i.test(s)) return `#${[...s].map((c) => c + c).join("")}`.toUpperCase();
  if (/^[0-9a-f]{6}$/i.test(s)) return `#${s}`.toUpperCase();
  return undefined;
}

export function newKit(name: string, from: Partial<BrandKit> = {}): BrandKit {
  return { id: crypto.randomUUID(), name: name.trim() || "Brand", colors: [], headingFont: "", bodyFont: "", ...from };
}

export function activeKit(state: BrandState): BrandKit | undefined {
  return state.kits.find((k) => k.id === state.active) ?? state.kits[0];
}

/** Move a color one place left (-1) or right (+1). */
export function moveColor(colors: string[], index: number, by: -1 | 1): string[] {
  const to = index + by;
  if (to < 0 || to >= colors.length) return colors;
  const out = [...colors];
  [out[index], out[to]] = [out[to], out[index]];
  return out;
}

/** Add a color unless the kit already has it. */
export function addColor(colors: string[], hex: string): string[] {
  return colors.includes(hex) ? colors : [...colors, hex];
}

/**
 * Colors and fonts from a theme part (ppt/theme/themeN.xml): the six accents first, since those
 * are the brand colors in most templates, then the dark and light text/background pair.
 */
export function kitFromTheme(theme: Document): Omit<BrandKit, "id"> {
  const all = Array.from(theme.getElementsByTagName("*"));
  const find = (name: string, within?: Element) => (within ? Array.from(within.getElementsByTagName("*")) : all).find((e) => e.localName === name);
  const scheme = find("clrScheme");
  const color = (name: string): string | undefined => {
    const slot = scheme ? find(name, scheme) : undefined;
    const c = slot?.firstElementChild;
    if (!c) return undefined;
    return normalizeHex(c.getAttribute("val") ?? "") ?? normalizeHex(c.getAttribute("lastClr") ?? "");
  };
  const colors = ["accent1", "accent2", "accent3", "accent4", "accent5", "accent6", "dk2", "lt2"].map(color).filter((c): c is string => Boolean(c));
  const font = (name: string) => {
    const f = find(name);
    const latin = f && Array.from(f.children).find((e) => e.localName === "latin");
    return latin?.getAttribute("typeface") ?? "";
  };
  const themeName = all.find((e) => e.localName === "theme")?.getAttribute("name") ?? "";
  return { name: themeName ? `${themeName} (from deck)` : "From deck", colors: [...new Set(colors)], headingFont: font("majorFont"), bodyFont: font("minorFont") };
}

// --- Storage ---

const KEY = "brand.kits";

export function createBrandStore(store: Store) {
  return {
    async load(): Promise<BrandState> {
      const state = await store.get<BrandState>(KEY);
      if (state?.kits.length) return state;
      const kit = newKit("My brand");
      return { active: kit.id, kits: [kit] };
    },
    save: (state: BrandState) => store.set(KEY, state),
  };
}
