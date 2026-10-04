// Deck check rules: pure functions over a snapshot of the deck (unit tested). Reading the deck
// and going to a problem are in deckCheckActions.ts.

import type { SlideSize } from "./ppt";

export type Rule = "fonts" | "colors" | "offslide" | "placeholders" | "alttext";

export const RULES: { rule: Rule; label: string; help: string }[] = [
  { rule: "fonts", label: "Fonts", help: "Fonts that aren't your kit's heading or body font." },
  { rule: "colors", label: "Colors", help: "Fill and text colors that aren't in your kit (white and black are always fine)." },
  { rule: "offslide", label: "Off-slide", help: "Objects that run past the edge of the slide." },
  { rule: "placeholders", label: "Placeholders", help: "Empty title, text and picture placeholders." },
  { rule: "alttext", label: "Alt text", help: "Pictures without alt text (unless marked decorative)." },
];

export interface ShapeSnap {
  /** The top-level shape to select for Go to (a shape inside a group selects its group). */
  selectId: string;
  name: string;
  type: string;
  left: number;
  top: number;
  width: number;
  height: number;
  /** Inside a group: its position is checked through the group. */
  nested?: boolean;
  /** "#RRGGBB" when the shape has a solid fill. */
  fill?: string;
  fonts: string[];
  textColors: string[];
  /** The start of its text, if it has any. */
  text?: string;
  placeholder?: { type: string; empty: boolean };
  /** Pictures only: their alt text, and whether they're marked decorative. */
  picture?: { alt: string; decorative: boolean };
}

export interface SlideSnap {
  id: string;
  index: number;
  title: string;
  shapes: ShapeSnap[];
}

export interface Kit {
  name: string;
  colors: string[];
  headingFont: string;
  bodyFont: string;
}

export interface Issue {
  rule: Rule;
  slideId: string;
  slideIndex: number;
  slideTitle: string;
  shapeId: string;
  /** What the shape is, e.g. Text box “Source: …”. */
  shape: string;
  /** e.g. "Font", "Fill". */
  label: string;
  /** e.g. "Arial", "#FF0000". */
  value?: string;
  /** The problem, e.g. "isn't in your brand kit". */
  text: string;
  /** Colors: the nearest kit color. */
  closest?: string;
}

/** Colors that are always fine. */
const NEUTRALS = ["#FFFFFF", "#000000"];
const TOLERANCE = 1; // points

/** The nearest color by the "redmean" approximation of how different colors look. */
export function closestColor(hex: string, palette: string[]): string | undefined {
  const rgb = (h: string) => [1, 3, 5].map((i) => parseInt(h.slice(i, i + 2), 16));
  const [r1, g1, b1] = rgb(hex);
  let best: string | undefined;
  let bestD = Infinity;
  for (const p of palette) {
    const [r2, g2, b2] = rgb(p);
    const rm = (r1 + r2) / 2;
    const d = (2 + rm / 256) * (r1 - r2) ** 2 + 4 * (g1 - g2) ** 2 + (2 + (255 - rm) / 256) * (b1 - b2) ** 2;
    if (d < bestD) {
      bestD = d;
      best = p;
    }
  }
  return best;
}

const PLACEHOLDER_WORDS: Record<string, string> = {
  Title: "title",
  CenterTitle: "title",
  VerticalTitle: "title",
  Subtitle: "subtitle",
  Body: "text",
  VerticalBody: "text",
  Content: "content",
  VerticalContent: "content",
  Picture: "picture",
  OnlinePicture: "picture",
  Chart: "chart",
  Table: "table",
  SmartArt: "SmartArt",
  Media: "media",
};

/** Text cut to about `max` characters at a word break. */
function clip(text: string, max: number): string {
  if (text.length <= max) return text;
  const cut = text.slice(0, max);
  const space = cut.lastIndexOf(" ");
  return `${(space > max / 2 ? cut.slice(0, space) : cut).trimEnd()}…`;
}

/** How a shape is named in results: its kind and the start of its text, or its name. */
export function describeShape(s: ShapeSnap): string {
  const kind = s.picture ? "Picture" : s.type === "TextBox" ? "Text box" : s.type === "Placeholder" ? "Placeholder" : s.type === "Group" ? "Group" : s.type === "Line" ? "Line" : s.type === "Table" ? "Table" : "Shape";
  if (s.text) return `${kind} “${clip(s.text, 28)}”`;
  return `${kind} “${s.name}”`;
}

function offSlide(s: ShapeSnap, size: SlideSize): string | undefined {
  const right = s.left + s.width;
  const bottom = s.top + s.height;
  if (right <= 0 || bottom <= 0 || s.left >= size.width || s.top >= size.height) return "is entirely off the slide";
  const over = [
    { edge: "left", by: -s.left },
    { edge: "top", by: -s.top },
    { edge: "right", by: right - size.width },
    { edge: "bottom", by: bottom - size.height },
  ].filter((o) => o.by > TOLERANCE);
  if (over.length === 0) return undefined;
  const worst = over.reduce((a, b) => (b.by > a.by ? b : a));
  const more = over.length > 1 ? " (and past other edges)" : "";
  return `runs ${Math.round(worst.by)} pt past the ${worst.edge} edge of the slide${more}`;
}

const same = (a: string, b: string) => a.trim().toLowerCase() === b.trim().toLowerCase();

/** Every problem the chosen rules find, in slide order. */
export function checkDeck(slides: SlideSnap[], kit: Kit | undefined, size: SlideSize, rules: ReadonlySet<Rule>): Issue[] {
  const fonts = [kit?.headingFont, kit?.bodyFont].filter((f): f is string => Boolean(f?.trim()));
  const palette = (kit?.colors ?? []).map((c) => c.toUpperCase());
  const allowed = new Set([...palette, ...NEUTRALS]);
  const issues: Issue[] = [];

  for (const slide of slides) {
    for (const s of slide.shapes) {
      const add = (rule: Rule, label: string, text: string, value?: string, closest?: string) =>
        issues.push({ rule, slideId: slide.id, slideIndex: slide.index, slideTitle: slide.title, shapeId: s.selectId, shape: describeShape(s), label, value, text, closest });

      if (rules.has("fonts") && fonts.length) {
        // Theme font references ("+mn-lt") follow the template, so they aren't flagged.
        for (const f of new Set(s.fonts)) if (!f.startsWith("+") && !fonts.some((k) => same(k, f))) add("fonts", "Font", "isn't in your brand kit", f);
      }
      if (rules.has("colors") && palette.length) {
        const check = (label: string, color: string | undefined) => {
          const c = color?.toUpperCase();
          if (c && /^#[0-9A-F]{6}$/.test(c) && !allowed.has(c)) add("colors", label, "isn't in your palette", c, closestColor(c, palette));
        };
        check("Fill", s.fill);
        for (const c of new Set(s.textColors.map((x) => x.toUpperCase()))) check("Text color", c);
      }
      if (rules.has("offslide") && !s.nested) {
        const off = offSlide(s, size);
        if (off) add("offslide", s.picture ? "Picture" : "Object", off);
      }
      if (rules.has("placeholders") && s.placeholder?.empty) {
        const word = PLACEHOLDER_WORDS[s.placeholder.type];
        if (word) add("placeholders", "Empty", `${word} placeholder`);
      }
      if (rules.has("alttext") && s.picture && !s.picture.decorative && !s.picture.alt.trim()) add("alttext", "Alt text", "is missing");
    }
  }
  return issues;
}

/** How many issues each rule found. */
export function countByRule(issues: Issue[]): Record<Rule, number> {
  const counts = { fonts: 0, colors: 0, offslide: 0, placeholders: 0, alttext: 0 };
  for (const i of issues) counts[i.rule]++;
  return counts;
}
