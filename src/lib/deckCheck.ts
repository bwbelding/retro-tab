// Deck check rules: pure functions over a snapshot of the deck (unit tested). Reading the deck,
// going to a problem and applying fixes are in deckCheckActions.ts.

import type { SlideSize } from "./ppt";

export type Rule = "unfinished" | "notes" | "fonts" | "colors" | "smalltext" | "offslide" | "placeholders" | "titles" | "alttext";

export const RULES: { rule: Rule; label: string; help: string }[] = [
  { rule: "unfinished", label: "Unfinished", help: "Text like XX, TBD, ??, [insert …] or lorem ipsum." },
  { rule: "notes", label: "Notes & stamps", help: "Retro sticky notes and DRAFT-style stamps still on slides." },
  { rule: "fonts", label: "Fonts", help: "Fonts that aren't your kit's heading or body font." },
  { rule: "colors", label: "Colors", help: "Fill and text colors that aren't in your kit (white and black are always fine)." },
  { rule: "smalltext", label: "Small text", help: "Text under 12 pt, hard to read when presented." },
  { rule: "offslide", label: "Off-slide", help: "Objects that run past the edge of the slide." },
  { rule: "placeholders", label: "Placeholders", help: "Empty title, text and picture placeholders." },
  { rule: "titles", label: "Titles", help: "Slides without a title, and titles used on more than one slide." },
  { rule: "alttext", label: "Alt text", help: "Pictures without alt text (unless marked decorative)." },
];

/** Text smaller than this (points) is flagged. */
export const MIN_TEXT = 12;

/** A stretch of a shape's text with one font, color and size (positions in the whole text). */
export interface Run {
  start: number;
  length: number;
  font?: string;
  color?: string;
  size?: number;
}

export interface ShapeSnap {
  /** Shape ids from the slide's top level down to this shape (one id unless it's in a group). */
  path: string[];
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
  runs: Run[];
  /** All of its text. */
  content?: string;
  placeholder?: { type: string; empty: boolean };
  /** Pictures only: their alt text, and whether they're marked decorative. */
  picture?: { alt: string; decorative: boolean };
  /** A sticky note or stamp Retro added. */
  retro?: "sticky" | "stamp";
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

type Span = { start: number; length: number };

/** A change that fixes an issue, applied by deckCheckActions.applyFixes. */
export type Fix =
  | { kind: "font"; path: string[]; runs: Span[]; to: string }
  | { kind: "textColor"; path: string[]; runs: Span[]; to: string }
  | { kind: "fill"; path: string[]; to: string }
  | { kind: "move"; path: string[]; left: number; top: number }
  | { kind: "delete"; path: string[] }
  /** Needs the alt text from you. */
  | { kind: "alt"; path: string[] };

export interface Issue {
  rule: Rule;
  slideId: string;
  slideIndex: number;
  slideTitle: string;
  /** The top-level shape to select for Go to; empty for the slide itself. */
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
  fix?: Fix;
  /** The fix button's label, e.g. "Use Arial". */
  fixLabel?: string;
}

/** Colors that are always fine. */
const NEUTRALS = ["#FFFFFF", "#000000"];
const TOLERANCE = 1; // points
const TITLE_TYPES = ["Title", "CenterTitle", "VerticalTitle"];
/** Words and marks people leave in text that isn't finished yet. */
const UNFINISHED = /\b(?:TBD|TBC|TODO)\b|\bX{2,}\b|\?{2,}|\[(?:insert|add|tbd|todo|placeholder|name|date|x+)\b[^\]]*\]|\blorem ipsum\b/i;

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
export function clip(text: string, max: number): string {
  if (text.length <= max) return text;
  const cut = text.slice(0, max);
  const space = cut.lastIndexOf(" ");
  return `${(space > max / 2 ? cut.slice(0, space) : cut).trimEnd()}…`;
}

const oneLine = (text: string) => text.replace(/\s+/g, " ").trim();

/** How a shape is named in results: its kind and the start of its text, or its name. */
export function describeShape(s: ShapeSnap): string {
  const kind = s.picture ? "Picture" : s.type === "TextBox" ? "Text box" : s.type === "Placeholder" ? "Placeholder" : s.type === "Group" ? "Group" : s.type === "Line" ? "Line" : s.type === "Table" ? "Table" : "Shape";
  const text = s.content ? oneLine(s.content) : "";
  return text ? `${kind} “${clip(text, 28)}”` : `${kind} “${s.name}”`;
}

function offSlide(s: ShapeSnap, size: SlideSize): { text: string; fix?: Fix } | undefined {
  const right = s.left + s.width;
  const bottom = s.top + s.height;
  const clamp = (v: number, extent: number, max: number) => Math.round(Math.min(Math.max(v, 0), max - extent) * 100) / 100;
  // It can move fully onto the slide only if it fits.
  const fix: Fix | undefined =
    s.width <= size.width && s.height <= size.height ? { kind: "move", path: s.path, left: clamp(s.left, s.width, size.width), top: clamp(s.top, s.height, size.height) } : undefined;
  if (right <= 0 || bottom <= 0 || s.left >= size.width || s.top >= size.height) return { text: "is entirely off the slide", fix };
  const over = [
    { edge: "left", by: -s.left },
    { edge: "top", by: -s.top },
    { edge: "right", by: right - size.width },
    { edge: "bottom", by: bottom - size.height },
  ].filter((o) => o.by > TOLERANCE);
  if (over.length === 0) return undefined;
  const worst = over.reduce((a, b) => (b.by > a.by ? b : a));
  const more = over.length > 1 ? " (and past other edges)" : "";
  return { text: `runs ${Math.round(worst.by)} pt past the ${worst.edge} edge of the slide${more}`, fix };
}

const same = (a: string, b: string) => a.trim().toLowerCase() === b.trim().toLowerCase();
const spans = (runs: Run[]) => runs.map(({ start, length }) => ({ start, length }));

/** Every problem the chosen rules find, in slide order. */
export function checkDeck(slides: SlideSnap[], kit: Kit | undefined, size: SlideSize, rules: ReadonlySet<Rule>): Issue[] {
  const heading = kit?.headingFont.trim() ?? "";
  const body = kit?.bodyFont.trim() ?? "";
  const fonts = [heading, body].filter(Boolean);
  const palette = (kit?.colors ?? []).map((c) => c.toUpperCase());
  const allowed = new Set([...palette, ...NEUTRALS]);
  const issues: Issue[] = [];
  const titles = new Map<string, number>();

  for (const slide of slides) {
    const at = { slideId: slide.id, slideIndex: slide.index, slideTitle: slide.title };

    for (const s of slide.shapes) {
      const add = (rule: Rule, label: string, text: string, more: Partial<Issue> = {}) => issues.push({ rule, ...at, shapeId: s.path[0], shape: describeShape(s), label, text, ...more });

      if (s.retro) {
        if (rules.has("notes")) add("notes", s.retro === "sticky" ? "Sticky note" : "Stamp", "is still on the slide", { fix: { kind: "delete", path: s.path }, fixLabel: "Delete" });
        continue; // Retro's notes and stamps are meant to look different: no other rule applies.
      }

      if (rules.has("unfinished") && s.content) {
        const match = s.content.match(UNFINISHED);
        if (match) add("unfinished", "Text", "looks unfinished", { value: match[0] });
      }

      const isTitle = Boolean(s.placeholder && TITLE_TYPES.includes(s.placeholder.type));
      if (rules.has("fonts") && fonts.length) {
        // Theme font references ("+mn-lt") follow the template, so they aren't flagged.
        const to = isTitle ? heading || body : body || heading;
        const off = s.runs.filter((r) => r.font && !r.font.startsWith("+") && !fonts.some((k) => same(k, r.font!)));
        for (const font of new Set(off.map((r) => r.font!))) {
          const runs = off.filter((r) => r.font === font);
          add("fonts", "Font", "isn't in your brand kit", { value: font, fix: { kind: "font", path: s.path, runs: spans(runs), to }, fixLabel: `Use ${to}` });
        }
      }

      if (rules.has("colors") && palette.length) {
        const bad = (c: string | undefined) => (c && /^#[0-9A-F]{6}$/i.test(c) && !allowed.has(c.toUpperCase()) ? c.toUpperCase() : undefined);
        const fill = bad(s.fill);
        if (fill) {
          const to = closestColor(fill, palette)!;
          add("colors", "Fill", "isn't in your palette", { value: fill, closest: to, fix: { kind: "fill", path: s.path, to }, fixLabel: "Use closest" });
        }
        const off = s.runs.filter((r) => bad(r.color));
        for (const color of new Set(off.map((r) => r.color!.toUpperCase()))) {
          const to = closestColor(color, palette)!;
          const runs = off.filter((r) => r.color!.toUpperCase() === color);
          add("colors", "Text color", "isn't in your palette", { value: color, closest: to, fix: { kind: "textColor", path: s.path, runs: spans(runs), to }, fixLabel: "Use closest" });
        }
      }

      if (rules.has("smalltext")) {
        const sizes = s.runs.map((r) => r.size).filter((v): v is number => v !== undefined && v > 0);
        const smallest = Math.min(...sizes);
        if (smallest < MIN_TEXT) add("smalltext", "Text", `is under ${MIN_TEXT} pt`, { value: `${Math.round(smallest * 10) / 10} pt` });
      }

      if (rules.has("offslide") && !s.nested) {
        const off = offSlide(s, size);
        if (off) add("offslide", s.picture ? "Picture" : "Object", off.text, off.fix ? { fix: off.fix, fixLabel: "Move onto slide" } : {});
      }

      if (rules.has("placeholders") && s.placeholder?.empty) {
        const word = PLACEHOLDER_WORDS[s.placeholder.type];
        if (word) add("placeholders", "Empty", `${word} placeholder`, { fix: { kind: "delete", path: s.path }, fixLabel: "Delete" });
      }

      if (rules.has("alttext") && s.picture && !s.picture.decorative && !s.picture.alt.trim()) add("alttext", "Alt text", "is missing", { fix: { kind: "alt", path: s.path } });
    }

    if (rules.has("titles")) {
      const title = oneLine(slide.title);
      const titleShape = slide.shapes.find((s) => s.placeholder && TITLE_TYPES.includes(s.placeholder.type));
      if (!title) {
        // An empty title placeholder is already reported as one.
        if (!(rules.has("placeholders") && titleShape?.placeholder?.empty)) issues.push({ rule: "titles", ...at, shapeId: titleShape?.path[0] ?? "", shape: "This slide", label: "Title", text: "is missing" });
      } else {
        const key = title.toLowerCase();
        const first = titles.get(key);
        if (first === undefined) titles.set(key, slide.index);
        else issues.push({ rule: "titles", ...at, shapeId: titleShape?.path[0] ?? "", shape: titleShape ? describeShape(titleShape) : "This slide", label: "Title", value: clip(title, 32), text: `is the same as slide ${first + 1}'s` });
      }
    }
  }
  return issues;
}

/** How many issues each rule found. */
export function countByRule(issues: Issue[]): Record<Rule, number> {
  const counts = Object.fromEntries(RULES.map((r) => [r.rule, 0])) as Record<Rule, number>;
  for (const i of issues) counts[i.rule]++;
  return counts;
}

/** Issues Retro can fix without asking (alt text needs your words). */
export const autoFixable = (issues: Issue[]) => issues.filter((i) => i.fix && i.fix.kind !== "alt");
