// Section tracker: sections, which slides they cover, and the tracker bar's shapes. Sections are
// marked on the slide where each one starts (a slide tag), so moving or adding slides keeps them
// right: a section runs until the next marked slide. Pure (unit tested); PowerPoint work is in
// trackerActions.ts.

import type { SlideSize } from "./ppt";

/** Slide tag on a section's first slide; the value is the section name, hex-encoded. */
export const SECTION_TAG = "RETRO_SECTION";
/** Slide tag on the agenda slide Retro built. */
export const AGENDA_TAG = "RETRO_AGENDA";
/** Presentation tag holding the tracker settings. */
export const TRACKER_TAG = "RETRO_TRACKER";

export type TrackerStyle = "tabs" | "bar" | "dots";
export type TrackerPosition = "top" | "bottom";

export interface TrackerSettings {
  style: TrackerStyle;
  position: TrackerPosition;
  agenda: boolean;
}

export const DEFAULT_SETTINGS: TrackerSettings = { style: "tabs", position: "top", agenda: true };

export interface Section {
  name: string;
  /** Slide indexes (0-based), inclusive. */
  start: number;
  end: number;
}

// Tag values may come back in capitals, so names are stored as hex (which survives that).
export function encodeName(name: string): string {
  return [...new TextEncoder().encode(name)].map((b) => b.toString(16).padStart(2, "0")).join("");
}

export function decodeName(hex: string): string {
  const pairs = hex.match(/[0-9a-f]{2}/gi) ?? [];
  return new TextDecoder().decode(new Uint8Array(pairs.map((p) => parseInt(p, 16))));
}

export function encodeSettings(s: TrackerSettings): string {
  return `${s.style}|${s.position}|${s.agenda ? 1 : 0}`;
}

export function decodeSettings(value: string | undefined): TrackerSettings {
  const [style, position, agenda] = (value ?? "").toLowerCase().split("|");
  return {
    style: style === "bar" || style === "dots" ? style : DEFAULT_SETTINGS.style,
    position: position === "bottom" ? "bottom" : "top",
    agenda: agenda === undefined ? DEFAULT_SETTINGS.agenda : agenda === "1",
  };
}

/** Sections from each slide's section tag (hex name, or undefined), in slide order. */
export function sectionsFrom(tags: (string | undefined)[]): Section[] {
  const sections: Section[] = [];
  tags.forEach((tag, i) => {
    if (tag === undefined) return;
    if (sections.length) sections[sections.length - 1].end = i - 1;
    sections.push({ name: decodeName(tag), start: i, end: tags.length - 1 });
  });
  return sections;
}

/** The section a slide is in, or -1 before the first section. */
export function sectionOf(sections: Section[], slide: number): number {
  return sections.findIndex((s) => slide >= s.start && slide <= s.end);
}

/** "5" or "3–6" (1-based, as PowerPoint numbers slides). */
export const slideRange = (s: Section) => (s.start === s.end ? `${s.start + 1}` : `${s.start + 1}–${s.end + 1}`);

export interface TrackerPart {
  kind: "box" | "dot" | "label";
  left: number;
  top: number;
  width: number;
  height: number;
  /** Text for boxes and labels. */
  text?: string;
  active: boolean;
}

const MARGIN = 24;
const EDGE = 8;

/**
 * The tracker's shapes for one slide. Tabs: one box per section across the slide. Bar: thin
 * segments with the names above them. Dots: a dot per section in the corner, with the current
 * section's name beside them.
 */
export function trackerParts(style: TrackerStyle, position: TrackerPosition, slide: SlideSize, names: string[], active: number): TrackerPart[] {
  const n = names.length;
  if (n === 0) return [];
  const span = slide.width - 2 * MARGIN;
  const at = (height: number) => (position === "top" ? EDGE : slide.height - EDGE - height);
  const round = (v: number) => Math.round(v * 100) / 100;

  if (style === "tabs") {
    const gap = 2;
    const width = (span - gap * (n - 1)) / n;
    const top = at(18);
    return names.map((text, i) => ({ kind: "box", left: round(MARGIN + i * (width + gap)), top, width: round(width), height: 18, text, active: i === active }));
  }

  if (style === "bar") {
    const gap = 3;
    const width = (span - gap * (n - 1)) / n;
    // An 18 pt block: the 4 pt bar at the slide's edge, the names on the inner side of it.
    const block = at(18);
    const barTop = position === "top" ? block : block + 14;
    const labelTop = position === "top" ? block + 4 : block;
    return names.flatMap((text, i): TrackerPart[] => {
      const left = round(MARGIN + i * (width + gap));
      return [
        { kind: "box", left, top: barTop, width: round(width), height: 4, active: i === active },
        { kind: "label", left, top: labelTop, width: round(width), height: 14, text, active: i === active },
      ];
    });
  }

  const size = 8;
  const gap = 6;
  const top = at(14);
  const dotsWidth = n * size + (n - 1) * gap;
  const left0 = slide.width - MARGIN - dotsWidth;
  const dots: TrackerPart[] = names.map((_, i) => ({ kind: "dot", left: left0 + i * (size + gap), top: top + 3, width: size, height: size, active: i === active }));
  const labelWidth = 200;
  return active < 0 ? dots : [...dots, { kind: "label", left: left0 - 8 - labelWidth, top, width: labelWidth, height: 14, text: names[active], active: true }];
}

/** Black or white text, whichever reads better on a color. */
export function textOn(hex: string): "#000000" | "#FFFFFF" {
  const [r, g, b] = [1, 3, 5].map((i) => parseInt(hex.slice(i, i + 2), 16) / 255).map((c) => (c <= 0.03928 ? c / 12.92 : ((c + 0.055) / 1.055) ** 2.4));
  return 0.2126 * r + 0.7152 * g + 0.0722 * b > 0.4 ? "#000000" : "#FFFFFF";
}

/** The agenda list: one numbered line per section. */
export const agendaText = (names: string[]) => names.map((name, i) => `${i + 1}   ${name}`).join("\n");
