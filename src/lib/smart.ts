// Smart Elements: small status graphics (numbered circles, Harvey balls, traffic lights, progress
// bars, star ratings, checkboxes, arrows and trend indicators) built from ordinary PowerPoint
// shapes, so they look right everywhere. Each element is tagged with its kind and value; changing
// the value rebuilds it in the same box. `partsFor` (what to draw) is pure and unit tested.

import { bounds, type Rect } from "./layout";
import { currentSlide, selectedShapes, slideSize, UserError } from "./ppt";

export const KIND_TAG = "RETRO_KIND";
export const SMART_TAG = "RETRO_SMART";
export const VALUE_TAG = "RETRO_VALUE";

export type SmartKind = "number" | "harvey" | "traffic" | "traffic3" | "progress" | "stars" | "checkbox" | "arrow" | "trend";

export const COLORS = {
  accent: "#1B4F8C",
  ink: "#333F48",
  track: "#E1E1E1",
  dim: "#6B6B6B",
  red: "#D93025",
  amber: "#F2A900",
  green: "#1E8E3E",
  grey: "#8C8C8C",
  gold: "#F2B400",
  white: "#FFFFFF",
} as const;

interface Choice {
  value: string;
  label: string;
}

export interface SmartInfo {
  label: string;
  size: { width: number; height: number };
  defaultValue: string;
  /** Fixed choices for the pane; kinds without them take a number. */
  choices?: Choice[];
}

const STATUS: Choice[] = [
  { value: "red", label: "Red" },
  { value: "amber", label: "Amber" },
  { value: "green", label: "Green" },
];

export const SMART: Record<SmartKind, SmartInfo> = {
  number: { label: "Numbered circle", size: { width: 28, height: 28 }, defaultValue: "1" },
  harvey: {
    label: "Harvey ball",
    size: { width: 24, height: 24 },
    defaultValue: "50",
    choices: ["0", "25", "50", "75", "100"].map((v) => ({ value: v, label: `${v}%` })),
  },
  traffic: { label: "Traffic light", size: { width: 18, height: 18 }, defaultValue: "green", choices: STATUS },
  traffic3: { label: "Traffic lights (3)", size: { width: 22, height: 58 }, defaultValue: "green", choices: STATUS },
  progress: { label: "Progress bar", size: { width: 120, height: 12 }, defaultValue: "60" },
  stars: {
    label: "Star rating",
    size: { width: 100, height: 18 },
    defaultValue: "3",
    choices: ["0", "1", "2", "3", "4", "5"].map((v) => ({ value: v, label: v })),
  },
  checkbox: {
    label: "Checkbox",
    size: { width: 18, height: 18 },
    defaultValue: "on",
    choices: [
      { value: "on", label: "Ticked" },
      { value: "off", label: "Empty" },
    ],
  },
  arrow: {
    label: "Arrow",
    size: { width: 24, height: 24 },
    defaultValue: "up",
    choices: [
      { value: "up", label: "Up" },
      { value: "down", label: "Down" },
      { value: "left", label: "Left" },
      { value: "right", label: "Right" },
    ],
  },
  trend: {
    label: "Trend",
    size: { width: 24, height: 24 },
    defaultValue: "up",
    choices: [
      { value: "up", label: "Up" },
      { value: "up-slight", label: "Slightly up" },
      { value: "flat", label: "Flat" },
      { value: "down-slight", label: "Slightly down" },
      { value: "down", label: "Down" },
    ],
  },
};

export const SMART_KINDS = Object.keys(SMART) as SmartKind[];

/** Clamp a value to what the kind accepts, falling back to its default. */
export function normalizeValue(kind: SmartKind, value: string): string {
  const info = SMART[kind];
  if (info.choices) return info.choices.some((c) => c.value === value) ? value : info.defaultValue;
  const n = Math.round(Number(value));
  if (!Number.isFinite(n)) return info.defaultValue;
  return kind === "progress" ? String(Math.min(100, Math.max(0, n))) : String(Math.max(0, n));
}

/** One PowerPoint shape to draw, in slide points. */
export interface Part {
  shape: "Ellipse" | "Pie" | "Rectangle" | "RoundRectangle" | "Star5" | "UpArrow" | "DownArrow" | "LeftArrow" | "RightArrow";
  left: number;
  top: number;
  width: number;
  height: number;
  fill?: string;
  line?: { color: string; weight: number };
  text?: { text: string; color: string; size: number; bold?: boolean };
  rotation?: number;
  /** Pie only: filled sweep as a fraction of the circle, clockwise from 12 o'clock. */
  pieFraction?: number;
}

type Box = Omit<Rect, "id">;

const STATUS_COLOR: Record<string, string> = { red: COLORS.red, amber: COLORS.amber, green: COLORS.green };
const TREND: Record<string, { rotation: number; color: string }> = {
  up: { rotation: -90, color: COLORS.green },
  "up-slight": { rotation: -45, color: COLORS.green },
  flat: { rotation: 0, color: COLORS.grey },
  "down-slight": { rotation: 45, color: COLORS.red },
  down: { rotation: 90, color: COLORS.red },
};
const ARROW_SHAPE = { up: "UpArrow", down: "DownArrow", left: "LeftArrow", right: "RightArrow" } as const;

/** The shapes that make up an element of this kind and value, drawn in `box`. */
export function partsFor(kind: SmartKind, rawValue: string, box: Box): Part[] {
  const value = normalizeValue(kind, rawValue);
  const { left, top, width, height } = box;
  const full = { left, top, width, height };
  switch (kind) {
    case "number":
      return [
        {
          shape: "Ellipse",
          ...full,
          fill: COLORS.accent,
          text: { text: value, color: COLORS.white, size: Math.max(6, Math.round(height * 0.5)), bold: true },
        },
      ];
    case "harvey": {
      const p = Number(value) / 100;
      const ring: Part = { shape: "Ellipse", ...full, fill: COLORS.white, line: { color: COLORS.ink, weight: 1 } };
      if (p === 0) return [ring];
      if (p === 1) return [{ ...ring, fill: COLORS.ink }];
      return [ring, { shape: "Pie", ...full, fill: COLORS.ink, pieFraction: p }];
    }
    case "traffic":
      return [{ shape: "Ellipse", ...full, fill: STATUS_COLOR[value] }];
    case "traffic3": {
      // Lights fit both ways, so a resized (even squat) housing still holds all three.
      const d = Math.min(width * 0.7, height / 3.6);
      const gap = (height - 3 * d) / 4;
      const lights = (["red", "amber", "green"] as const).map(
        (c, i): Part => ({
          shape: "Ellipse",
          left: left + (width - d) / 2,
          top: top + gap + i * (d + gap),
          width: d,
          height: d,
          fill: c === value ? STATUS_COLOR[c] : COLORS.dim,
        }),
      );
      return [{ shape: "RoundRectangle", ...full, fill: COLORS.ink }, ...lights];
    }
    case "progress": {
      const p = Number(value) / 100;
      const track: Part = { shape: "RoundRectangle", ...full, fill: COLORS.track };
      if (p === 0) return [track];
      return [track, { shape: "RoundRectangle", left, top, width: width * p, height, fill: COLORS.accent }];
    }
    case "stars": {
      const n = Number(value);
      const size = Math.min(height, width / 5.5);
      const gap = (width - 5 * size) / 4;
      return Array.from({ length: 5 }, (_, i): Part => ({
        shape: "Star5",
        left: left + i * (size + gap),
        top: top + (height - size) / 2,
        width: size,
        height: size,
        ...(i < n ? { fill: COLORS.gold } : {}),
        line: { color: COLORS.gold, weight: 1 },
      }));
    }
    case "checkbox":
      return [
        {
          shape: "RoundRectangle",
          ...full,
          fill: COLORS.white,
          line: { color: COLORS.ink, weight: 1.25 },
          ...(value === "on" ? { text: { text: "✓", color: COLORS.ink, size: Math.max(6, Math.round(height * 0.75)), bold: true } } : {}),
        },
      ];
    case "arrow":
      return [{ shape: ARROW_SHAPE[value as keyof typeof ARROW_SHAPE], ...full, fill: COLORS.accent }];
    case "trend":
      return [{ shape: "RightArrow", ...full, fill: TREND[value].color, rotation: TREND[value].rotation }];
  }
}

/**
 * Pie adjustment values for a sweep clockwise from 12 o'clock. PowerPoint stores pie angles
 * clockwise from 3 o'clock; `unit` is how many API units make one degree (1, or 60000 when the API
 * exposes raw OOXML values), detected from the pie's defaults at runtime.
 */
export function pieAngles(fraction: number, unit: number): [number, number] {
  const norm = (deg: number) => {
    // Degrees: (-180, 180], as PowerPoint reports them. Raw OOXML: [0, 360).
    if (unit === 1) {
      const d = ((deg % 360) + 360) % 360;
      return d > 180 ? d - 360 : d;
    }
    return (((deg % 360) + 360) % 360) * unit;
  };
  return [norm(270), norm(270 + 360 * fraction)];
}

/** A pie's default sweep is 0° to 270°; large magnitudes mean raw OOXML units (60000 per degree). */
export function detectAngleUnit(defaultEnd: number): number {
  return Math.abs(defaultEnd) > 1000 ? 60000 : 1;
}

/** Number for a new circle: one more than the highest numbered circle already on the slide. */
export function nextNumber(existing: string[]): number {
  const nums = existing.map(Number).filter((n) => Number.isFinite(n));
  return nums.length ? Math.max(...nums) + 1 : 1;
}

/** Reading order for renumbering: top to bottom, then left to right (rows within half a circle). */
export function readingOrder<T extends { left: number; top: number; height: number }>(items: T[]): T[] {
  return [...items].sort((a, b) => (Math.abs(a.top - b.top) > Math.min(a.height, b.height) / 2 ? a.top - b.top : a.left - b.left));
}

/** Where a new element goes: right of the selection, or centred on the slide. */
export function placeFor(kind: SmartKind, slide: { width: number; height: number }, selection: Rect[]): Box {
  const { width, height } = SMART[kind].size;
  if (selection.length === 0) return { left: (slide.width - width) / 2, top: (slide.height - height) / 2, width, height };
  const b = bounds(selection);
  const left = Math.min(b.left + b.width + 8, slide.width - width);
  return { left, top: b.top + (b.height - height) / 2, width, height };
}

// ---- PowerPoint ----

async function drawPart(context: PowerPoint.RequestContext, shapes: PowerPoint.ShapeCollection, part: Part): Promise<PowerPoint.Shape> {
  const s = shapes.addGeometricShape(part.shape, { left: part.left, top: part.top, width: part.width, height: part.height });
  if (part.fill) s.fill.setSolidColor(part.fill);
  else s.fill.clear();
  if (part.line) {
    s.lineFormat.visible = true;
    s.lineFormat.color = part.line.color;
    s.lineFormat.weight = part.line.weight;
  } else {
    s.lineFormat.visible = false;
  }
  if (part.text) {
    const tf = s.textFrame;
    tf.leftMargin = tf.rightMargin = tf.topMargin = tf.bottomMargin = 0;
    tf.wordWrap = false;
    tf.autoSizeSetting = PowerPoint.ShapeAutoSize.autoSizeNone;
    tf.verticalAlignment = PowerPoint.TextVerticalAlignment.middleCentered;
    tf.textRange.text = part.text.text;
    tf.textRange.paragraphFormat.horizontalAlignment = PowerPoint.ParagraphHorizontalAlignment.center;
    tf.textRange.font.color = part.text.color;
    tf.textRange.font.size = part.text.size;
    tf.textRange.font.bold = Boolean(part.text.bold);
  }
  if (part.rotation) s.rotation = part.rotation;
  if (part.pieFraction !== undefined) {
    const adj = s.adjustments;
    const end = adj.get(1);
    await context.sync();
    const [a1, a2] = pieAngles(part.pieFraction, detectAngleUnit(end.value));
    adj.set(0, a1);
    adj.set(1, a2);
  }
  return s;
}

function tag(shape: PowerPoint.Shape, kind: SmartKind, value: string) {
  shape.tags.add(KIND_TAG, "smart");
  shape.tags.add(SMART_TAG, kind);
  shape.tags.add(VALUE_TAG, value);
  shape.name = `Retro ${SMART[kind].label}`;
}

/** Draw an element on a slide and return its (single or group) shape. */
async function draw(context: PowerPoint.RequestContext, slide: PowerPoint.Slide, kind: SmartKind, value: string, box: Box): Promise<PowerPoint.Shape> {
  // Grouping needs PowerPointApi 1.8; trend rotation and Harvey ball angles need 1.10.
  if (!Office.context.requirements.isSetSupported("PowerPointApi", "1.10")) {
    throw new UserError("Smart elements need PowerPoint 16.105 or later.");
  }
  const v = normalizeValue(kind, value);
  const drawn: PowerPoint.Shape[] = [];
  for (const part of partsFor(kind, v, box)) drawn.push(await drawPart(context, slide.shapes, part));
  await context.sync();
  const shape = drawn.length === 1 ? drawn[0] : slide.shapes.addGroup(drawn);
  tag(shape, kind, v);
  shape.load("id");
  await context.sync();
  return shape;
}

export interface SmartSelection {
  id: string;
  kind: SmartKind;
  value: string;
}

/** Numbers of the numbered circles already on a slide. */
async function numbersOn(context: PowerPoint.RequestContext, slide: PowerPoint.Slide): Promise<{ shape: PowerPoint.Shape; value: string }[]> {
  const shapes = slide.shapes;
  shapes.load("items/id,items/left,items/top,items/height");
  await context.sync();
  const tagged = shapes.items.map((shape) => {
    const kind = shape.tags.getItemOrNullObject(SMART_TAG);
    const value = shape.tags.getItemOrNullObject(VALUE_TAG);
    kind.load("value");
    value.load("value");
    return { shape, kind, value };
  });
  await context.sync();
  return tagged
    .filter((t) => !t.kind.isNullObject && t.kind.value.toLowerCase() === "number")
    .map((t) => ({ shape: t.shape, value: t.value.isNullObject ? "" : t.value.value }));
}

/** Insert an element next to the selection (or centred), with its default value. */
export async function insertSmart(kind: SmartKind): Promise<void> {
  await PowerPoint.run(async (context) => {
    const slide = await currentSlide(context);
    const { rects } = await selectedShapes(context);
    const box = placeFor(kind, await slideSize(context), rects);
    let value = SMART[kind].defaultValue;
    if (kind === "number") value = String(nextNumber((await numbersOn(context, slide)).map((n) => n.value)));
    const shape = await draw(context, slide, kind, value, box);
    slide.setSelectedShapes([shape.id]);
    await context.sync();
  });
}

/** The selected Smart Element, if exactly one is selected. */
export async function readSmartSelection(): Promise<SmartSelection | undefined> {
  return PowerPoint.run(async (context) => {
    const sel = context.presentation.getSelectedShapes();
    sel.load("items/id");
    await context.sync();
    if (sel.items.length !== 1) return undefined;
    const shape = sel.items[0];
    const kind = shape.tags.getItemOrNullObject(SMART_TAG);
    const value = shape.tags.getItemOrNullObject(VALUE_TAG);
    kind.load("value");
    value.load("value");
    await context.sync();
    if (kind.isNullObject) return undefined;
    const k = kind.value.toLowerCase() as SmartKind;
    if (!SMART_KINDS.includes(k)) return undefined;
    return { id: shape.id, kind: k, value: value.isNullObject ? SMART[k].defaultValue : value.value.toLowerCase() };
  });
}

/** Change the selected element's value, rebuilding it in the same place and size. */
export async function updateSmart(target: SmartSelection, value: string): Promise<void> {
  await PowerPoint.run(async (context) => {
    const slide = await currentSlide(context);
    const old = slide.shapes.getItem(target.id);
    old.load("left,top,width,height");
    await context.sync();
    const box = { left: old.left, top: old.top, width: old.width, height: old.height };
    old.delete();
    const shape = await draw(context, slide, target.kind, value, box);
    slide.setSelectedShapes([shape.id]);
    await context.sync();
  });
}

/** Renumber the slide's numbered circles 1, 2, 3… in reading order. Returns how many. */
export async function renumberSlide(): Promise<number> {
  return PowerPoint.run(async (context) => {
    const slide = await currentSlide(context);
    const circles = await numbersOn(context, slide);
    const ordered = readingOrder(circles.map((c) => ({ ...c, left: c.shape.left, top: c.shape.top, height: c.shape.height })));
    ordered.forEach((c, i) => {
      c.shape.textFrame.textRange.text = String(i + 1);
      c.shape.tags.add(VALUE_TAG, String(i + 1));
    });
    await context.sync();
    return ordered.length;
  });
}
