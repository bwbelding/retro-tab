// Geometry for the Size & Position tools. Pure functions over shape rectangles (points), so they
// are unit tested without PowerPoint; src/lib/ppt.ts reads and writes the real shapes.

export interface Rect {
  id: string;
  left: number;
  top: number;
  width: number;
  height: number;
}

export type Dimension = "width" | "height" | "both";
/** Which selected shape the others are matched to. Selection order is as PowerPoint reports it. */
export type Reference = "first" | "last" | "largest" | "smallest";
export type Axis = "horizontal" | "vertical";
export type Unit = "pt" | "cm" | "in";

const POINTS_PER: Record<Unit, number> = { pt: 1, cm: 72 / 2.54, in: 72 };

export const toPoints = (value: number, unit: Unit) => value * POINTS_PER[unit];
export const fromPoints = (points: number, unit: Unit) => points / POINTS_PER[unit];

const area = (r: Rect) => r.width * r.height;

export function pickReference(rects: Rect[], ref: Reference): Rect {
  if (rects.length === 0) throw new Error("No shapes to match.");
  switch (ref) {
    case "first":
      return rects[0];
    case "last":
      return rects[rects.length - 1];
    case "largest":
      return rects.reduce((a, b) => (area(b) > area(a) ? b : a));
    case "smallest":
      return rects.reduce((a, b) => (area(b) < area(a) ? b : a));
  }
}

/** Give every shape the reference shape's width, height or both. Shapes keep their top-left corner. */
export function matchSize(rects: Rect[], dim: Dimension, ref: Reference = "first"): Rect[] {
  const target = pickReference(rects, ref);
  return rects.map((r) => ({
    ...r,
    width: dim === "height" ? r.width : target.width,
    height: dim === "width" ? r.height : target.height,
  }));
}

/** Swap two shapes' positions by their centers, so shapes of different sizes trade places cleanly. */
export function swapPositions(a: Rect, b: Rect): [Rect, Rect] {
  const ca = { x: a.left + a.width / 2, y: a.top + a.height / 2 };
  const cb = { x: b.left + b.width / 2, y: b.top + b.height / 2 };
  return [
    { ...a, left: cb.x - a.width / 2, top: cb.y - a.height / 2 },
    { ...b, left: ca.x - b.width / 2, top: ca.y - b.height / 2 },
  ];
}

/**
 * Space shapes with an exact gap along an axis. The shape nearest the start stays put; the rest
 * follow in their current order. Returned in the input order.
 */
export function distributeWithGap(rects: Rect[], gap: number, axis: Axis): Rect[] {
  const start = axis === "horizontal" ? "left" : "top";
  const size = axis === "horizontal" ? "width" : "height";
  const sorted = [...rects].sort((a, b) => a[start] - b[start]);
  const placed = new Map<string, Rect>();
  let next = sorted[0]?.[start] ?? 0;
  for (const r of sorted) {
    placed.set(r.id, { ...r, [start]: next });
    next += r[size] + gap;
  }
  return rects.map((r) => placed.get(r.id)!);
}

export function nudge(rects: Rect[], dx: number, dy: number): Rect[] {
  return rects.map((r) => ({ ...r, left: r.left + dx, top: r.top + dy }));
}

export interface ExactValues {
  left?: number;
  top?: number;
  width?: number;
  height?: number;
}

/** Set position and/or size on each shape. With lockAspect, a single given dimension scales the other. */
export function setEach(rects: Rect[], v: ExactValues, lockAspect = false): Rect[] {
  return rects.map((r) => {
    let { width, height } = r;
    if (v.width !== undefined && v.height !== undefined) {
      width = v.width;
      height = v.height;
    } else if (v.width !== undefined) {
      width = v.width;
      if (lockAspect && r.width > 0) height = (r.height * v.width) / r.width;
    } else if (v.height !== undefined) {
      height = v.height;
      if (lockAspect && r.height > 0) width = (r.width * v.height) / r.height;
    }
    return { ...r, left: v.left ?? r.left, top: v.top ?? r.top, width, height };
  });
}

export function bounds(rects: Rect[]): Omit<Rect, "id"> {
  const left = Math.min(...rects.map((r) => r.left));
  const top = Math.min(...rects.map((r) => r.top));
  const right = Math.max(...rects.map((r) => r.left + r.width));
  const bottom = Math.max(...rects.map((r) => r.top + r.height));
  return { left, top, width: right - left, height: bottom - top };
}

/**
 * Treat the selection as one block: move its bounding box to left/top and scale it (positions and
 * sizes) to width/height.
 */
export function setAsWhole(rects: Rect[], v: ExactValues, lockAspect = false): Rect[] {
  const b = bounds(rects);
  let sx = v.width !== undefined && b.width > 0 ? v.width / b.width : 1;
  let sy = v.height !== undefined && b.height > 0 ? v.height / b.height : 1;
  if (lockAspect) {
    if (v.width !== undefined && v.height === undefined) sy = sx;
    if (v.height !== undefined && v.width === undefined) sx = sy;
  }
  const left = v.left ?? b.left;
  const top = v.top ?? b.top;
  return rects.map((r) => ({
    ...r,
    left: left + (r.left - b.left) * sx,
    top: top + (r.top - b.top) * sy,
    width: r.width * sx,
    height: r.height * sy,
  }));
}

/** The value to show in an X/Y/W/H field: the shared value, or undefined when shapes differ. */
export function commonValue(rects: Rect[], key: keyof Omit<Rect, "id">): number | undefined {
  if (rects.length === 0) return undefined;
  const first = rects[0][key];
  return rects.every((r) => Math.abs(r[key] - first) < 0.01) ? first : undefined;
}

/** Columns for a logo grid when you don't pick: about twice as wide as tall, never more than 6. */
export function gridColumns(count: number, area: { width: number; height: number }): number {
  if (count <= 1) return 1;
  // Cells about 2:1 suit most logos: c columns, r rows with (w/c)/(h/r) close to 2.
  let best = 1;
  let bestScore = Infinity;
  for (let c = 1; c <= Math.min(6, count); c++) {
    const r = Math.ceil(count / c);
    const score = Math.abs(Math.log(area.width / c / (area.height / r) / 2)) + (r * c - count) * 0.05;
    if (score < bestScore) {
      bestScore = score;
      best = c;
    }
  }
  return best;
}

/**
 * Logos placed in an even grid, each centered in its cell and sized to the same area, so a wide
 * wordmark and a square icon look equally prominent. Each keeps its own proportions and stays
 * inside its cell with room around it.
 */
export function logoGrid(sizes: { width: number; height: number }[], area: { left: number; top: number; width: number; height: number }, columns = gridColumns(sizes.length, area)): Omit<Rect, "id">[] {
  const rows = Math.ceil(sizes.length / columns);
  const cellW = area.width / columns;
  const cellH = area.height / rows;
  const maxW = cellW * 0.75;
  const maxH = cellH * 0.6;
  // The shared area: the largest every logo fits at, but not under half the usual size because
  // of one extreme logo (that one is shrunk on its own below).
  const usual = maxW * maxH * 0.45;
  const fits = Math.min(...sizes.map((s) => Math.min((maxW * maxW * s.height) / s.width, (maxH * maxH * s.width) / s.height)));
  const target = Math.max(Math.min(usual, fits), usual / 2);
  const r2 = (v: number) => Math.round(v * 100) / 100;
  return sizes.map((s, i) => {
    const ratio = s.width / s.height;
    let width = Math.sqrt(target * ratio);
    let height = width / ratio;
    const shrink = Math.min(1, maxW / width, maxH / height);
    width *= shrink;
    height *= shrink;
    const cx = area.left + (i % columns) * cellW + cellW / 2;
    const cy = area.top + Math.floor(i / columns) * cellH + cellH / 2;
    return { left: r2(cx - width / 2), top: r2(cy - height / 2), width: r2(width), height: r2(height) };
  });
}
