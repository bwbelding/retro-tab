// Roadmap builder: pasted rows (lane, item, start, end) to an editable quarterly timeline. Parsing
// and layout are pure (unit tested); drawing it is in roadmapActions.ts.

export interface RoadmapItem {
  lane: string;
  name: string;
  /** Quarter numbers: fiscal year × 4 + quarter − 1. */
  start: number;
  /** Absent for a milestone (a single point in time). */
  end?: number;
}

export interface Parsed {
  items: RoadmapItem[];
  /** Rows that couldn't be read, with why. */
  errors: string[];
}

const MONTHS = ["jan", "feb", "mar", "apr", "may", "jun", "jul", "aug", "sep", "oct", "nov", "dec"];

/** A four-digit year from "2027", "27" or "'27". */
const fullYear = (y: string) => (y.replace("'", "").length <= 2 ? 2000 + Number(y.replace("'", "")) : Number(y));

/** The fiscal quarter a calendar month falls in. A fiscal year is named for the year it ends in. */
export function monthToQuarter(month: number, year: number, fyStart: number): number {
  const fy = fyStart > 1 && month >= fyStart ? year + 1 : year;
  const q = Math.floor(((month - fyStart + 12) % 12) / 3) + 1;
  return fy * 4 + q - 1;
}

/**
 * A quarter from text: "Q1 FY27", "FY27 Q1", "Q1 2027", "Q1'27", "2027 Q1", or a month that's
 * placed in its fiscal quarter: "Mar 2027", "March 2027", "2027-03", "3/2027".
 */
export function parseWhen(text: string, fyStart: number): number | undefined {
  const t = text.trim().toLowerCase();
  if (!t) return undefined;
  let m = t.match(/^q([1-4])\s*[-' ]?\s*(?:fy\s*)?('?\d{2}|\d{4})$/) ?? t.match(/^(?:fy\s*)?('?\d{2}|\d{4})\s*[- ]?\s*q([1-4])$/);
  if (m) {
    const [q, y] = /^q/.test(t) ? [m[1], m[2]] : [m[2], m[1]];
    return fullYear(y) * 4 + Number(q) - 1;
  }
  m = t.match(/^([a-z]{3})[a-z]*\.?\s+('?\d{2}|\d{4})$/);
  if (m && MONTHS.includes(m[1])) return monthToQuarter(MONTHS.indexOf(m[1]) + 1, fullYear(m[2]), fyStart);
  m = t.match(/^(\d{4})-(\d{1,2})$/) ?? t.match(/^(\d{1,2})\/(\d{4})$/);
  if (m) {
    const [month, year] = t.includes("-") ? [Number(m[2]), Number(m[1])] : [Number(m[1]), Number(m[2])];
    if (month >= 1 && month <= 12) return monthToQuarter(month, year, fyStart);
  }
  return undefined;
}

/** "Q1 FY27", or "Q1 2027" when the fiscal year is the calendar year. */
export function quarterLabel(quarter: number, fyStart: number): string {
  const fy = Math.floor(quarter / 4);
  const q = (quarter % 4) + 1;
  return fyStart === 1 ? `Q${q} ${fy}` : `Q${q} FY${String(fy).slice(-2)}`;
}

/** Rows from a paste: tabs (from Excel), pipes or commas between columns. A header row is skipped. */
export function parseRoadmap(text: string, fyStart: number): Parsed {
  const items: RoadmapItem[] = [];
  const errors: string[] = [];
  for (const [n, line] of text.split(/\r?\n/).entries()) {
    if (!line.trim()) continue;
    const cells = (line.includes("\t") ? line.split("\t") : line.includes("|") ? line.split("|") : line.split(",")).map((c) => c.trim());
    const [lane, name, from, to] = cells;
    if (n === 0 && /^(lane|stream|team|track)/i.test(lane) && !parseWhen(from ?? "", fyStart)) continue;
    const start = parseWhen(from ?? "", fyStart);
    const end = to ? parseWhen(to, fyStart) : undefined;
    if (!lane || !name) errors.push(`Row ${n + 1}: needs a lane and an item.`);
    else if (start === undefined) errors.push(`Row ${n + 1}: can't read the start “${from ?? ""}” (try Q1 FY27 or Mar 2027).`);
    else if (to && end === undefined) errors.push(`Row ${n + 1}: can't read the end “${to}”.`);
    else items.push({ lane, name, start: end !== undefined ? Math.min(start, end) : start, end: end !== undefined ? Math.max(start, end) : undefined });
  }
  return { items, errors };
}

export interface Area {
  left: number;
  top: number;
  width: number;
  height: number;
}

export interface RoadmapPart {
  kind: "header" | "laneLabel" | "laneLine" | "gridLine" | "bar" | "milestone" | "milestoneLabel";
  left: number;
  top: number;
  width: number;
  height: number;
  text?: string;
  /** Lane number, for picking its color. */
  lane?: number;
}

const LANE_LABEL = 110;
const HEADER = 24;
const BAR = 22;
const GAP = 4;
const DIAMOND = 12;

const r2 = (v: number) => Math.round(v * 100) / 100;

/** Items packed into rows within a lane so none overlap (first row each fits in). */
function packRows(items: RoadmapItem[]): RoadmapItem[][] {
  const rows: { end: number; items: RoadmapItem[] }[] = [];
  for (const item of [...items].sort((a, b) => a.start - b.start)) {
    const last = item.end ?? item.start;
    const row = rows.find((r) => r.end < item.start);
    if (row) {
      row.items.push(item);
      row.end = last;
    } else rows.push({ end: last, items: [item] });
  }
  return rows.map((r) => r.items);
}

/** The roadmap's shapes in an area of the slide: quarter headers, lanes, bars and milestones. */
export function layoutRoadmap(items: RoadmapItem[], area: Area, fyStart: number): RoadmapPart[] {
  if (items.length === 0) return [];
  const first = Math.min(...items.map((i) => i.start));
  const last = Math.max(...items.map((i) => i.end ?? i.start));
  const quarters = last - first + 1;
  const lanes = [...new Set(items.map((i) => i.lane))];
  const gridLeft = area.left + LANE_LABEL;
  const col = (area.width - LANE_LABEL) / quarters;
  const x = (quarter: number) => gridLeft + (quarter - first) * col;
  const parts: RoadmapPart[] = [];

  for (let q = 0; q < quarters; q++) {
    parts.push({ kind: "header", left: r2(x(first + q)), top: area.top, width: r2(col), height: HEADER, text: quarterLabel(first + q, fyStart) });
  }
  const bodyTop = area.top + HEADER;
  const bodyHeight = area.height - HEADER;
  const packed = lanes.map((lane) => packRows(items.filter((i) => i.lane === lane)));
  // Lanes share the height by how many rows each needs.
  const rowCount = packed.reduce((sum, rows) => sum + rows.length, 0);
  const rowHeight = Math.min(bodyHeight / rowCount, BAR + 2 * GAP + 8);
  for (let q = 1; q < quarters; q++) parts.push({ kind: "gridLine", left: r2(x(first + q)), top: bodyTop, width: 0, height: r2(rowHeight * rowCount) });

  let top = bodyTop;
  packed.forEach((rows, lane) => {
    const laneHeight = rowHeight * rows.length;
    parts.push({ kind: "laneLine", left: area.left, top: r2(top), width: area.width, height: 0, lane });
    parts.push({ kind: "laneLabel", left: area.left, top: r2(top), width: LANE_LABEL - 8, height: r2(laneHeight), text: lanes[lane], lane });
    rows.forEach((row, n) => {
      const mid = top + rowHeight * n + rowHeight / 2;
      for (const item of row) {
        if (item.end === undefined) {
          const cx = x(item.start) + col / 2;
          parts.push({ kind: "milestone", left: r2(cx - DIAMOND / 2), top: r2(mid - DIAMOND / 2), width: DIAMOND, height: DIAMOND, lane });
          parts.push({ kind: "milestoneLabel", left: r2(cx + DIAMOND / 2 + 3), top: r2(mid - 9), width: r2(Math.max(col - DIAMOND, 60)), height: 18, text: item.name, lane });
        } else {
          const height = Math.min(BAR, rowHeight - 2 * GAP);
          parts.push({ kind: "bar", left: r2(x(item.start) + 2), top: r2(mid - height / 2), width: r2((item.end - item.start + 1) * col - 4), height: r2(height), text: item.name, lane });
        }
      }
    });
    top += laneHeight;
  });
  parts.push({ kind: "laneLine", left: area.left, top: r2(top), width: area.width, height: 0 });
  return parts;
}

/** The pasted text and settings, kept on the roadmap so it can be edited later. */
export interface RoadmapSource {
  text: string;
  fyStart: number;
}
