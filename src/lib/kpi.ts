// KPI and scorecard builder: pasted rows (metric, value, change, status) to big-number tiles or
// a scorecard table. Parsing and layout are pure (unit tested); drawing is in kpiActions.ts.

import type { Area } from "./roadmap";

export type Status = "green" | "amber" | "red";
export type Direction = "up" | "down" | "flat";
export type KpiStyle = "tiles" | "table";

export interface Kpi {
  metric: string;
  value: string;
  change?: string;
  direction?: Direction;
  status?: Status;
}

export const STATUS_LABEL: Record<Status, string> = { green: "On track", amber: "At risk", red: "Off track" };
export const ARROW: Record<Direction, string> = { up: "▲", down: "▼", flat: "▶" };

/** "G", "green", "on track", "✓"… to a status. */
export function parseStatus(text: string): Status | undefined {
  const t = text.trim().toLowerCase();
  if (!t) return undefined;
  if (/^(g|green|on track|ok|good|done|✓|✔)$/.test(t)) return "green";
  if (/^(a|amber|y|yellow|at risk|risk|watch|caution)$/.test(t)) return "amber";
  if (/^(r|red|off track|behind|late|blocked|✗|✘)$/.test(t)) return "red";
  return undefined;
}

/** Which way a change goes, from its sign: "+12%", "−3 pts", "▲ 4", "0", "flat". */
export function directionOf(change: string): Direction | undefined {
  const t = change.trim();
  if (!t) return undefined;
  if (/^[+▲↑]/.test(t)) return "up";
  if (/^[-−–▼↓]/.test(t)) return "down";
  if (/^(0+([.,]0+)?\s*%?|flat|=|±0.*|no change)$/i.test(t)) return "flat";
  return /^\d/.test(t) ? "up" : undefined;
}

/** Rows from a paste: tabs (from Excel), pipes or commas between columns. A header row is skipped. */
export function parseKpis(text: string): { rows: Kpi[]; errors: string[] } {
  const rows: Kpi[] = [];
  const errors: string[] = [];
  for (const [n, line] of text.split(/\r?\n/).entries()) {
    if (!line.trim()) continue;
    // Commas only split when there are no tabs or pipes, since values like "$1,200" have commas.
    const cells = (line.includes("\t") ? line.split("\t") : line.includes("|") ? line.split("|") : line.split(/,\s+/)).map((c) => c.trim());
    const [metric, value, change = "", status = ""] = cells;
    if (n === 0 && /^(metric|kpi|measure|name)$/i.test(metric)) continue;
    if (!metric || !value) {
      errors.push(`Row ${n + 1}: needs a metric and a value.`);
      continue;
    }
    const s = parseStatus(status);
    if (status && !s) errors.push(`Row ${n + 1}: status “${status}” isn't green, amber or red; left blank.`);
    rows.push({ metric, value, change: change || undefined, direction: directionOf(change), status: s });
  }
  return { rows, errors };
}

/** Tiles in an area: up to 4 in a row (3 when there are 5 or 6), equal sizes with gaps. */
export function tileRects(count: number, area: Area, gap = 16): Area[] {
  if (count === 0) return [];
  const columns = count <= 4 ? count : count <= 6 ? 3 : 4;
  const rows = Math.ceil(count / columns);
  const width = (area.width - gap * (columns - 1)) / columns;
  // Tiles stay a sensible shape: no taller than they're wide.
  const height = Math.min((area.height - gap * (rows - 1)) / rows, width);
  const r2 = (v: number) => Math.round(v * 100) / 100;
  return Array.from({ length: count }, (_, i) => ({
    left: r2(area.left + (i % columns) * (width + gap)),
    top: r2(area.top + Math.floor(i / columns) * (height + gap)),
    width: r2(width),
    height: r2(height),
  }));
}

/** The table's text: a header row, then one row per KPI. */
export function tableText(rows: Kpi[]): string[][] {
  return [["Metric", "Value", "Change", "Status"], ...rows.map((r) => [r.metric, r.value, r.change ? `${ARROW[r.direction ?? "flat"]} ${r.change}` : "", r.status ? STATUS_LABEL[r.status] : ""])];
}

/** The pasted text and style, kept on the KPI shapes so they can be edited later. */
export interface KpiSource {
  text: string;
  style: KpiStyle;
}
