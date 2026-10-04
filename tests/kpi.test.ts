import { describe, expect, it } from "vitest";
import { directionOf, parseKpis, parseStatus, tableText, tileRects } from "../src/lib/kpi";

describe("reading KPI rows", () => {
  it("reads statuses and directions in the usual ways", () => {
    expect(["G", "green", "On track", "✓", "a", "At risk", "R", "behind", "", "maybe"].map(parseStatus)).toEqual(["green", "green", "green", "green", "amber", "amber", "red", "red", undefined, undefined]);
    expect(["+12%", "−3 pts", "-$0.4M", "▲ 4", "0%", "flat", "12%", "", "n/a"].map(directionOf)).toEqual(["up", "down", "down", "up", "flat", "flat", "up", undefined, undefined]);
  });

  it("reads tabs, pipes or commas, keeps commas inside numbers, and skips a header", () => {
    const text = ["Metric\tValue\tChange\tStatus", "ARR\t$4.2M\t+12%\tG", "Churn | 2.1% | -0.3 pts | amber", "Customers, 1,240, +85, ok", "NPS\t\t+4", "Uptime\t99.95%\t\tpurple"].join("\n");
    const { rows, errors } = parseKpis(text);
    expect(rows).toEqual([
      { metric: "ARR", value: "$4.2M", change: "+12%", direction: "up", status: "green" },
      { metric: "Churn", value: "2.1%", change: "-0.3 pts", direction: "down", status: "amber" },
      { metric: "Customers", value: "1,240", change: "+85", direction: "up", status: "green" },
      { metric: "Uptime", value: "99.95%", change: undefined, direction: undefined, status: undefined },
    ]);
    expect(errors).toEqual(["Row 5: needs a metric and a value.", "Row 6: status “purple” isn't green, amber or red; left blank."]);
  });
});

describe("laying out", () => {
  const area = { left: 36, top: 120, width: 888, height: 384 };

  it("puts up to four tiles in a row, three for five or six, and keeps them no taller than wide", () => {
    expect(tileRects(4, area).map((r) => r.top)).toEqual([120, 120, 120, 120]);
    const six = tileRects(6, area);
    expect(new Set(six.map((r) => r.top)).size).toBe(2);
    expect(six[2].left + six[2].width).toBeCloseTo(924, 1);
    const one = tileRects(1, area);
    expect(one[0].height).toBeLessThanOrEqual(one[0].width);
    expect(tileRects(0, area)).toEqual([]);
    for (const r of tileRects(8, area)) expect(r.top + r.height).toBeLessThanOrEqual(area.top + area.height + 0.01);
  });

  it("builds the table text with arrows and status words", () => {
    const { rows } = parseKpis("ARR\t$4.2M\t+12%\tG\nUptime\t99.9%");
    expect(tableText(rows)).toEqual([
      ["Metric", "Value", "Change", "Status"],
      ["ARR", "$4.2M", "▲ +12%", "On track"],
      ["Uptime", "99.9%", "", ""],
    ]);
  });
});
