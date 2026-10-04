import { describe, expect, it } from "vitest";
import { layoutRoadmap, monthToQuarter, parseRoadmap, parseWhen, quarterLabel } from "../src/lib/roadmap";

const q = (fy: number, quarter: number) => fy * 4 + quarter - 1;

describe("reading dates", () => {
  it("reads quarters in the usual ways", () => {
    for (const text of ["Q1 FY27", "q1 fy27", "FY27 Q1", "Q1 2027", "Q1'27", "2027 Q1", "Q1-27", "FY2027 Q1"]) expect(parseWhen(text, 7)).toBe(q(2027, 1));
    expect(parseWhen("Q5 2027", 1)).toBeUndefined();
    expect(parseWhen("soon", 1)).toBeUndefined();
  });

  it("puts months in their fiscal quarter, naming a fiscal year for the year it ends", () => {
    // Fiscal year starting in July: July 2026 is Q1 FY27, March 2027 is Q3 FY27.
    expect(monthToQuarter(7, 2026, 7)).toBe(q(2027, 1));
    expect(parseWhen("Mar 2027", 7)).toBe(q(2027, 3));
    expect(parseWhen("June 2027", 7)).toBe(q(2027, 4));
    expect(parseWhen("2026-10", 7)).toBe(q(2027, 2));
    expect(parseWhen("10/2026", 7)).toBe(q(2027, 2));
    // Calendar fiscal year.
    expect(parseWhen("Mar 2027", 1)).toBe(q(2027, 1));
    expect(parseWhen("2027-13", 1)).toBeUndefined();
  });

  it("labels quarters", () => {
    expect(quarterLabel(q(2027, 3), 7)).toBe("Q3 FY27");
    expect(quarterLabel(q(2027, 3), 1)).toBe("Q3 2027");
  });
});

describe("reading rows", () => {
  it("reads tabs, pipes or commas, skips a header, and explains bad rows", () => {
    const text = ["Lane\tItem\tStart\tEnd", "Platform\tNew API\tQ1 FY27\tQ2 FY27", "Platform | GA launch | Q3 FY27", "Data, Lakehouse, Q4 FY27, Q2 FY27", "Data\tMystery\tsoon", "\t\t", "Solo"].join("\n");
    const { items, errors } = parseRoadmap(text, 7);
    expect(items).toEqual([
      { lane: "Platform", name: "New API", start: q(2027, 1), end: q(2027, 2) },
      { lane: "Platform", name: "GA launch", start: q(2027, 3), end: undefined },
      { lane: "Data", name: "Lakehouse", start: q(2027, 2), end: q(2027, 4) }, // ends before it starts: swapped
    ]);
    expect(errors).toEqual(["Row 5: can't read the start “soon” (try Q1 FY27 or Mar 2027).", "Row 7: needs a lane and an item."]);
  });
});

describe("layout", () => {
  const area = { left: 36, top: 120, width: 888, height: 380 };
  const { items } = parseRoadmap(
    ["Platform\tNew API\tQ1 FY27\tQ2 FY27", "Platform\tSDK\tQ2 FY27\tQ3 FY27", "Platform\tGA\tQ4 FY27", "Data\tLakehouse\tQ2 FY27\tQ4 FY27"].join("\n"),
    7,
  );
  const parts = layoutRoadmap(items, area, 7);

  it("has a header per quarter across the area right of the lane labels", () => {
    const headers = parts.filter((p) => p.kind === "header");
    expect(headers.map((h) => h.text)).toEqual(["Q1 FY27", "Q2 FY27", "Q3 FY27", "Q4 FY27"]);
    expect(headers[0].left).toBe(146);
    expect(headers[3].left + headers[3].width).toBeCloseTo(924, 1);
  });

  it("stacks overlapping items in a lane and keeps the rest on one row", () => {
    const bars = parts.filter((p) => p.kind === "bar");
    const api = bars.find((b) => b.text === "New API")!;
    const sdk = bars.find((b) => b.text === "SDK")!;
    expect(sdk.top).toBeGreaterThan(api.top); // they overlap in Q2, so SDK goes below
    expect(sdk.left).toBeCloseTo(api.left + 194.5, 1); // one quarter later
    const ga = parts.find((p) => p.kind === "milestone")!;
    expect(ga.top + ga.height / 2).toBeCloseTo(api.top + api.height / 2, 1); // the milestone fits on New API's row
    expect(parts.filter((p) => p.kind === "laneLabel").map((l) => l.text)).toEqual(["Platform", "Data"]);
    // Every shape stays inside the area.
    for (const p of parts) expect(p.top + p.height).toBeLessThanOrEqual(area.top + area.height + 0.01);
  });

  it("draws nothing for no items", () => {
    expect(layoutRoadmap([], area, 7)).toEqual([]);
  });
});
