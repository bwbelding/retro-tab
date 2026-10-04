import { describe, expect, it } from "vitest";
import { agendaText, decodeName, decodeSettings, encodeName, encodeSettings, sectionOf, sectionsFrom, slideRange, textOn, trackerParts } from "../src/lib/tracker";

const slide = { width: 960, height: 540 };

describe("sections", () => {
  it("encodes names so capitalized tag values still decode", () => {
    for (const name of ["Intro", "Market & sizing", "Café — 2026", ""]) {
      expect(decodeName(encodeName(name))).toBe(name);
      expect(decodeName(encodeName(name).toUpperCase())).toBe(name);
    }
  });

  it("runs each section until the next marked slide", () => {
    const tags = [undefined, undefined, encodeName("Intro"), undefined, encodeName("Market"), undefined, undefined, encodeName("Ask")];
    const sections = sectionsFrom(tags);
    expect(sections).toEqual([
      { name: "Intro", start: 2, end: 3 },
      { name: "Market", start: 4, end: 6 },
      { name: "Ask", start: 7, end: 7 },
    ]);
    expect([0, 1, 2, 3, 4, 7].map((i) => sectionOf(sections, i))).toEqual([-1, -1, 0, 0, 1, 2]);
    expect(sections.map(slideRange)).toEqual(["3–4", "5–7", "8"]);
    expect(sectionsFrom([undefined, undefined])).toEqual([]);
  });

  it("reads settings back, whatever the case, with defaults for anything missing", () => {
    const s = { style: "dots", position: "bottom", agenda: false } as const;
    expect(decodeSettings(encodeSettings(s).toUpperCase())).toEqual(s);
    expect(decodeSettings(undefined)).toEqual({ style: "tabs", position: "top", agenda: true });
    expect(decodeSettings("weird|x")).toEqual({ style: "tabs", position: "top", agenda: true });
  });

  it("numbers the agenda", () => {
    expect(agendaText(["Intro", "Plan"])).toBe("1   Intro\n2   Plan");
  });
});

describe("tracker shapes", () => {
  const names = ["Intro", "Market", "Plan", "Ask"];
  const right = (p: { left: number; width: number }) => p.left + p.width;

  it("lays tabs across the slide inside the margins, highlighting the current section", () => {
    const tabs = trackerParts("tabs", "top", slide, names, 1);
    expect(tabs).toHaveLength(4);
    expect(tabs[0].left).toBe(24);
    expect(right(tabs[3])).toBeCloseTo(936, 1);
    expect(tabs.map((t) => t.active)).toEqual([false, true, false, false]);
    expect(tabs.every((t) => t.top === 8 && t.height === 18)).toBe(true);
    expect(trackerParts("tabs", "bottom", slide, names, 1)[0].top).toBe(540 - 8 - 18);
  });

  it("puts bar labels on the slide side of the bar", () => {
    const top = trackerParts("bar", "top", slide, names, 0);
    const [bar, label] = top;
    expect(bar).toMatchObject({ kind: "box", top: 8, height: 4 });
    expect(label.top).toBeGreaterThan(bar.top);
    const [bBar, bLabel] = trackerParts("bar", "bottom", slide, names, 0);
    expect(bBar.top + bBar.height).toBe(540 - 8);
    expect(bLabel.top).toBeLessThan(bBar.top);
  });

  it("puts dots in the corner with the current section named beside them", () => {
    const parts = trackerParts("dots", "top", slide, names, 2);
    const dots = parts.filter((p) => p.kind === "dot");
    expect(dots).toHaveLength(4);
    expect(right(dots[3])).toBe(936);
    expect(parts.find((p) => p.kind === "label")).toMatchObject({ text: "Plan" });
    expect(right(parts.find((p) => p.kind === "label")!)).toBeLessThan(dots[0].left);
    expect(trackerParts("tabs", "top", slide, [], 0)).toEqual([]);
  });

  it("picks readable text for a fill", () => {
    expect(textOn("#FFE680")).toBe("#000000");
    expect(textOn("#333F48")).toBe("#FFFFFF");
    expect(textOn("#C43E1C")).toBe("#FFFFFF");
  });
});
