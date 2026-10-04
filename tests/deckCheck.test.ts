import { describe, expect, it } from "vitest";
import { autoFixable, checkDeck, closestColor, countByRule, describeShape, RULES, type Rule, type Run, type ShapeSnap, type SlideSnap } from "../src/lib/deckCheck";

const size = { width: 960, height: 540 };
const kit = { name: "Company", colors: ["#E07A1F", "#333F48", "#5E8AB4"], headingFont: "Georgia", bodyFont: "Arial" };
const all = new Set<Rule>(RULES.map((r) => r.rule));
const only = (...rules: Rule[]) => new Set<Rule>(rules);

let n = 0;
const shape = (over: Partial<ShapeSnap>): ShapeSnap => ({ path: [`s${++n}`], name: "Shape 1", type: "GeometricShape", left: 100, top: 100, width: 100, height: 50, runs: [], ...over });
const run = (start: number, length: number, more: Partial<Run> = {}): Run => ({ start, length, font: "Arial", color: "#333F48", size: 14, ...more });
const deck = (...shapes: ShapeSnap[]): SlideSnap[] => [{ id: "slide1", index: 2, title: "Market size", shapes }];

describe("brand rules", () => {
  it("flags fonts outside the kit, with a fix for just those runs, ignoring case and theme fonts", () => {
    const s = shape({ content: "Source: 2025 survey of 400 buyers in Europe", runs: [run(0, 23, { font: "arial" }), run(23, 10, { font: "Calibri" }), run(33, 4, { font: "+mn-lt" }), run(37, 6, { font: "Calibri" })] });
    const issues = checkDeck(deck(s), kit, size, only("fonts"));
    expect(issues).toHaveLength(1);
    expect(issues[0]).toMatchObject({ label: "Font", value: "Calibri", text: "isn't in your brand kit", slideIndex: 2, slideTitle: "Market size", fixLabel: "Use Arial" });
    expect(issues[0].fix).toEqual({ kind: "font", path: s.path, runs: [{ start: 23, length: 10 }, { start: 37, length: 6 }], to: "Arial" });
    expect(issues[0].shape).toBe("Shape “Source: 2025 survey of 400…”");
  });

  it("fixes titles with the heading font", () => {
    const title = shape({ type: "Placeholder", placeholder: { type: "Title", empty: false }, content: "Plan", runs: [run(0, 4, { font: "Calibri" })] });
    expect(checkDeck(deck(title), kit, size, only("fonts"))[0].fix).toMatchObject({ to: "Georgia" });
  });

  it("flags off-palette fill and text colors with the closest kit color as the fix, but not white or black", () => {
    const s = shape({ fill: "#ff0000", runs: [run(0, 3, { color: "#3A3A3A" }), run(3, 2, { color: "#FFFFFF" }), run(5, 2, { color: "#000000" }), run(7, 2, { color: "#333f48" }), run(9, 2, { color: "#3a3a3a" })] });
    const issues = checkDeck(deck(s), kit, size, only("colors"));
    expect(issues.map((i) => [i.label, i.value, i.closest])).toEqual([
      ["Fill", "#FF0000", "#E07A1F"],
      ["Text color", "#3A3A3A", "#333F48"],
    ]);
    expect(issues[0].fix).toEqual({ kind: "fill", path: s.path, to: "#E07A1F" });
    expect(issues[1].fix).toMatchObject({ kind: "textColor", runs: [{ start: 0, length: 3 }, { start: 9, length: 2 }], to: "#333F48" });
  });

  it("skips font and color rules when the kit has none", () => {
    const s = shape({ fill: "#FF0000", runs: [run(0, 2, { font: "Comic Sans MS" })] });
    expect(checkDeck(deck(s), { name: "Empty", colors: [], headingFont: "", bodyFont: "" }, size, only("fonts", "colors"))).toEqual([]);
    expect(checkDeck(deck(s), undefined, size, only("fonts", "colors"))).toEqual([]);
  });
});

describe("layout rules", () => {
  it("says how far objects run off the slide and moves them back on when they fit", () => {
    const issues = checkDeck(
      deck(
        shape({ left: 900, width: 84, picture: { alt: "Map", decorative: false } }),
        shape({ left: -10, top: 500, width: 50, height: 60 }),
        shape({ left: 2000 }),
        shape({ left: 959.5, width: 1 }), // within tolerance
        shape({ left: -500, nested: true }),
        shape({ left: -5, width: 980 }), // bigger than the slide: can't be moved on
      ),
      kit,
      size,
      only("offslide"),
    );
    expect(issues.map((i) => `${i.label} ${i.text}`)).toEqual([
      "Picture runs 24 pt past the right edge of the slide",
      "Object runs 20 pt past the bottom edge of the slide (and past other edges)",
      "Object is entirely off the slide",
      "Object runs 15 pt past the right edge of the slide (and past other edges)",
    ]);
    expect(issues.map((i) => i.fix && i.fix.kind === "move" && [i.fix.left, i.fix.top])).toEqual([[876, 100], [0, 480], [860, 100], undefined]);
  });

  it("flags empty placeholders (fixed by deleting) and pictures without alt text unless decorative", () => {
    const issues = checkDeck(
      deck(
        shape({ type: "Placeholder", name: "Title 1", placeholder: { type: "Title", empty: true } }),
        shape({ type: "Placeholder", placeholder: { type: "Body", empty: false } }),
        shape({ type: "Placeholder", placeholder: { type: "SlideNumber", empty: true } }),
        shape({ type: "Image", name: "Picture 3", picture: { alt: "", decorative: false } }),
        shape({ type: "Image", picture: { alt: " ", decorative: true } }),
        shape({ type: "Image", picture: { alt: "Team photo", decorative: false } }),
      ),
      kit,
      size,
      only("placeholders", "alttext"),
    );
    expect(issues.map((i) => [i.rule, i.text, i.shape, i.fix?.kind])).toEqual([
      ["placeholders", "title placeholder", "Placeholder “Title 1”", "delete"],
      ["alttext", "is missing", "Picture “Picture 3”", "alt"],
    ]);
    expect(autoFixable(issues).map((i) => i.rule)).toEqual(["placeholders"]);
  });
});

describe("pre-send rules", () => {
  it("finds unfinished text", () => {
    const texts = ["Revenue grew XX% in Q3", "Launch date TBD", "Owner: ??", "[Insert customer quote]", "Lorem ipsum dolor", "Growth was 12% (todo: check)", "Sizes XL and XXL", "See [1] and [2]", "Done."];
    const found = texts.map((content) => checkDeck(deck(shape({ content })), kit, size, only("unfinished"))[0]?.value);
    expect(found).toEqual(["XX", "TBD", "??", "[Insert customer quote]", "Lorem ipsum", "todo", undefined, undefined, undefined]);
  });

  it("flags Retro's notes and stamps (fixed by deleting) and skips them for every other rule", () => {
    const sticky = shape({ retro: "sticky", content: "TBD", fill: "#FFE680", left: 2000, runs: [run(0, 3, { font: "Comic Sans MS", size: 8 })] });
    const stamp = shape({ retro: "stamp", content: "DRAFT" });
    const issues = checkDeck(deck(sticky, stamp), kit, size, all);
    expect(issues.map((i) => [i.label, i.text, i.fix?.kind])).toEqual([
      ["Sticky note", "is still on the slide", "delete"],
      ["Stamp", "is still on the slide", "delete"],
    ]);
  });

  it("flags text under 12 pt once per shape, with the smallest size", () => {
    const issues = checkDeck(deck(shape({ runs: [run(0, 5, { size: 14 }), run(5, 5, { size: 9 }), run(10, 5, { size: 10.5 })] }), shape({ runs: [run(0, 5, { size: 12 })] })), kit, size, only("smalltext"));
    expect(issues.map((i) => `${i.value} ${i.text}`)).toEqual(["9 pt is under 12 pt"]);
  });

  it("flags missing titles and titles used twice", () => {
    const title = (text: string) => shape({ type: "Placeholder", placeholder: { type: "Title", empty: !text }, content: text });
    const slides: SlideSnap[] = [
      { id: "a", index: 0, title: "Market size", shapes: [title("Market size")] },
      { id: "b", index: 1, title: "", shapes: [shape({ type: "Image" })] },
      { id: "c", index: 2, title: "Market  size", shapes: [title("Market  size")] },
      { id: "d", index: 3, title: "", shapes: [title("")] },
    ];
    expect(checkDeck(slides, kit, size, only("titles")).map((i) => `${i.slideIndex + 1}: ${i.label} ${i.value ?? ""} ${i.text}`)).toEqual([
      "2: Title  is missing",
      "3: Title Market size is the same as slide 1's",
      "4: Title  is missing",
    ]);
    // With the placeholder rule on, an empty title placeholder is reported once, as a placeholder.
    expect(checkDeck(slides, kit, size, only("titles", "placeholders")).map((i) => `${i.slideIndex + 1}:${i.rule}`)).toEqual(["2:titles", "3:titles", "4:placeholders"]);
  });
});

describe("helpers", () => {
  it("counts by rule, finds the nearest color and describes shapes", () => {
    expect(countByRule([])).toMatchObject({ fonts: 0, unfinished: 0, titles: 0 });
    expect(closestColor("#3A3A3A", kit.colors)).toBe("#333F48");
    expect(closestColor("#000000", [])).toBeUndefined();
    expect(describeShape(shape({ type: "TextBox", content: "Hi\nthere" }))).toBe("Text box “Hi there”");
  });

  it("runs only the chosen rules", () => {
    const s = shape({ fill: "#FF0000", runs: [run(0, 2, { font: "Calibri" })], left: -50 });
    expect(checkDeck(deck(s), kit, size, only("fonts")).map((i) => i.rule)).toEqual(["fonts"]);
    expect(checkDeck(deck(s), kit, size, new Set()).length).toBe(0);
  });
});
