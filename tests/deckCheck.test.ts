import { describe, expect, it } from "vitest";
import { checkDeck, closestColor, countByRule, describeShape, RULES, type Rule, type ShapeSnap, type SlideSnap } from "../src/lib/deckCheck";

const size = { width: 960, height: 540 };
const kit = { name: "Company", colors: ["#E07A1F", "#333F48", "#5E8AB4"], headingFont: "Georgia", bodyFont: "Arial" };
const all = new Set<Rule>(RULES.map((r) => r.rule));

const shape = (over: Partial<ShapeSnap>): ShapeSnap => ({ selectId: "s", name: "Shape 1", type: "GeometricShape", left: 100, top: 100, width: 100, height: 50, fonts: [], textColors: [], ...over });
const deck = (...shapes: ShapeSnap[]): SlideSnap[] => [{ id: "slide1", index: 2, title: "Market size", shapes }];

describe("deck check", () => {
  it("flags fonts outside the kit, ignoring case and theme font references", () => {
    const issues = checkDeck(deck(shape({ fonts: ["arial", "Calibri", "+mn-lt", "Calibri"], text: "Source: 2025 survey of 400 buyers in Europe" })), kit, size, all);
    expect(issues).toHaveLength(1);
    expect(issues[0]).toMatchObject({ rule: "fonts", label: "Font", value: "Calibri", text: "isn't in your brand kit", slideIndex: 2, slideTitle: "Market size" });
    expect(issues[0].shape).toBe("Shape “Source: 2025 survey of 400…”");
  });

  it("flags off-palette fill and text colors with the closest kit color, but not white or black", () => {
    const issues = checkDeck(deck(shape({ fill: "#ff0000", textColors: ["#3A3A3A", "#FFFFFF", "#000000", "#333f48"] })), kit, size, all);
    expect(issues.map((i) => [i.label, i.value, i.closest])).toEqual([
      ["Fill", "#FF0000", "#E07A1F"],
      ["Text color", "#3A3A3A", "#333F48"],
    ]);
  });

  it("skips font and color rules when the kit has none", () => {
    const empty = { name: "Empty", colors: [], headingFont: "", bodyFont: "" };
    expect(checkDeck(deck(shape({ fill: "#FF0000", fonts: ["Comic Sans MS"] })), empty, size, all)).toEqual([]);
    expect(checkDeck(deck(shape({ fill: "#FF0000", fonts: ["Comic Sans MS"] })), undefined, size, all)).toEqual([]);
  });

  it("says how far objects run off the slide, checking grouped shapes through their group", () => {
    const issues = checkDeck(
      deck(
        shape({ left: 900, width: 84, picture: { alt: "Map", decorative: false } }),
        shape({ left: -10, top: 500, width: 50, height: 60 }),
        shape({ left: 2000 }),
        shape({ left: 959.5, width: 1 }), // within tolerance
        shape({ left: -500, nested: true }),
      ),
      kit,
      size,
      new Set(["offslide"]),
    );
    expect(issues.map((i) => `${i.label} ${i.text}`)).toEqual([
      "Picture runs 24 pt past the right edge of the slide",
      "Object runs 20 pt past the bottom edge of the slide (and past other edges)",
      "Object is entirely off the slide",
    ]);
  });

  it("flags empty placeholders by kind, and pictures without alt text unless decorative", () => {
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
      all,
    );
    expect(issues.map((i) => [i.rule, i.text, i.shape])).toEqual([
      ["placeholders", "title placeholder", "Placeholder “Title 1”"],
      ["alttext", "is missing", "Picture “Picture 3”"],
    ]);
    expect(countByRule(issues)).toEqual({ fonts: 0, colors: 0, offslide: 0, placeholders: 1, alttext: 1 });
  });

  it("runs only the chosen rules", () => {
    const s = shape({ fill: "#FF0000", fonts: ["Calibri"], left: -50 });
    expect(checkDeck(deck(s), kit, size, new Set(["fonts"])).map((i) => i.rule)).toEqual(["fonts"]);
    expect(checkDeck(deck(s), kit, size, new Set()).length).toBe(0);
  });

  it("finds the nearest color and describes shapes", () => {
    expect(closestColor("#3A3A3A", kit.colors)).toBe("#333F48");
    expect(closestColor("#000000", [])).toBeUndefined();
    expect(describeShape(shape({ type: "TextBox", text: "Hi" }))).toBe("Text box “Hi”");
  });
});
