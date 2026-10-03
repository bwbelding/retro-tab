import { describe, expect, it } from "vitest";
import { API_GEOMETRIES, apiGeometry } from "../src/lib/geometry";
import type { Frame } from "../src/lib/ooxml";
import { boundsOf, lineFrame, rebuildIssues, shapeRotation, type Recipe, type RecipeNode } from "../src/lib/shapeCapture";
import { compareRecipes, copyOffset, describe as describeNode, lineEnds } from "../src/lib/shapeTest";

const frame = (f: Partial<Frame> = {}): Frame => ({ left: 10, top: 20, width: 100, height: 50, rotation: 0, flipH: false, flipV: false, ...f });

const chevron = (f: Partial<Frame> = {}): RecipeNode => ({
  kind: "shape",
  name: "Badge",
  geometry: "Chevron",
  frame: frame(f),
  issues: [],
  look: {
    fill: { type: "Solid", color: "#C43E1C", transparency: 0 },
    line: { visible: false },
    adjustments: [0.5],
    text: {
      text: "Hi there",
      margins: [7.2, 7.2, 3.6, 3.6],
      anchor: "Middle",
      wordWrap: true,
      autoSize: "AutoSizeNone",
      paragraphs: [{ start: 0, length: 8, align: "Center" }],
      runs: [
        { start: 0, length: 2, font: { name: "Aptos", size: 18, bold: false } },
        { start: 2, length: 6, font: { name: "Aptos", size: 18, bold: true } },
      ],
    },
  },
});

const recipe = (nodes: RecipeNode[]): Recipe => ({ version: 1, box: boundsOf(nodes), nodes });

describe("shape types", () => {
  it("maps slide XML names to the API's", () => {
    expect(apiGeometry("rect")).toBe("Rectangle");
    expect(apiGeometry("roundRect")).toBe("RoundRectangle");
    expect(apiGeometry("wedgeRoundRectCallout")).toBe("WedgeRRectCallout");
    expect(apiGeometry("chevron")).toBe("Chevron");
    expect(apiGeometry("flowChartProcess")).toBe("FlowChartProcess");
    expect(apiGeometry("star5")).toBe("Star5");
    expect(apiGeometry("notAShape")).toBeUndefined();
    expect(API_GEOMETRIES.size).toBe(177);
  });
});

describe("rebuilding", () => {
  it("turns a mirrored line into the same line without mirroring", () => {
    for (const f of [frame({ flipV: true }), frame({ flipH: true }), frame({ flipH: true, flipV: true }), frame({ flipV: true, rotation: 30 })]) {
      const g = lineFrame(f);
      expect(g.flipH || g.flipV).toBe(false);
      const sorted = (ends: [number, number][]) => ends.map((p) => p.map((v) => Math.round(v * 1000) / 1000)).sort();
      expect(sorted(lineEnds(g))).toEqual(sorted(lineEnds(f)));
    }
    // A line mirrored vertically runs from bottom-left to top-right.
    const ends = lineEnds(frame({ flipV: true }))
      .map((p) => p.map((v) => Math.round(v)))
      .sort((a, b) => a[0] - b[0]);
    expect(ends).toEqual([
      [10, 70],
      [110, 20],
    ]);
  });

  it("treats a shape mirrored both ways as turned 180°", () => {
    expect(shapeRotation(frame({ flipH: true, flipV: true, rotation: 270 }))).toBe(90);
    expect(shapeRotation(frame({ rotation: -90 }))).toBe(270);
  });

  it("lists why a shape needs the helper slide", () => {
    expect(rebuildIssues(chevron())).toEqual([]);
    const gradient = chevron();
    gradient.look.fill = { type: "Gradient" };
    expect(rebuildIssues(gradient)).toEqual(["gradient fill"]);
    const group: RecipeNode = { kind: "group", name: "G", frame: frame(), issues: [], look: {}, children: [gradient, { ...chevron(), issues: ["arrowheads"] }] };
    expect(rebuildIssues(group)).toEqual(["arrowheads", "gradient fill"]);
  });

  it("places copies beside the originals when there's room", () => {
    expect(copyOffset({ left: 10, top: 20, width: 100, height: 50 }, { width: 960, height: 540 })).toEqual({ dx: 124, dy: 0 });
    expect(copyOffset({ left: 500, top: 20, width: 400, height: 50 }, { width: 960, height: 540 })).toEqual({ dx: 0, dy: 74 });
    expect(copyOffset({ left: 0, top: 0, width: 900, height: 500 }, { width: 960, height: 540 })).toEqual({ dx: 24, dy: 24 });
  });
});

describe("comparing a copy with its original", () => {
  it("finds nothing when the copy matches apart from where it is", () => {
    const moved = chevron({ left: 134, top: 20 });
    expect(compareRecipes(recipe([chevron()]), recipe([moved]), 124, 0)).toEqual([]);
  });

  it("accepts a line rebuilt without mirroring", () => {
    const line = (f: Partial<Frame>): RecipeNode => ({ kind: "line", name: "L", frame: frame(f), issues: [], look: { line: { visible: true, color: "#000000", weight: 1 } } });
    const original = line({ flipV: true });
    const g = lineFrame(original.frame);
    expect(compareRecipes(recipe([original]), recipe([line({ ...g, left: g.left + 200 })]), 200, 0)).toEqual([]);
    expect(compareRecipes(recipe([original]), recipe([line({ left: 210 })]), 200, 0)).toEqual(["1. Line: line runs in a different direction or length"]);
  });

  it("reports what changed", () => {
    const copy = chevron();
    copy.geometry = "Rectangle";
    copy.look.adjustments = [50000];
    copy.look.fill = { type: "Solid", color: "#000000", transparency: 0 };
    copy.look.text!.runs = [{ start: 0, length: 8, font: { name: "Aptos", size: 18, bold: false } }];
    copy.frame.rotation = 90;
    expect(compareRecipes(recipe([chevron()]), recipe([copy]), 0, 0)).toEqual([
      "1. Chevron with text: shape type Chevron came back Rectangle",
      "1. Chevron with text: rotation 0° came back 90°",
      "1. Chevron with text: fill colour #C43E1C came back #000000",
      "1. Chevron with text: adjustment 1 was 0.5, came back 50000",
      '1. Chevron with text: font differs from character 3 ({"name":"Aptos","size":18,"bold":true} vs {"name":"Aptos","size":18,"bold":false})',
    ]);
  });

  it("reports paragraph spacing that changed in the copy", () => {
    const withStyles = (style: string): RecipeNode => ({
      ...chevron(),
      text: { length: 8, paragraphs: [{ start: 0, length: 8 }], runs: [], styles: [style] },
    });
    const a = withStyles("marL=0 lnSpc=pct:90000 bullet=none");
    expect(compareRecipes(recipe([a]), recipe([withStyles("marL=0 lnSpc=pct:90000 bullet=none")]), 0, 0)).toEqual([]);
    expect(compareRecipes(recipe([a]), recipe([withStyles("marL=0 lnSpc=pct:100000 bullet=none")]), 0, 0)).toEqual([
      "1. Chevron with text: paragraph 1 lnSpc pct:90000 → pct:100000",
    ]);
  });

  it("describes shapes in plain words", () => {
    expect(describeNode(chevron())).toBe("Chevron with text");
    expect(describeNode({ ...chevron(), geometry: "Round2SameRectangle", look: {} })).toBe("Round 2 Same Rectangle");
    expect(describeNode({ ...chevron(), textBox: true })).toBe("Text box with text");
  });
});
