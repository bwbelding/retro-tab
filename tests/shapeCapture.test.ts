// Runs capture and rebuild against a stand-in for PowerPoint's API, with a real exported .pptx,
// to check the two halves are wired together. What PowerPoint itself does is tested on the Mac.

import { beforeAll, describe, expect, it } from "vitest";
import { Window } from "happy-dom";
import { captureShapes, rebuildIssues, rebuildShapes, type RecipeNode } from "../src/lib/shapeCapture";
import { bytesToBase64 } from "../src/lib/zip";
import { pptxBytes } from "./helpers/pptx";

/* eslint-disable @typescript-eslint/no-explicit-any */

beforeAll(() => {
  globalThis.DOMParser = new Window().DOMParser as unknown as typeof DOMParser;
  (globalThis as any).PowerPoint = {
    ShapeAutoSize: { autoSizeNone: "AutoSizeNone" },
    ConnectorType: { straight: "Straight" },
    GeometricShapeType: { rectangle: "Rectangle" },
  };
});

function fakeContext() {
  const pending: (() => Promise<void>)[] = [];
  return { pending, sync: async () => { for (const p of pending.splice(0)) await p(); } };
}

const BADGE_TEXT = "Hi there\rA\vB";

function liveShape(name: string, type: string, children: any[] = []) {
  const tags = new Map<string, string>();
  return {
    name,
    id: `id-${name}`,
    type,
    tags: { add: (k: string, v: string) => tags.set(k, v), delete: (k: string) => tags.delete(k), map: tags },
    load() {},
    group: { shapes: { items: children, load() {} } },
    fill: { load: () => ({ type: "Solid", foregroundColor: "#C43E1C", transparency: 0 }) },
    lineFormat: { load: () => ({ visible: name === "Connector", color: "#000000", weight: 1, dashStyle: "Solid", style: "Single", transparency: 0 }) },
    adjustments: { load: () => ({ count: 1, get: () => ({ value: 0.5 }) }) },
    textFrame: {
      load: () => ({ leftMargin: 7.2, rightMargin: 7.2, topMargin: 3.6, bottomMargin: 3.6, verticalAlignment: "Middle", wordWrap: true, autoSizeSetting: "AutoSizeNone" }),
      textRange: {
        load: () => ({
          text: BADGE_TEXT,
          getSubstring: (start: number) => ({
            font: { load: () => ({ name: "Aptos", size: 18, bold: start === 2, italic: false, color: null }) },
            paragraphFormat: { load: () => ({ horizontalAlignment: "Center" }) },
          }),
        }),
      },
    },
  };
}

function liveSlide(context: ReturnType<typeof fakeContext>, all: ReturnType<typeof liveShape>[]) {
  return {
    id: "slide-1",
    exportAsBase64() {
      const result = { value: "" };
      // The export carries whatever capture tags are on the shapes when it runs.
      const order = ["Badge", "Group 1", "Inner", "Connector", "Icon", "Scribble", "Chart", "Shadowed"];
      const values = order.map((n) => all.find((s) => s.name === n)?.tags.map.get("RETRO_CAPTURE") ?? "");
      context.pending.push(async () => void (result.value = bytesToBase64(await pptxBytes(values))));
      return result;
    },
  };
}

async function captured() {
  const context = fakeContext();
  const inner = liveShape("Inner", "GeometricShape");
  const connector = liveShape("Connector", "Line");
  const group = liveShape("Group 1", "Group", [inner, connector]);
  const tops = [liveShape("Badge", "GeometricShape"), group, liveShape("Icon", "Image"), liveShape("Scribble", "Freeform"), liveShape("Chart", "Chart"), liveShape("Shadowed", "GeometricShape")];
  const all = [...tops, inner, connector];
  const capture = await captureShapes(context as any, liveSlide(context, all) as any, tops as any);
  return { capture, all };
}

describe("capturing shapes", () => {
  it("finds every selected shape and removes the temporary tags", async () => {
    const { capture, all } = await captured();
    expect(capture).toMatchObject({ found: 6, selected: 6 });
    expect(capture.recipe.nodes.map((n) => n.name)).toEqual(["Badge", "Group 1", "Icon", "Scribble", "Chart", "Shadowed"]);
    expect(all.every((s) => s.tags.map.size === 0)).toBe(true);
    expect(capture.recipe.box).toEqual({ left: 0, top: 0, width: 300, height: 200 });
  });

  it("combines the shape type from the file with formatting from PowerPoint", async () => {
    const badge = (await captured()).capture.recipe.nodes[0];
    expect(badge.geometry).toBe("Chevron");
    expect(badge.look.fill).toEqual({ type: "Solid", color: "#C43E1C", transparency: 0 });
    expect(badge.look.adjustments).toEqual([0.5]);
    expect(badge.look.text?.runs.map((r) => [r.start, r.length, r.font.bold])).toEqual([
      [0, 2, false],
      [2, 6, true],
      [9, 1, false],
      [11, 1, false],
    ]);
    expect(badge.look.text?.paragraphs).toEqual([
      { start: 0, length: 8, align: "Center" },
      { start: 9, length: 3, align: "Center" },
    ]);
    expect(rebuildIssues(badge)).toEqual([]);
  });

  it("keeps group children, pictures and the reasons others need the helper slide", async () => {
    const [, group, icon, scribble, chart, shadowed] = (await captured()).capture.recipe.nodes;
    expect(group.children?.map((c) => [c.name, c.kind])).toEqual([
      ["Inner", "shape"],
      ["Connector", "line"],
    ]);
    expect(group.children?.[1].look.line).toMatchObject({ visible: true, weight: 1 });
    expect(icon.image).toBe("iVBORw==");
    expect(rebuildIssues(group)).toEqual(["arrowheads"]);
    expect(rebuildIssues(icon)).toContain("SVG icon");
    expect(rebuildIssues(scribble)).toEqual(["freeform or edited points"]);
    expect(rebuildIssues(chart)).toEqual(["chart"]);
    expect(rebuildIssues(shadowed)).toEqual(["shadow or effect"]);
  });
});

describe("rebuilding shapes", () => {
  function recordingSlide() {
    const made: any[] = [];
    let selected: string[] = [];
    const shape = (how: string, arg?: unknown) => {
      const s: any = {
        how,
        arg,
        id: `new-${made.length + 1}`,
        load() {},
        adjustments: { values: [] as number[], set: (i: number, v: number) => (s.adjustments.values[i] = v) },
        fill: { setSolidColor: (c: string) => (s.fill.color = c), clear: () => (s.fill.cleared = true), setImage: (b: string) => (s.fill.image = b) },
        lineFormat: {},
        subs: [] as any[],
        textFrame: { textRange: { getSubstring: (start: number, length: number) => { const sub = { start, length, font: {}, paragraphFormat: {} }; s.subs.push(sub); return sub; } } },
      };
      made.push(s);
      return s;
    };
    const slide = {
      shapes: {
        addGeometricShape: (type: string) => shape("geometric", type),
        addTextBox: (text: string) => shape("textbox", text),
        addLine: (type: string, box: unknown) => shape("line", { type, box }),
        addGroup: (ids: string[]) => shape("group", ids),
      },
      setSelectedShapes: (ids: string[]) => (selected = ids),
    };
    return { slide, made, selected: () => selected };
  }

  it("rebuilds a captured shape exactly, offset and selected", async () => {
    const badge = (await captured()).capture.recipe.nodes[0];
    const { slide, made, selected } = recordingSlide();
    await rebuildShapes(fakeContext() as any, slide as any, { version: 1, box: { left: 0, top: 0, width: 0, height: 0 }, nodes: [badge] }, 100, 5);
    const s = made[0];
    expect([s.how, s.arg, s.left, s.top, s.width, s.height, s.rotation, s.name]).toEqual(["geometric", "Chevron", 110, 25, 100, 50, 90, "Badge"]);
    expect(s.fill.color).toBe("#C43E1C");
    expect(s.adjustments.values).toEqual([0.5]);
    expect(s.textFrame.textRange.text).toBe(BADGE_TEXT);
    expect(s.textFrame.autoSizeSetting).toBe("AutoSizeNone");
    expect(s.subs.filter((x: any) => x.font.bold).map((x: any) => [x.start, x.length])).toEqual([[2, 6]]);
    expect(selected()).toEqual(["new-1"]);
  });

  it("groups rebuilt children and turns mirrored lines into rotated ones", async () => {
    const line: RecipeNode = {
      kind: "line",
      name: "L",
      frame: { left: 0, top: 0, width: 100, height: 20, rotation: 0, flipH: false, flipV: true },
      issues: [],
      look: { line: { visible: true, color: "#000000", weight: 2 } },
    };
    const box: RecipeNode = { ...line, kind: "shape", name: "B", geometry: "Ellipse", look: { fill: { type: "NoFill" } } };
    const group: RecipeNode = { kind: "group", name: "G", frame: { ...line.frame, flipV: false }, issues: [], look: {}, children: [box, line] };
    const { slide, made } = recordingSlide();
    await rebuildShapes(fakeContext() as any, slide as any, { version: 1, box: { left: 0, top: 0, width: 0, height: 0 }, nodes: [group] }, 0, 0);
    expect(made.map((m) => m.how)).toEqual(["geometric", "line", "group"]);
    expect(made[0].fill.cleared).toBe(true);
    expect(made[1].arg).toEqual({ type: "Straight", box: { left: 40, top: -40, width: 20, height: 100 } });
    expect(made[1].rotation).toBe(90);
    expect(made[1].lineFormat).toMatchObject({ visible: true, color: "#000000", weight: 2 });
    expect(made[2].arg).toEqual(["new-1", "new-2"]);
  });
});
