import { describe, expect, it } from "vitest";
import {
  commonValue,
  distributeWithGap,
  fromPoints,
  matchSize,
  nudge,
  pickReference,
  setAsWhole,
  setEach,
  swapPositions,
  toPoints,
  type Rect,
} from "../src/lib/layout";

const r = (id: string, left: number, top: number, width: number, height: number): Rect => ({ id, left, top, width, height });

describe("units", () => {
  it("converts cm and inches to points and back", () => {
    expect(toPoints(1, "in")).toBe(72);
    expect(toPoints(2.54, "cm")).toBeCloseTo(72);
    expect(fromPoints(36, "in")).toBe(0.5);
  });
});

describe("pickReference", () => {
  const rects = [r("a", 0, 0, 10, 10), r("b", 0, 0, 30, 30), r("c", 0, 0, 5, 5)];
  it("picks by selection order or area", () => {
    expect(pickReference(rects, "first").id).toBe("a");
    expect(pickReference(rects, "last").id).toBe("c");
    expect(pickReference(rects, "largest").id).toBe("b");
    expect(pickReference(rects, "smallest").id).toBe("c");
  });
});

describe("matchSize", () => {
  const rects = [r("a", 10, 10, 100, 50), r("b", 200, 20, 40, 80)];
  it("matches width only, keeping positions", () => {
    expect(matchSize(rects, "width")).toEqual([r("a", 10, 10, 100, 50), r("b", 200, 20, 100, 80)]);
  });
  it("matches height to the last shape", () => {
    expect(matchSize(rects, "height", "last")[0].height).toBe(80);
  });
  it("matches both", () => {
    expect(matchSize(rects, "both")[1]).toEqual(r("b", 200, 20, 100, 50));
  });
});

describe("swapPositions", () => {
  it("swaps centers so different sizes trade places", () => {
    const [a, b] = swapPositions(r("a", 0, 0, 100, 100), r("b", 300, 0, 50, 50));
    expect(a).toEqual(r("a", 275, -25, 100, 100)); // center was (325, 25)
    expect(b).toEqual(r("b", 25, 25, 50, 50)); // center was (50, 50)
  });
});

describe("distributeWithGap", () => {
  it("keeps the first shape and spaces the rest by the gap, preserving input order", () => {
    const out = distributeWithGap([r("c", 500, 0, 20, 10), r("a", 0, 0, 100, 10), r("b", 300, 0, 50, 10)], 12, "horizontal");
    expect(out.map((x) => [x.id, x.left])).toEqual([
      ["c", 174],
      ["a", 0],
      ["b", 112],
    ]);
  });
  it("works vertically", () => {
    const out = distributeWithGap([r("a", 0, 0, 10, 40), r("b", 0, 90, 10, 20)], 6, "vertical");
    expect(out[1].top).toBe(46);
  });
});

describe("nudge", () => {
  it("moves every shape", () => {
    expect(nudge([r("a", 1, 2, 3, 4)], -1, 5)).toEqual([r("a", 0, 7, 3, 4)]);
  });
});

describe("setEach", () => {
  it("sets only the given values", () => {
    expect(setEach([r("a", 1, 2, 30, 40)], { left: 10, height: 80 })).toEqual([r("a", 10, 2, 30, 80)]);
  });
  it("keeps the aspect ratio when locked", () => {
    expect(setEach([r("a", 0, 0, 100, 50)], { width: 200 }, true)[0].height).toBe(100);
  });
});

describe("setAsWhole", () => {
  const rects = [r("a", 100, 100, 50, 50), r("b", 200, 100, 50, 50)]; // bounds 100,100 150x50
  it("moves the block's corner", () => {
    expect(setAsWhole(rects, { left: 0, top: 0 })).toEqual([r("a", 0, 0, 50, 50), r("b", 100, 0, 50, 50)]);
  });
  it("scales positions and sizes to the new width, locked aspect", () => {
    const out = setAsWhole(rects, { width: 300 }, true);
    expect(out).toEqual([r("a", 100, 100, 100, 100), r("b", 300, 100, 100, 100)]);
  });
});

describe("commonValue", () => {
  it("returns the shared value or undefined when mixed", () => {
    expect(commonValue([r("a", 1, 0, 5, 5), r("b", 1, 9, 6, 5)], "left")).toBe(1);
    expect(commonValue([r("a", 1, 0, 5, 5), r("b", 1, 9, 6, 5)], "width")).toBeUndefined();
  });
});
