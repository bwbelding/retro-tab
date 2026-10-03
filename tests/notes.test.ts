import { describe, expect, it } from "vitest";
import { STICKY_SIZE, stampRect, stickyPosition } from "../src/lib/notes";

const slide = { width: 960, height: 540 };
const r = (left: number, top: number, width: number, height: number) => ({ id: "x", left, top, width, height });

describe("stickyPosition", () => {
  it("goes in the top-right corner when nothing is selected", () => {
    expect(stickyPosition(slide, [])).toEqual({ left: 960 - STICKY_SIZE.width - 24, top: 24 });
  });

  it("goes just right of the selection when it fits", () => {
    expect(stickyPosition(slide, [r(100, 200, 200, 50)])).toEqual({ left: 312, top: 200 });
  });

  it("falls back to the corner when there's no room on the right", () => {
    expect(stickyPosition(slide, [r(700, 200, 200, 50)])).toEqual({ left: 756, top: 24 });
  });

  it("keeps the note on the slide vertically", () => {
    expect(stickyPosition(slide, [r(100, 500, 50, 20)]).top).toBe(540 - STICKY_SIZE.height);
  });
});

describe("stampRect", () => {
  it("sizes to the label and sits top-right", () => {
    const draft = stampRect(slide, "DRAFT");
    const conf = stampRect(slide, "CONFIDENTIAL");
    expect(conf.width).toBeGreaterThan(draft.width);
    expect(draft.left + draft.width).toBe(960 - 24);
    expect(draft.top).toBe(24);
  });
});
