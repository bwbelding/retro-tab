import { describe, expect, it } from "vitest";
import { appendixInPlace } from "../src/lib/appendix";

describe("appendix", () => {
  it("knows when the appendix already sits at the end behind its divider", () => {
    expect(appendixInPlace([8, 9], 7, 10)).toBe(true);
    expect(appendixInPlace([3, 9], 7, 10)).toBe(false); // one is still in the middle
    expect(appendixInPlace([8, 9], -1, 10)).toBe(false); // no divider yet
    expect(appendixInPlace([7, 8], 6, 10)).toBe(false); // not at the very end
    expect(appendixInPlace([], 7, 10)).toBe(false);
  });
});
