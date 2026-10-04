import { describe, expect, it } from "vitest";
import { titleState, wordCount } from "../src/lib/storyline";

const row = (title: string) => ({ slideId: "s", index: 0, shapeId: "t", title });

describe("storyline", () => {
  it("counts words, ignoring stray punctuation", () => {
    expect(wordCount("Market grew 30% in Q3")).toBe(5);
    expect(wordCount("Plan – next steps")).toBe(3);
    expect(wordCount("  ")).toBe(0);
  });

  it("flags topic labels and missing titles", () => {
    expect(titleState(row("Market size"))).toBe("short");
    expect(titleState(row("Q3 update: on track"))).toBe("short");
    expect(titleState(row("Market grew 30% in Q3"))).toBe("ok");
    expect(titleState(row(""))).toBe("missing");
  });
});
