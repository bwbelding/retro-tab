import { describe, expect, it } from "vitest";
import { findMatches, replaceIn, type TextSource } from "../src/lib/findReplace";

const src = (text: string, slideIndex = 0, more: Partial<TextSource> = {}): TextSource => ({ slideId: `s${slideIndex}`, slideIndex, path: ["a"], where: "Text box", text, ...more });

describe("find", () => {
  it("finds every match, ignoring case unless asked", () => {
    const sources = [src("ARR is $4.2M. Q3 ARR grew."), src("arr", 1)];
    expect(findMatches(sources, "ARR", { matchCase: false, wholeWord: false }).map((m) => `${m.source.slideIndex}:${m.start}:${m.found}`)).toEqual(["0:0:ARR", "0:17:ARR", "1:0:arr"]);
    expect(findMatches(sources, "ARR", { matchCase: true, wholeWord: false })).toHaveLength(2);
    expect(findMatches(sources, "", { matchCase: false, wholeWord: false })).toEqual([]);
  });

  it("matches whole words and numbers only, when asked", () => {
    const text = "$4.2M, 14.2%, 4.25x and 4.2";
    const starts = (wholeWord: boolean) => findMatches([src(text)], "4.2", { matchCase: false, wholeWord }).map((m) => m.start);
    expect(starts(false)).toEqual([1, 8, 14, 24]);
    expect(starts(true)).toEqual([1, 24]);
    expect(findMatches([src("Q3 FY26 and Q3FY26")], "Q3 FY26", { matchCase: false, wholeWord: true })).toHaveLength(1);
    expect(findMatches([src("Revenue 1,4.2 and 4.2,5 and 4.2, then")], "4.2", { matchCase: false, wholeWord: true }).map((m) => m.start)).toEqual([28]);
    expect(findMatches([src("plan, planned, Plan.")], "plan", { matchCase: false, wholeWord: true })).toHaveLength(2);
    // A query starting or ending with punctuation can touch a word on that side.
    expect(findMatches([src("($4.2M)")], "($4.2M)", { matchCase: false, wholeWord: true })).toHaveLength(1);
  });

  it("shows each match in context with distinct keys", () => {
    const text = "In the third quarter our annual recurring revenue reached $4.2M, up from $3.1M in the second quarter of the year.";
    const [m] = findMatches([src(text)], "$4.2M", { matchCase: false, wholeWord: true });
    expect(m.before).toBe("…curring revenue reached ");
    expect(m.after).toBe(", up from $3.1M in the s…");
    const keys = findMatches([src("a a a"), src("a", 1, { cell: { row: 1, column: 2 } })], "a", { matchCase: false, wholeWord: true }).map((x) => x.key);
    expect(new Set(keys).size).toBe(4);
  });
});

describe("replace", () => {
  it("replaces chosen matches from the end, so earlier positions stay right", () => {
    const text = "Q3: 4.2 then 4.2 then 4.2";
    const matches = findMatches([src(text)], "4.2", { matchCase: false, wholeWord: true });
    expect(replaceIn(text, [matches[0], matches[2]], "4.75")).toBe("Q3: 4.75 then 4.2 then 4.75");
    expect(replaceIn(text, [], "x")).toBe(text);
  });
});
