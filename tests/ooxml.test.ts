import { beforeAll, describe, expect, it } from "vitest";
import { Window } from "happy-dom";
import { allIssues, parseSlide, resolvePart, taggedNodes, textDefaults, textIssues, textLayout, type XmlNode } from "../src/lib/ooxml";
import { openZip } from "../src/lib/zip";
import { NS, pptxBytes } from "./helpers/pptx";

beforeAll(() => {
  // Only the XML parser is borrowed from happy-dom; zip streams stay Node's own.
  globalThis.DOMParser = new Window().DOMParser as unknown as typeof DOMParser;
});

const PLAIN = "marL=0 indent=0 lnSpc=pct:100000 spcBef=0 spcAft=0 bullet=none";
const xml = (text: string) => new DOMParser().parseFromString(text, "application/xml");

async function exported() {
  return openZip(await pptxBytes());
}

describe("exported slide XML", () => {
  it("resolves relationship targets", () => {
    expect(resolvePart("ppt/slides/slide1.xml", "../media/image1.png")).toBe("ppt/media/image1.png");
    expect(resolvePart("ppt/slides/slide1.xml", "/ppt/tags/tag1.xml")).toBe("ppt/tags/tag1.xml");
  });

  it("finds tagged shapes, outermost first, and skips untagged ones", async () => {
    const { nodes } = await parseSlide(await exported());
    expect(nodes.map((n) => n.name)).toEqual(["Plain", "Badge", "Group 1", "Icon", "Scribble", "Chart", "Shadowed"]);
    const tagged = taggedNodes(nodes);
    expect(tagged.map((n) => [n.name, n.cid])).toEqual([
      ["Badge", "1"],
      ["Group 1", "2"],
      ["Icon", "5"],
      ["Scribble", "6"],
      ["Chart", "7"],
      ["Shadowed", "8"],
    ]);
  });

  it("reads shape type, frame and text runs", async () => {
    const badge = taggedNodes((await parseSlide(await exported())).nodes)[0];
    expect(badge).toMatchObject({ kind: "shape", geometry: "Chevron", issues: [] });
    expect(badge.frame).toEqual({ left: 10, top: 20, width: 100, height: 50, rotation: 90, flipH: false, flipV: false });
    expect(badge.text).toEqual({
      length: 12,
      paragraphs: [
        { start: 0, length: 8 },
        { start: 9, length: 3 },
      ],
      runs: [
        { start: 0, length: 2 },
        { start: 2, length: 6 },
        { start: 9, length: 1 },
        { start: 11, length: 1 },
      ],
      styles: [PLAIN, PLAIN],
    });
  });

  it("maps group children onto the slide and collects their issues", async () => {
    const group = taggedNodes((await parseSlide(await exported())).nodes)[1];
    expect(group.frame).toMatchObject({ left: 100, top: 100, width: 200, height: 100 });
    const [inner, connector] = group.children as XmlNode[];
    expect(inner).toMatchObject({ geometry: "RoundRectangle", frame: { left: 120, top: 100, width: 40, height: 20 } });
    expect(connector).toMatchObject({ kind: "line", issues: ["arrowheads"], frame: { flipV: true, width: 20, height: 20 } });
    expect(allIssues(group)).toEqual(["arrowheads"]);
  });

  it("flags what the add-in API can't recreate", async () => {
    const [, , icon, scribble, chart, shadowed] = taggedNodes((await parseSlide(await exported())).nodes);
    expect(icon).toMatchObject({ kind: "picture", media: "ppt/media/image1.png" });
    expect(icon.issues).toContain("SVG icon");
    expect(scribble).toMatchObject({ kind: "other", issues: ["freeform or edited points"] });
    expect(chart).toMatchObject({ kind: "other", issues: ["chart"] });
    expect(shadowed.issues).toEqual(["shadow or effect"]);
  });

  it("flags text the API can't format, naming the values", () => {
    const doc = xml(`<p:txBody ${NS}><a:bodyPr vert="vert"/><a:p><a:pPr><a:lnSpc><a:spcPct val="90000"/></a:lnSpc></a:pPr><a:r><a:rPr baseline="30000"/><a:t>x</a:t></a:r></a:p></p:txBody>`);
    expect(textLayout(doc.documentElement)).toEqual({
      length: 1,
      paragraphs: [{ start: 0, length: 1 }],
      runs: [{ start: 0, length: 1 }],
      styles: ["marL=0 indent=0 lnSpc=pct:90000 spcBef=0 spcAft=0 bullet=none"],
    });
    expect(textIssues(doc.documentElement)).toEqual([
      "vertical or rotated text",
      "line or paragraph spacing (line spacing 90% where a new shape gets 100%)",
      "character spacing or superscript",
    ]);
  });

  it("accepts paragraph settings a new shape gets from the template anyway", () => {
    const defaultTextStyle = xml(`<p:defaultTextStyle ${NS}><a:lvl1pPr marL="0"><a:spcBef><a:spcPts val="0"/></a:spcBef></a:lvl1pPr></p:defaultTextStyle>`).documentElement;
    const theme = xml(
      `<a:theme ${NS}><a:objectDefaults><a:spDef><a:lstStyle><a:lvl1pPr marL="171450" indent="-171450"><a:lnSpc><a:spcPct val="90000"/></a:lnSpc><a:buNone/></a:lvl1pPr></a:lstStyle></a:spDef></a:objectDefaults></a:theme>`,
    );
    const defaults = textDefaults(defaultTextStyle, theme);
    const body = (pPr: string) =>
      xml(`<p:txBody ${NS}><a:bodyPr/><a:p>${pPr}<a:r><a:t>Label</a:t></a:r></a:p></p:txBody>`).documentElement;
    // Written out in full by PowerPoint, but exactly what a new shape gets: fine.
    const asTemplate = `<a:pPr marL="171450" indent="-171450"><a:lnSpc><a:spcPct val="90000"/></a:lnSpc><a:spcBef><a:spcPts val="0"/></a:spcBef><a:buNone/></a:pPr>`;
    expect(textIssues(body(asTemplate), defaults)).toEqual([]);
    // A setting that differs is named.
    const spaced = asTemplate.replace('<a:spcPts val="0"/>', '<a:spcPts val="600"/>');
    expect(textIssues(body(spaced), defaults)).toEqual(["line or paragraph spacing (space before 6 pt where a new shape gets 0)"]);
    // Relying on inherited settings while new shapes get the template's: also named.
    expect(textIssues(body("<a:pPr/>"), defaults)).toEqual([
      "paragraph indents (left indent 0 pt where a new shape gets 13.5 pt)",
      "paragraph indents (first-line indent 0 pt where a new shape gets -13.5 pt)",
      "line or paragraph spacing (line spacing 100% where a new shape gets 90%)",
    ]);
    // Text boxes use the template's text-box defaults, which this one doesn't set.
    expect(textIssues(body("<a:pPr/>"), defaults, true)).toEqual([]);
    expect(textIssues(body(`<a:pPr><a:buChar char="•"/></a:pPr>`), defaults, true)).toEqual(["bullets (bullet “•” where a new text box gets none)"]);
  });
});
