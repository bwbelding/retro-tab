import { beforeAll, describe, expect, it } from "vitest";
import { Window } from "happy-dom";
import { allIssues, parseSlide, resolvePart, taggedNodes, textIssues, textLayout, type XmlNode } from "../src/lib/ooxml";
import { openZip } from "../src/lib/zip";
import { NS, pptxBytes } from "./helpers/pptx";

beforeAll(() => {
  // Only the XML parser is borrowed from happy-dom; zip streams stay Node's own.
  globalThis.DOMParser = new Window().DOMParser as unknown as typeof DOMParser;
});

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

  it("flags text the API can't format", () => {
    const doc = new DOMParser().parseFromString(
      `<p:txBody ${NS}><a:bodyPr vert="vert"/><a:p><a:pPr><a:lnSpc/></a:pPr><a:r><a:rPr baseline="30000"/><a:t>x</a:t></a:r></a:p></p:txBody>`,
      "application/xml",
    );
    expect(textLayout(doc.documentElement)).toEqual({ length: 1, paragraphs: [{ start: 0, length: 1 }], runs: [{ start: 0, length: 1 }] });
    expect(textIssues(doc.documentElement)).toEqual(["vertical or rotated text", "line or paragraph spacing", "character spacing or superscript"]);
  });
});
