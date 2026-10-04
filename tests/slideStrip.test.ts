import { beforeAll, describe, expect, it } from "vitest";
import { Window } from "happy-dom";
import { cleanSlide, stripToCaptured } from "../src/lib/slideStrip";
import { openZip } from "../src/lib/zip";
import { makeZip } from "./helpers/makeZip";

beforeAll(() => {
  const w = new Window();
  globalThis.DOMParser = w.DOMParser as unknown as typeof DOMParser;
  globalThis.XMLSerializer = w.XMLSerializer as unknown as typeof XMLSerializer;
});

const NS = `xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"`;
const REL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
const rels = (...items: [string, string, string][]) =>
  `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">${items.map(([id, type, target]) => `<Relationship Id="${id}" Type="${REL}/${type}" Target="${target}"/>`).join("")}</Relationships>`;
const shape = (name: string, tag: string | undefined, inner = "") =>
  `<p:sp><p:nvSpPr><p:cNvPr id="2" name="${name}"/><p:cNvSpPr/><p:nvPr>${tag ? `<p:custDataLst><p:tags r:id="${tag}"/></p:custDataLst>` : ""}</p:nvPr></p:nvSpPr><p:spPr>${inner}</p:spPr></p:sp>`;
// Ink and equations come wrapped in mc:AlternateContent.
const ink = (name: string, tag: string | undefined) =>
  `<mc:AlternateContent xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"><mc:Choice Requires="p14">${shape(name, tag)}</mc:Choice><mc:Fallback>${shape(name, tag)}</mc:Fallback></mc:AlternateContent>`;
const override = (part: string) => `<Override PartName="/${part}" ContentType="x"/>`;
const PARTS = [
  "ppt/presentation.xml",
  "ppt/slides/slide1.xml",
  "ppt/notesSlides/notesSlide1.xml",
  "ppt/charts/chart1.xml",
  "ppt/slideLayouts/slideLayout1.xml",
  "ppt/slideMasters/slideMaster1.xml",
  "ppt/theme/theme1.xml",
  "ppt/tags/tag1.xml",
  "ppt/tags/tag2.xml",
  "ppt/tags/tag3.xml",
  "ppt/commentAuthors.xml",
  "ppt/comments/comment1.xml",
  "docProps/core.xml",
  "docProps/app.xml",
];

async function exportedSlide() {
  return makeZip({
    "[Content_Types].xml": `<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="png" ContentType="image/png"/>${PARTS.map(override).join("")}</Types>`,
    "_rels/.rels": rels(["rId1", "officeDocument", "ppt/presentation.xml"], ["rId2", "metadata/core-properties", "docProps/core.xml"], ["rId3", "extended-properties", "docProps/app.xml"], ["rId4", "metadata/thumbnail", "docProps/thumbnail.jpeg"]),
    "docProps/core.xml": "<cp:coreProperties>Client confidential deck, by Brett</cp:coreProperties>",
    "docProps/app.xml": "<Properties>Q3 numbers</Properties>",
    "docProps/thumbnail.jpeg": new Uint8Array([1, 2, 3]),
    "ppt/presentation.xml": `<p:presentation ${NS}/>`,
    "ppt/_rels/presentation.xml.rels": rels(
      ["rId1", "slideMaster", "slideMasters/slideMaster1.xml"],
      ["rId2", "slide", "slides/slide1.xml"],
      ["rId3", "commentAuthors", "commentAuthors.xml"],
      ["rId4", "customXml", "../customXml/item1.xml"],
      ["rId5", "printerSettings", "printerSettings/printerSettings1.bin"],
    ),
    "ppt/commentAuthors.xml": "<p:cmAuthorLst>Brett Belding</p:cmAuthorLst>",
    "customXml/item1.xml": "<SharePoint>Legal hold</SharePoint>",
    "ppt/printerSettings/printerSettings1.bin": new Uint8Array([7]),
    "ppt/comments/comment1.xml": "<p:cmLst>Fix this number</p:cmLst>",
    "ppt/slides/slide1.xml": `<p:sld ${NS}><p:cSld><p:spTree><p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr/>${shape(
      "Saved badge",
      "rId10",
      '<a:blipFill><a:blip r:embed="rId20"/></a:blipFill>',
    )}${shape("Confidential note", "rId11", '<a:blipFill><a:blip r:embed="rId21"/></a:blipFill>')}<p:graphicFrame><p:nvGraphicFramePr><p:cNvPr id="9" name="Revenue chart"/><p:cNvGraphicFramePr/><p:nvPr/></p:nvGraphicFramePr><a:graphic><a:graphicData><c:chart xmlns:c="c" r:id="rId30"/></a:graphicData></a:graphic></p:graphicFrame>${ink("Saved ink", "rId12")}${ink("Other ink", undefined)}</p:spTree></p:cSld><p:timing><p:tnLst/></p:timing></p:sld>`,
    "ppt/slides/_rels/slide1.xml.rels": rels(
      ["rId1", "slideLayout", "../slideLayouts/slideLayout1.xml"],
      ["rId10", "tags", "../tags/tag1.xml"],
      ["rId11", "tags", "../tags/tag2.xml"],
      ["rId20", "image", "../media/image1.png"],
      ["rId21", "image", "../media/image2.png"],
      ["rId30", "chart", "../charts/chart1.xml"],
      ["rId12", "tags", "../tags/tag3.xml"],
      ["rId40", "notesSlide", "../notesSlides/notesSlide1.xml"],
      ["rId41", "comments", "../comments/comment1.xml"],
    ),
    "ppt/tags/tag3.xml": `<p:tagLst ${NS}><p:tag name="RETRO_CAPTURE" val="3"/></p:tagLst>`,
    "ppt/tags/tag1.xml": `<p:tagLst ${NS}><p:tag name="RETRO_CAPTURE" val="1"/></p:tagLst>`,
    "ppt/tags/tag2.xml": `<p:tagLst ${NS}><p:tag name="OTHER" val="x"/></p:tagLst>`,
    "ppt/media/image1.png": new Uint8Array([1]),
    "ppt/media/image2.png": new Uint8Array([2]),
    "ppt/media/image3.png": new Uint8Array([3]),
    "ppt/charts/chart1.xml": "<c:chartSpace>revenue 12.4m</c:chartSpace>",
    "ppt/charts/_rels/chart1.xml.rels": rels(["rId1", "package", "../embeddings/Microsoft_Excel_Worksheet.xlsx"]),
    "ppt/embeddings/Microsoft_Excel_Worksheet.xlsx": new Uint8Array([9]),
    "ppt/notesSlides/notesSlide1.xml": "<p:notes>Speaker notes: do not share</p:notes>",
    "ppt/notesSlides/_rels/notesSlide1.xml.rels": rels(["rId1", "slide", "../slides/slide1.xml"]),
    "ppt/slideLayouts/slideLayout1.xml": `<p:sldLayout ${NS}/>`,
    "ppt/slideLayouts/_rels/slideLayout1.xml.rels": rels(["rId1", "slideMaster", "../slideMasters/slideMaster1.xml"]),
    "ppt/slideMasters/slideMaster1.xml": `<p:sldMaster ${NS}/>`,
    // The master uses image2 too, so it must survive even though the removed shape used it.
    "ppt/slideMasters/_rels/slideMaster1.xml.rels": rels(["rId1", "slideLayout", "../slideLayouts/slideLayout1.xml"], ["rId2", "theme", "../theme/theme1.xml"], ["rId3", "image", "../media/image2.png"]),
    "ppt/theme/theme1.xml": `<a:theme ${NS}/>`,
  });
}

describe("stripping a saved slide", () => {
  it("keeps only the saved shapes and what the slide's template needs", async () => {
    const zip = openZip(await stripToCaptured(await exportedSlide()));
    expect([...zip.names].sort()).toEqual(
      [
        "[Content_Types].xml",
        "_rels/.rels",
        "ppt/presentation.xml",
        "ppt/_rels/presentation.xml.rels",
        "ppt/slides/slide1.xml",
        "ppt/slides/_rels/slide1.xml.rels",
        "ppt/tags/tag1.xml",
        "ppt/tags/tag3.xml",
        "ppt/media/image1.png",
        "ppt/media/image2.png",
        "ppt/slideLayouts/slideLayout1.xml",
        "ppt/slideLayouts/_rels/slideLayout1.xml.rels",
        "ppt/slideMasters/slideMaster1.xml",
        "ppt/slideMasters/_rels/slideMaster1.xml.rels",
        "ppt/theme/theme1.xml",
      ].sort(),
    );
  });

  it("removes other shapes, animations, notes, charts and document properties", async () => {
    const zip = openZip(await stripToCaptured(await exportedSlide()));
    const slide = (await zip.text("ppt/slides/slide1.xml"))!;
    expect(slide).toContain("Saved badge");
    expect(slide).not.toContain("Confidential note");
    expect(slide).not.toContain("Revenue chart");
    expect(slide).not.toContain("timing");
    expect(slide).toContain("Saved ink");
    expect(slide).not.toContain("Other ink");
    const slideRels = (await zip.text("ppt/slides/_rels/slide1.xml.rels"))!;
    expect(slideRels.match(/Id="(rId\d+)"/g)).toEqual(['Id="rId1"', 'Id="rId10"', 'Id="rId20"', 'Id="rId12"']);
    const presRels = (await zip.text("ppt/_rels/presentation.xml.rels"))!;
    expect(presRels).toContain("slideMaster");
    expect(presRels).not.toContain("commentAuthors");
    expect(presRels).not.toContain("customXml");
    expect(presRels).not.toContain("printerSettings");
    const rootRels = (await zip.text("_rels/.rels"))!;
    expect(rootRels).toContain("officeDocument");
    expect(rootRels).not.toContain("docProps");
    const types = (await zip.text("[Content_Types].xml"))!;
    for (const gone of ["notesSlide1", "chart1", "docProps", "tag2", "commentAuthors", "comment1"]) expect(types).not.toContain(gone);
    expect(types).toContain("/ppt/slides/slide1.xml");
    for (const name of zip.names) expect(await zip.bytes(name)).toBeDefined();
  });

  it("keeps a whole saved slide and its speaker notes, but not comments or the deck's properties", async () => {
    const zip = openZip(await cleanSlide(await exportedSlide()));
    const slide = (await zip.text("ppt/slides/slide1.xml"))!;
    for (const kept of ["Saved badge", "Confidential note", "Revenue chart", "Saved ink", "Other ink", "timing"]) expect(slide).toContain(kept);
    for (const part of ["ppt/notesSlides/notesSlide1.xml", "ppt/charts/chart1.xml", "ppt/embeddings/Microsoft_Excel_Worksheet.xlsx", "ppt/tags/tag2.xml", "ppt/media/image2.png"]) {
      expect(zip.names).toContain(part);
    }
    for (const gone of ["ppt/comments/comment1.xml", "ppt/commentAuthors.xml", "docProps/core.xml", "docProps/thumbnail.jpeg", "customXml/item1.xml", "ppt/media/image3.png"]) {
      expect(zip.names).not.toContain(gone);
    }
    expect(await zip.text("ppt/slides/_rels/slide1.xml.rels")).not.toContain("comments");
  });
});
