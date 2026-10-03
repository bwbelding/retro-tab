// A small exported slide (.pptx) for tests: tagged shapes of every kind Retro handles, as
// PowerPoint writes them, plus the theme and tag parts they refer to.

import { makeZip } from "./makeZip";

export const NS = `xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"`;
const xfrm = (x: number, y: number, cx: number, cy: number, attrs = "") => `<a:xfrm${attrs}><a:off x="${x}" y="${y}"/><a:ext cx="${cx}" cy="${cy}"/></a:xfrm>`;
const nvPr = (tag?: string) => (tag ? `<p:nvPr><p:custDataLst><p:tags r:id="${tag}"/></p:custDataLst></p:nvPr>` : "<p:nvPr/>");
const sp = (id: number, name: string, tag: string | undefined, spPr: string, extra = "") =>
  `<p:sp><p:nvSpPr><p:cNvPr id="${id}" name="${name}"/><p:cNvSpPr/>${nvPr(tag)}</p:nvSpPr><p:spPr>${spPr}</p:spPr>${extra}</p:sp>`;
const PT = 12700;

export const SLIDE = `<p:sld ${NS}><p:cSld><p:spTree>
<p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr/>
${sp(2, "Plain", undefined, `${xfrm(0, 0, PT, PT)}<a:prstGeom prst="rect"/>`)}
${sp(
  3,
  "Badge",
  "rId10",
  `${xfrm(10 * PT, 20 * PT, 100 * PT, 50 * PT, ' rot="5400000"')}<a:prstGeom prst="chevron"><a:avLst/></a:prstGeom><a:solidFill><a:srgbClr val="FF0000"/></a:solidFill>`,
  `<p:txBody><a:bodyPr/><a:p><a:r><a:t>Hi</a:t></a:r><a:r><a:rPr b="1"/><a:t> there</a:t></a:r></a:p><a:p><a:r><a:t>A</a:t></a:r><a:br/><a:r><a:t>B</a:t></a:r></a:p></p:txBody>`,
)}
<p:grpSp><p:nvGrpSpPr><p:cNvPr id="4" name="Group 1"/><p:cNvGrpSpPr/>${nvPr("rId11")}</p:nvGrpSpPr>
<p:grpSpPr><a:xfrm><a:off x="${100 * PT}" y="${100 * PT}"/><a:ext cx="${200 * PT}" cy="${100 * PT}"/><a:chOff x="0" y="0"/><a:chExt cx="${100 * PT}" cy="${50 * PT}"/></a:xfrm></p:grpSpPr>
${sp(5, "Inner", "rId12", `${xfrm(10 * PT, 0, 20 * PT, 10 * PT)}<a:prstGeom prst="roundRect"/>`)}
<p:cxnSp><p:nvCxnSpPr><p:cNvPr id="6" name="Connector"/><p:cNvCxnSpPr/>${nvPr("rId13")}</p:nvCxnSpPr><p:spPr>${xfrm(0, 0, 10 * PT, 10 * PT, ' flipV="1"')}<a:prstGeom prst="straightConnector1"/><a:ln><a:tailEnd type="triangle"/></a:ln></p:spPr></p:cxnSp>
</p:grpSp>
<p:pic><p:nvPicPr><p:cNvPr id="7" name="Icon"/><p:cNvPicPr/>${nvPr("rId14")}</p:nvPicPr><p:blipFill><a:blip r:embed="rId20"><a:extLst><a:ext uri="{96DAC541-7B7A-43D3-8B79-37D633B846F1}"><asvg:svgBlip xmlns:asvg="http://schemas.microsoft.com/office/drawing/2016/SVG/main" r:embed="rId21"/></a:ext></a:extLst></a:blip><a:stretch/></p:blipFill><p:spPr>${xfrm(0, 0, 24 * PT, 24 * PT)}<a:prstGeom prst="rect"/></p:spPr></p:pic>
${sp(8, "Scribble", "rId15", `${xfrm(0, 0, PT, PT)}<a:custGeom/>`)}
<p:graphicFrame><p:nvGraphicFramePr><p:cNvPr id="9" name="Chart"/><p:cNvGraphicFramePr/>${nvPr("rId16")}</p:nvGraphicFramePr><p:xfrm><a:off x="0" y="0"/><a:ext cx="${PT}" cy="${PT}"/></p:xfrm><a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/chart"/></a:graphic></p:graphicFrame>
${sp(10, "Shadowed", "rId17", `${xfrm(0, 0, PT, PT)}<a:prstGeom prst="ellipse"/>`, `<p:style><a:lnRef idx="2"/><a:fillRef idx="1"/><a:effectRef idx="3"/><a:fontRef idx="minor"/></p:style>`)}
</p:spTree></p:cSld></p:sld>`;

const rel = (id: string, target: string) => `<Relationship Id="${id}" Type="x" Target="${target}"/>`;
const rels = (...items: string[]) => `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">${items.join("")}</Relationships>`;
const tag = (val: string) => `<p:tagLst ${NS}><p:tag name="RETRO_CAPTURE" val="${val}"/></p:tagLst>`;
const THEME = `<a:theme ${NS}><a:themeElements><a:fmtScheme><a:effectStyleLst>
<a:effectStyle><a:effectLst/></a:effectStyle><a:effectStyle><a:effectLst/></a:effectStyle>
<a:effectStyle><a:effectLst><a:outerShdw blurRad="57150"/></a:effectLst></a:effectStyle>
</a:effectStyleLst></a:fmtScheme></a:themeElements></a:theme>`;

/**
 * The exported file. `tagValues` are the RETRO_CAPTURE values, in slide order: Badge, Group 1,
 * Inner, Connector, Icon, Scribble, Chart, Shadowed. They default to 1–8.
 */
export async function pptxBytes(tagValues = ["1", "2", "3", "4", "5", "6", "7", "8"]): Promise<Uint8Array> {
  const files: Record<string, string | Uint8Array> = {
    "ppt/slides/slide1.xml": SLIDE,
    "ppt/slides/_rels/slide1.xml.rels": rels(
      rel("rId1", "../slideLayouts/slideLayout1.xml"),
      ...[10, 11, 12, 13, 14, 15, 16, 17].map((n, i) => rel(`rId${n}`, `../tags/tag${i + 1}.xml`)),
      rel("rId20", "../media/image1.png"),
      rel("rId21", "../media/image2.svg"),
    ),
    "ppt/slideLayouts/_rels/slideLayout1.xml.rels": rels(rel("rId1", "../slideMasters/slideMaster1.xml")),
    "ppt/slideMasters/_rels/slideMaster1.xml.rels": rels(rel("rId1", "../slideLayouts/slideLayout1.xml"), rel("rId2", "../theme/theme1.xml")),
    "ppt/theme/theme1.xml": THEME,
    "ppt/media/image1.png": new Uint8Array([137, 80, 78, 71]),
  };
  tagValues.forEach((v, i) => (files[`ppt/tags/tag${i + 1}.xml`] = tag(v)));
  return makeZip(files);
}

