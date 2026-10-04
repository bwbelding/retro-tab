// Reads the .pptx PowerPoint exports for a slide and describes the shapes Retro tagged before
// exporting: what each one is (the API can't say), where it sits, how its text is split into runs,
// and anything about it that the add-in API can't recreate. Colors and fonts are read from the
// live shapes instead (see shapeCapture.ts), because PowerPoint resolves theme colors for us there.

import { apiGeometry, BENT_LINES, STRAIGHT_LINES } from "./geometry";
import type { Zip } from "./zip";

export const CAPTURE_TAG = "RETRO_CAPTURE";
const EMU_PER_POINT = 12700;

export interface Frame {
  left: number;
  top: number;
  width: number;
  height: number;
  /** Degrees clockwise. */
  rotation: number;
  flipH: boolean;
  flipV: boolean;
}

export interface Span {
  start: number;
  length: number;
}

export interface TextLayout {
  /** Total length, counting one character between paragraphs and for each line break. */
  length: number;
  paragraphs: Span[];
  runs: Span[];
  /** Spacing, indents and bullets of each paragraph as rendered, for comparing copies. */
  styles: string[];
}

export type NodeKind = "shape" | "line" | "picture" | "group" | "other";

export interface XmlNode {
  /** Capture id from the RETRO_CAPTURE tag, if this node was tagged. */
  cid?: string;
  kind: NodeKind;
  name: string;
  frame: Frame;
  /** API geometric shape type, for kind "shape". */
  geometry?: string;
  textBox?: boolean;
  text?: TextLayout;
  /** Zip part holding the picture, for pictures and picture-filled shapes. */
  media?: string;
  /** Things the add-in API can't recreate. Empty means Retro can rebuild it exactly. */
  issues: string[];
  children?: XmlNode[];
}

// --- Small DOM helpers that ignore namespace prefixes ---

const kids = (el: Element | undefined, name: string): Element[] => (el ? Array.from(el.children).filter((c) => c.localName === name) : []);
const kid = (el: Element | undefined, name: string): Element | undefined => kids(el, name)[0];
const path = (el: Element | undefined, ...names: string[]): Element | undefined => names.reduce<Element | undefined>((e, n) => kid(e, n), el);
const num = (el: Element | undefined, attr: string, fallback = 0): number => {
  const v = el?.getAttribute(attr);
  return v == null || v === "" ? fallback : Number(v);
};
const flag = (el: Element | undefined, attr: string) => {
  const v = el?.getAttribute(attr);
  return v === "1" || v === "true";
};
/** r:id style attributes, whatever the prefix. */
const relAttr = (el: Element | undefined, local: string): string | undefined => {
  if (!el) return undefined;
  for (const a of Array.from(el.attributes)) if (a.name.endsWith(`:${local}`)) return a.value;
  return undefined;
};

export function parseXml(text: string): Document {
  return new DOMParser().parseFromString(text, "application/xml");
}

// --- Relationships ---

export type Rels = Map<string, string>;

function dirOf(part: string) {
  return part.slice(0, part.lastIndexOf("/") + 1);
}

/** Resolve a relationship target relative to the part that owns it ("../media/a.png" etc.). */
export function resolvePart(from: string, target: string): string {
  if (target.startsWith("/")) return target.slice(1);
  const parts = (dirOf(from) + target).split("/");
  const out: string[] = [];
  for (const p of parts) {
    if (p === "..") out.pop();
    else if (p !== "." && p !== "") out.push(p);
  }
  return out.join("/");
}

export async function readRels(zip: Zip, part: string): Promise<Rels> {
  const relsPart = `${dirOf(part)}_rels/${part.slice(part.lastIndexOf("/") + 1)}.rels`;
  const text = await zip.text(relsPart);
  const rels: Rels = new Map();
  if (!text) return rels;
  for (const r of Array.from(parseXml(text).getElementsByTagName("Relationship"))) {
    if (r.getAttribute("TargetMode") === "External") continue;
    rels.set(r.getAttribute("Id") ?? "", resolvePart(part, r.getAttribute("Target") ?? ""));
  }
  return rels;
}

// --- Text ---

/**
 * Paragraph settings the add-in API can't set: spacing, indents and bullets, as normalized strings
 * (spacing "pct:<1000ths of a percent>" or "pts:<100ths of a point>", indents in EMU).
 */
export type ParaStyle = Record<string, string>;

/** What PowerPoint uses when nothing sets a value. */
const BUILT_IN: ParaStyle = { marL: "0", indent: "0", lnSpc: "pct:100000", spcBef: "0", spcAft: "0", bullet: "none" };

/** Paragraph defaults: what a shape inherits, and what a newly created shape or text box gets. */
export interface TextDefaults {
  inherited: ParaStyle[];
  newShape: ParaStyle[];
  newTextBox: ParaStyle[];
}

const LEVELS = 9;
export const PLAIN_DEFAULTS: TextDefaults = {
  inherited: Array.from({ length: LEVELS }, () => BUILT_IN),
  newShape: Array.from({ length: LEVELS }, () => BUILT_IN),
  newTextBox: Array.from({ length: LEVELS }, () => BUILT_IN),
};

function spacing(el: Element | undefined): string | undefined {
  const pct = kid(el, "spcPct");
  const pts = kid(el, "spcPts");
  const value = pct ? num(pct, "val") : pts ? num(pts, "val") : undefined;
  if (value === undefined) return undefined;
  if (value === 0 && el?.localName !== "lnSpc") return "0";
  return `${pct ? "pct" : "pts"}:${value}`;
}

/** The settings a paragraph-properties element (pPr or lvlNpPr) sets explicitly. */
export function paraStyle(pPr: Element | undefined): ParaStyle {
  const out: ParaStyle = {};
  if (!pPr) return out;
  for (const a of ["marL", "indent"]) if (pPr.hasAttribute(a)) out[a] = String(num(pPr, a));
  for (const n of ["lnSpc", "spcBef", "spcAft"]) {
    const v = spacing(kid(pPr, n));
    if (v !== undefined) out[n] = v;
  }
  if (kid(pPr, "buNone")) out.bullet = "none";
  else if (kid(pPr, "buChar")) out.bullet = `char:${kid(pPr, "buChar")!.getAttribute("char") ?? ""}`;
  else if (kid(pPr, "buAutoNum")) out.bullet = `number:${kid(pPr, "buAutoNum")!.getAttribute("type") ?? ""}`;
  else if (kid(pPr, "buBlip")) out.bullet = "picture";
  return out;
}

/** Level n (0-based) of a list style (lstStyle, defaultTextStyle, ...). */
const levelOf = (list: Element | undefined, n: number) => paraStyle(kid(list, `lvl${n + 1}pPr`));

export function textDefaults(defaultTextStyle: Element | undefined, theme: Document | undefined): TextDefaults {
  const objectDefaults = theme ? Array.from(theme.getElementsByTagName("*")).find((e) => e.localName === "objectDefaults") : undefined;
  const spDef = path(objectDefaults, "spDef", "lstStyle");
  const txDef = path(objectDefaults, "txDef", "lstStyle");
  const inherited = Array.from({ length: LEVELS }, (_, n) => ({ ...BUILT_IN, ...paraStyle(kid(defaultTextStyle, "defPPr")), ...levelOf(defaultTextStyle, n) }));
  return {
    inherited,
    newShape: inherited.map((s, n) => ({ ...s, ...levelOf(spDef, n) })),
    newTextBox: inherited.map((s, n) => ({ ...s, ...levelOf(txDef, n) })),
  };
}

const KEY_NAME: Record<string, string> = { marL: "left indent", indent: "first-line indent", lnSpc: "line spacing", spcBef: "space before", spcAft: "space after", bullet: "bullet" };
const KEY_ISSUE: Record<string, string> = { marL: "paragraph indents", indent: "paragraph indents", lnSpc: "line or paragraph spacing", spcBef: "line or paragraph spacing", spcAft: "line or paragraph spacing", bullet: "bullets" };

function showValue(key: string, v: string): string {
  if (key === "marL" || key === "indent") return `${Math.round((Number(v) / EMU_PER_POINT) * 10) / 10} pt`;
  if (v === "0") return "0";
  const [unit, n] = v.split(":");
  if (unit === "pct") return `${Number(n) / 1000}%`;
  if (unit === "pts") return `${Number(n) / 100} pt`;
  return v.replace("char:", "“").replace(/^“(.*)$/, "“$1”");
}

/** Each paragraph's rendered spacing, indents and bullets, with its level. */
function paragraphStyles(txBody: Element | undefined, inherited: ParaStyle[]): { level: number; style: ParaStyle }[] {
  const lstStyle = kid(txBody, "lstStyle");
  return kids(txBody, "p").map((p) => {
    const pPr = kid(p, "pPr");
    const level = Math.min(LEVELS - 1, Math.max(0, num(pPr, "lvl")));
    return { level, style: { ...inherited[level], ...levelOf(lstStyle, level), ...paraStyle(pPr) } };
  });
}

export function textLayout(txBody: Element | undefined, defaults: TextDefaults = PLAIN_DEFAULTS): TextLayout | undefined {
  const paras = kids(txBody, "p");
  if (paras.length === 0) return undefined;
  const paragraphs: Span[] = [];
  const runs: Span[] = [];
  let pos = 0;
  paras.forEach((p, i) => {
    if (i > 0) pos += 1; // paragraph separator
    const start = pos;
    for (const c of Array.from(p.children)) {
      if (c.localName === "r" || c.localName === "fld") {
        const length = (kid(c, "t")?.textContent ?? "").length;
        if (length) runs.push({ start: pos, length });
        pos += length;
      } else if (c.localName === "br") {
        pos += 1;
      }
    }
    paragraphs.push({ start, length: pos - start });
  });
  if (pos === 0) return undefined;
  const styles = paragraphStyles(txBody, defaults.inherited).map(({ style }) =>
    Object.keys(BUILT_IN)
      .map((k) => `${k}=${style[k]}`)
      .join(" "),
  );
  return { length: pos, paragraphs, runs, styles };
}

/** What a paragraph setting is, and what a newly created shape would get instead. */
export function textIssues(txBody: Element | undefined, defaults: TextDefaults = PLAIN_DEFAULTS, textBox = false): string[] {
  const issues = new Set<string>();
  const body = kid(txBody, "bodyPr");
  const vert = body?.getAttribute("vert");
  if ((vert && vert !== "horz") || num(body, "rot") !== 0) issues.add("vertical or rotated text");
  if (num(body, "numCol", 1) > 1) issues.add("text in columns");
  const fresh = textBox ? defaults.newTextBox : defaults.newShape;
  for (const { level, style } of paragraphStyles(txBody, defaults.inherited)) {
    for (const key of Object.keys(BUILT_IN)) {
      if (style[key] === fresh[level][key]) continue;
      issues.add(`${KEY_ISSUE[key]} (${KEY_NAME[key]} ${showValue(key, style[key])} where a new ${textBox ? "text box" : "shape"} gets ${showValue(key, fresh[level][key])})`);
    }
  }
  for (const p of kids(txBody, "p")) {
    for (const r of [...kids(p, "r"), ...kids(p, "fld")]) {
      const rPr = kid(r, "rPr");
      if (!rPr) continue;
      if (["effectLst", "highlight", "gradFill", "pattFill", "blipFill"].some((n) => kid(rPr, n))) issues.add("text effects or highlight");
      if (num(rPr, "baseline") !== 0 || num(rPr, "spc") !== 0) issues.add("character spacing or superscript");
      if (kid(rPr, "ln")) issues.add("outlined text");
    }
  }
  return [...issues];
}

// --- Shapes ---

interface Ctx {
  rels: Rels;
  zip: Zip;
  /** Effects defined by the theme's effect styles, indexed from 1 like effectRef. */
  themeEffects: boolean[];
  text: TextDefaults;
  tags: Map<string, string | undefined>;
}

interface Transform {
  /** Maps child coordinates (EMU) into slide coordinates (EMU). */
  x: (v: number) => number;
  y: (v: number) => number;
  sx: number;
  sy: number;
}

const IDENTITY: Transform = { x: (v) => v, y: (v) => v, sx: 1, sy: 1 };

function frameOf(xfrm: Element | undefined, t: Transform): Frame {
  const off = kid(xfrm, "off");
  const ext = kid(xfrm, "ext");
  const pt = (emu: number) => Math.round((emu / EMU_PER_POINT) * 100) / 100;
  return {
    left: pt(t.x(num(off, "x"))),
    top: pt(t.y(num(off, "y"))),
    width: pt(num(ext, "cx") * t.sx),
    height: pt(num(ext, "cy") * t.sy),
    rotation: num(xfrm, "rot") / 60000,
    flipH: flag(xfrm, "flipH"),
    flipV: flag(xfrm, "flipV"),
  };
}

function childTransform(xfrm: Element | undefined, t: Transform): Transform {
  const off = kid(xfrm, "off");
  const ext = kid(xfrm, "ext");
  const chOff = kid(xfrm, "chOff");
  const chExt = kid(xfrm, "chExt");
  const sx = num(chExt, "cx") ? num(ext, "cx") / num(chExt, "cx") : 1;
  const sy = num(chExt, "cy") ? num(ext, "cy") / num(chExt, "cy") : 1;
  return {
    x: (v) => t.x(num(off, "x") + (v - num(chOff, "x")) * sx),
    y: (v) => t.y(num(off, "y") + (v - num(chOff, "y")) * sy),
    sx: t.sx * sx,
    sy: t.sy * sy,
  };
}

function hasEffects(spPr: Element | undefined, style: Element | undefined, ctx: Ctx): boolean {
  if (kid(spPr, "scene3d") || kid(spPr, "sp3d")) return true;
  const own = kid(spPr, "effectLst") ?? kid(spPr, "effectDag");
  if (own) return own.children.length > 0;
  const idx = num(kid(style, "effectRef"), "idx");
  return idx > 0 && Boolean(ctx.themeEffects[idx]);
}

function arrowheads(ln: Element | undefined): boolean {
  return ["headEnd", "tailEnd"].some((n) => {
    const type = kid(ln, n)?.getAttribute("type");
    return Boolean(type && type !== "none");
  });
}

function fillIssues(spPr: Element | undefined): string[] {
  if (kid(spPr, "gradFill")) return ["gradient fill"];
  if (kid(spPr, "pattFill")) return ["pattern fill"];
  const blip = kid(spPr, "blipFill");
  if (blip && (kid(blip, "tile") || kid(blip, "srcRect")?.attributes.length)) return ["tiled or cropped picture fill"];
  return [];
}

function lineIssues(ln: Element | undefined): string[] {
  const issues = [];
  if (arrowheads(ln)) issues.push("arrowheads");
  if (kid(ln, "gradFill") || kid(ln, "pattFill")) issues.push("gradient or pattern line");
  if (kid(ln, "custDash")) issues.push("custom dashes");
  return issues;
}

/** The picture a blip points at; formats PowerPoint can't take back from an add-in are an issue. */
function mediaPart(blip: Element | undefined, ctx: Ctx, issues: string[]): string | undefined {
  const embed = relAttr(blip, "embed");
  const part = embed ? ctx.rels.get(embed) : undefined;
  if (blip && !part) issues.push("linked picture");
  else if (part && !/\.(png|jpe?g|gif|bmp)$/i.test(part)) issues.push("picture format");
  return part;
}

function captureId(nvPr: Element | undefined, ctx: Ctx): string | undefined {
  const rid = relAttr(kid(kid(nvPr, "custDataLst"), "tags"), "id");
  return rid ? ctx.tags.get(rid) : undefined;
}

function nonVisual(el: Element): { cNvPr?: Element; nvPr?: Element; cNv?: Element } {
  const nv = Array.from(el.children).find((c) => c.localName.startsWith("nv"));
  return {
    cNvPr: kid(nv, "cNvPr"),
    nvPr: kid(nv, "nvPr"),
    cNv: nv ? Array.from(nv.children).find((c) => c.localName.startsWith("cNv") && c.localName !== "cNvPr") : undefined,
  };
}

function readNode(el: Element, ctx: Ctx, t: Transform): XmlNode | undefined {
  const { cNvPr, nvPr, cNv } = nonVisual(el);
  const base = { cid: captureId(nvPr, ctx), name: cNvPr?.getAttribute("name") ?? "" };
  const issues: string[] = [];
  if (kid(nvPr, "ph")) issues.push("slide placeholder");

  switch (el.localName) {
    case "grpSp": {
      const grpSpPr = kid(el, "grpSpPr");
      const xfrm = kid(grpSpPr, "xfrm");
      const inner = childTransform(xfrm, t);
      const children = Array.from(el.children)
        .map((c) => readNode(c, ctx, inner))
        .filter((n): n is XmlNode => Boolean(n));
      const frame = frameOf(xfrm, t);
      if (frame.flipH || frame.flipV) issues.push("mirrored group");
      if (hasEffects(grpSpPr, undefined, ctx)) issues.push("shadow or effect");
      return { ...base, kind: "group", frame, issues, children };
    }
    case "sp":
    case "cxnSp": {
      const spPr = kid(el, "spPr");
      const style = kid(el, "style");
      const frame = frameOf(kid(spPr, "xfrm"), t);
      const prst = kid(spPr, "prstGeom")?.getAttribute("prst") ?? "";
      const txBody = kid(el, "txBody");
      if (hasEffects(spPr, style, ctx)) issues.push("shadow or effect");
      issues.push(...lineIssues(kid(spPr, "ln")));
      if (kid(spPr, "custGeom")) return { ...base, kind: "other", frame, issues: [...issues, "freeform or edited points"] };
      if (STRAIGHT_LINES.has(prst)) return { ...base, kind: "line", frame, issues };
      if (BENT_LINES.test(prst)) return { ...base, kind: "line", frame, issues: [...issues, "elbow or curved connector"] };
      const geometry = apiGeometry(prst);
      if (!geometry) issues.push(prst ? `shape type “${prst}”` : "unknown shape type");
      if (frame.flipH !== frame.flipV) issues.push("mirrored");
      const textBox = flag(cNv, "txBox");
      issues.push(...fillIssues(spPr), ...textIssues(txBody, ctx.text, textBox));
      const blip = kid(kid(spPr, "blipFill"), "blip");
      const media = mediaPart(blip, ctx, issues);
      return { ...base, kind: "shape", frame, geometry, textBox, text: textLayout(txBody, ctx.text), media, issues };
    }
    case "pic": {
      const spPr = kid(el, "spPr");
      const blipFill = kid(el, "blipFill");
      const blip = kid(blipFill, "blip");
      const frame = frameOf(kid(spPr, "xfrm"), t);
      if (hasEffects(spPr, kid(el, "style"), ctx)) issues.push("shadow or effect");
      if (kid(spPr, "ln")) issues.push("picture border");
      if (kid(blipFill, "srcRect")?.attributes.length) issues.push("cropped picture");
      if (blip && Array.from(blip.getElementsByTagName("*")).some((e) => e.localName === "svgBlip")) issues.push("SVG icon");
      if (blip && Array.from(blip.children).some((c) => c.localName !== "extLst")) issues.push("picture color adjustments");
      const prst = kid(spPr, "prstGeom")?.getAttribute("prst");
      if (prst && prst !== "rect") issues.push("picture cropped to a shape");
      if (frame.flipH || frame.flipV) issues.push("mirrored");
      return { ...base, kind: "picture", frame, media: mediaPart(blip, ctx, issues), issues };
    }
    case "graphicFrame": {
      const uri = path(el, "graphic", "graphicData")?.getAttribute("uri") ?? "";
      const what = uri.endsWith("/table") ? "table" : uri.includes("chart") ? "chart" : uri.includes("diagram") ? "SmartArt" : "embedded object";
      return { ...base, kind: "other", frame: frameOf(kid(el, "xfrm"), t), issues: [...issues, what] };
    }
    case "contentPart":
    case "AlternateContent":
      return { ...base, kind: "other", frame: frameOf(undefined, t), issues: [...issues, "ink, 3-D model or other special object"] };
    default:
      return undefined; // nvGrpSpPr, grpSpPr and other non-shape children
  }
}

/** The first slide part in an exported presentation. */
export function slidePart(zip: Zip): string | undefined {
  return zip.names
    .filter((n) => /^ppt\/slides\/slide\d+\.xml$/.test(n))
    .sort((a, b) => Number(a.match(/\d+/)![0]) - Number(b.match(/\d+/)![0]))[0];
}

async function readTheme(zip: Zip, slide: string, slideRels: Rels): Promise<Document | undefined> {
  // slide → layout → master → theme
  const follow = async (part: string | undefined, rels: Rels | undefined, type: string) => {
    if (!part || !rels) return [undefined, undefined] as const;
    const target = [...rels.values()].find((t) => t.includes(`/${type}`));
    return [target, target ? await readRels(zip, target) : undefined] as const;
  };
  const [layout, layoutRels] = await follow(slide, slideRels, "slideLayouts/");
  const [master, masterRels] = await follow(layout, layoutRels, "slideMasters/");
  const [theme] = await follow(master, masterRels, "theme/");
  const text = theme ? await zip.text(theme) : undefined;
  return text ? parseXml(text) : undefined;
}

/** The theme used by the first slide of an exported presentation. */
export async function deckTheme(zip: Zip): Promise<Document | undefined> {
  const part = slidePart(zip);
  return part ? readTheme(zip, part, await readRels(zip, part)) : undefined;
}

function themeEffects(theme: Document | undefined): boolean[] {
  if (!theme) return [];
  const styles = Array.from(theme.getElementsByTagName("*")).find((e) => e.localName === "effectStyleLst");
  return [false, ...kids(styles, "effectStyle").map((s) => kids(s, "effectLst").some((l) => l.children.length > 0) || Boolean(kid(s, "scene3d") || kid(s, "sp3d")))];
}

async function defaultTextStyle(zip: Zip): Promise<Element | undefined> {
  const text = await zip.text("ppt/presentation.xml");
  return text ? Array.from(parseXml(text).getElementsByTagName("*")).find((e) => e.localName === "defaultTextStyle") : undefined;
}

export interface ParsedSlide {
  part: string;
  /** Every top-level shape on the slide, in z-order (back to front). */
  nodes: XmlNode[];
  rels: Rels;
}

export async function parseSlide(zip: Zip): Promise<ParsedSlide> {
  const part = slidePart(zip);
  if (!part) throw new Error("The exported file has no slide in it.");
  const rels = await readRels(zip, part);
  const tags = new Map<string, string | undefined>();
  for (const [id, target] of rels) {
    if (!target.includes("/tags/")) continue;
    const text = await zip.text(target);
    if (!text) continue;
    const tag = Array.from(parseXml(text).getElementsByTagName("*")).find((e) => e.localName === "tag" && e.getAttribute("name")?.toUpperCase() === CAPTURE_TAG);
    tags.set(id, tag?.getAttribute("val") ?? undefined);
  }
  const theme = await readTheme(zip, part, rels);
  const ctx: Ctx = { rels, zip, tags, themeEffects: themeEffects(theme), text: textDefaults(await defaultTextStyle(zip), theme) };
  const doc = parseXml((await zip.text(part)) ?? "");
  const spTree = Array.from(doc.getElementsByTagName("*")).find((e) => e.localName === "spTree");
  const nodes = Array.from(spTree?.children ?? [])
    .map((c) => readNode(c, ctx, IDENTITY))
    .filter((n): n is XmlNode => Boolean(n));
  return { part, nodes, rels };
}

/** The tagged nodes, outermost first: a tagged group's tagged children stay inside it. */
export function taggedNodes(nodes: XmlNode[]): XmlNode[] {
  const out: XmlNode[] = [];
  const walk = (list: XmlNode[]) => {
    for (const n of list) {
      if (n.cid !== undefined) out.push(n);
      else if (n.children) walk(n.children);
    }
  };
  walk(nodes);
  return out;
}

/** Every issue in a node and its children, without repeats. */
export function allIssues(node: XmlNode): string[] {
  const set = new Set(node.issues);
  for (const c of node.children ?? []) for (const i of allIssues(c)) set.add(i);
  return [...set];
}
