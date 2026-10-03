// Captures selected shapes as a "recipe" Retro can rebuild later, and rebuilds them.
//
// Capturing combines two sources, because neither has everything:
//  - PowerPoint's add-in API: colours (with theme colours already resolved), line, fonts,
//    text settings and shape adjustments, all in the units the API takes back when rebuilding.
//  - The slide exported as a .pptx: which shape type each one is, flips, how text is split
//    into runs, embedded pictures, and effects the API can't recreate (see ooxml.ts).
// The two are matched with a temporary RETRO_CAPTURE tag on every captured shape.

import { allIssues, CAPTURE_TAG, parseSlide, taggedNodes, type Frame, type Span, type XmlNode } from "./ooxml";
import { base64ToBytes, bytesToBase64, openZip } from "./zip";

export interface FontLook {
  name?: string;
  size?: number;
  color?: string;
  bold?: boolean;
  italic?: boolean;
  underline?: string;
  strikethrough?: boolean;
  allCaps?: boolean;
  smallCaps?: boolean;
}

export interface TextLook {
  text: string;
  margins: [number, number, number, number];
  anchor: string;
  wordWrap: boolean;
  autoSize: string;
  paragraphs: (Span & { align?: string })[];
  runs: (Span & { font: FontLook })[];
}

export interface Look {
  fill?: { type: string; color?: string; transparency?: number };
  line?: { visible: boolean; color?: string; weight?: number; dash?: string; style?: string; transparency?: number };
  adjustments?: number[];
  text?: TextLook;
}

export interface RecipeNode extends XmlNode {
  look: Look;
  /** Base64 picture data for pictures and picture-filled shapes. */
  image?: string;
  children?: RecipeNode[];
}

export interface Recipe {
  version: 1;
  /** Bounding box of the captured shapes when they were saved, in points. */
  box: { left: number; top: number; width: number; height: number };
  nodes: RecipeNode[];
}

export interface Capture {
  recipe: Recipe;
  /** The exported slide (.pptx, base64), still carrying the capture tags, for the helper-slide route. */
  pptx: string;
  /** How many selected shapes were found in the exported file. */
  found: number;
  selected: number;
  exportMs: number;
}

// --- Walking the live shapes ---

interface Live {
  cid: string;
  shape: PowerPoint.Shape;
  children: Live[];
}

async function loadTree(context: PowerPoint.RequestContext, shapes: PowerPoint.Shape[], next: { n: number }): Promise<Live[]> {
  shapes.forEach((s) => s.load("id,type"));
  await context.sync();
  const out: Live[] = [];
  for (const shape of shapes) {
    let children: Live[] = [];
    if (shape.type === "Group") {
      const items = shape.group.shapes;
      items.load("items/id");
      await context.sync();
      children = await loadTree(context, items.items, next);
    }
    out.push({ cid: String(next.n++), shape, children });
  }
  return out;
}

const flat = (live: Live[]): Live[] => live.flatMap((l) => [l, ...flat(l.children)]);

const clean = <T extends object>(o: T): T => Object.fromEntries(Object.entries(o).filter(([, v]) => v !== null && v !== undefined)) as T;

const FONT_PROPS = "name,size,color,bold,italic,underline,strikethrough,allCaps,smallCaps";

/** Queues reads of everything the API can tell us about one shape, batched with the others. */
function queueLook(shape: PowerPoint.Shape, node: XmlNode) {
  const isShape = node.kind === "shape";
  return {
    fill: isShape ? shape.fill.load("type,foregroundColor,transparency") : undefined,
    line: isShape || node.kind === "line" ? shape.lineFormat.load("visible,color,weight,dashStyle,style,transparency") : undefined,
    adj: isShape && !node.textBox ? shape.adjustments.load("count") : undefined,
    frame: node.text ? shape.textFrame.load("leftMargin,rightMargin,topMargin,bottomMargin,verticalAlignment,wordWrap,autoSizeSetting") : undefined,
    range: node.text ? shape.textFrame.textRange.load("text") : undefined,
  };
}

/** Runs and paragraphs to read fonts and alignment from: the XML's, if its text matches the API's. */
function spans(layout: XmlNode["text"], text: string) {
  if (layout && layout.length === text.length) return { runs: layout.runs, paragraphs: layout.paragraphs.filter((p) => p.length > 0) };
  const whole = text.length ? [{ start: 0, length: text.length }] : [];
  return { runs: whole, paragraphs: whole };
}

async function readLooks(context: PowerPoint.RequestContext, pairs: { live: Live; node: XmlNode }[]): Promise<Look[]> {
  const queued = pairs.map(({ live, node }) => queueLook(live.shape, node));
  await context.sync();

  // Second pass, now that adjustment counts and text lengths are known.
  const adjValues = queued.map((q) => (q.adj ? Array.from({ length: q.adj.count }, (_, k) => q.adj!.get(k)) : undefined));
  const textReads = queued.map((q, i) => {
    if (!q.range) return undefined;
    const { runs, paragraphs } = spans(pairs[i].node.text, q.range.text);
    return {
      runs,
      paragraphs,
      fonts: runs.map((r) => q.range!.getSubstring(r.start, r.length).font.load(FONT_PROPS)),
      formats: paragraphs.map((p) => q.range!.getSubstring(p.start, p.length).paragraphFormat.load("horizontalAlignment")),
    };
  });
  await context.sync();

  return queued.map((q, i) => {
    const look: Look = {};
    if (q.fill) look.fill = clean({ type: q.fill.type, color: q.fill.type === "Solid" ? q.fill.foregroundColor : undefined, transparency: q.fill.transparency });
    if (q.line)
      look.line = q.line.visible
        ? clean({ visible: true, color: q.line.color, weight: q.line.weight, dash: q.line.dashStyle, style: q.line.style, transparency: q.line.transparency })
        : { visible: false };
    if (adjValues[i]) look.adjustments = adjValues[i]!.map((v) => v.value);
    const t = textReads[i];
    if (t && q.frame && q.range) {
      look.text = {
        text: q.range.text,
        margins: [q.frame.leftMargin, q.frame.rightMargin, q.frame.topMargin, q.frame.bottomMargin],
        anchor: q.frame.verticalAlignment,
        wordWrap: q.frame.wordWrap,
        autoSize: q.frame.autoSizeSetting,
        paragraphs: t.paragraphs.map((p, k) => clean({ ...p, align: t.formats[k].horizontalAlignment ?? undefined })),
        runs: t.runs.map((r, k) => {
          const f = t.fonts[k];
          const font = { name: f.name, size: f.size, color: f.color, bold: f.bold, italic: f.italic, underline: f.underline, strikethrough: f.strikethrough, allCaps: f.allCaps, smallCaps: f.smallCaps };
          return { ...r, font: clean(font) as FontLook };
        }),
      };
    }
    return look;
  });
}

// --- Capture ---

/** Tag, export, untag. Returns the exported slide. */
async function exportTagged(context: PowerPoint.RequestContext, slide: PowerPoint.Slide, live: Live[]): Promise<string> {
  const all = flat(live);
  all.forEach((l) => l.shape.tags.add(CAPTURE_TAG, l.cid));
  await context.sync();
  try {
    const b64 = slide.exportAsBase64();
    await context.sync();
    return b64.value;
  } finally {
    all.forEach((l) => l.shape.tags.delete(CAPTURE_TAG));
    await context.sync();
  }
}

export function boundsOf(nodes: { frame: Frame }[]) {
  const left = Math.min(...nodes.map((n) => n.frame.left));
  const top = Math.min(...nodes.map((n) => n.frame.top));
  const right = Math.max(...nodes.map((n) => n.frame.left + n.frame.width));
  const bottom = Math.max(...nodes.map((n) => n.frame.top + n.frame.height));
  return { left, top, width: right - left, height: bottom - top };
}

/** Capture shapes on one slide. */
export async function captureShapes(context: PowerPoint.RequestContext, slide: PowerPoint.Slide, shapes: PowerPoint.Shape[]): Promise<Capture> {
  const live = await loadTree(context, shapes, { n: 1 });
  const started = Date.now();
  const pptx = await exportTagged(context, slide, live);
  const exportMs = Date.now() - started;

  const zip = openZip(base64ToBytes(pptx));
  const parsed = await parseSlide(zip);
  const byCid = new Map(flat(live).map((l) => [l.cid, l]));
  const tops = taggedNodes(parsed.nodes);

  // Pair every tagged XML node (at any depth inside the captured shapes) with its live shape.
  const pairs: { live: Live; node: XmlNode }[] = [];
  const pairUp = (nodes: XmlNode[]) =>
    nodes.forEach((n) => {
      const l = n.cid !== undefined ? byCid.get(n.cid) : undefined;
      if (l) pairs.push({ live: l, node: n });
      pairUp(n.children ?? []);
    });
  pairUp(tops);
  const looks = new Map<XmlNode, Look>();
  (await readLooks(context, pairs)).forEach((look, i) => looks.set(pairs[i].node, look));

  const images = new Map<string, string>();
  const toRecipe = async (n: XmlNode): Promise<RecipeNode> => {
    let image: string | undefined;
    if (n.media) {
      if (!images.has(n.media)) {
        const bytes = await zip.bytes(n.media);
        if (bytes) images.set(n.media, bytesToBase64(bytes));
      }
      image = images.get(n.media);
    }
    const children = n.children ? await Promise.all(n.children.map(toRecipe)) : undefined;
    return clean({ ...n, look: looks.get(n) ?? {}, image, children });
  };
  const nodes = await Promise.all(tops.map(toRecipe));
  const topLevel = new Set(live.map((l) => l.cid));
  return {
    recipe: { version: 1, box: nodes.length ? boundsOf(nodes) : { left: 0, top: 0, width: 0, height: 0 }, nodes },
    pptx,
    found: tops.filter((n) => topLevel.has(n.cid!)).length,
    selected: live.length,
    exportMs,
  };
}

/** Why a node can't be rebuilt exactly with one click (empty when it can). */
export function rebuildIssues(node: RecipeNode): string[] {
  const issues = new Set(allIssues(node));
  const walk = (n: RecipeNode) => {
    const fill = n.look.fill?.type;
    if (fill === "Gradient") issues.add("gradient fill");
    if (fill === "Pattern") issues.add("pattern fill");
    if (fill === "PictureAndTexture" && !n.image) issues.add("texture fill");
    if (fill === "SlideBackground") issues.add("slide background fill");
    if (n.kind === "picture" && !n.image) issues.add("picture data missing");
    if (n.kind === "shape" && n.text && !n.look.text) issues.add("text couldn't be read");
    n.children?.forEach(walk);
  };
  walk(node);
  return [...issues];
}

// --- Rebuild ---

function applyFrame(s: PowerPoint.Shape, f: Frame, dx: number, dy: number) {
  s.left = f.left + dx;
  s.top = f.top + dy;
  s.width = Math.max(f.width, 0.01);
  s.height = Math.max(f.height, 0.01);
}

function applyLine(s: PowerPoint.Shape, line: Look["line"]) {
  if (!line) return;
  const lf = s.lineFormat;
  lf.visible = line.visible;
  if (!line.visible) return;
  if (line.color) lf.color = line.color;
  if (line.weight !== undefined) lf.weight = line.weight;
  if (line.dash) lf.dashStyle = line.dash as PowerPoint.ShapeLineDashStyle;
  if (line.style) lf.style = line.style as PowerPoint.ShapeLineStyle;
  if (line.transparency !== undefined) lf.transparency = line.transparency;
}

function applyFont(font: PowerPoint.ShapeFont, f: FontLook) {
  if (f.name) font.name = f.name;
  if (f.size) font.size = f.size;
  if (f.color) font.color = f.color;
  if (f.bold !== undefined) font.bold = f.bold;
  if (f.italic !== undefined) font.italic = f.italic;
  if (f.underline) font.underline = f.underline as PowerPoint.ShapeFontUnderlineStyle;
  if (f.strikethrough !== undefined) font.strikethrough = f.strikethrough;
  if (f.allCaps !== undefined) font.allCaps = f.allCaps;
  if (f.smallCaps !== undefined) font.smallCaps = f.smallCaps;
}

function applyText(s: PowerPoint.Shape, t: TextLook | undefined) {
  if (!t) return;
  const tf = s.textFrame;
  tf.autoSizeSetting = PowerPoint.ShapeAutoSize.autoSizeNone;
  [tf.leftMargin, tf.rightMargin, tf.topMargin, tf.bottomMargin] = t.margins;
  tf.verticalAlignment = t.anchor as PowerPoint.TextVerticalAlignment;
  tf.wordWrap = t.wordWrap;
  tf.textRange.text = t.text;
  for (const p of t.paragraphs) if (p.align) tf.textRange.getSubstring(p.start, p.length).paragraphFormat.horizontalAlignment = p.align as PowerPoint.ParagraphHorizontalAlignment;
  for (const r of t.runs) applyFont(tf.textRange.getSubstring(r.start, r.length).font, r.font);
  tf.autoSizeSetting = t.autoSize as PowerPoint.ShapeAutoSize;
}

/**
 * A straight line runs from the top-left to the bottom-right of its box. Mirroring it one way
 * gives the other diagonal, which is the same as a 90° turn of a box with width and height swapped.
 */
export function lineFrame(f: Frame): Frame {
  if (f.flipH === f.flipV) return { ...f, rotation: f.flipH ? f.rotation + 180 : f.rotation, flipH: false, flipV: false };
  const cx = f.left + f.width / 2;
  const cy = f.top + f.height / 2;
  return { left: cx - f.height / 2, top: cy - f.width / 2, width: f.height, height: f.width, rotation: f.rotation + 90, flipH: false, flipV: false };
}

/** Shapes mirrored both ways look the same turned 180°. */
export function shapeRotation(f: Frame): number {
  const r = f.flipH && f.flipV ? f.rotation + 180 : f.rotation;
  return ((r % 360) + 360) % 360;
}

async function build(context: PowerPoint.RequestContext, slide: PowerPoint.Slide, n: RecipeNode, dx: number, dy: number): Promise<PowerPoint.Shape | undefined> {
  const shapes = slide.shapes;
  let s: PowerPoint.Shape;
  switch (n.kind) {
    case "group": {
      const kids = [];
      for (const c of n.children ?? []) {
        const k = await build(context, slide, c, dx, dy);
        if (k) kids.push(k);
      }
      if (kids.length === 0) return undefined;
      if (kids.length === 1) return kids[0];
      kids.forEach((k) => k.load("id"));
      await context.sync();
      s = shapes.addGroup(kids.map((k) => k.id));
      if (n.frame.rotation) s.rotation = shapeRotation(n.frame);
      break;
    }
    case "line": {
      const f = lineFrame(n.frame);
      s = shapes.addLine(PowerPoint.ConnectorType.straight, { left: f.left + dx, top: f.top + dy, width: f.width, height: f.height });
      if (f.rotation % 360) s.rotation = ((f.rotation % 360) + 360) % 360;
      applyLine(s, n.look.line);
      break;
    }
    case "picture": {
      s = shapes.addGeometricShape(PowerPoint.GeometricShapeType.rectangle);
      applyFrame(s, n.frame, dx, dy);
      if (n.image) s.fill.setImage(n.image);
      s.lineFormat.visible = false;
      if (n.frame.rotation) s.rotation = n.frame.rotation;
      break;
    }
    case "shape": {
      if (n.textBox) {
        s = shapes.addTextBox(n.look.text?.text ?? "");
      } else {
        s = shapes.addGeometricShape((n.geometry ?? "Rectangle") as PowerPoint.GeometricShapeType);
      }
      applyFrame(s, n.frame, dx, dy);
      const rot = shapeRotation(n.frame);
      if (rot) s.rotation = rot;
      const fill = n.look.fill;
      if (fill?.type === "NoFill") s.fill.clear();
      else if (fill?.type === "PictureAndTexture" && n.image) s.fill.setImage(n.image);
      else if (fill?.color) s.fill.setSolidColor(fill.color);
      if (fill?.transparency !== undefined && fill.type === "Solid") s.fill.transparency = fill.transparency;
      applyLine(s, n.look.line);
      n.look.adjustments?.forEach((v, i) => s.adjustments.set(i, v));
      applyText(s, n.look.text);
      break;
    }
    default:
      return undefined;
  }
  s.name = n.name;
  return s;
}

/**
 * Rebuild captured shapes on a slide, offset by (dx, dy) from where they were saved, and select them.
 * Only call for recipes whose rebuildIssues are empty; anything else won't come back exactly.
 */
export async function rebuildShapes(context: PowerPoint.RequestContext, slide: PowerPoint.Slide, recipe: Recipe, dx: number, dy: number): Promise<PowerPoint.Shape[]> {
  const out: PowerPoint.Shape[] = [];
  for (const n of recipe.nodes) {
    const s = await build(context, slide, n, dx, dy);
    if (s) out.push(s);
  }
  out.forEach((s) => s.load("id"));
  await context.sync();
  slide.setSelectedShapes(out.map((s) => s.id));
  await context.sync();
  return out;
}

// --- Helper-slide route: an exact copy for shapes the API can't rebuild ---

export const HELPER_TAG = "RETRO_HELPER";

/**
 * Insert the exported slide right after `after`, keep only the captured shapes on it, and select
 * them, so they can be cut and pasted onto the slide being worked on. Returns the helper slide's id.
 */
export async function insertHelperSlide(context: PowerPoint.RequestContext, pptx: string, after: string): Promise<string> {
  const before = context.presentation.slides;
  before.load("items/id");
  await context.sync();
  const existing = new Set(before.items.map((s) => s.id));
  context.presentation.insertSlidesFromBase64(pptx, { formatting: PowerPoint.InsertSlideFormatting.useDestinationTheme, targetSlideId: after });
  await context.sync();
  const now = context.presentation.slides;
  now.load("items/id");
  await context.sync();
  const helper = now.items.find((s) => !existing.has(s.id));
  if (!helper) throw new Error("PowerPoint didn't add the helper slide.");
  helper.tags.add(HELPER_TAG, "1");
  const shapes = helper.shapes;
  shapes.load("items/id");
  await context.sync();
  const tags = shapes.items.map((s) => s.tags.getItemOrNullObject(CAPTURE_TAG).load("value"));
  await context.sync();
  const keep: string[] = [];
  shapes.items.forEach((s, i) => {
    if (tags[i].isNullObject) s.delete();
    else {
      keep.push(s.id);
      s.tags.delete(CAPTURE_TAG);
    }
  });
  await context.sync();
  context.presentation.setSelectedSlides([helper.id]);
  await context.sync();
  helper.setSelectedShapes(keep);
  await context.sync();
  return helper.id;
}
