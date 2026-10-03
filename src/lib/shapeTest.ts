// The Phase 3 capability test: on the user's own PowerPoint, capture the selected shapes, rebuild
// them beside the originals, capture the copies and compare. The differences tell us which parts
// of the shape library work on this Mac before the library is built around them.

import type { Frame } from "./ooxml";
import { currentSlide, slideSize, UserError } from "./ppt";
import { captureShapes, insertHelperSlide, lineFrame, rebuildIssues, rebuildShapes, shapeRotation, type Recipe, type RecipeNode } from "./shapeCapture";
import type { Status } from "./diagnostics";

// --- Describing and comparing recipes (pure, unit tested) ---

const KIND_NAME = { shape: "Shape", line: "Line", picture: "Picture", group: "Group", other: "Object" } as const;

export function describe(n: RecipeNode): string {
  if (n.kind === "group") return `Group of ${n.children?.length ?? 0}`;
  if (n.kind === "shape") {
    const what = n.textBox ? "Text box" : (n.geometry ?? "Shape").replace(/([a-z])(?=[A-Z0-9])|([0-9])(?=[A-Z])/g, "$1$2 ");
    return n.look.text?.text ? `${what} with text` : what;
  }
  return KIND_NAME[n.kind];
}

const close = (a: number | undefined, b: number | undefined, tol: number) => a === b || (a !== undefined && b !== undefined && Math.abs(a - b) <= tol);
const angle = (deg: number) => ((Math.round(deg * 10) / 10) % 360 + 360) % 360;
const sameAngle = (a: number, b: number) => Math.min(Math.abs(angle(a) - angle(b)), 360 - Math.abs(angle(a) - angle(b))) <= 0.5;

/** End points of a straight line, in either order. */
export function lineEnds(f: Frame): [number, number][] {
  const g = lineFrame(f);
  const cx = g.left + g.width / 2;
  const cy = g.top + g.height / 2;
  const r = (g.rotation * Math.PI) / 180;
  const at = (x: number, y: number): [number, number] => [cx + x * Math.cos(r) - y * Math.sin(r), cy + x * Math.sin(r) + y * Math.cos(r)];
  return [at(-g.width / 2, -g.height / 2), at(g.width / 2, g.height / 2)];
}

function compareFrames(where: string, a: Frame, b: Frame, dx: number, dy: number, line: boolean, out: string[]) {
  if (line) {
    const [a1, a2] = lineEnds(a);
    const [b1, b2] = lineEnds({ ...b, left: b.left - dx, top: b.top - dy });
    const near = (p: [number, number], q: [number, number]) => Math.hypot(p[0] - q[0], p[1] - q[1]) <= 1;
    if (!((near(a1, b1) && near(a2, b2)) || (near(a1, b2) && near(a2, b1)))) out.push(`${where}: line runs in a different direction or length`);
    return;
  }
  const moved = !close(a.left, b.left - dx, 0.5) || !close(a.top, b.top - dy, 0.5);
  if (moved) out.push(`${where}: position off by ${Math.round(b.left - dx - a.left)}, ${Math.round(b.top - dy - a.top)} pt`);
  if (!close(a.width, b.width, 0.5) || !close(a.height, b.height, 0.5))
    out.push(`${where}: size ${Math.round(a.width)}×${Math.round(a.height)} came back ${Math.round(b.width)}×${Math.round(b.height)} pt`);
  if (!sameAngle(shapeRotation(a), shapeRotation(b))) out.push(`${where}: rotation ${angle(shapeRotation(a))}° came back ${angle(shapeRotation(b))}°`);
  if ((a.flipH !== a.flipV) !== (b.flipH !== b.flipV)) out.push(`${where}: mirroring lost`);
}

function compareNode(where: string, a: RecipeNode, b: RecipeNode, dx: number, dy: number, out: string[]) {
  const pictureAsShape = a.kind === "picture" && b.kind === "shape" && b.look.fill?.type === "PictureAndTexture";
  if (a.kind !== b.kind && !pictureAsShape) {
    out.push(`${where}: ${describe(a)} came back as ${describe(b)}`);
    return;
  }
  if (a.geometry !== b.geometry && !pictureAsShape) out.push(`${where}: shape type ${a.geometry} came back ${b.geometry}`);
  compareFrames(where, a.frame, b.frame, dx, dy, a.kind === "line", out);

  const fa = a.look.fill;
  const fb = b.look.fill;
  if (fa && fb && !pictureAsShape) {
    if (fa.type !== fb.type) out.push(`${where}: fill ${fa.type} came back ${fb.type}`);
    else if (fa.color?.toUpperCase() !== fb.color?.toUpperCase()) out.push(`${where}: fill color ${fa.color} came back ${fb.color}`);
    if (!close(fa.transparency, fb.transparency, 0.01)) out.push(`${where}: fill transparency ${fa.transparency} came back ${fb.transparency}`);
  }
  const la = a.look.line;
  const lb = b.look.line;
  if (la && lb) {
    if (la.visible !== lb.visible) out.push(`${where}: outline ${la.visible ? "lost" : "added"}`);
    else if (la.visible) {
      if (la.color?.toUpperCase() !== lb.color?.toUpperCase()) out.push(`${where}: outline color ${la.color} came back ${lb.color}`);
      if (!close(la.weight, lb.weight, 0.05)) out.push(`${where}: outline weight ${la.weight} came back ${lb.weight}`);
      if (la.dash !== lb.dash) out.push(`${where}: outline dashes ${la.dash} came back ${lb.dash}`);
    }
  }
  const aa = a.look.adjustments ?? [];
  const ab = b.look.adjustments ?? [];
  aa.forEach((v, i) => {
    if (!close(v, ab[i], Math.max(1e-3, Math.abs(v) * 1e-3))) out.push(`${where}: adjustment ${i + 1} was ${v}, came back ${ab[i]}`);
  });

  const ta = a.look.text;
  const tb = b.look.text;
  if (ta || tb) {
    if (ta?.text !== tb?.text) out.push(`${where}: text changed`);
    else if (ta && tb) {
      if (ta.margins.some((m, i) => !close(m, tb.margins[i], 0.1))) out.push(`${where}: text margins changed`);
      if (ta.anchor !== tb.anchor) out.push(`${where}: vertical alignment ${ta.anchor} came back ${tb.anchor}`);
      if (ta.wordWrap !== tb.wordWrap) out.push(`${where}: word wrap changed`);
      if (ta.autoSize !== tb.autoSize) out.push(`${where}: autofit ${ta.autoSize} came back ${tb.autoSize}`);
      const align = (t: typeof ta) => t.paragraphs.map((p) => p.align ?? "").join(",");
      if (align(ta) !== align(tb)) out.push(`${where}: paragraph alignment changed`);
      // Compare fonts character by character, since the copy may split runs differently.
      const fontAt = (t: typeof ta, i: number) => JSON.stringify(t.runs.find((r) => i >= r.start && i < r.start + r.length)?.font ?? {});
      const firstDiff = Array.from({ length: ta.text.length }, (_, i) => i).find((i) => fontAt(ta, i) !== fontAt(tb, i));
      if (firstDiff !== undefined) out.push(`${where}: font differs from character ${firstDiff + 1} (${fontAt(ta, firstDiff)} vs ${fontAt(tb, firstDiff)})`);
    }
  }

  // Spacing, indents and bullets, read from the slide file (the API can't report them).
  const sa = a.text?.styles ?? [];
  const sb = b.text?.styles ?? [];
  sa.forEach((style, i) => {
    if (sb[i] === undefined || style === sb[i]) return;
    const before = new Map(style.split(" ").map((kv) => kv.split("=") as [string, string]));
    const changed = sb[i]
      .split(" ")
      .map((kv) => kv.split("="))
      .filter(([k, v]) => before.get(k) !== v)
      .map(([k, v]) => `${k} ${before.get(k)} → ${v}`);
    out.push(`${where}: paragraph ${i + 1} ${changed.join(", ")}`);
  });

  const ca = a.children ?? [];
  const cb = b.children ?? [];
  if (ca.length !== cb.length) out.push(`${where}: ${ca.length} shapes in the group came back as ${cb.length}`);
  else ca.forEach((c, i) => compareNode(`${where} › ${describe(c)}`, c, cb[i], dx, dy, out));
}

/** Differences between an original capture and a capture of its copy rebuilt (dx, dy) away. */
export function compareRecipes(original: Recipe, copy: Recipe, dx: number, dy: number): string[] {
  const out: string[] = [];
  if (original.nodes.length !== copy.nodes.length) out.push(`${original.nodes.length} shapes came back as ${copy.nodes.length}`);
  original.nodes.forEach((n, i) => {
    if (copy.nodes[i]) compareNode(`${i + 1}. ${describe(n)}`, n, copy.nodes[i], dx, dy, out);
  });
  return out;
}

/** Where to put the copies: right of the originals if they fit, else below, else a small offset. */
export function copyOffset(box: Recipe["box"], slide: { width: number; height: number }): { dx: number; dy: number } {
  const gap = 24;
  if (box.left + box.width * 2 + gap <= slide.width) return { dx: box.width + gap, dy: 0 };
  if (box.top + box.height * 2 + gap <= slide.height) return { dx: 0, dy: box.height + gap };
  return { dx: gap, dy: gap };
}

// --- Running the test in PowerPoint ---

export interface TestStep {
  label: string;
  status: Status;
  detail: string;
}

export interface TestItem {
  name: string;
  what: string;
  issues: string[];
}

export interface ShapeTestResult {
  steps: TestStep[];
  items: TestItem[];
  /** PNG preview of the first selected shape, as a data URL. */
  preview?: string;
  differences: string[];
}

const ms = (started: number) => `${Date.now() - started} ms`;

export async function runShapeTest(update: (r: ShapeTestResult) => void): Promise<ShapeTestResult> {
  const r: ShapeTestResult = { steps: [], items: [], differences: [] };
  const step = (s: TestStep) => {
    r.steps.push(s);
    update({ ...r, steps: [...r.steps] });
  };
  if (!Office.context.requirements.isSetSupported("PowerPointApi", "1.10")) {
    throw new UserError("The shape library needs PowerPoint 16.105 or later (PowerPointApi 1.10).");
  }

  await PowerPoint.run(async (context) => {
    const slide = await currentSlide(context);
    const selected = context.presentation.getSelectedShapes();
    selected.load("items/id");
    await context.sync();
    if (selected.items.length === 0) throw new UserError("Select one or more of your shapes on the slide first.");
    const size = await slideSize(context);

    let started = Date.now();
    const original = await captureShapes(context, slide, selected.items);
    step({ label: "Export the slide", status: "ok", detail: `${Math.round((original.pptx.length * 3) / 4 / 1024)} KB in ${original.exportMs} ms` });
    step({
      label: "Find the selected shapes in it",
      status: original.found === original.selected ? "ok" : "fail",
      detail: `${original.found} of ${original.selected} found · ${ms(started)} including reading formatting`,
    });

    r.items = original.recipe.nodes.map((n) => ({ name: n.name, what: describe(n), issues: rebuildIssues(n) }));
    update({ ...r });

    started = Date.now();
    try {
      const img = selected.items[0].getImageAsBase64({ format: PowerPoint.ShapeGetImageFormatType.png, width: 240 });
      await context.sync();
      r.preview = `data:image/png;base64,${img.value}`;
      step({ label: "Make a preview picture", status: "ok", detail: ms(started) });
    } catch (e) {
      step({ label: "Make a preview picture", status: "fail", detail: e instanceof Error ? e.message : String(e) });
    }

    const rebuildable: Recipe = { ...original.recipe, nodes: original.recipe.nodes.filter((n) => rebuildIssues(n).length === 0) };
    if (rebuildable.nodes.length === 0) {
      step({ label: "Rebuild with one click", status: "warn", detail: "Nothing selected can be rebuilt exactly; these would use the helper slide." });
      return;
    }
    const { dx, dy } = copyOffset(original.recipe.box, size);
    started = Date.now();
    const copies = await rebuildShapes(context, slide, rebuildable, dx, dy);
    step({ label: "Rebuild with one click", status: "ok", detail: `${copies.length} rebuilt beside the originals and selected · ${ms(started)}` });

    started = Date.now();
    const check = await captureShapes(context, slide, copies);
    r.differences = compareRecipes(rebuildable, check.recipe, dx, dy);
    step({
      label: "Compare the copies with the originals",
      status: r.differences.length ? "warn" : "ok",
      detail: r.differences.length ? `${r.differences.length} difference${r.differences.length > 1 ? "s" : ""} (listed below)` : `Identical · ${ms(started)}`,
    });
  });
  return r;
}

/** Put an exact copy of the selection on a helper slide after the current one. */
export async function runHelperSlideTest(): Promise<string> {
  return PowerPoint.run(async (context) => {
    const slide = await currentSlide(context);
    const selected = context.presentation.getSelectedShapes();
    selected.load("items/id");
    await context.sync();
    if (selected.items.length === 0) throw new UserError("Select one or more of your shapes on the slide first.");
    const capture = await captureShapes(context, slide, selected.items);
    const started = Date.now();
    await insertHelperSlide(context, capture.pptx, slide.id);
    return `Helper slide added after this one in ${ms(started)}, with your shapes selected. Check they look identical, then delete the helper slide.`;
  });
}

export function formatShapeTest(r: ShapeTestResult, version: string, office: string): string {
  const mark = { ok: "OK  ", warn: "WARN", fail: "FAIL", pending: "...." } as const;
  return [
    `Retro shape library test — v${version} · PowerPoint ${office}`,
    "",
    ...r.steps.map((s) => `${mark[s.status]} ${s.label}: ${s.detail}`),
    "",
    "Selected shapes:",
    ...r.items.map((i, k) => `  ${k + 1}. ${i.what} “${i.name}” — ${i.issues.length ? `helper slide (${i.issues.join(", ")})` : "one click"}`),
    ...(r.differences.length ? ["", "Differences after rebuilding:", ...r.differences.map((d) => `  - ${d}`)] : []),
  ].join("\n");
}
