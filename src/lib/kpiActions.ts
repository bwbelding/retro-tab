// Drawing KPI tiles or a scorecard table in PowerPoint. Parsing and layout are in kpi.ts.

import { activeKit } from "./brand";
import { brandStore } from "./brandActions";
import { contentArea, readBuiltSelection, tagBuilt } from "./built";
import { ARROW, parseKpis, STATUS_LABEL, tableText, tileRects, type Kpi, type KpiSource } from "./kpi";
import { currentSlide, UserError } from "./ppt";
import type { Area } from "./roadmap";
import { COLORS } from "./smart";
import { textOn } from "./tracker";

const KIND = "kpi";
const RULE = "#BFBFBF";
/** Status colors match the Smart Elements' traffic lights. */
const STATUS_COLOR = { green: COLORS.green, amber: COLORS.amber, red: COLORS.red };
const DIRECTION_COLOR = { up: COLORS.green, down: COLORS.red, flat: COLORS.dim };

export const readKpiSelection = () => readBuiltSelection<KpiSource>(KIND);

/** The change line's color: the status if there is one, otherwise up green, down red. */
const changeColor = (k: Kpi) => (k.status ? STATUS_COLOR[k.status] : k.direction ? DIRECTION_COLOR[k.direction] : "#000000");

function textStyle(s: PowerPoint.Shape, size: number, color: string, font: string, bold = false) {
  const t = s.textFrame;
  t.textRange.font.size = size;
  t.textRange.font.color = color;
  t.textRange.font.bold = bold;
  if (font) t.textRange.font.name = font;
  t.wordWrap = true;
  t.autoSizeSetting = PowerPoint.ShapeAutoSize.autoSizeNone;
  t.leftMargin = t.rightMargin = t.topMargin = t.bottomMargin = 0;
}

/** One tile: a frame, the big number, the metric and the change. */
function drawTile(slide: PowerPoint.Slide, k: Kpi, r: Area, accent: string, font: string): PowerPoint.Shape[] {
  const pad = 14;
  const valueSize = Math.round(Math.min(Math.max(r.height * 0.3, 28), 48));
  const frame = slide.shapes.addGeometricShape(PowerPoint.GeometricShapeType.roundRectangle, r);
  frame.fill.clear();
  frame.lineFormat.color = RULE;
  frame.lineFormat.weight = 1;
  const inner = { left: r.left + pad, width: r.width - 2 * pad };
  const value = slide.shapes.addTextBox(k.value, { ...inner, top: r.top + pad, height: valueSize * 1.25 });
  textStyle(value, valueSize, accent, font, true);
  const metric = slide.shapes.addTextBox(k.metric, { ...inner, top: r.top + pad + valueSize * 1.3, height: 20 });
  textStyle(metric, 14, "#000000", font);
  const shapes = [frame, value, metric];
  const line = [k.change ? `${ARROW[k.direction ?? "flat"]} ${k.change}` : "", k.status ? `● ${STATUS_LABEL[k.status]}` : ""].filter(Boolean).join("   ");
  if (line) {
    const change = slide.shapes.addTextBox(line, { ...inner, top: r.top + r.height - pad - 20, height: 20 });
    textStyle(change, 14, changeColor(k), font, true);
    shapes.push(change);
  }
  shapes.forEach((s) => (s.name = "Retro KPI part"));
  return shapes;
}

/** A scorecard table: a header row in the brand color, then one row per KPI with colored change and status. */
function drawTable(slide: PowerPoint.Slide, rows: Kpi[], area: Area, accent: string, font: string): PowerPoint.Shape {
  const text = tableText(rows);
  const rowHeight = 30;
  const header = { fill: { color: accent }, font: { color: textOn(accent), bold: true, size: 14, ...(font ? { name: font } : {}) } };
  const body = (color = "#000000", bold = false) => ({
    font: { color, bold, size: 14, ...(font ? { name: font } : {}) },
    borders: { bottom: { color: RULE, weight: 0.75 } },
  });
  const cells = text.map((row, r) =>
    row.map((cell, c) => {
      if (r === 0) return { ...header, text: cell };
      const k = rows[r - 1];
      if (c === 2 && cell) return { ...body(changeColor({ ...k, status: undefined })), text: cell };
      if (c === 3 && k.status) return { ...body(), textRuns: [{ text: "● ", font: { color: STATUS_COLOR[k.status] } }, { text: cell }] };
      return { ...body("#000000", c === 1), text: cell };
    }),
  );
  return slide.shapes.addTable(text.length, 4, {
    left: area.left,
    top: area.top,
    width: area.width,
    height: Math.min(rowHeight * text.length, area.height),
    style: "NoStyleNoGrid",
    specificCellProperties: cells,
  });
}

/**
 * Draw KPI tiles or a scorecard table from pasted rows: in place of the selected KPIs when they're
 * selected (so editing keeps their spot), otherwise below the current slide's title.
 */
export async function insertKpis(source: KpiSource, replaceId?: string): Promise<string> {
  const { rows, errors } = parseKpis(source.text);
  if (rows.length === 0) throw new UserError(errors[0] ?? "Paste or type the KPI rows first: metric, value, change, status.");
  const kit = activeKit(await brandStore.load());
  const accent = kit?.colors[0] ?? "#333F48";
  const font = kit?.bodyFont ?? "";
  await PowerPoint.run(async (context) => {
    const slide = await currentSlide(context);
    let area: Area;
    if (replaceId) {
      const old = slide.shapes.getItem(replaceId).load("left,top,width,height");
      await context.sync();
      area = { left: old.left, top: old.top, width: old.width, height: old.height };
      old.delete();
    } else {
      area = await contentArea(context, slide);
    }
    let made: PowerPoint.Shape;
    if (source.style === "table") {
      made = drawTable(slide, rows, area, accent, font);
      made.name = "Retro scorecard";
    } else {
      const shapes = tileRects(rows.length, area).flatMap((r, i) => drawTile(slide, rows[i], r, accent, font));
      await context.sync();
      made = slide.shapes.addGroup(shapes);
      made.name = "Retro KPI tiles";
    }
    tagBuilt(made, KIND, source);
    await context.sync();
    slide.setSelectedShapes([made.id]);
    await context.sync();
  });
  const done = `${replaceId ? "Updated" : "Made"} ${source.style === "table" ? "a scorecard table" : `${rows.length} KPI tile${rows.length === 1 ? "" : "s"}`} with ${rows.length} metric${rows.length === 1 ? "" : "s"}.`;
  return errors.length ? `${done} ${errors.join(" ")}` : done;
}
