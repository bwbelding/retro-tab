// Drawing a roadmap in PowerPoint, and reading one back for editing. Parsing and layout are in
// roadmap.ts.

import { activeKit } from "./brand";
import { brandStore } from "./brandActions";
import { contentArea, readBuiltSelection, tagBuilt } from "./built";
import { currentSlide, UserError } from "./ppt";
import { layoutRoadmap, parseRoadmap, type Area, type RoadmapPart, type RoadmapSource } from "./roadmap";
import { textOn } from "./tracker";

const KIND = "roadmap";
/** Lane colors when the brand kit has none. */
const FALLBACK = ["#333F48", "#C43E1C", "#5E8AB4", "#7A7A7A"];
const RULE = "#BFBFBF";

/** The selected roadmap, if exactly one Retro roadmap is selected. */
export const readRoadmapSelection = () => readBuiltSelection<RoadmapSource>(KIND);

function draw(slide: PowerPoint.Slide, part: RoadmapPart, colors: string[], font: string): PowerPoint.Shape {
  const rect = { left: part.left, top: part.top, width: part.width, height: part.height };
  const color = colors[(part.lane ?? 0) % colors.length];
  const text = (s: PowerPoint.Shape, size: number, bold: boolean, fill: string, align: PowerPoint.ParagraphHorizontalAlignment) => {
    const t = s.textFrame;
    t.textRange.font.size = size;
    t.textRange.font.bold = bold;
    t.textRange.font.color = fill;
    if (font) t.textRange.font.name = font;
    t.textRange.paragraphFormat.horizontalAlignment = align;
    t.verticalAlignment = PowerPoint.TextVerticalAlignment.middle;
    t.wordWrap = part.kind === "laneLabel";
    t.autoSizeSetting = PowerPoint.ShapeAutoSize.autoSizeNone;
    t.leftMargin = t.rightMargin = part.kind === "bar" ? 4 : 0;
    t.topMargin = t.bottomMargin = 0;
  };
  let s: PowerPoint.Shape;
  switch (part.kind) {
    case "gridLine":
    case "laneLine":
      s = slide.shapes.addLine(PowerPoint.ConnectorType.straight, rect);
      s.lineFormat.color = RULE;
      s.lineFormat.weight = 0.75;
      break;
    case "bar":
      s = slide.shapes.addGeometricShape(PowerPoint.GeometricShapeType.roundRectangle, rect);
      s.fill.setSolidColor(color);
      s.lineFormat.visible = false;
      s.textFrame.textRange.text = part.text ?? "";
      text(s, 12, true, textOn(color), PowerPoint.ParagraphHorizontalAlignment.left);
      break;
    case "milestone":
      s = slide.shapes.addGeometricShape(PowerPoint.GeometricShapeType.diamond, rect);
      s.fill.setSolidColor(color);
      s.lineFormat.visible = false;
      break;
    default:
      s = slide.shapes.addTextBox(part.text ?? "", rect);
      text(s, 12, part.kind !== "milestoneLabel", "#000000", part.kind === "header" ? PowerPoint.ParagraphHorizontalAlignment.center : PowerPoint.ParagraphHorizontalAlignment.left);
  }
  s.name = `Retro roadmap ${part.kind}`;
  return s;
}

/**
 * Draw a roadmap from pasted rows: in place of the selected roadmap when one is selected (so
 * editing keeps its spot), otherwise below the current slide's title. Returns a note for the pane.
 */
export async function insertRoadmap(source: RoadmapSource, replaceId?: string): Promise<string> {
  const { items, errors } = parseRoadmap(source.text, source.fyStart);
  if (items.length === 0) throw new UserError(errors[0] ?? "Paste or type the roadmap rows first: lane, item, start, end.");
  const kit = activeKit(await brandStore.load());
  const colors = kit?.colors.length ? kit.colors : FALLBACK;
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
    const shapes = layoutRoadmap(items, area, source.fyStart).map((part) => draw(slide, part, colors, font));
    await context.sync();
    const group = slide.shapes.addGroup(shapes);
    group.name = "Retro roadmap";
    tagBuilt(group, KIND, source);
    await context.sync();
    slide.setSelectedShapes([group.id]);
    await context.sync();
  });
  const lanes = new Set(items.map((i) => i.lane)).size;
  const done = `${replaceId ? "Updated" : "Drew"} the roadmap: ${items.length} item${items.length === 1 ? "" : "s"} in ${lanes} lane${lanes === 1 ? "" : "s"}.`;
  return errors.length ? `${done} Skipped: ${errors.join(" ")}` : done;
}
