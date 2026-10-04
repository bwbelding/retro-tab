// Arranging logos in an even grid below the slide's title. The geometry is logoGrid in layout.ts.

import { contentArea } from "./built";
import { logoGrid } from "./layout";
import { currentSlide, UserError } from "./ppt";

/**
 * Arrange the selected shapes (or the given ones) as a logo grid, in reading order. `columns`
 * left out picks a count that gives logo-shaped cells. Returns how many were arranged.
 */
export async function arrangeLogos(columns?: number, ids?: string[]): Promise<number> {
  return PowerPoint.run(async (context) => {
    const slide = await currentSlide(context);
    let shapes: PowerPoint.Shape[];
    if (ids) {
      shapes = ids.map((id) => slide.shapes.getItem(id).load("left,top,width,height"));
    } else {
      const sel = context.presentation.getSelectedShapes().load("items/left,items/top,items/width,items/height");
      await context.sync();
      // Reading order: rows top to bottom (within half a logo's height), then left to right.
      shapes = [...sel.items].sort((a, b) => (Math.abs(a.top - b.top) > Math.min(a.height, b.height) / 2 ? a.top - b.top : a.left - b.left));
    }
    await context.sync();
    if (shapes.length < 2) throw new UserError("Select two or more logos to arrange.");
    const area = await contentArea(context, slide);
    const rects = logoGrid(shapes, area, columns);
    shapes.forEach((s, i) => {
      s.left = rects[i].left;
      s.top = rects[i].top;
      s.width = rects[i].width;
      s.height = rects[i].height;
    });
    await context.sync();
    return shapes.length;
  });
}
