// Applying brand kits in PowerPoint: colors to fill, outline or text; fonts to text; and reading
// colors and fonts from the selection or the deck's theme. The model is in brand.ts.

import { activeKit, createBrandStore, kitFromTheme, newKit, normalizeHex, type BrandKit, type ColorTarget } from "./brand";
import { kv } from "./idb";
import { deckTheme } from "./ooxml";
import { currentSlide, UserError } from "./ppt";
import { base64ToBytes, openZip } from "./zip";

export const brandStore = createBrandStore(kv);

const TEXT_TYPES = new Set(["GeometricShape", "TextBox", "Placeholder", "Callout"]);
const FILL_TYPES = new Set(["GeometricShape", "TextBox", "Placeholder", "Callout"]);
const LINE_TYPES = new Set(["GeometricShape", "TextBox", "Placeholder", "Callout", "Line", "Image"]);

/** The selected shapes, with groups opened up into the shapes inside them. */
async function selectedLeaves(context: PowerPoint.RequestContext): Promise<PowerPoint.Shape[]> {
  const sel = context.presentation.getSelectedShapes();
  sel.load("items/id,items/type");
  await context.sync();
  const out: PowerPoint.Shape[] = [];
  const open = async (shapes: PowerPoint.Shape[]) => {
    for (const s of shapes) {
      if (s.type === "Group") {
        const inner = s.group.shapes;
        inner.load("items/id,items/type");
        await context.sync();
        await open(inner.items);
      } else out.push(s);
    }
  };
  await open(sel.items);
  return out;
}

async function kitOrThrow(): Promise<BrandKit> {
  const kit = activeKit(await brandStore.load());
  if (!kit) throw new UserError("Set up a brand kit first (Retro Polish → Brand Kit).");
  return kit;
}

/** Apply a color to the selection's fill, outline or text. Returns how many shapes changed. */
export async function applyColor(hex: string, target: ColorTarget): Promise<number> {
  return PowerPoint.run(async (context) => {
    const shapes = await selectedLeaves(context);
    if (shapes.length === 0) throw new UserError("Select the shapes to color first.");
    const types = target === "fill" ? FILL_TYPES : target === "line" ? LINE_TYPES : TEXT_TYPES;
    const fits = shapes.filter((s) => types.has(s.type));
    if (fits.length === 0) throw new UserError(target === "text" ? "None of the selected shapes has text." : `Can't set the ${target === "fill" ? "fill" : "outline"} of the selected shapes.`);
    for (const s of fits) {
      if (target === "fill") s.fill.setSolidColor(hex);
      else if (target === "line") {
        s.lineFormat.visible = true;
        s.lineFormat.color = hex;
      } else s.textFrame.textRange.font.color = hex;
    }
    await context.sync();
    return fits.length;
  });
}

/** Ribbon Fill 1/2/3: fill the selection with the active kit's nth color. */
export async function applyFill(n: number): Promise<void> {
  const kit = await kitOrThrow();
  const hex = kit.colors[n - 1];
  if (!hex) throw new UserError(`The “${kit.name}” kit has no color ${n} yet. Add colors in Retro Polish → Brand Kit.`);
  await applyColor(hex, "fill");
}

const TITLE_PLACEHOLDERS = new Set(["Title", "CenterTitle", "VerticalTitle"]);

/**
 * Brand fonts: titles get the heading font and everything else the body font. Works on the
 * selection, or on the whole slide when nothing is selected. `only` forces one font everywhere.
 */
export async function applyFonts(only?: "heading" | "body"): Promise<string> {
  const kit = await kitOrThrow();
  if (!kit.headingFont && !kit.bodyFont) throw new UserError(`The “${kit.name}” kit has no fonts yet. Add them in Retro Polish → Brand Kit.`);
  return PowerPoint.run(async (context) => {
    let shapes = await selectedLeaves(context);
    const whole = shapes.length === 0;
    if (whole && only) throw new UserError("Select the text to change first.");
    if (whole) {
      const slide = await currentSlide(context);
      slide.shapes.load("items/id,items/type");
      await context.sync();
      shapes = slide.shapes.items.filter((s) => s.type !== "Group");
    }
    const texty = shapes.filter((s) => TEXT_TYPES.has(s.type));
    const kinds = texty.map((s) => (s.type === "Placeholder" ? s.placeholderFormat.load("type") : undefined));
    await context.sync();
    let changed = 0;
    texty.forEach((s, i) => {
      const isTitle = Boolean(kinds[i] && TITLE_PLACEHOLDERS.has(kinds[i]!.type));
      const font = only === "heading" ? kit.headingFont : only === "body" ? kit.bodyFont : isTitle ? kit.headingFont || kit.bodyFont : kit.bodyFont || kit.headingFont;
      if (!font) return;
      s.textFrame.textRange.font.name = font;
      changed++;
    });
    await context.sync();
    if (changed === 0) throw new UserError(whole ? "There's no text on this slide." : "None of the selected shapes has text.");
    return `Applied brand fonts to ${changed} shape${changed === 1 ? "" : "s"}${whole ? " on this slide" : ""}.`;
  });
}

/** The selected shape's fill color (or its text color when it has no solid fill). */
export async function colorFromSelection(): Promise<string> {
  return PowerPoint.run(async (context) => {
    const [s] = await selectedLeaves(context);
    if (!s) throw new UserError("Select a shape to take its color.");
    const fill = FILL_TYPES.has(s.type) ? s.fill.load("type,foregroundColor") : undefined;
    const font = TEXT_TYPES.has(s.type) ? s.textFrame.textRange.font.load("color") : undefined;
    await context.sync();
    const hex = (fill?.type === "Solid" ? normalizeHex(fill.foregroundColor) : undefined) ?? (font?.color ? normalizeHex(font.color) : undefined);
    if (!hex) throw new UserError("That shape has no solid fill or text color to take.");
    return hex;
  });
}

/** The font used in the selected shape's text. */
export async function fontFromSelection(): Promise<string> {
  return PowerPoint.run(async (context) => {
    const [s] = (await selectedLeaves(context)).filter((x) => TEXT_TYPES.has(x.type));
    if (!s) throw new UserError("Select a shape with text to take its font.");
    const font = s.textFrame.textRange.font.load("name");
    await context.sync();
    if (!font.name) throw new UserError("That text uses more than one font. Select text in a single font.");
    return font.name;
  });
}

/** A new kit holding this deck's theme colors and fonts. */
export async function kitFromDeck(): Promise<BrandKit> {
  const pptx = await PowerPoint.run(async (context) => {
    const slide = await currentSlide(context);
    const b64 = slide.exportAsBase64();
    await context.sync();
    return b64.value;
  });
  const theme = await deckTheme(openZip(base64ToBytes(pptx)));
  if (!theme) throw new UserError("Couldn't find this deck's theme.");
  return newKit("", kitFromTheme(theme));
}
