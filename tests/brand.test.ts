import { beforeAll, describe, expect, it } from "vitest";
import { Window } from "happy-dom";
import { activeKit, addColor, createBrandStore, kitFromTheme, moveColor, newKit, normalizeHex } from "../src/lib/brand";
import type { Store } from "../src/lib/library";

beforeAll(() => {
  globalThis.DOMParser = new Window().DOMParser as unknown as typeof DOMParser;
});

const THEME = `<a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" name="Acme"><a:themeElements>
<a:clrScheme name="Acme"><a:dk1><a:sysClr val="windowText" lastClr="000000"/></a:dk1><a:lt1><a:sysClr val="window" lastClr="FFFFFF"/></a:lt1>
<a:dk2><a:srgbClr val="1F2A44"/></a:dk2><a:lt2><a:srgbClr val="e7e6e6"/></a:lt2>
<a:accent1><a:srgbClr val="C43E1C"/></a:accent1><a:accent2><a:srgbClr val="5E8AB4"/></a:accent2><a:accent3><a:srgbClr val="333F48"/></a:accent3>
<a:accent4><a:srgbClr val="C43E1C"/></a:accent4><a:accent5><a:srgbClr val="8C8C8C"/></a:accent5><a:accent6><a:srgbClr val="2E7D32"/></a:accent6>
<a:hlink><a:srgbClr val="0563C1"/></a:hlink></a:clrScheme>
<a:fontScheme name="Acme"><a:majorFont><a:latin typeface="Georgia"/><a:ea typeface=""/></a:majorFont><a:minorFont><a:latin typeface="Arial"/></a:minorFont></a:fontScheme>
</a:themeElements></a:theme>`;

describe("brand kits", () => {
  it("normalizes hex colors", () => {
    expect(normalizeHex("#c43e1c")).toBe("#C43E1C");
    expect(normalizeHex("abc")).toBe("#AABBCC");
    expect(normalizeHex(" 5E8AB4 ")).toBe("#5E8AB4");
    expect(normalizeHex("red")).toBeUndefined();
    expect(normalizeHex("#12345")).toBeUndefined();
  });

  it("adds colors once and reorders them", () => {
    expect(addColor(["#111111"], "#111111")).toEqual(["#111111"]);
    expect(addColor(["#111111"], "#222222")).toEqual(["#111111", "#222222"]);
    expect(moveColor(["#A", "#B", "#C"], 2, -1)).toEqual(["#A", "#C", "#B"]);
    expect(moveColor(["#A", "#B"], 0, -1)).toEqual(["#A", "#B"]);
  });

  it("reads a deck theme's accents, dark/light pair and fonts", () => {
    const kit = kitFromTheme(new DOMParser().parseFromString(THEME, "application/xml"));
    expect(kit).toEqual({
      name: "Acme (from deck)",
      colors: ["#C43E1C", "#5E8AB4", "#333F48", "#8C8C8C", "#2E7D32", "#1F2A44", "#E7E6E6"],
      headingFont: "Georgia",
      bodyFont: "Arial",
    });
  });

  it("starts with one empty kit and finds the active one", async () => {
    const map = new Map<string, unknown>();
    const store: Store = { get: async <T>(k: string) => map.get(k) as T | undefined, set: async (k, v) => void map.set(k, v), del: async (k) => void map.delete(k) };
    const brand = createBrandStore(store);
    const first = await brand.load();
    expect(first.kits).toHaveLength(1);
    expect(activeKit(first)?.name).toBe("My brand");
    const client = newKit("Client A", { colors: ["#000000"] });
    await brand.save({ active: client.id, kits: [...first.kits, client] });
    expect(activeKit(await brand.load())?.name).toBe("Client A");
  });
});
