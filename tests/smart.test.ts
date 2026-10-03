import { describe, expect, it } from "vitest";
import { detectAngleUnit, nextNumber, normalizeValue, partsFor, pieAngles, placeFor, readingOrder, SMART, SMART_KINDS } from "../src/lib/smart";

const box = { left: 100, top: 50, width: 40, height: 40 };

describe("normalizeValue", () => {
  it("keeps valid choices and falls back to the default", () => {
    expect(normalizeValue("harvey", "75")).toBe("75");
    expect(normalizeValue("harvey", "60")).toBe("50"); // quarters only
    expect(normalizeValue("trend", "sideways")).toBe("up");
  });
  it("clamps numbers", () => {
    expect(normalizeValue("progress", "140")).toBe("100");
    expect(normalizeValue("progress", "-3")).toBe("0");
    expect(normalizeValue("number", "4.6")).toBe("5");
    expect(normalizeValue("number", "abc")).toBe("1");
  });
});

describe("partsFor", () => {
  it("draws every kind's default inside its box", () => {
    for (const kind of SMART_KINDS) {
      const parts = partsFor(kind, SMART[kind].defaultValue, box);
      expect(parts.length, kind).toBeGreaterThan(0);
      for (const p of parts) {
        expect(p.left, kind).toBeGreaterThanOrEqual(box.left - 0.001);
        expect(p.top, kind).toBeGreaterThanOrEqual(box.top - 0.001);
        expect(p.left + p.width, kind).toBeLessThanOrEqual(box.left + box.width + 0.001);
        expect(p.top + p.height, kind).toBeLessThanOrEqual(box.top + box.height + 0.001);
      }
    }
  });

  it("builds Harvey balls from a ring and a pie in quarters", () => {
    expect(partsFor("harvey", "0", box)).toHaveLength(1);
    expect(partsFor("harvey", "100", box)[0].fill).toBe("#333F48");
    const half = partsFor("harvey", "50", box);
    expect(half.map((p) => p.shape)).toEqual(["Ellipse", "Pie"]);
    expect(half[1].pieFraction).toBe(0.5);
  });

  it("fills the progress bar to the percentage", () => {
    const parts = partsFor("progress", "25", { left: 0, top: 0, width: 200, height: 10 });
    expect(parts[1].width).toBe(50);
    expect(partsFor("progress", "0", box)).toHaveLength(1);
  });

  it("fills the right number of stars", () => {
    const stars = partsFor("stars", "3", { left: 0, top: 0, width: 110, height: 20 });
    expect(stars).toHaveLength(5);
    expect(stars.filter((s) => s.fill).length).toBe(3);
  });

  it("lights only the chosen traffic light", () => {
    const lights = partsFor("traffic3", "amber", { left: 0, top: 0, width: 22, height: 58 }).slice(1);
    expect(lights.map((l) => l.fill)).toEqual(["#6B6B6B", "#F2A900", "#6B6B6B"]);
  });

  it("rotates and colours trends", () => {
    expect(partsFor("trend", "up", box)[0]).toMatchObject({ shape: "RightArrow", rotation: -90, fill: "#1E8E3E" });
    expect(partsFor("trend", "down-slight", box)[0]).toMatchObject({ rotation: 45, fill: "#D93025" });
    expect(partsFor("trend", "flat", box)[0].rotation).toBe(0);
  });

  it("points arrows the chosen way and ticks checkboxes", () => {
    expect(partsFor("arrow", "left", box)[0].shape).toBe("LeftArrow");
    expect(partsFor("checkbox", "on", box)[0].text?.text).toBe("✓");
    expect(partsFor("checkbox", "off", box)[0].text).toBeUndefined();
  });
});

describe("pie angles", () => {
  it("detects degrees versus raw OOXML units from the default end angle", () => {
    expect(detectAngleUnit(270)).toBe(1);
    expect(detectAngleUnit(-90)).toBe(1);
    expect(detectAngleUnit(16200000)).toBe(60000);
  });
  it("sweeps clockwise from 12 o'clock", () => {
    expect(pieAngles(0.25, 1)).toEqual([-90, 0]);
    expect(pieAngles(0.5, 1)).toEqual([-90, 90]);
    expect(pieAngles(0.75, 1)).toEqual([-90, 180]);
    expect(pieAngles(0.25, 60000)).toEqual([16200000, 0]);
    expect(pieAngles(0.75, 60000)).toEqual([16200000, 10800000]);
  });
});

describe("numbering", () => {
  it("continues after the highest number on the slide", () => {
    expect(nextNumber([])).toBe(1);
    expect(nextNumber(["1", "4", "2", ""])).toBe(5);
  });
  it("orders circles in rows, left to right", () => {
    const order = readingOrder([
      { id: "c", left: 10, top: 100, height: 20 },
      { id: "b", left: 200, top: 12, height: 20 },
      { id: "a", left: 10, top: 10, height: 20 },
    ]);
    expect(order.map((o) => o.id)).toEqual(["a", "b", "c"]);
  });
});

describe("placeFor", () => {
  it("centres on the slide without a selection", () => {
    expect(placeFor("harvey", { width: 960, height: 540 }, [])).toEqual({ left: 468, top: 258, width: 24, height: 24 });
  });
  it("sits right of the selection, inside the slide", () => {
    const at = placeFor("number", { width: 960, height: 540 }, [{ id: "x", left: 100, top: 100, width: 200, height: 50 }]);
    expect(at).toEqual({ left: 308, top: 111, width: 28, height: 28 });
  });
});
