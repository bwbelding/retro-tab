import { describe, expect, it } from "vitest";
import { asIsRect, cropFor, filterPhotos, folderOf, isPhoto, outputSize, outputType, planScan, splitPath, topFolders, type PhotoEntry, type PickedFile } from "../src/lib/photos";

const file = (path: string, size = 100, lastModified = 1): PickedFile => ({ name: path.split("/").pop()!, type: "", size, lastModified, webkitRelativePath: path });
const entry = (id: string, size = 100, modified = 1): PhotoEntry => ({ id, name: id.split("/").pop()!, folder: folderOf(id), size, modified, width: 4000, height: 3000 });

describe("photo folder", () => {
  it("recognizes photos and skips hidden and non-image files", () => {
    expect(isPhoto({ name: "Paris.JPG", type: "" })).toBe(true);
    expect(isPhoto({ name: "scan.heic", type: "" })).toBe(true);
    expect(isPhoto({ name: "x.dat", type: "image/png" })).toBe(true);
    expect(isPhoto({ name: "._Paris.jpg", type: "image/jpeg" })).toBe(false);
    expect(isPhoto({ name: "notes.txt", type: "text/plain" })).toBe(false);
    expect(isPhoto({ name: "logo.svg", type: "image/svg+xml" })).toBe(false);
  });

  it("splits paths into the folder and the path inside it", () => {
    expect(splitPath("Brand Photos/Cities/Paris.jpg")).toEqual({ root: "Brand Photos", path: "Cities/Paris.jpg" });
    expect(folderOf("Cities/Europe/Paris.jpg")).toBe("Cities/Europe");
    expect(folderOf("Paris.jpg")).toBe("");
  });

  it("only re-reads new or changed photos on a rescan", () => {
    const known = [entry("Cities/Paris.jpg"), entry("Cities/Rome.jpg"), entry("Old.jpg")];
    const plan = planScan([file("Photos/Cities/Paris.jpg"), file("Photos/Cities/Rome.jpg", 200), file("Photos/New.png"), file("Photos/readme.txt")], known);
    expect(plan.root).toBe("Photos");
    expect(plan.unchanged.map((p) => p.id)).toEqual(["Cities/Paris.jpg"]);
    expect(plan.fresh.map((f) => f.webkitRelativePath)).toEqual(["Photos/Cities/Rome.jpg", "Photos/New.png"]);
    expect(plan.removed).toEqual(["Old.jpg"]);
  });

  it("filters by top-level subfolder and file name, in natural order", () => {
    const photos = [entry("Cities/Paris 10.jpg"), entry("Cities/Paris 2.jpg"), entry("Cities/Europe/Rome.jpg"), entry("People/Team.jpg"), entry("Cover.jpg")];
    expect(topFolders(photos)).toEqual(["Cities", "People"]);
    expect(filterPhotos(photos, "", "Cities").map((p) => p.id)).toEqual(["Cities/Europe/Rome.jpg", "Cities/Paris 2.jpg", "Cities/Paris 10.jpg"]);
    expect(filterPhotos(photos, "team", "").map((p) => p.id)).toEqual(["People/Team.jpg"]);
    expect(filterPhotos(photos, "", "")).toHaveLength(5);
  });
});

describe("photo geometry", () => {
  it("crops a wide photo evenly from both sides to fit a narrower box", () => {
    expect(cropFor(4000, 2000, 1, "top")).toEqual({ sx: 1000, sy: 0, sw: 2000, sh: 2000 });
  });

  it("crops a tall photo from the top, center or bottom", () => {
    expect(cropFor(1600, 2000, 16 / 9, "top")).toEqual({ sx: 0, sy: 0, sw: 1600, sh: 900 });
    expect(cropFor(1600, 2000, 16 / 9, "center")).toEqual({ sx: 0, sy: 550, sw: 1600, sh: 900 });
    expect(cropFor(1600, 2000, 16 / 9, "bottom")).toEqual({ sx: 0, sy: 1100, sw: 1600, sh: 900 });
  });

  it("saves at 2 pixels per point, never enlarging and never over 3000 px", () => {
    expect(outputSize({ sw: 6000, sh: 3375 }, { width: 960, height: 540 })).toEqual({ width: 1920, height: 1080 });
    expect(outputSize({ sw: 800, sh: 450 }, { width: 960, height: 540 })).toEqual({ width: 800, height: 450 });
    expect(outputSize({ sw: 8000, sh: 8000 }, { width: 4000, height: 4000 })).toEqual({ width: 3000, height: 3000 });
  });

  it("fits an uncropped photo in 80% of the slide, centered", () => {
    expect(asIsRect(4000, 3000, { width: 960, height: 540 })).toEqual({ left: 192, top: 54, width: 576, height: 432 });
  });

  it("keeps formats that can be transparent", () => {
    expect(outputType("logo.PNG")).toBe("image/png");
    expect(outputType("photo.heic")).toBe("image/jpeg");
  });
});
