import { describe, expect, it } from "vitest";
import { asIsRect, childFolders, cropFor, filterPhotos, fitRect, folderOf, isPhoto, isVector, outputSize, outputType, planScan, scanStats, splitPath, svgSize, type PhotoEntry, type PickedFile } from "../src/lib/photos";

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

  it("recognizes SVG graphics separately from photos", () => {
    expect(isVector({ name: "pin.SVG", type: "" })).toBe(true);
    expect(isVector({ name: "x", type: "image/svg+xml" })).toBe(true);
    expect(isVector({ name: "._pin.svg", type: "" })).toBe(false);
    expect(isPhoto({ name: "pin.svg", type: "" })).toBe(false);
  });

  it("accepts other common photo extensions", () => {
    for (const name of ["a.jfif", "a.jpe", "a.avif", "a.TIF"]) expect(isPhoto({ name, type: "" })).toBe(true);
  });

  it("reports what the folder picker returned, so missing photos can be traced", () => {
    const stats = scanStats([
      file("Photos/Cover.jpg"),
      file("Photos/Cities/Paris.jpg"),
      file("Photos/Cities/Europe/Rome.JPG"),
      file("Photos/Cities/Europe/clip.mov"),
      file("Photos/Raw/shot.CR2"),
      file("Photos/Raw/shot2.cr2"),
      file("Photos/Icons/pin.svg"),
      file("Photos/.DS_Store"),
      file("Photos/README"),
    ]);
    expect(stats).toEqual({
      files: 9,
      folders: 5,
      depth: 2,
      photos: 4,
      skipped: [
        { ext: ".cr2", count: 2 },
        { ext: ".mov", count: 1 },
        { ext: "(no extension)", count: 1 },
      ],
    });
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

  it("lists the subfolders inside any folder, with what's in each", () => {
    const photos = [entry("Cities/Paris 10.jpg"), entry("Cities/Europe/Rome.jpg"), entry("Cities/Europe/Italy/Milan.jpg"), entry("People/Team.jpg"), entry("Cover.jpg")];
    expect(childFolders(photos, "")).toEqual([
      { name: "Cities", path: "Cities", count: 3 },
      { name: "People", path: "People", count: 1 },
    ]);
    expect(childFolders(photos, "Cities")).toEqual([{ name: "Europe", path: "Cities/Europe", count: 2 }]);
    expect(childFolders(photos, "Cities/Europe/Italy")).toEqual([]);
  });

  it("filters by folder (and everything below it), kind and file name, in natural order", () => {
    const photos = [entry("Cities/Paris 10.jpg"), entry("Cities/Paris 2.jpg"), entry("Cities/Europe/Rome.jpg"), entry("People/Team.jpg"), entry("Cover.jpg"), { ...entry("Cities/pin.svg"), vector: true }];
    expect(filterPhotos(photos, "", "Cities", "graphics").map((p) => p.id)).toEqual(["Cities/pin.svg"]);
    expect(filterPhotos(photos, "", "Cities", "photos")).toHaveLength(3);
    expect(filterPhotos(photos, "", "Cities", "photos").map((p) => p.id)).toEqual(["Cities/Europe/Rome.jpg", "Cities/Paris 2.jpg", "Cities/Paris 10.jpg"]);
    expect(filterPhotos(photos, "", "Cities/Europe").map((p) => p.id)).toEqual(["Cities/Europe/Rome.jpg"]);
    expect(filterPhotos(photos, "team", "").map((p) => p.id)).toEqual(["People/Team.jpg"]);
    expect(filterPhotos(photos, "", "")).toHaveLength(6);
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

  it("reads an SVG's size from width, height and viewBox", () => {
    expect(svgSize("48", "24", null)).toEqual({ width: 48, height: 24 });
    expect(svgSize("36pt", "18pt", null)).toEqual({ width: 48, height: 24 });
    expect(svgSize(null, null, "0 0 512 256")).toEqual({ width: 512, height: 256 });
    expect(svgSize("100%", "100%", "0,0,64,64")).toEqual({ width: 64, height: 64 });
    expect(svgSize("200", null, "0 0 100 50")).toEqual({ width: 200, height: 100 });
    expect(svgSize(null, null, null)).toEqual({ width: 100, height: 100 });
  });

  it("fits a graphic inside part of an area, centered", () => {
    expect(fitRect(100, 100, { left: 0, top: 0, width: 960, height: 540 }, 0.4)).toEqual({ left: 372, top: 162, width: 216, height: 216 });
    expect(fitRect(200, 100, { left: 100, top: 100, width: 100, height: 100 }, 0.8)).toEqual({ left: 110, top: 130, width: 80, height: 40 });
  });

  it("keeps formats that can be transparent", () => {
    expect(outputType("logo.PNG")).toBe("image/png");
    expect(outputType("photo.heic")).toBe("image/jpeg");
  });
});
