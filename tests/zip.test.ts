import { describe, expect, it } from "vitest";
import { base64ToBytes, bytesToBase64, openZip } from "../src/lib/zip";
import { makeZip } from "./helpers/makeZip";

describe("zip reader", () => {
  it("reads deflated and stored entries", async () => {
    for (const compress of [true, false]) {
      const zip = openZip(await makeZip({ "ppt/slides/slide1.xml": "<a>hello</a>", "ppt/media/image1.png": new Uint8Array([1, 2, 3]) }, compress));
      expect(zip.names).toEqual(["ppt/slides/slide1.xml", "ppt/media/image1.png"]);
      expect(await zip.text("ppt/slides/slide1.xml")).toBe("<a>hello</a>");
      expect(Array.from((await zip.bytes("ppt/media/image1.png"))!)).toEqual([1, 2, 3]);
      expect(await zip.text("missing.xml")).toBeUndefined();
    }
  });

  it("rejects files that aren't zips", () => {
    expect(() => openZip(new Uint8Array(40))).toThrow(/isn't a readable/);
  });

  it("round-trips base64", () => {
    const bytes = new Uint8Array(70000).map((_, i) => i % 256);
    expect(Array.from(base64ToBytes(bytesToBase64(bytes)))).toEqual(Array.from(bytes));
  });
});
