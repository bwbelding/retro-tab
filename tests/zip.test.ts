import { describe, expect, it } from "vitest";
import { base64ToBytes, bytesToBase64, crc32, openZip, storedEntry, writeZip } from "../src/lib/zip";
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

describe("zip writer", () => {
  it("copies entries into a new zip that reads back the same", async () => {
    const source = openZip(await makeZip({ "a.xml": "<a>hello</a>", "b.bin": new Uint8Array([9, 8, 7]) }));
    const copy = openZip(writeZip(source.names.map((n) => source.raw(n)!)));
    expect(copy.names).toEqual(["a.xml", "b.bin"]);
    expect(await copy.text("a.xml")).toBe("<a>hello</a>");
    expect(Array.from((await copy.bytes("b.bin"))!)).toEqual([9, 8, 7]);
    expect(copy.raw("a.xml")).toMatchObject({ method: 8, size: 12 });
  });
});

describe("new entries", () => {
  it("computes zip CRC-32", () => {
    expect(crc32(new TextEncoder().encode("123456789"))).toBe(0xcbf43926);
    expect(crc32(new Uint8Array())).toBe(0);
  });

  it("writes uncompressed entries that read back", async () => {
    const zip = openZip(writeZip([storedEntry("a.txt", new TextEncoder().encode("hello"))]));
    expect(await zip.text("a.txt")).toBe("hello");
    expect(zip.raw("a.txt")).toMatchObject({ method: 0, crc: 0x3610a686, size: 5 });
  });
});
