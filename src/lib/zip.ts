// Just enough of a zip reader to look inside the .pptx that PowerPoint exports for a slide.
// Uses the browser's built-in DecompressionStream, so no library is needed.

export interface Zip {
  names: string[];
  bytes(name: string): Promise<Uint8Array | undefined>;
  text(name: string): Promise<string | undefined>;
  /** An entry exactly as stored (still compressed), for copying into another zip. */
  raw(name: string): RawEntry | undefined;
}

/** A zip entry as stored: compressed data plus what's needed to write it back unchanged. */
export interface RawEntry {
  name: string;
  method: number;
  crc: number;
  /** Uncompressed size. */
  size: number;
  data: Uint8Array;
}

interface Entry {
  method: number;
  crc: number;
  compressed: number;
  size: number;
  offset: number;
}

const EOCD = 0x06054b50;
const CENTRAL = 0x02014b50;
const LOCAL = 0x04034b50;

export function base64ToBytes(b64: string): Uint8Array {
  const bin = atob(b64);
  const out = new Uint8Array(bin.length);
  for (let i = 0; i < bin.length; i++) out[i] = bin.charCodeAt(i);
  return out;
}

export function bytesToBase64(bytes: Uint8Array): string {
  let bin = "";
  for (let i = 0; i < bytes.length; i += 0x8000) bin += String.fromCharCode(...bytes.subarray(i, i + 0x8000));
  return btoa(bin);
}

async function inflate(data: Uint8Array): Promise<Uint8Array> {
  const stream = new Blob([data as BlobPart]).stream().pipeThrough(new DecompressionStream("deflate-raw"));
  return new Uint8Array(await new Response(stream).arrayBuffer());
}

export function openZip(data: Uint8Array): Zip {
  const view = new DataView(data.buffer, data.byteOffset, data.byteLength);
  let end = -1;
  for (let i = data.length - 22; i >= Math.max(0, data.length - 65557); i--) {
    if (view.getUint32(i, true) === EOCD) {
      end = i;
      break;
    }
  }
  if (end < 0) throw new Error("The exported slide isn't a readable .pptx file.");

  const count = view.getUint16(end + 10, true);
  const entries = new Map<string, Entry>();
  const decoder = new TextDecoder();
  let p = view.getUint32(end + 16, true);
  for (let i = 0; i < count; i++) {
    if (view.getUint32(p, true) !== CENTRAL) throw new Error("The exported slide's file list is damaged.");
    const nameLen = view.getUint16(p + 28, true);
    const name = decoder.decode(data.subarray(p + 46, p + 46 + nameLen));
    entries.set(name, {
      method: view.getUint16(p + 10, true),
      crc: view.getUint32(p + 16, true),
      compressed: view.getUint32(p + 20, true),
      size: view.getUint32(p + 24, true),
      offset: view.getUint32(p + 42, true),
    });
    p += 46 + nameLen + view.getUint16(p + 30, true) + view.getUint16(p + 32, true);
  }

  const raw = (name: string): RawEntry | undefined => {
    const e = entries.get(name);
    if (!e) return undefined;
    if (view.getUint32(e.offset, true) !== LOCAL) throw new Error(`Can't read ${name} from the exported slide.`);
    const start = e.offset + 30 + view.getUint16(e.offset + 26, true) + view.getUint16(e.offset + 28, true);
    return { name, method: e.method, crc: e.crc, size: e.size, data: data.subarray(start, start + e.compressed) };
  };

  const bytes = async (name: string) => {
    const r = raw(name);
    if (!r) return undefined;
    if (r.method === 0) return r.data;
    if (r.method === 8) return inflate(r.data);
    throw new Error(`Unsupported compression in ${name}.`);
  };

  return {
    names: [...entries.keys()],
    raw,
    bytes,
    text: async (name) => {
      const b = await bytes(name);
      return b && decoder.decode(b);
    },
  };
}

/** Write entries (already compressed, as read with Zip.raw) into a new zip file. */
export function writeZip(entries: RawEntry[]): Uint8Array {
  const enc = new TextEncoder();
  const names = entries.map((e) => enc.encode(e.name));
  const localSize = entries.reduce((n, e, i) => n + 30 + names[i].length + e.data.length, 0);
  const centralSize = names.reduce((n, name) => n + 46 + name.length, 0);
  const out = new Uint8Array(localSize + centralSize + 22);
  const view = new DataView(out.buffer);
  let p = 0;
  const offsets: number[] = [];
  entries.forEach((e, i) => {
    offsets.push(p);
    view.setUint32(p, LOCAL, true);
    view.setUint16(p + 4, 20, true); // version needed
    view.setUint16(p + 8, e.method, true);
    view.setUint32(p + 14, e.crc, true);
    view.setUint32(p + 18, e.data.length, true);
    view.setUint32(p + 22, e.size, true);
    view.setUint16(p + 26, names[i].length, true);
    out.set(names[i], p + 30);
    out.set(e.data, p + 30 + names[i].length);
    p += 30 + names[i].length + e.data.length;
  });
  const central = p;
  entries.forEach((e, i) => {
    view.setUint32(p, CENTRAL, true);
    view.setUint16(p + 4, 20, true); // version made by
    view.setUint16(p + 6, 20, true); // version needed
    view.setUint16(p + 10, e.method, true);
    view.setUint32(p + 16, e.crc, true);
    view.setUint32(p + 20, e.data.length, true);
    view.setUint32(p + 24, e.size, true);
    view.setUint16(p + 28, names[i].length, true);
    view.setUint32(p + 42, offsets[i], true);
    out.set(names[i], p + 46);
    p += 46 + names[i].length;
  });
  view.setUint32(p, EOCD, true);
  view.setUint16(p + 8, entries.length, true);
  view.setUint16(p + 10, entries.length, true);
  view.setUint32(p + 12, centralSize, true);
  view.setUint32(p + 16, central, true);
  return out;
}

let crcTable: Uint32Array | undefined;

/** CRC-32 as zip files use it. */
export function crc32(bytes: Uint8Array): number {
  if (!crcTable) {
    crcTable = new Uint32Array(256);
    for (let n = 0; n < 256; n++) {
      let c = n;
      for (let k = 0; k < 8; k++) c = c & 1 ? 0xedb88320 ^ (c >>> 1) : c >>> 1;
      crcTable[n] = c >>> 0;
    }
  }
  let crc = 0xffffffff;
  for (const b of bytes) crc = crcTable[(crc ^ b) & 0xff] ^ (crc >>> 8);
  return (crc ^ 0xffffffff) >>> 0;
}

/** A new, uncompressed zip entry. */
export function storedEntry(name: string, data: Uint8Array): RawEntry {
  return { name, method: 0, crc: crc32(data), size: data.length, data };
}
