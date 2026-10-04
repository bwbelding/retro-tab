// Strips an exported slide down to the shapes being saved, so the shape library never keeps the
// rest of the slide: other shapes, speaker notes, comments, pictures and charts only they used,
// animations, and the deck's document properties and thumbnail. What's left is still a valid
// .pptx that inserts as a helper slide with just the saved shapes.

import { CAPTURE_TAG, parseXml, resolvePart, slidePart } from "./ooxml";
import { openZip, storedEntry, writeZip, type RawEntry } from "./zip";

interface Rel {
  el: Element;
  id: string;
  type: string;
  target: string;
  external: boolean;
}

const XML_HEADER = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n';
/** Children of the shape tree that aren't shapes and must stay. */
const TREE_PARTS = new Set(["nvGrpSpPr", "grpSpPr", "extLst"]);

const relsPartOf = (part: string) => {
  const i = part.lastIndexOf("/");
  return `${part.slice(0, i + 1)}_rels/${part.slice(i + 1)}.rels`;
};

/** The part a .rels file belongs to ("" for the package's own _rels/.rels). */
const sourceOf = (relsPart: string) => relsPart.replace(/_rels\/([^/]*)\.rels$/, "$1");

function readRels(text: string | undefined, source: string): { doc?: Document; rels: Rel[] } {
  if (!text) return { rels: [] };
  const doc = parseXml(text);
  const rels = Array.from(doc.getElementsByTagName("Relationship")).map((el) => {
    const external = el.getAttribute("TargetMode") === "External";
    const raw = el.getAttribute("Target") ?? "";
    return { el, id: el.getAttribute("Id") ?? "", type: el.getAttribute("Type") ?? "", external, target: external ? raw : resolvePart(source || "x", raw) };
  });
  return { doc, rels };
}

const serialize = (doc: Document) => new TextEncoder().encode(XML_HEADER + new XMLSerializer().serializeToString(doc).replace(/^<\?xml[^>]*\?>\s*/, ""));

const kid = (el: Element | undefined, name: string) => (el ? Array.from(el.children).find((c) => c.localName === name) : undefined);

/** Package-level parts that hold people's names, printers or document metadata rather than slide content. */
const PEOPLE_AND_METADATA = ["/commentAuthors", "/authors", "/customXml", "/printerSettings"];

/** The rId of a shape's own tags (in its non-visual properties, not a nested shape's). */
function ownTagId(shape: Element): string | undefined {
  // Ink, equations and newer content come wrapped in mc:AlternateContent; look at the wrapped shape.
  if (shape.localName === "AlternateContent") {
    for (const branch of Array.from(shape.children)) {
      for (const inner of Array.from(branch.children)) {
        const id = ownTagId(inner);
        if (id) return id;
      }
    }
    return undefined;
  }
  const nv = Array.from(shape.children).find((c) => c.localName.startsWith("nv"));
  const tags = kid(kid(kid(nv, "nvPr"), "custDataLst"), "tags");
  return tags ? Array.from(tags.attributes).find((a) => a.name.endsWith(":id"))?.value : undefined;
}

/** Attribute values used anywhere in a part (which include the relationship ids it refers to). */
function usedIds(doc: Document): Set<string> {
  const ids = new Set<string>();
  for (const el of Array.from(doc.getElementsByTagName("*"))) {
    for (const a of Array.from(el.attributes)) ids.add(a.value);
  }
  return ids;
}

export async function stripToCaptured(pptx: Uint8Array): Promise<Uint8Array> {
  const zip = openZip(pptx);
  const slide = slidePart(zip);
  if (!slide) throw new Error("The exported file has no slide in it.");
  const slideRelsPart = relsPartOf(slide);
  const slideRels = readRels(await zip.text(slideRelsPart), slide);

  // Which tag parts mark captured shapes.
  const capturedTags = new Set<string>();
  for (const r of slideRels.rels) {
    if (r.external || !r.target.includes("/tags/")) continue;
    const text = await zip.text(r.target);
    if (text && Array.from(parseXml(text).getElementsByTagName("*")).some((e) => e.localName === "tag" && e.getAttribute("name")?.toUpperCase() === CAPTURE_TAG)) {
      capturedTags.add(r.id);
    }
  }

  // Keep only captured shapes in the shape tree, and drop animations (they name removed shapes).
  const slideDoc = parseXml((await zip.text(slide)) ?? "");
  const all = Array.from(slideDoc.getElementsByTagName("*"));
  const tree = all.find((e) => e.localName === "spTree");
  for (const child of Array.from(tree?.children ?? [])) {
    if (TREE_PARTS.has(child.localName)) continue;
    const id = ownTagId(child);
    if (!id || !capturedTags.has(id)) child.remove();
  }
  for (const timing of all.filter((e) => e.localName === "timing" && e.parentElement === slideDoc.documentElement)) timing.remove();

  // The slide keeps its layout and whatever its remaining shapes refer to.
  const used = usedIds(slideDoc);
  for (const r of slideRels.rels) if (!r.type.endsWith("/slideLayout") && !used.has(r.id)) r.el.remove();

  // The package drops its document properties and thumbnail (deck title, authors, slide titles).
  const rootRels = readRels(await zip.text("_rels/.rels"), "");
  for (const r of rootRels.rels) if (!r.type.endsWith("/officeDocument")) r.el.remove();

  const changed = new Map<string, Uint8Array>([
    [slide, serialize(slideDoc)],
    [slideRelsPart, serialize(slideRels.doc!)],
    ["_rels/.rels", serialize(rootRels.doc!)],
  ]);

  // The presentation drops its comment authors, printer settings and document-management metadata.
  const presentation = rootRels.rels.find((r) => r.type.endsWith("/officeDocument"))?.target;
  if (presentation) {
    const presRels = readRels(await zip.text(relsPartOf(presentation)), presentation);
    for (const r of presRels.rels) if (PEOPLE_AND_METADATA.some((t) => r.type.endsWith(t))) r.el.remove();
    if (presRels.doc) changed.set(relsPartOf(presentation), serialize(presRels.doc));
  }

  // Keep every part still reachable from the package root; everything else goes.
  const relsText = async (part: string) => (changed.has(part) ? new TextDecoder().decode(changed.get(part)) : await zip.text(part));
  const reachable = new Set<string>();
  const queue = [""];
  while (queue.length) {
    const part = queue.shift()!;
    const { rels } = readRels(await relsText(part ? relsPartOf(part) : "_rels/.rels"), part);
    for (const r of rels) {
      if (r.external || reachable.has(r.target)) continue;
      reachable.add(r.target);
      queue.push(r.target);
    }
  }
  const keep = (name: string) => {
    if (name === "[Content_Types].xml" || name === "_rels/.rels") return true;
    if (name.endsWith(".rels")) return reachable.has(sourceOf(name));
    return reachable.has(name);
  };

  // Content types must not list parts that are gone.
  const typesDoc = parseXml((await zip.text("[Content_Types].xml")) ?? "");
  for (const o of Array.from(typesDoc.getElementsByTagName("Override"))) {
    if (!keep((o.getAttribute("PartName") ?? "").replace(/^\//, ""))) o.remove();
  }
  changed.set("[Content_Types].xml", serialize(typesDoc));

  const entries: RawEntry[] = zip.names.filter(keep).map((name) => (changed.has(name) ? storedEntry(name, changed.get(name)!) : zip.raw(name)!));
  return writeZip(entries);
}
