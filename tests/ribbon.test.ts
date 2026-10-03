import { describe, expect, it } from "vitest";
import { XMLParser, XMLValidator } from "fast-xml-parser";
import { ADDINS, MAX_CONTROLS_PER_GROUP, VIEWS as RIBBON_VIEWS } from "../manifests/ribbon.mjs";
import { ICONS } from "../manifests/icons.mjs";
import { OFFICE_CONTROL_IDS } from "../manifests/office-control-ids.mjs";
import { buildManifest, iconsUsed } from "../scripts/build-manifests.mjs";
import { VIEWS } from "../src/taskpane/views";
import pkg from "../package.json";

type Control = { office?: string; id?: string; view?: string; label?: string; tip?: string; menu?: Control[] };

describe("ribbon definition", () => {
  it("defines two add-ins with distinct GUIDs", () => {
    expect(ADDINS).toHaveLength(2);
    const ids = ADDINS.map((a) => a.id);
    for (const id of ids) expect(id).toMatch(/^[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}$/);
    expect(new Set(ids).size).toBe(2);
  });

  it.each(ADDINS)("$name keeps every group within Office's control limit", (addin) => {
    for (const g of addin.groups) expect(g.controls.length, g.label).toBeLessThanOrEqual(MAX_CONTROLS_PER_GROUP);
  });

  it.each(ADDINS)("$name uses only published built-in control IDs", (addin) => {
    const office = addin.groups.flatMap((g) => g.controls as Control[]).filter((c) => c.office);
    for (const c of office) expect(OFFICE_CONTROL_IDS.has(c.office!), c.office).toBe(true);
  });

  it.each(ADDINS)("$name has unique control ids, existing icons and known views", (addin) => {
    const custom = addin.groups.flatMap((g) => g.controls as Control[]).filter((c) => !c.office);
    const all = [...addin.groups.map((g) => g.id), ...custom.flatMap((c) => [c.id, ...(c.menu ?? []).map((m) => m.id)])];
    expect(new Set(all).size).toBe(all.length);
    for (const name of iconsUsed(addin)) expect(ICONS, name).toHaveProperty(name);
    for (const c of custom.flatMap((c) => (c.menu ? c.menu : [c]))) {
      expect(RIBBON_VIEWS, c.id).toContain(c.view);
      expect(c.label!.length).toBeLessThanOrEqual(125);
      expect(c.tip!.length).toBeLessThanOrEqual(250);
    }
  });

  it("has a version Microsoft's validator accepts (1.0 or higher)", () => {
    expect(Number(pkg.version.split(".")[0])).toBeGreaterThanOrEqual(1);
  });

  it("matches the task pane's views", () => {
    expect([...RIBBON_VIEWS].sort()).toEqual(VIEWS.map((v) => v.key).sort());
  });
});

describe.each(ADDINS)("$name manifest", (addin) => {
  const xml = buildManifest(addin, { baseUrl: "https://example.github.io/retro-tab/", version: "1.2.3" });
  const doc = new XMLParser({ ignoreAttributes: false }).parse(xml);

  it("is well-formed XML", () => {
    expect(XMLValidator.validate(xml)).toBe(true);
  });

  it("has exactly one custom tab", () => {
    expect(xml.match(/<CustomTab /g)).toHaveLength(1);
  });

  it("uses resource ids of at most 32 characters", () => {
    const ids = [...xml.matchAll(/(?:resid|bt:\w+ id)="([^"]+)"/g)].map((m) => m[1]);
    expect(ids.length).toBeGreaterThan(0);
    for (const id of ids) expect(id.length, id).toBeLessThanOrEqual(32);
  });

  it("points every URL at the hosting site over HTTPS", () => {
    const urls = [...xml.matchAll(/DefaultValue="(https?:[^"]+)"/g)].map((m) => m[1]);
    for (const u of urls.filter((u) => !u.startsWith("https://github.com"))) {
      expect(u.startsWith("https://example.github.io/retro-tab/"), u).toBe(true);
    }
  });

  it("sets the four-part version and the tab label", () => {
    expect(doc.OfficeApp.Version).toBe("1.2.3.0");
    expect(xml).toContain(`DefaultValue="${addin.name}"`);
  });
});
