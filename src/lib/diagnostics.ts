// Checks that tell us which of the company's restrictions actually apply on this Mac.
// The pure helpers (interpretFolderPick, summarize, formatReport) are unit tested.

import type { RequirementResult } from "./capabilities";

export type Status = "ok" | "warn" | "fail" | "pending";

export interface Check {
  id: string;
  label: string;
  status: Status;
  detail: string;
}

export interface Environment {
  host: string;
  platform: string;
  officeVersion: string;
  tab: string;
  retroVersion: string;
  origin: string;
  userAgent: string;
}

export interface PickedFile {
  name: string;
  type: string;
  webkitRelativePath: string;
}

const IMAGE_EXT = /\.(jpe?g|png|gif|webp|heic|heif|tiff?|bmp|svg)$/i;

/** Classifies what the Finder dialog returned, so we learn whether folder picking is allowed. */
export function interpretFolderPick(files: PickedFile[]): Check {
  const id = "folder";
  const label = "Photo folder picking";
  if (files.length === 0) {
    return { id, label, status: "warn", detail: "Nothing was picked, or the Finder dialog didn't open. Try again." };
  }
  const images = files.filter((f) => f.type.startsWith("image/") || IMAGE_EXT.test(f.name)).length;
  const fromFolder = files.some((f) => f.webkitRelativePath.includes("/"));
  if (fromFolder) {
    const root = files[0].webkitRelativePath.split("/")[0];
    return { id, label, status: "ok", detail: `Folder picking works: “${root}” has ${files.length} files, ${images} of them images.` };
  }
  return {
    id,
    label,
    status: "warn",
    detail: `Only individual files can be picked (${files.length} picked, ${images} images). Photos will use “select all” in Finder instead of a folder.`,
  };
}

export function summarize(checks: Check[], requirements: RequirementResult[]): { status: Status; text: string } {
  if (checks.some((c) => c.status === "pending")) return { status: "pending", text: "Running checks…" };
  const failed = checks.filter((c) => c.status === "fail").length;
  const warned = checks.filter((c) => c.status === "warn").length;
  const missing = requirements.filter((r) => !r.supported).length;
  if (failed) return { status: "fail", text: `${failed} check${failed > 1 ? "s" : ""} failed. Copy the report and send it to me.` };
  if (warned || missing) {
    const parts = [];
    if (missing) parts.push(`${missing} PowerPoint feature${missing > 1 ? "s" : ""} unavailable`);
    if (warned) parts.push(`${warned} warning${warned > 1 ? "s" : ""}`);
    return { status: "warn", text: `Retro works, with ${parts.join(" and ")}.` };
  }
  return { status: "ok", text: "Everything Retro needs is available." };
}

const MARK: Record<Status, string> = { ok: "OK  ", warn: "WARN", fail: "FAIL", pending: "...." };

export function formatReport(env: Environment, requirements: RequirementResult[], checks: Check[], now = new Date()): string {
  const lines = [
    `Retro diagnostics report — ${now.toISOString()}`,
    "",
    `PowerPoint: ${env.officeVersion} (${env.host}, ${env.platform})`,
    `Retro: v${env.retroVersion}, opened from the ${env.tab} tab`,
    `Loaded from: ${env.origin}`,
    `Browser engine: ${env.userAgent}`,
    "",
    "PowerPoint features:",
    ...requirements.map(
      (r) => `  ${r.supported ? "OK  " : "MISS"} ${r.set} ${r.version} — ${r.label}${!r.supported && r.macMin ? ` (needs ${r.macMin}+)` : ""}`,
    ),
    "",
    "Company setup checks:",
    ...checks.map((c) => `  ${MARK[c.status]} ${c.label}: ${c.detail}`),
  ];
  return lines.join("\n");
}

// ---- Live checks (need a browser / PowerPoint) ----

const HEARTBEAT_KEY = (tab: string) => `diag.seen.${tab}`;

/** Write, read back and delete a value; also report how much space Retro may use. */
export async function checkStorage(kv: {
  get: <T>(k: string) => Promise<T | undefined>;
  set: (k: string, v: unknown) => Promise<void>;
  del: (k: string) => Promise<void>;
}): Promise<Check> {
  const id = "storage";
  const label = "Saving on this Mac";
  try {
    const token = `t${Date.now()}`;
    await kv.set("diag.probe", token);
    const back = await kv.get<string>("diag.probe");
    await kv.del("diag.probe");
    if (back !== token) return { id, label, status: "fail", detail: "Saved data didn't read back correctly." };
    let quota = "";
    if (navigator.storage?.estimate) {
      const { usage = 0, quota: q = 0 } = await navigator.storage.estimate();
      quota = ` Using ${mb(usage)} of ${mb(q)} allowed.`;
    }
    return { id, label, status: "ok", detail: `Works.${quota}` };
  } catch (e) {
    return { id, label, status: "fail", detail: `Blocked: ${errorText(e)}` };
  }
}

/** Each tab records when it last opened, so either tab can show whether storage is shared. */
export async function checkSharedStorage(
  kv: { get: <T>(k: string) => Promise<T | undefined>; set: (k: string, v: unknown) => Promise<void> },
  tab: string,
  otherTab: string,
): Promise<Check> {
  const id = "shared";
  const label = "Both tabs share saved data";
  const otherName = otherTab === "polish" ? "Retro Polish" : "Retro Build";
  try {
    await kv.set(HEARTBEAT_KEY(tab), new Date().toISOString());
    const other = await kv.get<string>(HEARTBEAT_KEY(otherTab));
    if (other) return { id, label, status: "ok", detail: `Yes. The ${otherName} tab last opened Retro at ${new Date(other).toLocaleString()}.` };
    return { id, label, status: "warn", detail: `Not confirmed yet. Open this page from the ${otherName} tab too, then click Run again.` };
  } catch (e) {
    return { id, label, status: "fail", detail: `Couldn't check: ${errorText(e)}` };
  }
}

export async function checkHosting(fetchFn: typeof fetch = fetch): Promise<Check> {
  const id = "hosting";
  const label = "Hosting site reachable";
  const started = performance.now();
  try {
    const res = await fetchFn("version.json", { cache: "no-store" });
    const ms = Math.round(performance.now() - started);
    if (!res.ok) return { id, label, status: "fail", detail: `${location.host} answered ${res.status}.` };
    return { id, label, status: "ok", detail: `${location.host} answered in ${ms} ms.` };
  } catch (e) {
    return { id, label, status: "fail", detail: `Couldn't reach ${location.host}: ${errorText(e)}` };
  }
}

/** A real PowerPoint call, not just a version check. */
export async function checkPowerPoint(): Promise<Check> {
  const id = "powerpoint";
  const label = "PowerPoint responds to Retro";
  if (typeof PowerPoint === "undefined") {
    return { id, label, status: "fail", detail: "Not running inside PowerPoint." };
  }
  try {
    const counts = await PowerPoint.run(async (context) => {
      const slides = context.presentation.slides;
      slides.load("items");
      await context.sync();
      return { slides: slides.items.length };
    });
    return { id, label, status: "ok", detail: `Yes. This presentation has ${counts.slides} slide${counts.slides === 1 ? "" : "s"}.` };
  } catch (e) {
    return { id, label, status: "fail", detail: `PowerPoint refused: ${errorText(e)}` };
  }
}

function mb(bytes: number): string {
  return bytes >= 1024 ** 3 ? `${(bytes / 1024 ** 3).toFixed(1)} GB` : `${Math.round(bytes / 1024 ** 2)} MB`;
}

function errorText(e: unknown): string {
  return e instanceof Error ? e.message : String(e);
}
