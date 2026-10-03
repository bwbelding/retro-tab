import { describe, expect, it } from "vitest";
import { formatReport, interpretFolderPick, summarize, type Check } from "../src/lib/diagnostics";
import { probeRequirements } from "../src/lib/capabilities";

const file = (name: string, webkitRelativePath = "", type = "") => ({ name, webkitRelativePath, type });

describe("interpretFolderPick", () => {
  it("reports when nothing was picked", () => {
    expect(interpretFolderPick([]).status).toBe("warn");
  });

  it("recognises a picked folder and counts images", () => {
    const r = interpretFolderPick([
      file("a.jpg", "Licensed Photos/a.jpg", "image/jpeg"),
      file("b.HEIC", "Licensed Photos/People/b.HEIC"),
      file("notes.txt", "Licensed Photos/notes.txt", "text/plain"),
    ]);
    expect(r.status).toBe("ok");
    expect(r.detail).toContain("“Licensed Photos” has 3 files, 2 of them images");
  });

  it("warns when only individual files come back", () => {
    const r = interpretFolderPick([file("a.jpg", "", "image/jpeg")]);
    expect(r.status).toBe("warn");
    expect(r.detail).toContain("Only individual files");
  });
});

describe("summarize", () => {
  const ok: Check = { id: "a", label: "A", status: "ok", detail: "" };
  const reqs = probeRequirements(() => true);

  it("is ok when everything passes", () => {
    expect(summarize([ok], reqs).status).toBe("ok");
  });

  it("warns about unavailable PowerPoint features", () => {
    const partial = probeRequirements((set, v) => !(set === "PowerPointApi" && v === "1.10"));
    expect(summarize([ok], partial)).toEqual({ status: "warn", text: "Retro works, with 1 PowerPoint feature unavailable." });
  });

  it("fails when any check fails, and waits on pending checks", () => {
    expect(summarize([ok, { ...ok, status: "fail" }], reqs).status).toBe("fail");
    expect(summarize([{ ...ok, status: "pending" }], reqs).status).toBe("pending");
  });
});

describe("probeRequirements", () => {
  it("treats a throwing check as unsupported", () => {
    const r = probeRequirements(() => {
      throw new Error("no Office");
    });
    expect(r.every((x) => !x.supported)).toBe(true);
  });
});

describe("formatReport", () => {
  it("lists environment, features and checks", () => {
    const text = formatReport(
      { host: "PowerPoint", platform: "Mac", officeVersion: "16.102", tab: "Retro Build", retroVersion: "1.0.0", origin: "https://x", userAgent: "UA" },
      probeRequirements((set, v) => !(set === "PowerPointApi" && v === "1.10")),
      [{ id: "s", label: "Saving on this Mac", status: "ok", detail: "Works." }],
      new Date("2026-10-03T00:00:00Z"),
    );
    expect(text).toContain("PowerPoint: 16.102 (PowerPoint, Mac)");
    expect(text).toContain("MISS PowerPointApi 1.10 — Alt text and slide backgrounds (needs 16.105+)");
    expect(text).toContain("OK   Saving on this Mac: Works.");
  });
});

describe("activity log", () => {
  it("keeps the newest entries and formats them for the report", async () => {
    const { appendEntry, formatActivity, describeError } = await import("../src/lib/activity");
    let log: Parameters<typeof appendEntry>[0] = [];
    for (let i = 0; i < 30; i++) log = appendEntry(log, { at: `2026-10-03T19:00:${String(i).padStart(2, "0")}.000Z`, action: `a${i}`, status: "ok", ms: i });
    expect(log).toHaveLength(25);
    expect(log[0].action).toBe("a5");
    const err = Object.assign(new Error("InvalidArgument"), { code: "InvalidArgument", debugInfo: { errorLocation: "Adjustments.set" } });
    const line = formatActivity([{ at: "2026-10-03T19:00:01.000Z", action: "ribbon insertHarvey", status: "error", ms: 40, detail: describeError(err) }])[0];
    expect(line).toBe("19:00:01 ERROR   ribbon insertHarvey (40 ms) — InvalidArgument · at Adjustments.set");
  });
});
