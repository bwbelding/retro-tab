// A short log of recent Retro actions (ribbon buttons and pane buttons), kept so a failure that
// shows nothing on screen still turns up in Diagnostics and its copied report. Stored in IndexedDB,
// which both tabs share, so Diagnostics on either tab sees actions from both.

import { kv } from "./idb";

export interface ActivityEntry {
  at: string;
  action: string;
  status: "started" | "ok" | "error";
  ms?: number;
  detail?: string;
}

const KEY = "diag.activity";
const MAX = 25;

/** Everything useful in an error, including Office.js debug info (where in a batch it failed). */
export function describeError(e: unknown): string {
  if (!(e instanceof Error)) return String(e);
  const parts = [e.message];
  const o = e as Error & { code?: string; debugInfo?: { errorLocation?: string; statement?: string; code?: string } };
  if (o.code && o.code !== e.message) parts.push(`code ${o.code}`);
  if (o.debugInfo?.errorLocation) parts.push(`at ${o.debugInfo.errorLocation}`);
  if (o.debugInfo?.statement) parts.push(`statement ${o.debugInfo.statement}`);
  return parts.join(" · ");
}

/** Add an entry, keeping the newest MAX. Pure, for tests. */
export function appendEntry(log: ActivityEntry[], entry: ActivityEntry): ActivityEntry[] {
  return [...log, entry].slice(-MAX);
}

export function formatActivity(log: ActivityEntry[]): string[] {
  return log.map((e) => {
    const time = e.at.slice(11, 19);
    const ms = e.ms !== undefined ? ` (${e.ms} ms)` : "";
    return `${time} ${e.status.toUpperCase().padEnd(7)} ${e.action}${ms}${e.detail ? ` — ${e.detail}` : ""}`;
  });
}

async function add(entry: ActivityEntry) {
  try {
    await kv.set(KEY, appendEntry((await kv.get<ActivityEntry[]>(KEY)) ?? [], entry));
  } catch {
    // Logging must never break an action.
  }
}

export const activity = {
  read: async (): Promise<ActivityEntry[]> => {
    try {
      return (await kv.get<ActivityEntry[]>(KEY)) ?? [];
    } catch {
      return [];
    }
  },
  /** Run work, logging its start, finish and any error (which is re-thrown). */
  async track<T>(action: string, work: () => Promise<T>): Promise<T> {
    const started = Date.now();
    await add({ at: new Date().toISOString(), action, status: "started" });
    try {
      const result = await work();
      await add({ at: new Date().toISOString(), action, status: "ok", ms: Date.now() - started });
      return result;
    } catch (e) {
      await add({ at: new Date().toISOString(), action, status: "error", ms: Date.now() - started, detail: describeError(e) });
      throw e;
    }
  },
};
