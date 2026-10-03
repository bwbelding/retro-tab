// Small per-Mac preferences, stored next to the rest of Retro's data.

import type { Reference } from "./layout";
import { kv } from "./idb";

const MATCH_REF = "pref.matchReference";

export const prefs = {
  async matchReference(): Promise<Reference> {
    try {
      return (await kv.get<Reference>(MATCH_REF)) ?? "first";
    } catch {
      return "first";
    }
  },
  setMatchReference: (ref: Reference) => kv.set(MATCH_REF, ref),
};
