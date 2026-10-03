// Small per-Mac preferences, stored next to the rest of Retro's data.

import type { Reference } from "./layout";
import type { Anchor, Placement } from "./photos";
import { kv } from "./idb";

const MATCH_REF = "pref.matchReference";
const PHOTO = "pref.photo";

async function read<T>(key: string, fallback: T): Promise<T> {
  try {
    return (await kv.get<T>(key)) ?? fallback;
  } catch {
    return fallback;
  }
}

export const prefs = {
  async matchReference(): Promise<Reference> {
    try {
      return (await kv.get<Reference>(MATCH_REF)) ?? "first";
    } catch {
      return "first";
    }
  },
  setMatchReference: (ref: Reference) => kv.set(MATCH_REF, ref),
  /** How photos were last placed and cropped. */
  photo: () => read<{ placement: Placement; anchor: Anchor }>(PHOTO, { placement: "full", anchor: "center" }),
  setPhoto: (value: { placement: Placement; anchor: Anchor }) => kv.set(PHOTO, value),
};
