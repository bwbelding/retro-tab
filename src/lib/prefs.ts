// Small per-Mac preferences, stored next to the rest of Retro's data.

import { RULES, type Rule } from "./deckCheck";
import type { Reference } from "./layout";
import type { Anchor, Placement } from "./photos";
import { kv } from "./idb";

const MATCH_REF = "pref.matchReference";
const PHOTO = "pref.photo";
// The rules switched off (not on), so rules added later start switched on.
const CHECK_RULES_OFF = "pref.checkRulesOff";
const FY_START = "pref.fyStart";

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
  /** The month (1–12) the fiscal year starts, for roadmaps. */
  fyStart: () => read<number>(FY_START, 1),
  setFyStart: (month: number) => kv.set(FY_START, month),
  /** Which Deck Check rules run. */
  checkRules: async () => {
    const off = await read<Rule[]>(CHECK_RULES_OFF, []);
    return RULES.map((r) => r.rule).filter((r) => !off.includes(r));
  },
  setCheckRules: (rules: Rule[]) =>
    kv.set(
      CHECK_RULES_OFF,
      RULES.map((r) => r.rule).filter((r) => !rules.includes(r)),
    ),
};
