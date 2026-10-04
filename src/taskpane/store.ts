// Pane state shared by the React app and the ribbon actions. With the shared runtime, ribbon
// buttons run in the same page as the pane, so an action can switch the pane's view or show a
// message and then bring the pane up.

import { useSyncExternalStore } from "react";
import type { ViewKey } from "./views";

export interface PaneMessage {
  intent: "info" | "success" | "warning" | "error";
  text: string;
}

interface PaneState {
  view: ViewKey;
  message?: PaneMessage;
  /** Which library the Shapes view shows. */
  library: "shapes" | "slides";
  /** The ribbon asked for the library's save form; cleared when that form closes. */
  saveRequested?: boolean;
  /** What the Check view shows. */
  checkMode: "check" | "find";
}

let state: PaneState = { view: "diagnostics", library: "shapes", checkMode: "check" };
const listeners = new Set<() => void>();

function set(next: Partial<PaneState>) {
  state = { ...state, ...next };
  listeners.forEach((l) => l());
}

export const pane = {
  init(view: ViewKey) {
    state = { view, library: "shapes", checkMode: "check" };
  },
  setCheckMode(checkMode: PaneState["checkMode"]) {
    set({ checkMode });
  },
  setLibrary(library: PaneState["library"], saveRequested = false) {
    set({ library, saveRequested });
  },
  setView(view: ViewKey) {
    set({ view, message: undefined });
  },
  showMessage(message: PaneMessage | undefined) {
    set({ message });
  },
  /** Switch to a view (optionally with a message) and bring the pane up. */
  async open(view: ViewKey, message?: PaneMessage) {
    set({ view, message });
    if (typeof Office === "undefined" || !Office.addin?.showAsTaskpane) return;
    try {
      await Office.addin.showAsTaskpane();
    } catch {
      // The pane couldn't be shown (e.g. PowerPoint is busy); the message waits for the next time it opens.
    }
  },
};

const subscribe = (l: () => void) => {
  listeners.add(l);
  return () => listeners.delete(l);
};

export function usePane(): PaneState {
  return useSyncExternalStore(subscribe, () => state);
}
