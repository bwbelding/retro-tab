// What each Retro ribbon button does. Registered with Office.actions.associate in the shared
// runtime, so buttons act in one click; the pane only opens to show a result or a problem.

import { matchSize, swapPositions, type Dimension } from "../lib/layout";
import { addStamp, addSticky, scanNotes, type StampLabel } from "../lib/notes";
import { applyRects, requireSelection, selectedShapes, UserError } from "../lib/ppt";
import { prefs } from "../lib/prefs";
import { pane } from "../taskpane/store";
import type { ViewKey } from "../taskpane/views";
import type { ActionName } from "./actionNames";

interface Action {
  /** Section the pane opens at if the action needs to report something. */
  view: ViewKey;
  run: () => Promise<void>;
}

const show = (view: ViewKey): Action => ({ view, run: () => pane.open(view) });

const match = (dim: Dimension): Action => ({
  view: "layout",
  run: async () => {
    const ref = await prefs.matchReference();
    await PowerPoint.run(async (context) => {
      const { shapes, rects } = await requireSelection(context, 2, "match their sizes");
      applyRects(shapes, matchSize(rects, dim, ref));
      await context.sync();
    });
  },
});

const stamp = (label: StampLabel): Action => ({
  view: "tools",
  run: async () => {
    await addStamp(label);
  },
});

export const ACTIONS: Record<ActionName, Action> = {
  showShapes: show("shapes"),
  showPhotos: show("photos"),
  showLayout: show("layout"),
  showBrand: show("brand"),
  showTools: show("tools"),
  showCheck: show("check"),
  showDiagnostics: show("diagnostics"),
  addSticky: {
    view: "tools",
    run: async () => {
      const selection = await PowerPoint.run(async (context) => (await selectedShapes(context)).rects);
      await addSticky("Note", "yellow", selection);
    },
  },
  stampDraft: stamp("DRAFT"),
  stampConfidential: stamp("CONFIDENTIAL"),
  stampReview: stamp("FOR REVIEW"),
  removeNotes: {
    view: "tools",
    run: async () => {
      const { notes, slides } = await scanNotes(true);
      await pane.open("tools", {
        intent: notes ? "success" : "info",
        text: notes
          ? `Removed ${notes} note${notes === 1 ? "" : "s"} and stamp${notes === 1 ? "" : "s"} from ${slides} slide${slides === 1 ? "" : "s"}.`
          : "There were no Retro notes or stamps to remove.",
      });
    },
  },
  matchWidth: match("width"),
  matchHeight: match("height"),
  matchSize: match("both"),
  swapPositions: {
    view: "layout",
    run: () =>
      PowerPoint.run(async (context) => {
        const { shapes, rects } = await requireSelection(context, 2, "swap their positions");
        if (shapes.length !== 2) throw new UserError("Select exactly 2 shapes to swap their positions.");
        applyRects(shapes, swapPositions(rects[0], rects[1]));
        await context.sync();
      }),
  },
};

/** Run an action for a ribbon button, reporting problems in the pane. */
export async function runAction(action: Action): Promise<void> {
  try {
    await action.run();
  } catch (e) {
    const text = e instanceof UserError ? e.message : `That didn't work: ${e instanceof Error ? e.message : String(e)}`;
    await pane.open(action.view, { intent: e instanceof UserError ? "warning" : "error", text });
  }
}

export function registerActions(): void {
  for (const [name, action] of Object.entries(ACTIONS)) {
    Office.actions.associate(name, (event: Office.AddinCommands.Event) => {
      void runAction(action).finally(() => event.completed());
    });
  }
}
