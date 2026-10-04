import { useCallback, useEffect, useState } from "react";
import { Button, Text } from "@fluentui/react-components";
import { activity, describeError } from "../../lib/activity";
import { markAppendix, moveAppendix, readAppendix, type AppendixState } from "../../lib/appendix";
import { UserError } from "../../lib/ppt";
import { pane } from "../store";
import { Section, useUi } from "../ui";

export function AppendixSection() {
  const ui = useUi();
  const [state, setState] = useState<AppendixState | undefined>();

  const reload = useCallback(() => {
    if (typeof PowerPoint === "undefined") return;
    readAppendix()
      .then(setState)
      .catch(() => undefined);
  }, []);

  useEffect(() => {
    // Follow the selected slides, so the buttons fit what's selected.
    reload();
    const doc = typeof Office === "undefined" ? undefined : Office.context?.document;
    if (!doc?.addHandlerAsync) return;
    doc.addHandlerAsync(Office.EventType.DocumentSelectionChanged, reload);
    return () => doc.removeHandlerAsync(Office.EventType.DocumentSelectionChanged, { handler: reload });
  }, [reload]);

  const act = async (work: () => Promise<string>) => {
    try {
      pane.showMessage({ intent: "success", text: await activity.track("pane appendix", work) });
    } catch (e) {
      pane.showMessage({ intent: e instanceof UserError ? "warning" : "error", text: e instanceof UserError ? e.message : describeError(e) });
    }
    reload();
  };

  const n = state?.marked.length ?? 0;
  const selected = state?.selected.length ?? 0;
  const selectedMarked = state ? state.selected.filter((i) => state.marked.includes(i)).length : 0;
  const plural = (k: number) => `${k} slide${k === 1 ? "" : "s"}`;

  return (
    <Section title="Appendix">
      <Text size={200}>
        {n === 0 ? "No appendix slides yet." : `${plural(n)} marked as appendix${state?.inPlace ? ", at the end behind the divider." : ", not all at the end yet."}`}
      </Text>
      <div className={ui.grid2}>
        <Button disabled={selected === 0 || selectedMarked === selected} onClick={() => void act(async () => `Marked ${plural(await markAppendix(true))} as appendix.`)}>
          Mark as appendix
        </Button>
        <Button disabled={selectedMarked === 0} onClick={() => void act(async () => `Unmarked ${plural(await markAppendix(false))}.`)}>
          Unmark
        </Button>
      </div>
      <Button appearance="primary" style={{ alignSelf: "flex-start" }} disabled={n === 0 || state?.inPlace} onClick={() => void act(moveAppendix)}>
        Move appendix to end
      </Button>
      <Text size={200} className={ui.muted}>
        Select slides in the thumbnails and mark them. Moving puts them at the end, in their order, behind an “Appendix” divider made from your template&apos;s section header. The tracker skips appendix slides.
      </Text>
    </Section>
  );
}
