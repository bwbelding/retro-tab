import { useCallback, useEffect, useState } from "react";
import { readSmartSelection, type SmartSelection } from "../lib/smart";

/** The selected Smart Element (if exactly one is selected), kept current as the selection changes. */
export function useSmartSelection(): { selected?: SmartSelection; refresh: () => void } {
  const [selected, setSelected] = useState<SmartSelection | undefined>();

  const refresh = useCallback(() => {
    if (typeof PowerPoint === "undefined") return;
    readSmartSelection()
      .then(setSelected)
      .catch(() => setSelected(undefined));
  }, []);

  useEffect(() => {
    // Read once now; the handler keeps it current.
    refresh();
    const doc = typeof Office !== "undefined" ? Office.context?.document : undefined;
    if (!doc?.addHandlerAsync) return;
    const handler = () => refresh();
    doc.addHandlerAsync(Office.EventType.DocumentSelectionChanged, handler);
    return () => doc.removeHandlerAsync(Office.EventType.DocumentSelectionChanged, { handler });
  }, [refresh]);

  return { selected, refresh };
}
