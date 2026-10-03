import { useCallback, useEffect, useState } from "react";
import type { Rect } from "../lib/layout";
import { selectedShapes } from "../lib/ppt";

/** The selected shapes' rectangles, kept up to date as the selection changes in PowerPoint. */
export function useSelection(): { rects: Rect[]; refresh: () => void } {
  const [rects, setRects] = useState<Rect[]>([]);

  const refresh = useCallback(() => {
    if (typeof PowerPoint === "undefined") return;
    PowerPoint.run(async (context) => (await selectedShapes(context)).rects)
      .then(setRects)
      .catch(() => setRects([]));
  }, []);

  useEffect(() => {
    // Read the selection once now; the handler below keeps it current.
    refresh();
    const doc = typeof Office !== "undefined" ? Office.context?.document : undefined;
    if (!doc?.addHandlerAsync) return;
    const handler = () => refresh();
    doc.addHandlerAsync(Office.EventType.DocumentSelectionChanged, handler);
    return () => doc.removeHandlerAsync(Office.EventType.DocumentSelectionChanged, { handler });
  }, [refresh]);

  return { rects, refresh };
}
