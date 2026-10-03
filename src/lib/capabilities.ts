// Which Office API versions this PowerPoint supports, and what each one unlocks in Retro.
// Features check these at runtime and show a "needs PowerPoint 16.xx" note instead of failing.

export interface Requirement {
  set: string;
  version: string;
  label: string;
  unlocks: string;
  /** Minimum PowerPoint for Mac version, where Microsoft publishes one. */
  macMin?: string;
}

export interface RequirementResult extends Requirement {
  supported: boolean;
}

export const REQUIREMENTS: Requirement[] = [
  { set: "AddinCommands", version: "1.1", label: "Custom ribbon tabs", unlocks: "Retro Build and Retro Polish tabs" },
  { set: "PowerPointApi", version: "1.3", label: "Tags on slides and shapes", unlocks: "Remove All notes, section tracker" },
  { set: "PowerPointApi", version: "1.4", label: "Create and style shapes", unlocks: "Stickies, stamps, brand fills" },
  { set: "PowerPointApi", version: "1.5", label: "Read and change the selection", unlocks: "Size & Position tools, Deck Check “Go to”", macMin: "16.64" },
  { set: "PowerPointApi", version: "1.8", label: "Groups, tables, slide export", unlocks: "Shape library groups and slide snippets", macMin: "16.96" },
  { set: "PowerPointApi", version: "1.10", label: "Alt text and slide backgrounds", unlocks: "Smart Elements, Deck Check alt-text rule, Photos → Background", macMin: "16.105" },
  { set: "ImageCoercion", version: "1.1", label: "Insert pictures", unlocks: "Photos" },
  { set: "ImageCoercion", version: "1.2", label: "Insert SVG images", unlocks: "SVG items in the shape library" },
  { set: "SharedRuntime", version: "1.1", label: "Shared runtime", unlocks: "One-click ribbon buttons without opening the pane" },
];

export type IsSetSupported = (set: string, version: string) => boolean;

export function probeRequirements(isSetSupported: IsSetSupported): RequirementResult[] {
  return REQUIREMENTS.map((r) => {
    try {
      return { ...r, supported: isSetSupported(r.set, r.version) };
    } catch {
      return { ...r, supported: false };
    }
  });
}

/** Office's own check, or "nothing supported" when the page is opened outside PowerPoint. */
export function officeIsSetSupported(): IsSetSupported {
  return (set, version) =>
    typeof Office !== "undefined" && !!Office.context?.requirements?.isSetSupported(set, version);
}
