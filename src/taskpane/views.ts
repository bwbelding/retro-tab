import type { FluentIcon } from "@fluentui/react-icons";
import {
  ColorRegular,
  ImageRegular,
  MoreHorizontalRegular,
  NoteRegular,
  ResizeRegular,
  ShapesRegular,
  ShieldCheckmarkRegular,
} from "@fluentui/react-icons";

// Keys must match VIEWS in manifests/ribbon.mjs (tests/ribbon.test.ts checks this).
export type ViewKey = "shapes" | "photos" | "layout" | "brand" | "tools" | "check" | "diagnostics";

export interface ViewInfo {
  key: ViewKey;
  /** Label in the pane's section bar. */
  nav: string;
  title: string;
  icon: FluentIcon;
  /** What's coming, shown until the view is built. */
  coming?: string;
}

export const VIEWS: ViewInfo[] = [
  { key: "shapes", nav: "Shapes", title: "Shape library", icon: ShapesRegular, coming: "Save your own shapes and insert them in one click. Arrives in Phase 3." },
  { key: "photos", nav: "Photos", title: "Photos", icon: ImageRegular, coming: "Browse your licensed photo folder and place photos as full bleeds, boxes or backgrounds. Arrives in Phase 4." },
  { key: "layout", nav: "Layout", title: "Size & position", icon: ResizeRegular },
  { key: "brand", nav: "Brand", title: "Brand kit", icon: ColorRegular, coming: "Saved brand colors and fonts with one-click apply. Arrives in Phase 5." },
  { key: "tools", nav: "Tools", title: "Slide tools", icon: NoteRegular },
  { key: "check", nav: "Check", title: "Deck check", icon: ShieldCheckmarkRegular, coming: "Find off-brand fonts and colors, objects off the slide and empty placeholders. Arrives in Phase 6." },
  { key: "diagnostics", nav: "More", title: "Diagnostics", icon: MoreHorizontalRegular },
];

export function viewFromUrl(search: string): ViewKey {
  const v = new URLSearchParams(search).get("view");
  return VIEWS.some((x) => x.key === v) ? (v as ViewKey) : "diagnostics";
}

export function tabFromUrl(search: string): "build" | "polish" {
  return new URLSearchParams(search).get("tab") === "polish" ? "polish" : "build";
}
