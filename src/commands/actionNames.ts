// Every function a Retro ribbon button can run (the manifest's ExecuteFunction names). Kept free of
// Office.js so the manifest tests can check the ribbon against it.

/** Buttons that only open the pane at a section are named show<View>, e.g. showLayout. */
export const SHOW_ACTIONS = ["showShapes", "showPhotos", "showLayout", "showBrand", "showTools", "showCheck", "showDiagnostics"] as const;

export const DO_ACTIONS = [
  "addSticky",
  "stampDraft",
  "stampConfidential",
  "stampReview",
  "removeNotes",
  "matchWidth",
  "matchHeight",
  "matchSize",
  "swapPositions",
] as const;

export type ActionName = (typeof SHOW_ACTIONS)[number] | (typeof DO_ACTIONS)[number];
export const ACTION_NAMES: readonly ActionName[] = [...SHOW_ACTIONS, ...DO_ACTIONS];

export const showActionFor = (view: string) => `show${view[0].toUpperCase()}${view.slice(1)}`;
