// Built-in PowerPoint control IDs that add-ins may place on a custom tab, from Microsoft's
// published list (github.com/OfficeDev/office-control-ids, PowerPoint folder). Only the IDs
// Retro uses are listed; add new ones here after checking that list.
export const OFFICE_CONTROL_IDS = new Set([
  "ShapesInsertGallery",
  "TextBoxInsert",
  "TableInsertGallery",
  "FontSizeIncrease",
  "FontSizeDecrease", // Published as "FontSizeDecrese", which showed nothing on PowerPoint for Mac.
  "Bold",
  "BulletsGallery",
  "LineSpacingGalleryPowerPoint",
  "FontColorPicker",
  "ObjectAlignMenu",
  "ObjectsGroupMenu",
  "ObjectRotateGallery",
  "SelectionPane",
  "ObjectBringToFront",
  "ObjectSendToBack",
  "FormatPainter",
  "ShapeFillColorPicker",
  "ShapeOutlineColorPicker",
  "PictureCropTools",
]);

// Unconfirmed IDs being tried on PowerPoint for Mac because the published ID showed nothing there
// (Icons: "IconInsertFromFile"; Picture: "FlyoutAnchorInsertPicture"). Mac renders nothing for
// an ID it doesn't know, so trying several is harmless. Once one is confirmed, move it into
// OFFICE_CONTROL_IDS and remove the rest.
export const TRIAL_CONTROL_IDS = new Set([
  "IconInsert",
  "IconsInsert",
  "PictureInsertFromFilePowerPoint",
  "PictureInsertFromFile",
  "PicturesInsertMenu",
]);
