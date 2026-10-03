// Built-in PowerPoint control IDs that add-ins may place on a custom tab, from Microsoft's
// published list (github.com/OfficeDev/office-control-ids, PowerPoint folder). Only the IDs
// Retro uses are listed; add new ones here after checking that list.
export const OFFICE_CONTROL_IDS = new Set([
  "ShapesInsertGallery",
  "TextBoxInsert",
  "TableInsertGallery",
  "PictureInsertFromFilePowerPoint", // Not on the published list; confirmed as "Picture" on PowerPoint for Mac 16.113.
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

// Unconfirmed IDs being tried on PowerPoint for Mac when a published ID shows nothing there. Mac
// renders nothing for an ID it doesn't know, so trying several is harmless. Once one is confirmed,
// move it into OFFICE_CONTROL_IDS and remove the rest. Tried and not shown on Mac 16.113:
// Icons (IconInsertFromFile, IconInsert, IconsInsert); Picture (FlyoutAnchorInsertPicture).
export const TRIAL_CONTROL_IDS = new Set([]);
