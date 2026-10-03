// Built-in PowerPoint control IDs that add-ins may place on a custom tab, from Microsoft's
// published list (github.com/OfficeDev/office-control-ids, PowerPoint folder). Only the IDs
// Retro uses are listed; add new ones here after checking that list.
export const OFFICE_CONTROL_IDS = new Set([
  "ShapesInsertGallery",
  "IconInsertFromFile",
  "TextBoxInsert",
  "TableInsertGallery",
  "FlyoutAnchorInsertPicture",
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
