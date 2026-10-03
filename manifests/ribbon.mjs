// Single source of truth for both ribbon tabs. scripts/build-manifests.mjs turns this into
// one add-in manifest per tab (Office allows only one custom tab per add-in), and
// tests/ribbon.test.ts enforces Office's limits on it.
//
// Control kinds:
//   { office: "Bold" }                          PowerPoint's own button (OfficeControl).
//                                               IDs come from github.com/OfficeDev/office-control-ids.
//   { id, label, tip, icon, view }              Retro button that opens the task pane at `view`.
//   { id, label, tip, icon, menu: [items] }     Retro dropdown; each item opens the pane at its view.

/** Views the task pane can open at (src/taskpane/views.ts must list the same keys). */
export const VIEWS = ["shapes", "photos", "layout", "brand", "tools", "check", "diagnostics"];

/** Max controls per group, enforced by Office. */
export const MAX_CONTROLS_PER_GROUP = 6;

/** @type {Array<{key: string, id: string, name: string, description: string, defaultView: string, groups: any[]}>} */
export const ADDINS = [
  {
    key: "build",
    id: "2cb84a06-4c49-421a-8781-d95ae2ddb2ab",
    name: "Retro Build",
    description: "Make, note and arrange slide content: stickies, shapes, photos, text and layout tools.",
    defaultView: "shapes",
    groups: [
      {
        id: "Notes",
        label: "Notes",
        icon: "sticky",
        controls: [
          { id: "StickyNote", label: "Sticky Note", tip: "Add a sticky note to this slide.", icon: "sticky", view: "tools" },
          {
            id: "Stamp",
            label: "Stamp",
            tip: "Stamp this slide or every slide.",
            icon: "stamp",
            menu: [
              { id: "StampDraft", label: "DRAFT", tip: "Stamp DRAFT on this slide.", icon: "stamp", view: "tools" },
              { id: "StampConfidential", label: "CONFIDENTIAL", tip: "Stamp CONFIDENTIAL on this slide.", icon: "stamp", view: "tools" },
              { id: "StampReview", label: "FOR REVIEW", tip: "Stamp FOR REVIEW on this slide.", icon: "stamp", view: "tools" },
            ],
          },
          { id: "RemoveNotes", label: "Remove All", tip: "Remove every sticky note and stamp Retro added.", icon: "removeNotes", view: "tools" },
        ],
      },
      {
        id: "Shapes",
        label: "Shapes",
        icon: "library",
        controls: [
          { office: "ShapesInsertGallery" },
          { id: "ShapeLibrary", label: "My Library", tip: "Browse and insert your saved shapes.", icon: "library", view: "shapes" },
          { id: "SaveShape", label: "Save to Library", tip: "Save the selected shapes to your library.", icon: "saveShape", view: "shapes" },
          {
            id: "RecentShapes",
            label: "Recent",
            tip: "Shapes you inserted recently.",
            icon: "recent",
            menu: [{ id: "RecentOpen", label: "Show recent shapes", tip: "Open the library at Recent.", icon: "recent", view: "shapes" }],
          },
          { office: "IconInsertFromFile" },
        ],
      },
      {
        id: "Insert",
        label: "Insert",
        icon: "photos",
        controls: [
          { id: "Photos", label: "Photos", tip: "Browse your licensed photo folder.", icon: "photos", view: "photos" },
          { office: "TextBoxInsert" },
          { office: "TableInsertGallery" },
          { office: "FlyoutAnchorInsertPicture" },
        ],
      },
      {
        id: "Text",
        label: "Text",
        icon: "text",
        controls: [
          { office: "FontSizeIncrease" },
          { office: "FontSizeDecrese" }, // sic: Microsoft's published ID is misspelled.
          { office: "Bold" },
          { office: "BulletsGallery" },
          { office: "LineSpacingGalleryPowerPoint" },
          { office: "FontColorPicker" },
        ],
      },
      {
        id: "Arrange",
        label: "Arrange",
        icon: "arrange",
        controls: [
          { office: "ObjectAlignMenu" },
          { office: "ObjectsGroupMenu" },
          { office: "ObjectRotateGallery" },
          { office: "SelectionPane" },
          { office: "ObjectBringToFront" },
          { office: "ObjectSendToBack" },
        ],
      },
      {
        id: "SizePosition",
        label: "Size & Position",
        icon: "matchSize",
        controls: [
          { id: "MatchWidth", label: "Match Width", tip: "Make selected shapes the same width.", icon: "matchWidth", view: "layout" },
          { id: "MatchHeight", label: "Match Height", tip: "Make selected shapes the same height.", icon: "matchHeight", view: "layout" },
          { id: "MatchSize", label: "Match Size", tip: "Make selected shapes the same size.", icon: "matchSize", view: "layout" },
          { id: "Swap", label: "Swap Positions", tip: "Swap the positions of two shapes.", icon: "swap", view: "layout" },
          { id: "GapDistribute", label: "Gap Distribute…", tip: "Space shapes with an exact gap.", icon: "gapDistribute", view: "layout" },
          { id: "ExactSize", label: "Exact Size…", tip: "Set exact size and position.", icon: "exactSize", view: "layout" },
        ],
      },
    ],
  },
  {
    key: "polish",
    id: "f75ca145-5a6d-42b3-9ca6-0c150b1b94a6",
    name: "Retro Polish",
    description: "Style and review slides: formatting, brand kit, deck check and Retro diagnostics.",
    defaultView: "brand",
    groups: [
      {
        id: "Format",
        label: "Format",
        icon: "format",
        controls: [
          { office: "FormatPainter" },
          { office: "ShapeFillColorPicker" },
          { office: "ShapeOutlineColorPicker" },
          { office: "PictureCropTools" },
        ],
      },
      {
        id: "Brand",
        label: "Brand",
        icon: "brandKit",
        controls: [
          { id: "BrandKit", label: "Brand Kit", tip: "Your brand colors and fonts.", icon: "brandKit", view: "brand" },
          { id: "Fill1", label: "Fill 1", tip: "Fill the selection with brand color 1.", icon: "fill1", view: "brand" },
          { id: "Fill2", label: "Fill 2", tip: "Fill the selection with brand color 2.", icon: "fill2", view: "brand" },
          { id: "Fill3", label: "Fill 3", tip: "Fill the selection with brand color 3.", icon: "fill3", view: "brand" },
          { id: "ApplyFonts", label: "Apply Fonts", tip: "Apply brand heading and body fonts.", icon: "applyFonts", view: "brand" },
        ],
      },
      {
        id: "Review",
        label: "Review",
        icon: "deckCheck",
        controls: [
          { id: "DeckCheck", label: "Deck Check", tip: "Find off-brand fonts and colors, off-slide objects and empty placeholders.", icon: "deckCheck", view: "check" },
          {
            id: "Tracker",
            label: "Tracker",
            tip: "Section tracker and agenda slide.",
            icon: "tracker",
            menu: [{ id: "TrackerOpen", label: "Set up tracker…", tip: "Define sections and apply the tracker.", icon: "tracker", view: "tools" }],
          },
        ],
      },
      {
        id: "Retro",
        label: "Retro",
        icon: "diagnostics",
        controls: [
          { id: "Diagnostics", label: "Diagnostics", tip: "See what your Mac and company settings allow.", icon: "diagnostics", view: "diagnostics" },
          {
            id: "Backup",
            label: "Backup",
            tip: "Back up or restore your library and kits.",
            icon: "backup",
            menu: [{ id: "BackupOpen", label: "Back up or restore…", tip: "Export or import a backup file.", icon: "backup", view: "diagnostics" }],
          },
        ],
      },
    ],
  },
];
