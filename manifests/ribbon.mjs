// Single source of truth for both ribbon tabs. scripts/build-manifests.mjs turns this into
// one add-in manifest per tab (Office allows only one custom tab per add-in), and
// tests/ribbon.test.ts enforces Office's limits on it.
//
// Control kinds:
//   { office: "Bold" }                          PowerPoint's own button (OfficeControl).
//                                               IDs come from github.com/OfficeDev/office-control-ids.
//   { office: "X", trial: true }                An unconfirmed ID being tried on Mac (see office-control-ids.mjs).
//   { id, label, tip, icon, run }               Retro button that acts in one click: `run` names a
//                                               function in src/commands/actions.ts.
//   { id, label, tip, icon, view }              Retro button that opens the task pane at `view`.
//   { id, label, tip, icon, menu: [items] }     Retro dropdown; each item has its own run or view.

/** Views the task pane can open at (src/taskpane/views.ts must list the same keys). */
export const VIEWS = ["shapes", "photos", "layout", "brand", "tools", "check", "diagnostics"];

/** Max controls per group, enforced by Office. */
export const MAX_CONTROLS_PER_GROUP = 6;

/** @type {Array<{key: string, id: string, name: string, description: string, groups: any[]}>} */
export const ADDINS = [
  {
    key: "build",
    id: "2cb84a06-4c49-421a-8781-d95ae2ddb2ab",
    name: "Retro Build",
    description: "Make, note and arrange slide content: stickies, shapes, photos, text and layout tools.",
    groups: [
      {
        id: "Notes",
        label: "Notes",
        icon: "sticky",
        controls: [
          { id: "StickyNote", label: "Sticky Note", tip: "Add a sticky note to this slide.", icon: "sticky", run: "addSticky" },
          {
            id: "Stamp",
            label: "Stamp",
            tip: "Stamp this slide or every slide.",
            icon: "stamp",
            menu: [
              { id: "StampDraft", label: "DRAFT", tip: "Stamp DRAFT on this slide.", icon: "stamp", run: "stampDraft" },
              { id: "StampConfidential", label: "CONFIDENTIAL", tip: "Stamp CONFIDENTIAL on this slide.", icon: "stamp", run: "stampConfidential" },
              { id: "StampReview", label: "FOR REVIEW", tip: "Stamp FOR REVIEW on this slide.", icon: "stamp", run: "stampReview" },
            ],
          },
          { id: "RemoveNotes", label: "Remove All", tip: "Remove every sticky note and stamp Retro added, on every slide.", icon: "removeNotes", run: "removeNotes" },
        ],
      },
      {
        id: "Shapes",
        label: "Shapes",
        icon: "library",
        // Retro's buttons first and PowerPoint's last: PowerPoint for Mac picks button sizes itself and
        // shows built-in controls small, so this keeps large buttons on the left of each group.
        controls: [
          { id: "ShapeLibrary", label: "My Library", tip: "Browse and insert your saved shapes.", icon: "library", view: "shapes" },
          { id: "SaveShape", label: "Save to Library", tip: "Save the selected shapes to your library.", icon: "saveShape", view: "shapes" },
          {
            id: "RecentShapes",
            label: "Recent",
            tip: "Shapes you inserted recently.",
            icon: "recent",
            menu: [{ id: "RecentOpen", label: "Show recent shapes", tip: "Open the library at Recent.", icon: "recent", view: "shapes" }],
          },
          {
            id: "Smart",
            label: "Smart",
            tip: "Insert a Smart Element: numbered circles, Harvey balls, traffic lights and more. Edit the selected one in the pane.",
            icon: "smart",
            menu: [
              { id: "SmartNumber", label: "Numbered circle", tip: "Next number on this slide.", icon: "smartNumber", run: "insertNumber" },
              { id: "SmartNumberSet", label: "Numbered circles 1–10", tip: "Ten numbered circles in a row.", icon: "smartNumberSet", run: "insertNumberSet" },
              { id: "SmartHarvey", label: "Harvey ball", tip: "0, 25, 50, 75 or 100%.", icon: "smartHarvey", run: "insertHarvey" },
              { id: "SmartHarveySet", label: "Harvey balls 0–100%", tip: "All five Harvey balls in a row: 0, 25, 50, 75 and 100%.", icon: "smartHarveySet", run: "insertHarveySet" },
              { id: "SmartTraffic", label: "Traffic light", tip: "Red, amber or green dot.", icon: "smartTraffic", run: "insertTraffic" },
              { id: "SmartTraffic3", label: "Traffic lights (3)", tip: "Three lights with one lit.", icon: "smartTraffic3", run: "insertTraffic3" },
              { id: "SmartProgress", label: "Progress bar", tip: "Any percentage.", icon: "smartProgress", run: "insertProgress" },
              { id: "SmartStars", label: "Star rating", tip: "0 to 5 stars.", icon: "smartStars", run: "insertStars" },
              { id: "SmartCheckbox", label: "Checkbox", tip: "Ticked or empty.", icon: "smartCheckbox", run: "insertCheckbox" },
              { id: "SmartArrow", label: "Arrow", tip: "Up, down, left or right.", icon: "smartArrow", run: "insertArrow" },
              { id: "SmartTrend", label: "Trend", tip: "Up, slightly up, flat, slightly down or down.", icon: "smartTrend", run: "insertTrend" },
              { id: "SmartRenumber", label: "Renumber circles", tip: "Number this slide's circles 1, 2, 3… in reading order.", icon: "renumber", run: "renumberCircles" },
              { id: "SmartEdit", label: "Edit selected…", tip: "Change the selected element's value in the pane.", icon: "smart", view: "tools" },
            ],
          },
          { office: "ShapesInsertGallery" },
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
          { office: "PictureInsertFromFilePowerPoint" },
        ],
      },
      {
        id: "Text",
        label: "Text",
        icon: "text",
        controls: [
          { office: "FontSizeIncrease" },
          // Microsoft's published list spells this "FontSizeDecrese", but that ID showed nothing on
          // PowerPoint for Mac 16.113. "FontSizeDecrease" is the standard Office ID.
          { office: "FontSizeDecrease" },
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
          { id: "MatchWidth", label: "Match Width", tip: "Make selected shapes the same width.", icon: "matchWidth", run: "matchWidth" },
          { id: "MatchHeight", label: "Match Height", tip: "Make selected shapes the same height.", icon: "matchHeight", run: "matchHeight" },
          { id: "MatchSize", label: "Match Size", tip: "Make selected shapes the same size.", icon: "matchSize", run: "matchSize" },
          { id: "Swap", label: "Swap Positions", tip: "Swap the positions of two shapes.", icon: "swap", run: "swapPositions" },
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
          { id: "ApplyFonts", label: "Apply Fonts", tip: "Apply brand heading and body fonts.", icon: "applyFonts", view: "brand" },
          { id: "Fill1", label: "Fill 1", tip: "Fill the selection with brand color 1.", icon: "fill1", view: "brand" },
          { id: "Fill2", label: "Fill 2", tip: "Fill the selection with brand color 2.", icon: "fill2", view: "brand" },
          { id: "Fill3", label: "Fill 3", tip: "Fill the selection with brand color 3.", icon: "fill3", view: "brand" },
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
