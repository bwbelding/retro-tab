# Retro Tab — PowerPoint for Mac add-in: project & feature plan

## Context
You want one custom PowerPoint ribbon tab that collects your most-used commands, keeps a library of shapes you've built yourself, and adds tools for layout, branding, slide utilities and deck checks. The main constraint is your Mac: your company manages it, you have no admin rights, and you don't know which policies are set. So the plan avoids anything that needs admin rights or local tooling, and the first thing it ships finds out which restrictions you actually have.

The repo (`bwbelding/retro-tab`) is empty apart from a README and LICENSE.

**Decisions made:** Office.js web add-in · repo made public + hosted on GitHub Pages · ribbon includes Arrange, Text, Insert and Format commands · v1 = shape library + layout tools + brand kit + slide utilities + deck quality check.

**Workflow you asked for:** (1) this plan → (2) UI design in Claude Design → (3) code, delivered phase by phase.

---

## Working around the managed Mac

| Restriction | Workaround |
|---|---|
| No admin, can't install Node/certificates | Nothing gets built or run on your Mac. The code is built in GitHub Actions and published to GitHub Pages over HTTPS, so there's no localhost server and no dev certificates. |
| Installing the add-in | You copy the two manifest files into `~/Library/Containers/com.microsoft.Powerpoint/Data/Documents/wef` using Finder (⌘⇧G to go to the folder), then restart PowerPoint. This is your own user folder, so no admin is needed. |
| IT has disabled add-ins (`OfficeWebAddinDisableOMEXCatalog`, which also blocks sideloading) | Phase 1 is a tiny test build that tells us right away. Fallbacks: (a) ask IT to deploy the same manifest through Microsoft 365 *Integrated Apps* (a standard request with low effort for them); (b) PowerPoint on the web → Add-ins → Upload My Add-in; (c) a small VBA `.ppam` as a last resort. |
| Company proxy blocks `*.github.io` | The test build checks this. Fallback: move hosting to Cloudflare Pages (no code changes, only the URLs in the manifest). |
| PowerPoint version lags behind (managed update channel) | Every feature checks `Office.context.requirements.isSetSupported('PowerPointApi', x)`. Unsupported features appear greyed out with a "needs PowerPoint 16.xx" message instead of failing. Reference: API 1.8 needs Mac version 16.96 or later, 1.9 needs 16.100, 1.10 needs 16.105. |
| Your saved data is lost when the Office cache is cleared | The library and brand kit are stored in IndexedDB, and **Export/Import backup (.json)** is built in. You get a reminder if you haven't backed up in 30 days. |
| Sign-in or Graph access may be blocked | No Microsoft account sign-in, no Azure app registration, no server. Everything runs on your Mac. |

---

## Architecture
- **Manifests:** XML add-in-only manifests. Mac sideloading uses XML; the newer unified JSON manifest isn't reliably supported there. There are two (one per tab, see Ribbon below), each with its own add-in ID. Both are generated from one template by a small build script, so the URLs and version stay in step. Each contains:
  - `OfficeControl` elements: PowerPoint's own buttons placed on your tab (this is supported on Mac).
  - Custom buttons: either run a function directly (`ExecuteFunction`) or open a task pane (`ShowTaskpane`).
- **Stack:** TypeScript + React + Fluent UI v9, so it looks native to Office. Built with Vite. Unit tests with Vitest. ESLint + `tsc`.
- **Code layout**
  - `manifests/ribbon.mjs`: the single definition of both tabs (groups, buttons, built-in controls). `scripts/build-manifests.mjs` generates `manifest-build.xml` and `manifest-polish.xml` from it.
  - `manifests/icons.mjs` plus `scripts/build-icons.mjs`: ribbon icon artwork, rendered to PNG at 16/32/64/80 px
  - `manifests/office-control-ids.mjs`: the built-in control IDs Retro uses, checked against Microsoft's published list
  - `src/taskpane/`: one React app with a section bar and a view per section: Shapes, Photos, Layout, Brand, Tools, Check, More (Diagnostics)
  - `src/commands/commands.ts`: handlers for buttons that run without the pane (from Phase 2)
  - `src/lib/`: `capabilities.ts` (API version checks), `diagnostics.ts`, `idb.ts` (IndexedDB storage); later `layout.ts`, `shapeRecipe.ts`, `photos.ts`, `quality.ts`, all unit tested
  - `install.html`: the install page served next to the add-in, with download buttons for both manifests
  - `.github/workflows/ci.yml`: lint → typecheck → test → build → Microsoft manifest validation on every PR; deploy to GitHub Pages on `main`

## Ribbon: two tabs grouped by workflow
**Office limits:** each add-in can have only **one** custom tab, and each group holds at most 6 controls (a dropdown menu counts as one). Two tabs therefore means **two manifests** for the same hosted code, both installed with one copy into the `wef` folder. Because both load from the same site, they **share storage**: one shape library, one brand kit. **Rule:** a feature never gets cut for lack of space. It moves to the other tab, into a dropdown menu, or into the task pane.

Groups mix PowerPoint's built-in buttons (B) with Retro's own (R), organised by workflow rather than by where a button comes from.

**Retro Build: make, note and arrange content**
1. **Notes:** Sticky Note (R) · Stamp ▾ (R) · Remove All (R)
2. **Shapes:** Shapes gallery (B) · My Library (R) · Save to Library (R) · Recent ▾ (R) · Icons (B)
3. **Insert:** Photos (R) · Text Box (B) · Table (B) · Picture (B)
4. **Text:** Font Size Up/Down, Bold, Bullets, Line Spacing, Font Color (all B, icon-only)
5. **Arrange:** Align & Distribute ▾, Group ▾, Rotate ▾, Selection Pane, Bring to Front, Send to Back (all B)
6. **Size & Position:** Match Width · Match Height · Match Size · Swap Positions · Gap Distribute… · Exact Size… (all R)

**Retro Polish: style and review it**
1. **Format:** Format Painter, Shape Fill, Shape Outline, Crop (all B)
2. **Brand:** Brand Kit · Fill 1 / 2 / 3 · Apply Fonts (all R)
3. **Review:** Deck Check · Tracker ▾ (R)
4. **Retro:** Diagnostics · Backup ▾ (R)

Built-in control IDs come from Microsoft's published list (OfficeDev/office-control-ids). "Text Box Options" isn't on that list, so **Font Color** takes its slot. Phase 1 confirms on your Mac which built-in buttons actually render; any that don't get a Retro-built replacement.

## Features

### Shape library (the hardest part technically)
- **Save:** select shapes, click *Save to Library*, enter a name and category. A preview thumbnail is stored with it.
- **Primary save method, the "recipe":** the add-in reads the shape's type, geometry, size, fill, line, text and font, and for groups, all of their children (group APIs need API 1.8). It saves that as JSON and rebuilds the shape exactly on insert. This covers styled shapes, callouts, badges, labels, lines and groups.
- **Fallback, "slide snippet":** for things the API can't read (freeforms, SmartArt, charts), the add-in exports the slide with `slide.exportAsBase64()` (API 1.8) and puts it back with `insertSlidesFromBase64`. That inserts a ready-made slide; Office.js has no way to move a shape between slides.
- **Import SVG/PNG** files as library items. They're inserted with `setSelectedDataAsync`. SVGs can then be turned into editable shapes with *Convert to Shape*.
- **Library pane:** search, categories, favourites, drag to reorder, rename or delete, export/import.
- **Phase 3 starts with a short test on your actual PowerPoint version** to confirm which shape properties can be read back. The recipe-versus-snippet split above is adjusted to whatever that test shows.

### Photos (licensed photo picker)
- **Source:** one folder you choose: local, or OneDrive / Box through their Finder-synced folders (no sign-in or IT approval). Subfolders become filters. **Change** switches folders.
- **Scale (2,000+ photos):** Retro stores a small preview of each photo in IndexedDB (~60 MB for 2,500) and reads the full-size file only when inserting. Rescans pick up new or changed files.
- **Placement:** Full bleed (cover the slide), Fill box (replace the selected shape at its size and position), Background (API 1.10; older versions get a full bleed sent to back), Insert as-is. Crop from Top / Center / Bottom. Retro crops and downsizes to slide resolution before inserting, so decks stay small.
- **Reconnect:** after PowerPoint restarts, browsing works straight away; inserting full-size photos may need one click on **Reconnect** (re-confirm the same folder in Finder). Phase 4 tests whether this can be avoided.
- **Folder picking:** if PowerPoint's add-in window doesn't allow choosing a folder, the fallback is selecting all photos in the folder (⌘A) in the same Finder dialog. Phase 1 Diagnostics tests this.
- **Licensing:** all photos are cleared for use, so no credit or expiry tracking.

### Layout tools
Match width, height or size (to the first or last shape selected), swap positions, distribute with a fixed gap in points or cm, nudge by an exact amount, set exact X/Y/W/H for several shapes at once. All the geometry is pure functions in `layout.ts`.

### Brand kit
Saved palette (hex values, with a "pick from selected shape" option) and saved heading and body fonts. One-click apply to fill, outline or text. Supports several kits (for example, Company, Client A).

### Slide utilities
- **Sticky Note and DRAFT stamp:** shapes tagged with `retro:sticky` using the tags API, so **Remove All** can strip them before you send a deck out.
- **Tracker/agenda builder:** you assign slides to sections (stored in tags). It generates an agenda slide and a section tracker bar on each slide, highlighting the current section. *Refresh* rebuilds them.

### Deck quality check
Scans every slide and flags: fonts outside the brand kit, fill or text colours outside the palette, objects partly off the slide, empty placeholders, and missing alt text (API 1.10). Each result has a **Go to** button that selects the problem shape. Rules are pure functions over a snapshot of the deck.

### Diagnostics (ships first)
Shows PowerPoint version and platform, a table of supported API versions, whether IndexedDB storage works, and whether the hosting site is reachable. A **Copy report** button helps if you need to send it to IT.

---

## Delivery phases
0. **Design:** done. Canvas "Retro Tab UI" (ribbon tabs, all task panes, light and dark).
1. **Skeleton + pipeline + restriction test:** both tabs (built-in buttons working, Retro buttons open the pane), the shared task pane with Diagnostics, the install page, CI and GitHub Pages deploy. **Checkpoint:** you install it and send me the Diagnostics report, including the folder-picking test.
2. **Notes + Size & Position:** stickies, stamps, Remove All; Match Width/Height/Size, Swap, Gap Distribute, Nudge, Exact Size.
3. **Shape library:** the capability test first, then the full feature, plus Backup/Restore.
4. **Photos.**
5. **Brand kit.**
6. **Section tracker + Deck check.**

Each phase is its own PR with a short manual test checklist for your Mac.

## Verification
- **CI on every PR:** `eslint`, `tsc --noEmit`, `vitest` (layout math, recipe save/rebuild, quality rules and storage export/import, using a mocked `PowerPoint` namespace), `vite build`, and `office-addin-manifest validate` run on both generated manifests. There's also a test that fails if any group has more than 6 controls or either manifest has more than one custom tab.
- **Task pane UI:** rendered here in headless Chromium with a stubbed Office.js. Screenshots go in each PR.
- **On your Mac (can't be automated here):** a checklist per phase. Phase 1's checklist: both tabs appear, their built-in buttons work, Diagnostics opens from both tabs, and you paste the report back to me. Later phases add one checkbox per button or feature.
- **Deploy check:** after merge, the Pages URL serves `taskpane.html` and the icons with HTTP 200.
