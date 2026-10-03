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
  - `manifests/template.xml` plus `scripts/build-manifests.ts`: produce `manifest.commands.xml` and `manifest.studio.xml`, both pointing at GitHub Pages URLs
  - `src/commands/commands.ts`: handlers for buttons that run without a task pane (match size, swap, sticky note, …)
  - `src/taskpane/`: one React app with a view per pane: Shapes, Layout, Brand, Utilities, Check, Diagnostics
  - `src/lib/ppt.ts`: thin wrappers around `PowerPoint.run` (getSelectedShapes, add shape, tags, select)
  - `src/lib/layout.ts`: pure geometry math (unit tested)
  - `src/lib/shapeRecipe.ts`: captures a shape to JSON and rebuilds it (unit tested)
  - `src/lib/storage.ts`: IndexedDB plus JSON export/import
  - `src/lib/quality.ts`: deck-check rules (pure, unit tested)
  - `src/lib/capabilities.ts`: checks which API versions are supported and gates features
  - `assets/`: icons at 16/32/80 px (the ribbon needs PNGs)
  - `.github/workflows/deploy.yml`: lint → test → build → validate manifest → deploy to Pages
  - `docs/INSTALL.md`: step-by-step install with screenshots, no Terminal needed

## Ribbon: two tabs, nothing dropped (exact split finalised in Design)
**Office limits:** each add-in can have only **one** custom tab, and each group holds at most 6 controls (a dropdown menu counts as one control and can hold many items). Two tabs therefore means **two manifests** (`manifest.commands.xml` and `manifest.studio.xml`) for the same hosted code. Both are installed with the same copy into the `wef` folder. Because both load from the same site, they **share storage**: one shape library, one brand kit. **Rule:** a feature never gets cut for lack of space. It moves to the other tab, into a dropdown menu, or into a task pane.

**Tab 1, "Retro" (everyday commands)**
1. **Arrange** (built-in): Align ▾, Distribute H/V, Group/Ungroup, Bring to Front/Send to Back, Rotate ▾, Selection Pane
2. **Layout+** (custom): Match Width · Match Height · Match Size · Swap Positions · Distribute with Gap… · Exact Size/Position… (opens the pane)
3. **Text** (built-in): Font Size +/−, Bold, Bullets, Line Spacing, Text Box options
4. **Insert** (built-in): Shapes, Text Box, Table, Picture, Icons
5. **Format** (built-in): Format Painter, Shape Fill, Shape Outline, Crop

**Tab 2, "Retro Studio" (your own tools)**
6. **My Shapes** (custom): Open Library · Save Selection to Library · Recent ▾
7. **Brand** (custom): Brand Kit pane · Quick-apply Fill 1/2/3 · Apply Fonts
8. **Slide Tools** (custom): Sticky Note · DRAFT Stamp · Remove All Stickies · Tracker ▾
9. **Check** (custom): Deck Check · Diagnostics · Backup ▾

Control IDs (from the OfficeDev/office-control-ids list) get checked on your Mac in Phase 2. If a built-in control doesn't render on Mac, its group falls back to a version I write in Office.js; align and distribute are easy to rebuild that way.

## Features

### Shape library (the hardest part technically)
- **Save:** select shapes, click *Save to Library*, enter a name and category. A preview thumbnail is stored with it.
- **Primary save method, the "recipe":** the add-in reads the shape's type, geometry, size, fill, line, text and font, and for groups, all of their children (group APIs need API 1.8). It saves that as JSON and rebuilds the shape exactly on insert. This covers styled shapes, callouts, badges, labels, lines and groups.
- **Fallback, "slide snippet":** for things the API can't read (freeforms, SmartArt, charts), the add-in exports the slide with `slide.exportAsBase64()` (API 1.8) and puts it back with `insertSlidesFromBase64`. That inserts a ready-made slide; Office.js has no way to move a shape between slides.
- **Import SVG/PNG** files as library items. They're inserted with `setSelectedDataAsync`. SVGs can then be turned into editable shapes with *Convert to Shape*.
- **Library pane:** search, categories, favourites, drag to reorder, rename or delete, export/import.
- **Phase 3 starts with a short test on your actual PowerPoint version** to confirm which shape properties can be read back. The recipe-versus-snippet split above is adjusted to whatever that test shows.

### Layout tools
Match width, height or size (to the first or last shape selected), swap positions, distribute with a fixed gap in points or cm, nudge by an exact amount, set exact X/Y/W/H for several shapes at once. All the geometry is pure functions in `layout.ts`.

### Brand kit
Saved palette (hex values, with a "pick from selected shape" option) and saved heading and body fonts. One-click apply to fill, outline or text. Supports several kits (for example, Company, Client A).

### Slide utilities
- **Sticky Note and DRAFT stamp:** shapes tagged with `retro:sticky` using the tags API, so **Remove All Stickies** can strip them before you send a deck out.
- **Tracker/agenda builder:** you assign slides to sections (stored in tags). It generates an agenda slide and a section tracker bar on each slide, highlighting the current section. *Refresh* rebuilds them.

### Deck quality check
Scans every slide and flags: fonts outside the brand kit, fill or text colours outside the palette, objects partly off the slide, empty placeholders, and missing alt text (API 1.10). Each result has a **Go to** button that selects the problem shape. Rules are pure functions over a snapshot of the deck.

### Diagnostics (ships first)
Shows PowerPoint version and platform, a table of supported API versions, whether IndexedDB storage works, and whether the hosting site is reachable. A **Copy report** button helps if you need to send it to IT.

---

## Delivery phases
0. **Design (next step after this plan):** in Claude Design, mock up the ribbon tab and each task pane in light and dark mode. You approve before any code is written.
1. **Skeleton + pipeline + restriction test:** make the repo public, set up the GitHub Actions deploy to Pages, and ship both manifests (two tabs, mostly placeholder buttons) plus the Diagnostics pane. This also checks that two sideloaded add-ins and shared storage work under your company's policies. **Checkpoint:** you install it and send me the Diagnostics report. This tells us which restrictions really apply.
2. **Ribbon commands + Layout+.**
3. **Shape library:** the capability test first, then the full feature.
4. **Brand kit.**
5. **Slide utilities.**
6. **Deck quality check.**

Each phase is its own PR with a short manual test checklist for your Mac.

## Verification
- **CI on every PR:** `eslint`, `tsc --noEmit`, `vitest` (layout math, recipe save/rebuild, quality rules and storage export/import, using a mocked `PowerPoint` namespace), `vite build`, and `office-addin-manifest validate` run on both generated manifests. There's also a test that fails if any group has more than 6 controls or either manifest has more than one custom tab.
- **Task pane UI:** rendered here in headless Chromium with a stubbed Office.js. Screenshots go in each PR.
- **On your Mac (can't be automated here):** a checklist per phase. Phase 1's checklist: the tab appears, Diagnostics opens, and you paste the report back to me. Later phases add one checkbox per button or feature.
- **Deploy check:** after merge, the Pages URL serves `taskpane.html` and the icons with HTTP 200.
