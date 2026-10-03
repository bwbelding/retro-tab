# Retro Tab

Creating a modern PowerPoint for Mac add-in that includes all of my most used features and some custom ones as well.

Two ribbon tabs, **Retro Build** and **Retro Polish**, built as an Office.js add-in that installs without admin rights. See [docs/PLAN.md](docs/PLAN.md) for the plan and feature set.

## Install

Open the install page at <https://bwbelding.github.io/retro-tab/install.html> and follow the three steps. Nothing is installed on the Mac beyond two small files in PowerPoint's add-in folder.

## Develop

Nothing runs locally on the Mac; GitHub Actions builds, checks and deploys every change.

```sh
npm ci
npm run lint && npm run typecheck && npm test
npm run build      # dist/: task pane, icons, manifest-build.xml, manifest-polish.xml
npm run validate   # Microsoft's manifest validator (needs internet access)
```

The ribbon is defined once in `manifests/ribbon.mjs`; both manifests are generated from it.
