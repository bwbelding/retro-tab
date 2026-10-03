// Renders every icon in manifests/icons.mjs to dist/assets/icons/<name>-<size>.png.
import { mkdirSync, writeFileSync } from "node:fs";
import { Resvg } from "@resvg/resvg-js";
import { ICONS, RENDER_SIZES, svgFor } from "../manifests/icons.mjs";

const out = new URL("../dist/assets/icons/", import.meta.url);
mkdirSync(out, { recursive: true });
for (const name of Object.keys(ICONS)) {
  for (const size of RENDER_SIZES) {
    const png = new Resvg(svgFor(name), { fitTo: { mode: "width", value: size } }).render().asPng();
    writeFileSync(new URL(`${name}-${size}.png`, out), png);
  }
}
console.log(`rendered ${Object.keys(ICONS).length} icons x ${RENDER_SIZES.length} sizes`);
