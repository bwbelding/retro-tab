import { defineConfig } from "vitest/config";
import react from "@vitejs/plugin-react";
import { readFileSync } from "node:fs";
import { fileURLToPath } from "node:url";

const { version } = JSON.parse(readFileSync(new URL("./package.json", import.meta.url), "utf8"));
const page = (name: string) => fileURLToPath(new URL(`./${name}.html`, import.meta.url));

// Relative base so the same build works at https://<user>.github.io/retro-tab/ or any other host.
export default defineConfig({
  base: "./",
  plugins: [react()],
  define: { __RETRO_VERSION__: JSON.stringify(version) },
  build: {
    outDir: "dist",
    emptyOutDir: true,
    rollupOptions: {
      input: { taskpane: page("taskpane"), commands: page("commands"), install: page("install") },
    },
  },
  test: {
    include: ["tests/**/*.test.ts"],
  },
});
