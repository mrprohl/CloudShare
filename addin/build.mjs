// Bundles the taskpane and the Smart Alert handlers as classic scripts (IIFE).
// The JavaScript-only runtime of classic Outlook does not support ES modules.
// docker/entrypoint.sh replaces the placeholders %%PUBLIC_BASE_URL%% and %%ENTRA_CLIENT_ID%% at container start.
import { build } from "esbuild";
import { cpSync, mkdirSync, rmSync } from "node:fs";

rmSync("dist", { recursive: true, force: true });
mkdirSync("dist", { recursive: true });

await build({
  entryPoints: { taskpane: "src/taskpane/taskpane.ts", commands: "src/commands/commands.ts" },
  outdir: "dist",
  bundle: true,
  format: "iife",
  target: "es2017",
  minify: true,
  sourcemap: false,
});

for (const file of ["src/taskpane/taskpane.html", "src/taskpane/taskpane.css", "src/commands/commands.html", "manifest.xml"]) {
  cpSync(file, `dist/${file.split("/").pop()}`);
}
cpSync("assets", "dist/assets", { recursive: true });
