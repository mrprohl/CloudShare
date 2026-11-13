// Bundles the taskpane and the Smart Alert handlers as classic scripts (IIFE).
// The JavaScript-only runtime of classic Outlook does not support ES modules.
// docker/entrypoint.sh replaces the placeholders %%PUBLIC_BASE_URL%% and %%ENTRA_CLIENT_ID%% at container start.
import { build } from "esbuild";
import { cpSync, mkdirSync, rmSync } from "node:fs";

rmSync("dist", { recursive: true, force: true });
mkdirSync("dist", { recursive: true });

const common = { outdir: "dist", bundle: true, format: "iife", minify: true, sourcemap: false };
await build({ ...common, entryPoints: { taskpane: "src/taskpane/taskpane.ts" }, target: "es2017" });
// The event handlers run in classic Outlook's JavaScript-only runtime (no DOM, older engine): conservative syntax,
// async/await is compiled to generators. test/bundle.test.ts checks the result.
await build({ ...common, entryPoints: { commands: "src/commands/commands.ts" }, target: "es2016" });

for (const file of ["src/taskpane/taskpane.html", "src/taskpane/taskpane.css", "src/commands/commands.html", "manifest.xml"]) {
  cpSync(file, `dist/${file.split("/").pop()}`);
}
cpSync("assets", "dist/assets", { recursive: true });
