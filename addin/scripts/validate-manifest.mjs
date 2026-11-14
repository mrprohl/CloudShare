// Validates manifest.xml with Microsoft's office-addin-manifest tool (schema + Office validation service), once per
// language in locales/. The manifest contains placeholders that docker/entrypoint.sh and the server (texts of the
// configured language) replace at container start – they are filled with the production address, a dummy client ID
// and the language's texts here, so the validator sees a deployable manifest.
import { execFileSync } from "node:child_process";
import { mkdtempSync, readdirSync, readFileSync, writeFileSync } from "node:fs";
import { tmpdir } from "node:os";
import path from "node:path";
import { loadLocale, localize } from "./localize.mjs";

const baseUrl = process.env.PUBLIC_BASE_URL ?? "https://cloudshare.example.com";
const clientId = process.env.ENTRA_CLIENT_ID ?? "11111111-2222-3333-4444-555555555555";
const languages = readdirSync(new URL("../../locales/", import.meta.url))
  .filter((file) => file.endsWith(".json"))
  .map((file) => file.slice(0, -5));

const template = readFileSync("manifest.xml", "utf8")
  .replaceAll("%%PUBLIC_BASE_URL%%", baseUrl)
  .replaceAll("%%PUBLIC_HOST%%", new URL(baseUrl).host)
  .replaceAll("%%ENTRA_CLIENT_ID%%", clientId);

for (const language of languages) {
  const manifest = localize(template, ".xml", loadLocale(language), language);
  if (manifest.includes("%%")) {
    console.error(`manifest.xml (${language}) contains an unknown placeholder`);
    process.exit(1);
  }
  console.log(`--- manifest.xml, language ${language}`);
  const file = path.join(mkdtempSync(path.join(tmpdir(), "cloudshare-manifest-")), "manifest.xml");
  writeFileSync(file, manifest);
  execFileSync("npx", ["--no-install", "office-addin-manifest", "validate", file], {
    stdio: "inherit",
    shell: process.platform === "win32",
  });
}
