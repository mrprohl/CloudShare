// Validates manifest.xml with Microsoft's office-addin-manifest tool (schema + Office validation service).
// The manifest contains placeholders that docker/entrypoint.sh replaces at container start – they are filled
// with the production address and a dummy client ID here, so the validator sees a deployable manifest.
import { execFileSync } from "node:child_process";
import { mkdtempSync, readFileSync, writeFileSync } from "node:fs";
import { tmpdir } from "node:os";
import path from "node:path";

const baseUrl = process.env.PUBLIC_BASE_URL ?? "https://cloudshare.example.com";
const clientId = process.env.ENTRA_CLIENT_ID ?? "11111111-2222-3333-4444-555555555555";

const manifest = readFileSync("manifest.xml", "utf8")
  .replaceAll("%%PUBLIC_BASE_URL%%", baseUrl)
  .replaceAll("%%PUBLIC_HOST%%", new URL(baseUrl).host)
  .replaceAll("%%ENTRA_CLIENT_ID%%", clientId);

if (manifest.includes("%%")) {
  console.error("manifest.xml contains an unknown placeholder");
  process.exit(1);
}

const file = path.join(mkdtempSync(path.join(tmpdir(), "cloudshare-manifest-")), "manifest.xml");
writeFileSync(file, manifest);
execFileSync("npx", ["--no-install", "office-addin-manifest", "validate", file], {
  stdio: "inherit",
  shell: process.platform === "win32",
});
