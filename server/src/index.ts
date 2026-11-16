import { existsSync } from "node:fs";
import path from "node:path";
import { fileURLToPath } from "node:url";
import { buildApp } from "./app.js";
import { config } from "./config.js";
import { startCleanup } from "./cleanup.js";
import { language, locale, localizePublic } from "./i18n.js";
import { startNotices } from "./notices.js";
import { startRetention } from "./db.js";
import { onUnsentExpired, startJobWatcher } from "./jobs.js";
import { revokeUnsentLinks } from "./shares.js";
import { initVault } from "./vaultwarden.js";

// Texts of the configured language into the add-in files (copied from the image by docker/entrypoint.sh) – before
// the server answers the first request for them.
const publicDir = path.resolve(path.dirname(fileURLToPath(import.meta.url)), "../public");
const localized = existsSync(publicDir) ? localizePublic(publicDir) : 0;

const app = await buildApp();
app.log.info({ language, locale: locale(), files: localized }, "Language");
await initVault((msg) => app.log.info(msg));
await app.listen({ host: "0.0.0.0", port: config.port });
// Links of drafts that were never sent are revoked when their job expires (shares.ts).
onUnsentExpired(revokeUnsentLinks);
startJobWatcher(app.log);
startCleanup(app.log);
startNotices(app.log);
startRetention(app.log);
