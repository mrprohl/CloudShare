import { existsSync } from "node:fs";
import path from "node:path";
import { fileURLToPath } from "node:url";
import { buildApp } from "./app.js";
import { config } from "./config.js";
import { startCleanup } from "./cleanup.js";
import { addinLanguages, language, locale, localizePublic } from "./i18n.js";
import { startNotices } from "./notices.js";
import { startProofWatcher } from "./proofWatcher.js";
import { startRetention } from "./db.js";
import { onUnsentExpired, startJobWatcher } from "./jobs.js";
import { revokeUnsentLinks } from "./shares.js";
import { initVault } from "./vaultwarden.js";

// Texts of all languages into the add-in files (copied from the image by docker/entrypoint.sh) – before the server
// answers the first request for them. The add-in picks the client's language, mails and SMS use LANGUAGE.
const publicDir = path.resolve(path.dirname(fileURLToPath(import.meta.url)), "../public");
const localized = existsSync(publicDir) ? localizePublic(publicDir) : 0;

const app = await buildApp();
app.log.info({ language, locale: locale(), addin: addinLanguages(), files: localized }, "Language");
await initVault((msg) => app.log.info(msg));
await app.listen({ host: "0.0.0.0", port: config.port });
// Links of drafts that were never sent are revoked when their job expires (shares.ts).
onUnsentExpired(revokeUnsentLinks);
startJobWatcher(app.log);
startCleanup(app.log);
startNotices(app.log);
startProofWatcher(app.log);
startRetention(app.log);
