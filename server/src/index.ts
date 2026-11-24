import { existsSync } from "node:fs";
import path from "node:path";
import { fileURLToPath } from "node:url";
import { buildApp } from "./app.js";
import { config } from "./config.js";
import { startCleanup } from "./cleanup.js";
import { addinLanguages, language, locale, localizePublic } from "./i18n.js";
import { startNotices } from "./notices.js";
import { prepareSigningIdentity, signingFilesProblem } from "./pdfSign.js";
import { startProofRetention } from "./proof.js";
import { startProofWatcher } from "./proofWatcher.js";
import { startRetention } from "./db.js";
import { onUnsentExpired, startJobWatcher } from "./jobs.js";
import { revokeUnsentLinks } from "./shares.js";
import { initVault, startLocalPasswordPurge } from "./vaultwarden.js";

// Texts of all languages into the add-in files (copied from the image by docker/entrypoint.sh) – before the server
// answers the first request for them. The add-in picks the client's language, mails and SMS use LANGUAGE.
const publicDir = path.resolve(path.dirname(fileURLToPath(import.meta.url)), "../public");
const localized = existsSync(publicDir) ? localizePublic(publicDir) : 0;

const app = await buildApp();
app.log.info({ language, locale: locale(), addin: addinLanguages(), files: localized }, "Language");
await initVault((msg) => app.log.info(msg));
// Own PEM files for the proof PDF that cannot be used: told now, not with the first proof (signed with the own certificate meanwhile).
const signingProblem = config.deliveryProof ? signingFilesProblem() : null;
if (signingProblem) {
  app.log.error({ problem: signingProblem }, "Proof signing files not usable – signing with CloudShare's own certificate instead");
}
// Which certificate signs the proofs – a new own one right away when PROOF_SIGN_NAME / PROOF_SIGN_ORGANIZATION changed.
if (config.deliveryProof) {
  const signer = prepareSigningIdentity();
  const message = {
    files: "Proof signing: own PEM files (PROOF_SIGN_CERT_FILE)",
    kept: "Proof signing: CloudShare's own certificate",
    made: "Proof signing: CloudShare made its own certificate",
    renewed: "Proof signing: PROOF_SIGN_NAME / PROOF_SIGN_ORGANIZATION changed – CloudShare made a new own certificate (earlier PDFs stay valid)",
  }[signer.source];
  app.log.info({ name: signer.name, organization: signer.organization }, message);
}
await app.listen({ host: "0.0.0.0", port: config.port });
// Links of drafts that were never sent are revoked when their job expires (shares.ts).
onUnsentExpired(revokeUnsentLinks);
startJobWatcher(app.log);
startCleanup(app.log);
startNotices(app.log);
startProofWatcher(app.log);
startRetention(app.log);
startProofRetention(app.log);
// Passwords kept without Vaultwarden go when their links expire (also after switching Vaultwarden on again).
startLocalPasswordPurge(app.log);
