// CI only – runs inside the PREVIOUS image (ci/stack.sh upgrade) and writes persistent data with that version's own
// code: two Nextcloud connections and two password jobs (one open with two files, one armed for a shared mailbox).
// verify.mjs then reads them with the NEW image.
import { writeFileSync } from "node:fs";

const { saveCredentials } = await import("/app/dist/credentials.js");
const jobs = await import("/app/dist/jobs.js");

const credentials = {
  "oid-upgrade-anna": { loginName: "anna", appPassword: "Anna-AppPw-1 äöü" },
  "oid-upgrade-bernd": { loginName: "bernd@example.com", appPassword: "Bernd-AppPw-2" },
};
for (const [oid, creds] of Object.entries(credentials)) await saveCredentials(oid, creds);

const open = await jobs.createJob("oid-upgrade-anna", "anna@example.com", "Bericht.pdf", "https://vault.example.com/#/send/a", 3);
await jobs.addFileToJob(open, "Anhang 2.zip");

const armed = await jobs.createJob("oid-upgrade-bernd", "bernd@example.com", "Vertrag.docx", "https://vault.example.com/#/send/b", 5);
const armedAt = Date.now();
await jobs.armJob(armed, {
  graphToken: "graph-token-upgrade",
  tokenExpiresAt: armedAt + 3_600_000,
  armedAt,
  mailbox: "info@example.com",
});

writeFileSync(
  "/data/ci-upgrade-expected.json",
  JSON.stringify({
    credentials,
    jobs: {
      [open]: { owner: "oid-upgrade-anna", fileNames: ["Bericht.pdf", "Anhang 2.zip"], maxAccessCount: 3, armed: false },
      [armed]: { owner: "oid-upgrade-bernd", fileNames: ["Vertrag.docx"], maxAccessCount: 5, armed: true, mailbox: "info@example.com" },
    },
  }),
);
console.log("seeded 2 connections and 2 password jobs");
