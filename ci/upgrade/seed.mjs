// CI only – runs inside the PREVIOUS image (ci/stack.sh upgrade) and writes persistent data with that version's own
// code, then notes what it wrote in /data/ci-upgrade-expected.json; verify.mjs reads it all with the NEW image.
//   always: two Nextcloud connections, two password jobs of the first format (one open, one armed for a shared mailbox)
//   if the previous version has them: a job with one password per step, an SMS delivery status, a pending download
//   confirmation, an upload folder in the cleanup list and a password entry for sending the password again
// Stores an older image does not know yet are skipped (and then not checked).
import { writeFileSync } from "node:fs";

const { saveCredentials } = await import("/app/dist/credentials.js");
const jobs = await import("/app/dist/jobs.js");
const optional = (path) => import(path).catch(() => null);

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

const expected = {
  credentials,
  jobs: {
    [open]: { owner: "oid-upgrade-anna", fileNames: ["Bericht.pdf", "Anhang 2.zip"], maxAccessCount: 3, armed: false },
    [armed]: { owner: "oid-upgrade-bernd", fileNames: ["Vertrag.docx"], maxAccessCount: 5, armed: true, mailbox: "info@example.com" },
  },
};
const seeded = ["2 connections", "2 password jobs"];

const owner = "oid-upgrade-anna";
const token = "tok-upgrade-plan";
const share = { token, name: "Plan.pdf", expireDate: "2099-12-01" };

if (typeof jobs.createSeparateJob === "function") {
  const id = await jobs.createSeparateJob(owner, "anna@example.com", {
    step: "step-upgrade", fileNames: ["Plan.pdf"], sendUrl: "https://vault.example.com/#/send/c", sendId: "send-upgrade",
    maxAccessCount: 3, shares: [share],
  });
  expected.separate = { id, owner, step: "step-upgrade", sendId: "send-upgrade", token };
  seeded.push("a job with one password per step");
}

const deliveries = await optional("/app/dist/deliveries.js");
if (deliveries?.recordDelivery) {
  await deliveries.recordDelivery([token], {
    owner, channel: "sms", to: "+49 170 •••• 67", at: Date.now(), messageId: "msg-upgrade", status: "submitted",
  });
  expected.delivery = { token, owner, channel: "sms", to: "+49 170 •••• 67", messageId: "msg-upgrade" };
  seeded.push("an SMS delivery status");
}

const notices = await optional("/app/dist/notices.js");
if (notices?.registerNotices) {
  await notices.registerNotices([[token, {
    owner, ownerUpn: "anna@example.com", notifyTo: "anna@example.com", fileName: "Plan.pdf", subject: "Upgrade äöü",
    recipients: "erika@example.com", expireDate: "2099-12-01", registeredAt: Date.now(),
  }]]);
  expected.notice = { token, fileName: "Plan.pdf", subject: "Upgrade äöü" };
  seeded.push("a pending download confirmation");
}

const cleanup = await optional("/app/dist/cleanup.js");
if (cleanup?.registerForCleanup) {
  await cleanup.registerForCleanup(owner, ["CloudShare", "2026-11-15_0a1b2c3d", "Plan.pdf"], "2099-12-01");
  expected.cleanup = { owner, folder: ["CloudShare", "2026-11-15_0a1b2c3d"], expireDate: "2099-12-01" };
  seeded.push("an upload folder in the cleanup list");
}

const refs = await optional("/app/dist/passwordRefs.js");
if (refs?.registerPasswordRef) {
  await refs.registerPasswordRef({
    owner, ownerUpn: "anna@example.com", sendId: "send-upgrade", tokens: [token], fileNames: ["Plan.pdf"],
    keepUntil: Date.now() + 20 * 86_400_000, subject: "Upgrade äöü", mailTo: ["erika@example.com"], sms: "+491701234567",
    maxAccessCount: 3,
  });
  expected.passwordRef = { token, owner, sendId: "send-upgrade", sms: "+491701234567" };
  seeded.push("a password entry for sending it again");
}

writeFileSync("/data/ci-upgrade-expected.json", JSON.stringify(expected));
console.log(`seeded ${seeded.join(", ")}`);
