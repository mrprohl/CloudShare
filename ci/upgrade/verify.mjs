// CI only – runs inside the NEW image (ci/stack.sh upgrade) on the data written by seed.mjs with the previous image:
// everything the previous version wrote must still be readable and unchanged.
import { readFileSync } from "node:fs";

const expected = JSON.parse(readFileSync("/data/ci-upgrade-expected.json", "utf8"));
// The switch from JSON files to SQLite takes no data over (accepted during the test phase): nothing to compare once.
// Exit code 3 = skipped (ci/stack.sh). From the next release on, both versions use the database again.
if (expected.storage !== "sqlite") {
  console.log("  – the previous version kept its data in JSON files, this one uses SQLite and takes none of them over");
  process.exit(3);
}
const { getCredentials } = await import("/app/dist/credentials.js");
const { getJob } = await import("/app/dist/jobs.js");

let failures = 0;
const check = (ok, text) => {
  console.log(`  ${ok ? "✓" : "✗"} ${text}`);
  if (!ok) failures++;
};

for (const [oid, creds] of Object.entries(expected.credentials)) {
  const actual = getCredentials(oid);
  check(
    actual?.loginName === creds.loginName && actual?.appPassword === creds.appPassword,
    `Nextcloud connection of ${oid} readable (app password decrypts)`,
  );
}

for (const [id, job] of Object.entries(expected.jobs)) {
  const actual = getJob(id);
  check(Boolean(actual), `password job ${id.slice(0, 8)} present`);
  if (!actual) continue;
  check(actual.owner === job.owner, `  owner unchanged`);
  check(JSON.stringify(actual.fileNames) === JSON.stringify(job.fileNames), `  file names unchanged`);
  check(actual.maxAccessCount === job.maxAccessCount, `  number of accesses unchanged`);
  check(actual.dispatched === false, `  not dispatched`);
  check(Boolean(actual.armed) === job.armed, `  armed state unchanged`);
  if (job.mailbox) check(actual.armed?.mailbox === job.mailbox, `  shared mailbox unchanged`);
}

if (expected.separate) {
  const { id, owner, step, sendId, token } = expected.separate;
  const part = getJob(id)?.separate?.[0];
  check(getJob(id)?.owner === owner && part?.step === step && part?.sendId === sendId, "job with one password per step readable");
  check(part?.shares?.[0]?.token === token, "  its links unchanged");
}

if (expected.delivery) {
  const { deliveryFor } = await import("/app/dist/deliveries.js");
  const { token, owner, ...fields } = expected.delivery;
  const actual = deliveryFor(token, owner);
  check(Object.entries(fields).every(([key, value]) => actual?.[key] === value), "SMS delivery status readable");
}

if (expected.notice) {
  const { pendingNotices } = await import("/app/dist/notices.js");
  const actual = new Map(pendingNotices()).get(expected.notice.token);
  check(actual?.fileName === expected.notice.fileName && actual?.subject === expected.notice.subject,
    "pending download confirmation readable");
}

if (expected.cleanup) {
  const { cleanupEntries } = await import("/app/dist/cleanup.js");
  const { owner, folder, expireDate } = expected.cleanup;
  check(cleanupEntries().some((e) => e.owner === owner && e.folder.join("/") === folder.join("/") && e.expireDate === expireDate),
    "upload folder in the cleanup list readable");
}

if (expected.passwordRef) {
  const { passwordRefFor } = await import("/app/dist/passwordRefs.js");
  const { token, owner, sendId, sms } = expected.passwordRef;
  const actual = passwordRefFor(token, owner);
  check(actual?.sendId === sendId && actual?.sms === sms, "password entry for sending it again readable");
}

process.exit(failures ? 1 : 0);
