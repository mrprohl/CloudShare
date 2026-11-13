// CI only – runs inside the NEW image (ci/stack.sh upgrade) on the data written by seed.mjs with the previous image:
// every connection and password job must still be readable and unchanged.
import { readFileSync } from "node:fs";

const expected = JSON.parse(readFileSync("/data/ci-upgrade-expected.json", "utf8"));
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

process.exit(failures ? 1 : 0);
