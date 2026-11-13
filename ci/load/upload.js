// CI only – k6 load test (ci/stack.sh load): parallel users upload large files through Caddy into CloudShare,
// exactly like the taskpane does (connect once, start upload, 10 MiB chunks, complete with password).
import { check, fail } from "k6";
import http from "k6/http";

const BASE = "https://localhost";
const MOCKS = "http://127.0.0.1:8099";
const MiB = 1024 * 1024;

const USERS = Number(__ENV.USERS || 10);
const FILES_PER_USER = Number(__ENV.FILES_PER_USER || 2);
const FILE_SIZE = Number(__ENV.FILE_MIB || 45) * MiB; // 45 MiB = 4 full chunks + one half chunk

const CHUNK = 10 * MiB;
const fullChunk = new Uint8Array(CHUNK).fill(0x43).buffer;

export const options = {
  insecureSkipTLSVerify: true, // Caddy's internal CA for localhost
  scenarios: {
    uploads: { executor: "per-vu-iterations", vus: USERS, iterations: FILES_PER_USER, maxDuration: "5m" },
  },
  thresholds: {
    checks: ["rate==1"],
    http_req_failed: ["rate==0"],
    "http_req_duration{step:chunk}": ["p(95)<3000"],
    "http_req_duration{step:complete}": ["p(95)<3000"],
  },
};

const json = (token) => ({ headers: { Authorization: `Bearer ${token}`, "Content-Type": "application/json" } });

function connect(token) {
  const start = http.post(`${BASE}/api/nextcloud/connect`, null, { ...json(token), tags: { step: "connect" } });
  if (!check(start, { "connect started": (r) => r.status === 200 })) fail(`connect: ${start.status} ${start.body}`);
  const poll = http.post(`${BASE}/api/nextcloud/connect/${start.json("connectId")}/poll`, null, {
    ...json(token),
    tags: { step: "connect" },
  });
  if (!check(poll, { "connected": (r) => r.json("connected") === true })) fail(`poll: ${poll.status} ${poll.body}`);
}

export default function () {
  const token = http.get(`${MOCKS}/token?oid=load-user-${__VU}`, { tags: { step: "token" } }).body;
  if (__ITER === 0) connect(token);

  const fileName = `Lasttest ${__VU}-${__ITER}.bin`;
  const start = http.post(`${BASE}/api/uploads`, JSON.stringify({ fileName, size: FILE_SIZE }), {
    ...json(token),
    tags: { step: "start" },
  });
  if (!check(start, { "upload started": (r) => r.status === 200 })) fail(`start: ${start.status} ${start.body}`);
  const { uploadId, totalChunks, chunkSize } = start.json();

  for (let index = 1; index <= totalChunks; index++) {
    const length = Math.min(chunkSize, FILE_SIZE - (index - 1) * chunkSize);
    const res = http.put(`${BASE}/api/uploads/${uploadId}/chunks/${index}`, length === CHUNK ? fullChunk : fullChunk.slice(0, length), {
      headers: { Authorization: `Bearer ${token}`, "Content-Type": "application/octet-stream" },
      tags: { step: "chunk" },
      timeout: "60s",
    });
    check(res, { "chunk accepted": (r) => r.status === 200 });
  }

  const done = http.post(
    `${BASE}/api/uploads/${uploadId}/complete`,
    JSON.stringify({ password: "kT7#pQ2m!xV9wR4e_Hn3", expireDays: 14, maxAccess: 3 }),
    { ...json(token), tags: { step: "complete" } },
  );
  check(done, {
    "share created": (r) => r.status === 200 && String(r.json("shareUrl")).includes("/s/"),
    "password job created": (r) => typeof r.json("jobId") === "string",
  });
}
