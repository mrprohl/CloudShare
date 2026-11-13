/**
 * API tests with fastify.inject(): authentication, ownership boundaries between users, input validation.
 * Entra ID, Nextcloud and Vaultwarden are replaced by test doubles; credentials and jobs use the real modules.
 * Test users: "Authorization: Bearer user-<name>" → oid "oid-<name>".
 */
import { afterAll, beforeAll, beforeEach, describe, expect, it, vi } from "vitest";

const nc = vi.hoisted(() => ({
  userInfo: vi.fn(),
  ensureFolder: vi.fn(),
  startUpload: vi.fn(),
  putChunk: vi.fn(),
  finishUpload: vi.fn(),
  abortUpload: vi.fn(),
  createPublicShare: vi.fn(),
  listPublicShares: vi.fn(),
  deleteShare: vi.fn(),
  publicDownloads: vi.fn(),
  countFiles: vi.fn(),
  revokeAppPassword: vi.fn(),
}));

vi.mock("../src/auth.js", async (importOriginal) => {
  const actual = await importOriginal<typeof import("../src/auth.js")>();
  return {
    ...actual,
    authenticate: vi.fn(async (req: { headers: Record<string, string | undefined> }) => {
      const header = req.headers.authorization ?? "";
      if (!header.startsWith("Bearer user-")) throw new actual.HttpError(401, "Anmeldung erforderlich");
      const name = header.slice("Bearer user-".length);
      return { oid: `oid-${name}`, upn: `${name}@example.test`, token: `sso-${name}` };
    }),
    onBehalfOf: vi.fn(async () => ({ token: "graph-token", expiresAt: Date.now() + 3_600_000 })),
  };
});
vi.mock("../src/vaultwarden.js", () => ({
  createPasswordSend: vi.fn(async () => "https://vault.example.test/#/send/abc"),
  initVault: vi.fn(),
}));
vi.mock("../src/nextcloud.js", async (importOriginal) => {
  const actual = await importOriginal<typeof import("../src/nextcloud.js")>();
  // `new NextcloudClient(...)` in the app returns the shared test double.
  class FakeNextcloudClient {
    constructor() {
      return nc;
    }
  }
  return { ...actual, NextcloudClient: FakeNextcloudClient };
});

const { buildApp } = await import("../src/app.js");
const { saveCredentials } = await import("../src/credentials.js");
const { createJob, getJob } = await import("../src/jobs.js");
const vaultwarden = await import("../src/vaultwarden.js");
const createPasswordSend = vi.mocked(vaultwarden.createPasswordSend);

const app = await buildApp({ logger: false });
const as = (user: string) => ({ authorization: `Bearer user-${user}` });
const STRONG = "kT7#pQ2m!xV9wR4e_Hn3";
const MiB = 1024 * 1024;

beforeAll(async () => {
  await saveCredentials("oid-anna", { loginName: "anna", appPassword: "pw-anna" });
  await saveCredentials("oid-bernd", { loginName: "bernd", appPassword: "pw-bernd" });
});
afterAll(() => app.close());

beforeEach(() => {
  for (const fn of Object.values(nc)) fn.mockReset();
  nc.userInfo.mockResolvedValue({ id: "uid", free: null });
  nc.createPublicShare.mockResolvedValue("https://cloud.example.test/s/abc");
  createPasswordSend.mockClear();
});

async function startUpload(user: string, fileName = "a.pdf", size = 100) {
  const res = await app.inject({ method: "POST", url: "/api/uploads", headers: as(user), payload: { fileName, size } });
  return { res, body: res.json() as { uploadId: string; totalChunks: number } };
}

const putChunk = (user: string, uploadId: string, index: string | number, payload = Buffer.from("x".repeat(100))) =>
  app.inject({
    method: "PUT",
    url: `/api/uploads/${uploadId}/chunks/${index}`,
    headers: { ...as(user), "content-type": "application/octet-stream" },
    payload,
  });

const complete = (user: string, uploadId: string, body: Record<string, unknown> = {}) =>
  app.inject({
    method: "POST",
    url: `/api/uploads/${uploadId}/complete`,
    headers: as(user),
    payload: { password: STRONG, ...body },
  });

describe("authentication", () => {
  it("serves the public configuration without login", async () => {
    const res = await app.inject({ method: "GET", url: "/api/config" });
    expect(res.statusCode).toBe(200);
    expect(res.json()).toMatchObject({ thresholdMb: 1, sendMaxAccessCount: 3, sendMaxAccessLimit: 20 });
  });

  it.each([
    ["GET", "/api/profile"],
    ["GET", "/api/nextcloud/status"],
    ["POST", "/api/nextcloud/connect"],
    ["DELETE", "/api/nextcloud/connection"],
    ["POST", "/api/uploads"],
    ["POST", "/api/requests"],
    ["GET", "/api/shares"],
    ["DELETE", "/api/shares/1"],
    ["POST", "/api/jobs/x/arm"],
  ] as const)("rejects %s %s without a valid token", async (method, url) => {
    const res = await app.inject({ method, url, headers: { authorization: "Bearer fremd" } });
    expect(res.statusCode).toBe(401);
  });

  it("asks users without Nextcloud connection to connect first", async () => {
    const { res } = await startUpload("carla");
    expect(res.statusCode).toBe(409);
    expect(res.json().code).toBe("NEXTCLOUD_NOT_CONNECTED");
  });
});

describe("uploads", () => {
  it("neutralises path characters in file names", async () => {
    const { res } = await startUpload("anna", "../../etc/passwd");
    expect(res.statusCode).toBe(200);
    const target: string[] = nc.startUpload.mock.calls[0][2];
    expect(target.at(-1)).toBe(".._.._etc_passwd");
    expect(target.every((segment) => segment !== ".." && !segment.includes("/"))).toBe(true);
  });

  it.each([".", "..", "", "   "])("rejects the file name %j", async (name) => {
    expect((await startUpload("anna", name)).res.statusCode).toBe(400);
  });

  it.each([0, -1, 10_000 * 10 * MiB + 1, "viel"])("rejects the size %j", async (size) => {
    expect((await startUpload("anna", "a.pdf", size as number)).res.statusCode).toBe(400);
  });

  it("refuses uploads larger than the free Nextcloud quota", async () => {
    nc.userInfo.mockResolvedValue({ id: "uid", free: 50 });
    const { res } = await startUpload("anna", "a.pdf", 100);
    expect(res.statusCode).toBe(507);
    expect(res.json().code).toBe("NEXTCLOUD_QUOTA_EXCEEDED");
  });

  it("validates chunk numbers and bodies", async () => {
    const { body } = await startUpload("anna"); // 100 bytes → 1 chunk
    expect((await putChunk("anna", body.uploadId, 0)).statusCode).toBe(400);
    expect((await putChunk("anna", body.uploadId, 2)).statusCode).toBe(400);
    expect((await putChunk("anna", body.uploadId, "abc")).statusCode).toBe(400);
    expect((await putChunk("anna", body.uploadId, 1, Buffer.alloc(0))).statusCode).toBe(400);
    expect((await putChunk("anna", body.uploadId, 1)).statusCode).toBe(200);
  });

  it("does not let another user touch an upload", async () => {
    const { body } = await startUpload("anna");
    expect((await putChunk("bernd", body.uploadId, 1)).statusCode).toBe(404);
    expect((await complete("bernd", body.uploadId)).statusCode).toBe(404);
    const del = await app.inject({ method: "DELETE", url: `/api/uploads/${body.uploadId}`, headers: as("bernd") });
    expect(del.statusCode).toBe(404);
    expect(nc.putChunk).not.toHaveBeenCalled();
    expect(nc.abortUpload).not.toHaveBeenCalled();
  });

  it("only completes uploads with all chunks", async () => {
    const { body } = await startUpload("anna");
    expect((await complete("anna", body.uploadId)).statusCode).toBe(409);
  });

  it("rejects weak passwords before anything is shared", async () => {
    const { body } = await startUpload("anna");
    await putChunk("anna", body.uploadId, 1);
    const res = await complete("anna", body.uploadId, { password: "kurz" });
    expect(res.statusCode).toBe(400);
    expect(res.json().error).toMatch(/12 Zeichen/);
    expect(nc.createPublicShare).not.toHaveBeenCalled();
  });

  it("limits validity and password-link accesses to the configured maximum", async () => {
    const { body } = await startUpload("anna");
    await putChunk("anna", body.uploadId, 1);
    const res = await complete("anna", body.uploadId, { expireDays: 1000, maxAccess: 500 });
    expect(res.statusCode).toBe(200);
    const maxDate = new Date(Date.now() + 90 * 86_400_000).toISOString().slice(0, 10);
    expect(res.json().expireDate).toBe(maxDate);
    expect(nc.createPublicShare.mock.calls[0][2]).toBe(maxDate);
    expect(createPasswordSend.mock.calls[0][2]).toBe(20);
  });

  it("adds further files to the user's own job, never to someone else's", async () => {
    const first = await startUpload("anna");
    await putChunk("anna", first.body.uploadId, 1);
    const { jobId } = (await complete("anna", first.body.uploadId)).json();

    const second = await startUpload("anna", "b.pdf");
    await putChunk("anna", second.body.uploadId, 1);
    expect((await complete("anna", second.body.uploadId, { jobId })).json().jobId).toBe(jobId);
    expect(getJob(jobId)?.fileNames).toEqual(["a.pdf", "b.pdf"]);

    const foreign = await startUpload("bernd", "c.pdf");
    await putChunk("bernd", foreign.body.uploadId, 1);
    const bernd = (await complete("bernd", foreign.body.uploadId, { jobId })).json();
    expect(bernd.jobId).not.toBe(jobId);
    expect(getJob(jobId)?.fileNames).toEqual(["a.pdf", "b.pdf"]);
  });
});

describe("upload links", () => {
  it("creates an upload-only share in its own request folder", async () => {
    const res = await app.inject({ method: "POST", url: "/api/requests", headers: as("anna"), payload: { password: STRONG } });
    expect(res.statusCode).toBe(200);
    const [folder, , , permissions, label] = nc.createPublicShare.mock.calls[0];
    expect(folder[0]).toBe("Outlook-Dateianforderungen");
    expect(permissions).toBe(4);
    expect(label).toBe("Dateianforderung");
  });
});

describe("my shares", () => {
  const shares = [
    { id: "1", path: "/Outlook-Anhaenge/2026-11-13_ab/a.pdf", url: "u1", expiration: "2026-11-27", sharedAt: 2, fileId: 11, isFolder: false, permissions: 1, label: "" },
    { id: "2", path: "/Outlook-Dateianforderungen/2026-11-13_cd", url: "u2", expiration: null, sharedAt: 3, fileId: 12, isFolder: true, permissions: 4, label: "Dateianforderung" },
    { id: "3", path: "/Privat/urlaub.jpg", url: "u3", expiration: null, sharedAt: 4, fileId: 13, isFolder: false, permissions: 1, label: "" },
  ];

  it("lists only CloudShare links with their status", async () => {
    nc.listPublicShares.mockResolvedValue(shares);
    nc.publicDownloads.mockResolvedValue({ count: 2, last: "2026-11-13T10:20:00+00:00" });
    nc.countFiles.mockResolvedValue(3);
    const res = await app.inject({ method: "GET", url: "/api/shares", headers: as("anna") });
    expect(res.statusCode).toBe(200);
    expect(res.json()).toEqual([
      expect.objectContaining({ id: "2", kind: "request", name: "Dateianforderung", received: 3, downloads: null }),
      expect.objectContaining({ id: "1", kind: "file", name: "a.pdf", downloads: { count: 2, last: "2026-11-13T10:20:00+00:00" } }),
    ]);
  });

  it("only revokes CloudShare links", async () => {
    nc.listPublicShares.mockResolvedValue(shares);
    const other = await app.inject({ method: "DELETE", url: "/api/shares/3", headers: as("anna") });
    expect(other.statusCode).toBe(404);
    const unknown = await app.inject({ method: "DELETE", url: "/api/shares/99", headers: as("anna") });
    expect(unknown.statusCode).toBe(404);
    expect(nc.deleteShare).not.toHaveBeenCalled();

    const own = await app.inject({ method: "DELETE", url: "/api/shares/1", headers: as("anna") });
    expect(own.statusCode).toBe(200);
    expect(nc.deleteShare).toHaveBeenCalledWith("1");
  });
});

describe("arming the password mail", () => {
  it("is only possible for the owner of the job", async () => {
    const id = await createJob("oid-anna", "anna@example.test", "a.pdf", "https://vault/#/send/x");
    const res = await app.inject({ method: "POST", url: `/api/jobs/${id}/arm`, headers: as("bernd"), payload: {} });
    expect(res.statusCode).toBe(404);
    expect(getJob(id)?.armed).toBeUndefined();
  });

  it("stores a valid shared mailbox and ignores anything else", async () => {
    const shared = await createJob("oid-anna", "anna@example.test", "a.pdf", "https://vault/#/send/x");
    await app.inject({ method: "POST", url: `/api/jobs/${shared}/arm`, headers: as("anna"), payload: { from: "Info@Example.test" } });
    expect(getJob(shared)?.armed?.mailbox).toBe("info@example.test");

    const odd = await createJob("oid-anna", "anna@example.test", "a.pdf", "https://vault/#/send/x");
    await app.inject({ method: "POST", url: `/api/jobs/${odd}/arm`, headers: as("anna"), payload: { from: "../users/x" } });
    expect(getJob(odd)?.armed).toBeDefined();
    expect(getJob(odd)?.armed?.mailbox).toBeUndefined();
  });
});

// Last on purpose: exhausts the strict limit of POST /api/nextcloud/connect for the test client address.
describe("rate limiting", () => {
  /** Unsigned token with an oid – only decoded for the rate-limit bucket, authenticate() still rejects it. */
  const tokenFor = (oid: string) =>
    ["{\"alg\":\"none\"}", JSON.stringify({ oid })].map((part) => Buffer.from(part).toString("base64url")).join(".") + ".";

  it("limits expensive actions per client and answers 429 with a German message", async () => {
    const statuses: number[] = [];
    let last;
    for (let i = 0; i < 12; i++) {
      last = await app.inject({ method: "POST", url: "/api/nextcloud/connect" });
      statuses.push(last.statusCode);
    }
    expect(statuses).toContain(401); // the first requests reach the route (no login)
    expect(last!.statusCode).toBe(429);
    expect(last!.json()).toMatchObject({ code: "RATE_LIMITED", error: expect.stringContaining("Zu viele Anfragen") });
    expect(last!.headers["retry-after"]).toBeDefined();
  });

  it("keeps separate buckets per user (oid from the token)", async () => {
    const other = await app.inject({
      method: "POST",
      url: "/api/nextcloud/connect",
      headers: { authorization: `Bearer ${tokenFor("oid-someone-else")}` },
    });
    expect(other.statusCode).toBe(401); // not limited – own bucket, then rejected by authenticate()
  });

  it("ignores X-Forwarded-For from clients outside private networks", async () => {
    const res = await app.inject({
      method: "POST",
      url: "/api/nextcloud/connect",
      remoteAddress: "127.0.0.1",
      headers: { "x-forwarded-for": "203.0.113.7" },
    });
    expect(res.statusCode).toBe(401); // loopback proxy is trusted → new client address, own bucket
    // A public sender cannot pick a fresh address per request: all of them count for 198.51.100.9.
    let spoofed;
    for (let i = 0; i < 12; i++) {
      spoofed = await app.inject({
        method: "POST",
        url: "/api/nextcloud/connect",
        remoteAddress: "198.51.100.9",
        headers: { "x-forwarded-for": `203.0.113.${100 + i}` },
      });
    }
    expect(spoofed!.statusCode).toBe(429);
  });
});
