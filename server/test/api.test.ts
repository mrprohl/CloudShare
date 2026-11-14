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
  deletePath: vi.fn(),
  linkPasswordEnforced: vi.fn(),
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
vi.mock("../src/graph.js", async (importOriginal) => ({
  ...(await importOriginal<typeof import("../src/graph.js")>()),
  fetchProfile: vi.fn(async () => ({
    displayName: "Anna Muster", givenName: "Anna", surname: "Muster", jobTitle: "", department: "Vertrieb", company: "",
    email: "anna@example.test", phone: "", mobile: "", office: "",
  })),
}));
vi.mock("../src/vaultwarden.js", () => ({
  createPasswordSend: vi.fn(async () => ({ url: "https://vault.example.test/#/send/abc", id: "send-1" })),
  createMessageSend: vi.fn(async () => ({ url: "https://vault.example.test/#/send/msg/key123", id: "send-msg" })),
  MESSAGE_MAX_LENGTH: 1000,
  setSendMaxAccess: vi.fn(async () => undefined),
  setSendExpiry: vi.fn(async () => undefined),
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
const setSendMaxAccess = vi.mocked(vaultwarden.setSendMaxAccess);
const setSendExpiry = vi.mocked(vaultwarden.setSendExpiry);
const createMessageSend = vi.mocked(vaultwarden.createMessageSend);

const app = await buildApp({ logger: false });
const as = (user: string) => ({ authorization: `Bearer user-${user}` });
const STRONG = "kT7#pQ2m!xV9wR4e_Hn3";
const MiB = 1024 * 1024;
const DAY = 86_400_000;

beforeAll(async () => {
  await saveCredentials("oid-anna", { loginName: "anna", appPassword: "pw-anna" });
  await saveCredentials("oid-bernd", { loginName: "bernd", appPassword: "pw-bernd" });
});
afterAll(() => app.close());

beforeEach(() => {
  for (const fn of Object.values(nc)) fn.mockReset();
  nc.userInfo.mockResolvedValue({ id: "uid", free: null });
  nc.linkPasswordEnforced.mockResolvedValue(false);
  nc.createPublicShare.mockResolvedValue("https://cloud.example.test/s/abc");
  createPasswordSend.mockClear();
  setSendMaxAccess.mockClear();
  setSendExpiry.mockClear();
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
    expect(res.json()).toMatchObject({ thresholdMb: 5, sendMaxAccessCount: 3, sendMaxAccessLimit: 20 });
  });

  it("allows the send handler script in classic Outlook's JavaScript-only runtime (well-known URI)", async () => {
    const res = await app.inject({ method: "GET", url: "/.well-known/microsoft-officeaddins-allowed.json" });
    expect(res.statusCode).toBe(200);
    expect(res.json()).toEqual({ allowed: ["https://cloudshare.example.test/commands.js"] });
  });

  it.each([
    ["GET", "/api/profile"],
    ["GET", "/api/nextcloud/status"],
    ["POST", "/api/nextcloud/connect"],
    ["DELETE", "/api/nextcloud/connection"],
    ["POST", "/api/uploads"],
    ["POST", "/api/requests"],
    ["POST", "/api/messages"],
    ["GET", "/api/shares"],
    ["DELETE", "/api/shares/1"],
    ["POST", "/api/jobs/x/arm"],
  ] as const)("rejects %s %s without a valid token", async (method, url) => {
    const res = await app.inject({ method, url, headers: { authorization: "Bearer fremd" } });
    expect(res.statusCode).toBe(401);
  });

  it("has no trace route unless CLIENT_TRACE is set", async () => {
    const res = await app.inject({ method: "POST", url: "/api/trace", payload: { run: "abc", step: "start" } });
    expect(res.statusCode).toBe(404);
    expect((await app.inject({ method: "GET", url: "/api/trace?run=abc&step=start" })).statusCode).toBe(404);
    expect((await app.inject({ method: "GET", url: "/api/config" })).json().trace).toBe(false);
  });

  it("reports storage and whether Nextcloud enforces link passwords", async () => {
    nc.userInfo.mockResolvedValue({ id: "uid", free: 30, used: 20, total: 50 });
    nc.linkPasswordEnforced.mockResolvedValue(true);
    const res = await app.inject({ method: "GET", url: "/api/nextcloud/status", headers: as("anna") });
    expect(res.json()).toEqual({
      connected: true, loginName: "anna", freeBytes: 30, usedBytes: 20, totalBytes: 50, passwordEnforced: true,
    });

    nc.linkPasswordEnforced.mockRejectedValue(new Error("capabilities unavailable"));
    const unknown = await app.inject({ method: "GET", url: "/api/nextcloud/status", headers: as("anna") });
    expect(unknown.json().passwordEnforced).toBe(false);
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

  it("shares without password on request – no Send, no password; only a job for the download confirmation", async () => {
    const { body } = await startUpload("anna");
    await putChunk("anna", body.uploadId, 1);
    const res = await complete("anna", body.uploadId, { password: "", noPassword: true, jobId: "job-x" });
    expect(res.statusCode).toBe(200);
    expect(res.json().noPassword).toBe(true);
    const job = getJob(res.json().jobId)!; // DOWNLOAD_NOTICE=true in the tests
    expect(job).toMatchObject({ owner: "oid-anna", noPassword: true, sendUrl: "", fileNames: ["a.pdf"] });
    expect(job.noticeShares?.map((s) => s.name)).toEqual(["a.pdf"]);
    expect(job.separate).toBeUndefined();

    // The next link without password of the same mail joins that job.
    const second = await startUpload("anna", "b.pdf");
    await putChunk("anna", second.body.uploadId, 1);
    const next = await complete("anna", second.body.uploadId, { noPassword: true, jobId: res.json().jobId });
    expect(next.json().jobId).toBe(res.json().jobId);
    expect(getJob(res.json().jobId)!.noticeShares?.map((s) => s.name)).toEqual(["a.pdf", "b.pdf"]);
    expect(res.json().shareUrl).toBe("https://cloud.example.test/s/abc");
    expect(nc.createPublicShare.mock.calls[0][1]).toBeUndefined();
    expect(createPasswordSend).not.toHaveBeenCalled();
  });

  it("gives every 'Als Link einfügen' step its own password and Send; files of one step share them", async () => {
    const upload = async (name: string, body: Record<string, unknown>) => {
      const started = await startUpload("anna", name);
      await putChunk("anna", started.body.uploadId, 1);
      return (await complete("anna", started.body.uploadId, body)).json();
    };
    // Step 1: two files selected together.
    const { jobId } = await upload("a.pdf", { step: "s1", maxAccess: 2 });
    await upload("b.pdf", { step: "s1", jobId, maxAccess: 2 });
    // Step 2: another click with a new password.
    const OTHER = "Zr8!mQ4#vT1x_Lp9Wk2s";
    const res = await upload("c.pdf", { step: "s2", jobId, password: OTHER, maxAccess: 5 });
    expect(res.jobId).toBe(jobId);

    expect(createPasswordSend.mock.calls.map((c) => [c[1], c[2]])).toEqual([[STRONG, 2], [OTHER, 5]]);
    expect(getJob(jobId)?.separate?.map((p) => [p.step, p.fileNames, p.maxAccessCount])).toEqual([
      ["s1", ["a.pdf", "b.pdf"], 2],
      ["s2", ["c.pdf"], 5],
    ]);
    expect(getJob(jobId)?.fileNames).toEqual(["a.pdf", "b.pdf", "c.pdf"]);
  });

  it("keeps one shared password for mails started without steps (drafts from before)", async () => {
    const first = await startUpload("anna");
    await putChunk("anna", first.body.uploadId, 1);
    const { jobId } = (await complete("anna", first.body.uploadId)).json();

    const second = await startUpload("anna", "b.pdf");
    await putChunk("anna", second.body.uploadId, 1);
    expect((await complete("anna", second.body.uploadId, { step: "s9", jobId })).json().jobId).toBe(jobId);
    expect(getJob(jobId)?.separate).toBeUndefined();
    expect(getJob(jobId)?.fileNames).toEqual(["a.pdf", "b.pdf"]);
    expect(createPasswordSend).toHaveBeenCalledOnce();
  });

  it("only accepts noPassword as a real boolean", async () => {
    const { body } = await startUpload("anna");
    await putChunk("anna", body.uploadId, 1);
    const res = await complete("anna", body.uploadId, { password: "", noPassword: "true" });
    expect(res.statusCode).toBe(400);
    expect(nc.createPublicShare).not.toHaveBeenCalled();
  });

  it("limits validity and password-link accesses to the configured maximum", async () => {
    const { body } = await startUpload("anna");
    await putChunk("anna", body.uploadId, 1);
    const res = await complete("anna", body.uploadId, { expireDays: 1000, maxAccess: 500, passwordDays: 99 });
    expect(res.statusCode).toBe(200);
    const maxDate = new Date(Date.now() + 90 * 86_400_000).toISOString().slice(0, 10);
    expect(res.json().expireDate).toBe(maxDate);
    expect(nc.createPublicShare.mock.calls[0][2]).toBe(maxDate);
    expect(createPasswordSend.mock.calls[0][2]).toBe(20);
    expect(createPasswordSend.mock.calls[0][3]).toBe(30); // password link: Vaultwarden keeps Sends at most 31 days
  });

  it("keeps the password link for the chosen days, 5 by default", async () => {
    const first = await startUpload("anna");
    await putChunk("anna", first.body.uploadId, 1);
    await complete("anna", first.body.uploadId, { passwordDays: 12, step: "d1" });
    expect(createPasswordSend.mock.calls[0][3]).toBe(12);
    const second = await startUpload("anna");
    await putChunk("anna", second.body.uploadId, 1);
    await complete("anna", second.body.uploadId, { step: "d2" });
    expect(createPasswordSend.mock.calls[1][3]).toBe(5);
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
    expect(folder[0]).toBe("CloudShare-Dateianforderungen");
    expect(permissions).toBe(4);
    expect(label).toBe("Dateianforderung");
  });
});

describe("password by SMS: the Send only holds the password until sending", () => {
  it("keeps the Sends of a job with SMS as long as possible, also those created before the switch", async () => {
    const first = await startUpload("anna");
    await putChunk("anna", first.body.uploadId, 1);
    const { jobId } = (await complete("anna", first.body.uploadId, { password: STRONG, passwordDays: 3, step: "s1" })).json();
    expect(createPasswordSend.mock.calls[0][3]).toBe(3);

    const res = await app.inject({ method: "PATCH", url: `/api/jobs/${jobId}`, headers: as("anna"), payload: { passwordSms: "0170 1234567" } });
    expect(res.statusCode).toBe(200);
    expect(setSendExpiry).toHaveBeenCalledOnce();
    expect(Math.round((setSendExpiry.mock.calls[0][1] - Date.now()) / DAY)).toBe(30);
    expect(Math.round((getJob(jobId)!.separate![0].sendExpiresAt! - Date.now()) / DAY)).toBe(30);

    const second = await startUpload("anna", "b.pdf");
    await putChunk("anna", second.body.uploadId, 1);
    await complete("anna", second.body.uploadId, { password: STRONG, passwordDays: 3, step: "s2", jobId });
    expect(createPasswordSend.mock.calls[1][3]).toBe(30);
  });
});

describe("confidential messages", () => {
  const send = (payload: Record<string, unknown>) =>
    app.inject({ method: "POST", url: "/api/messages", headers: as("anna"), payload });

  it("stores the text only in a password-protected Send and attaches it to the mail's job", async () => {
    createMessageSend.mockClear();
    const res = await send({ text: "  Zugang: anna / Sommer2026  ", password: STRONG, messageDays: 7, step: "msg1" });
    expect(res.statusCode).toBe(200);
    const [label, text, password, expiresAt] = createMessageSend.mock.calls[0];
    expect(label).toBe("Vertrauliche Nachricht");
    expect(text).toBe("Zugang: anna / Sommer2026"); // trimmed
    expect(password).toBe(STRONG);
    expect(Math.round((expiresAt.getTime() - Date.now()) / DAY)).toBe(7);
    const body = res.json();
    expect(body.shareUrl).toBe("https://vault.example.test/#/send/msg/key123");
    const job = getJob(body.jobId)!;
    expect(job.owner).toBe("oid-anna");
    expect(job.separate?.[0].fileNames).toEqual(["Vertrauliche Nachricht"]);
  });

  it("is available for the chosen days: 5 by default, at most 30 (independent of the link validity)", async () => {
    const days = async (payload: Record<string, unknown>) => {
      createMessageSend.mockClear();
      await send({ text: "x", password: STRONG, ...payload });
      return Math.round((createMessageSend.mock.calls[0][3].getTime() - Date.now()) / DAY);
    };
    expect(await days({})).toBe(5);
    expect(await days({ messageDays: 60, expireDays: 90 })).toBe(30);
    expect(await days({ messageDays: 2, expireDays: 90 })).toBe(2);
  });

  it.each([
    ["an empty text", { text: "   ", password: STRONG }],
    ["a text over 1000 characters", { text: "x".repeat(1001), password: STRONG }],
    ["no text", { password: STRONG }],
    ["a weak password", { text: "geheim", password: "kurz" }],
    ["no password", { text: "geheim", noPassword: true }],
  ])("refuses %s", async (_name, payload) => {
    createMessageSend.mockClear();
    const res = await send(payload);
    expect(res.statusCode).toBe(400);
    expect(createMessageSend).not.toHaveBeenCalled();
  });

  it("tells the add-in the limits", async () => {
    const config = (await app.inject({ method: "GET", url: "/api/config" })).json();
    expect(config).toMatchObject({ messageMaxLength: 1000, sendExpireDays: 5, sendMaxExpireDays: 30 });
  });
});

describe("my shares", () => {
  const shares = [
    { id: "1", path: "/CloudShare/2026-11-13_ab/a.pdf", url: "u1", expiration: "2026-11-27", sharedAt: Date.now() - 2 * DAY, fileId: 11, isFolder: false, permissions: 1, label: "", size: 4096, hasPassword: true },
    { id: "2", path: "/CloudShare-Dateianforderungen/2026-11-13_cd", url: "u2", expiration: null, sharedAt: Date.now() - DAY, fileId: 12, isFolder: true, permissions: 4, label: "Dateianforderung", size: null, hasPassword: true },
    { id: "3", path: "/Privat/urlaub.jpg", url: "u3", expiration: null, sharedAt: Date.now(), fileId: 13, isFolder: false, permissions: 1, label: "", size: 4096, hasPassword: false },
    { id: "4", path: "/CloudShare/2026-09-01_ef/alt.pdf", url: "u4", expiration: null, sharedAt: Date.now() - 31 * DAY, fileId: 14, isFolder: false, permissions: 1, label: "", size: 1, hasPassword: false },
    { id: "5", path: "/CloudShare/2026-11-13_gh/offen.pdf", url: "u5", expiration: null, sharedAt: Date.now() - 3 * DAY, fileId: 15, isFolder: false, permissions: 1, label: "", size: 1, hasPassword: false },
  ];

  it("lists only CloudShare links with their status", async () => {
    nc.listPublicShares.mockResolvedValue(shares);
    nc.publicDownloads.mockResolvedValue({ count: 2, last: "2026-11-13T10:20:00+00:00" });
    nc.countFiles.mockResolvedValue(3);
    const res = await app.inject({ method: "GET", url: "/api/shares", headers: as("anna") });
    expect(res.statusCode).toBe(200);
    // Newest first, only CloudShare folders, only the last 30 days ("alt.pdf" is 31 days old).
    expect(res.json()).toEqual([
      expect.objectContaining({ id: "2", kind: "request", name: "Dateianforderung", received: 3, downloads: null, protected: true }),
      expect.objectContaining({ id: "1", kind: "file", name: "a.pdf", downloads: { count: 2, last: "2026-11-13T10:20:00+00:00" }, protected: true }),
      expect.objectContaining({ id: "5", kind: "file", name: "offen.pdf", protected: false }),
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

  describe("revoking counts as expired (CLEANUP_AFTER_DAYS=0: file to the trash right away)", () => {
    const fileShare = (id: string, path: string, expiration: string | null = "2099-01-01") => ({
      id, path, url: `u${id}`, expiration, sharedAt: Date.now(), fileId: 20, isFolder: false, permissions: 1, label: "",
      size: 1, hasPassword: true,
    });
    const revoke = (id: string) => app.inject({ method: "DELETE", url: `/api/shares/${id}`, headers: as("anna") });

    it("deletes the upload folder when no other valid link points to it", async () => {
      nc.listPublicShares.mockResolvedValue([fileShare("7", "/CloudShare/2026-11-13_ab12cd34/a.pdf")]);
      const res = await revoke("7");
      expect(res.json()).toMatchObject({ ok: true, file: "deleted", deleteAfterDays: 0 });
      expect(nc.deletePath).toHaveBeenCalledWith("uid", ["CloudShare", "2026-11-13_ab12cd34"]);
    });

    it("keeps the file while another valid link (e.g. inserted again) points to it", async () => {
      nc.listPublicShares.mockResolvedValue([
        fileShare("7", "/CloudShare/2026-11-13_ab12cd34/a.pdf"),
        fileShare("8", "/CloudShare/2026-11-13_ab12cd34/a.pdf"),
      ]);
      expect((await revoke("7")).json().file).toBe("kept");
      expect(nc.deletePath).not.toHaveBeenCalled();
    });

    it("never deletes received files of an upload link", async () => {
      nc.listPublicShares.mockResolvedValue([
        { ...fileShare("9", "/CloudShare-Dateianforderungen/2026-11-13_cd34ef56"), isFolder: true, permissions: 4 },
      ]);
      expect((await revoke("9")).json().file).toBe("kept");
      expect(nc.deletePath).not.toHaveBeenCalled();
    });
  });

  const link = (id: string, body: Record<string, unknown>) =>
    app.inject({ method: "POST", url: `/api/shares/${id}/link`, headers: as("anna"), payload: body });

  it("links a shared file again: new link with this mail's password and a password job", async () => {
    nc.listPublicShares.mockResolvedValue(shares);
    const res = await link("1", { password: STRONG, expireDays: 7 });
    expect(res.statusCode).toBe(200);
    expect(res.json()).toMatchObject({ kind: "file", fileName: "a.pdf", size: 4096 });
    const [segments, password, , permissions] = nc.createPublicShare.mock.calls[0];
    expect(segments).toEqual(["CloudShare", "2026-11-13_ab", "a.pdf"]);
    expect(password).toBe(STRONG);
    expect(permissions).toBeUndefined(); // read-only default
    expect(getJob(res.json().jobId)?.fileNames).toEqual(["a.pdf"]);
    expect(nc.deleteShare).not.toHaveBeenCalled(); // the old link stays valid
  });

  it("links a file again without password – no password job, only the download confirmation", async () => {
    nc.listPublicShares.mockResolvedValue(shares);
    const res = await link("1", { noPassword: true });
    expect(res.statusCode).toBe(200);
    expect(res.json().noPassword).toBe(true);
    expect(getJob(res.json().jobId)).toMatchObject({ noPassword: true });
    expect(nc.createPublicShare.mock.calls[0][1]).toBeUndefined();
    expect(createPasswordSend).not.toHaveBeenCalled();
  });

  it("links an upload folder again only with password, as a create-only share", async () => {
    nc.listPublicShares.mockResolvedValue(shares);
    expect((await link("2", { noPassword: true })).statusCode).toBe(400);
    const res = await link("2", { password: STRONG });
    expect(res.statusCode).toBe(200);
    expect(res.json().kind).toBe("request");
    expect(nc.createPublicShare.mock.calls[0][3]).toBe(4);
  });

  it("only links CloudShare shares again and checks the password", async () => {
    nc.listPublicShares.mockResolvedValue(shares);
    expect((await link("3", { password: STRONG })).statusCode).toBe(404);
    expect((await link("1", { password: "kurz" })).statusCode).toBe(400);
    expect(nc.createPublicShare).not.toHaveBeenCalled();
  });
});

describe("arming the password mail", () => {
  it("is only possible for the owner of the job", async () => {
    const id = await createJob("oid-anna", "anna@example.test", "a.pdf", "https://vault/#/send/x");
    const res = await app.inject({ method: "POST", url: `/api/jobs/${id}/arm`, headers: as("bernd"), payload: {} });
    expect(res.statusCode).toBe(404);
    expect(getJob(id)?.armed).toBeUndefined();
  });

  it("takes the other password recipient from the arm call, only if valid", async () => {
    const id = await createJob("oid-anna", "anna@example.test", "a.pdf", "https://vault/#/send/x");
    const arm = (passwordTo: unknown) =>
      app.inject({ method: "POST", url: `/api/jobs/${id}/arm`, headers: as("anna"), payload: { passwordTo } });
    await arm("keine-adresse");
    expect(getJob(id)?.passwordTo).toBeUndefined();
    await arm("kollege@example.test");
    expect(getJob(id)?.passwordTo).toBe("kollege@example.test");
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

describe("password by SMS (seven.io)", () => {
  const patch = (id: string, body: Record<string, unknown>) =>
    app.inject({ method: "PATCH", url: `/api/jobs/${id}`, headers: as("anna"), payload: body });

  it("offers SMS in the configuration when seven.io is set up", async () => {
    expect((await app.inject({ method: "GET", url: "/api/config" })).json()).toMatchObject({ smsEnabled: true, smsDefaultCountry: "49" });
  });

  it("stores the number normalized, answers it masked and switches between SMS and mail", async () => {
    const id = await createJob("oid-anna", "anna@example.test", "a.pdf", "https://vault/#/send/x", 3, "send-9");
    const res = await patch(id, { passwordSms: "0170 1234567" });
    expect(res.statusCode).toBe(200);
    expect(res.json().passwordSms).toBe("+49 170 •••• 67");
    expect(getJob(id)?.passwordSms).toBe("+491701234567");

    expect((await patch(id, { passwordSms: "12" })).statusCode).toBe(400);
    expect(getJob(id)?.passwordSms).toBe("+491701234567");

    await patch(id, { passwordTo: "kollege@example.test" }); // another address replaces the SMS
    expect(getJob(id)?.passwordSms).toBeUndefined();
    expect(getJob(id)?.passwordTo).toBe("kollege@example.test");
  });

  it("takes the number and the sender's profile when arming", async () => {
    const id = await createJob("oid-anna", "anna@example.test", "a.pdf", "https://vault/#/send/x", 3, "send-9");
    await app.inject({ method: "POST", url: `/api/jobs/${id}/arm`, headers: as("anna"), payload: { passwordSms: "+49 170 7654321" } });
    expect(getJob(id)).toMatchObject({ passwordSms: "+491707654321", senderName: "Anna Muster", department: "Vertrieb" });
  });

  it("accepts delivery reports only with the secret and updates the status idempotently", async () => {
    const { recordDelivery: record, deliveryFor: delivery } = await import("../src/deliveries.js");
    await record(["tok-dlr"], { owner: "oid-anna", channel: "sms", to: "+49 170 •••• 67", at: Date.now(), messageId: "771", status: "submitted" });
    const report = { webhook_event: "dlr", data: { msg_id: "771", status: "DELIVERED" } };

    expect((await app.inject({ method: "POST", url: "/webhooks/seven?token=falsch", payload: report })).statusCode).toBe(404);
    expect((await app.inject({ method: "POST", url: "/webhooks/seven", payload: report })).statusCode).toBe(404);
    expect(delivery("tok-dlr", "oid-anna")?.status).toBe("submitted");

    for (let i = 0; i < 2; i++) {
      const ok = await app.inject({ method: "POST", url: "/webhooks/seven?token=test-webhook-secret", payload: report });
      expect(ok.statusCode).toBe(200);
    }
    expect(delivery("tok-dlr", "oid-anna")).toMatchObject({ status: "delivered" });

    // A late intermediate report does not overwrite the final state; the classic form callback works too.
    await app.inject({
      method: "POST",
      url: "/webhooks/seven?token=test-webhook-secret",
      headers: { "content-type": "application/x-www-form-urlencoded" },
      payload: "msg_id=771&status=BUFFERED",
    });
    expect(delivery("tok-dlr", "oid-anna")?.status).toBe("delivered");
  });
});

describe("changing 'Kennwort abrufbar' after the first link", () => {
  const change = (user: string, id: string, maxAccess: unknown) =>
    app.inject({ method: "PATCH", url: `/api/jobs/${id}`, headers: as(user), payload: { maxAccess } });

  it("updates the Send and the number in the password mail, within the limit", async () => {
    const id = await createJob("oid-anna", "anna@example.test", "a.pdf", "https://vault/#/send/x", 3, "send-9");
    const res = await change("anna", id, 500);
    expect(res.statusCode).toBe(200);
    expect(res.json()).toEqual({ maxAccess: 20, passwordTo: "", passwordSms: "" });
    expect(setSendMaxAccess).toHaveBeenCalledWith("send-9", 20);
    expect(getJob(id)?.maxAccessCount).toBe(20);
  });

  it("refuses other users, sent jobs and jobs without a known Send", async () => {
    const id = await createJob("oid-anna", "anna@example.test", "a.pdf", "https://vault/#/send/x", 3, "send-9");
    expect((await change("bernd", id, 5)).statusCode).toBe(404);

    const old = await createJob("oid-anna", "anna@example.test", "a.pdf", "https://vault/#/send/x", 3);
    const fixed = await change("anna", old, 5);
    expect(fixed.statusCode).toBe(409);
    expect(fixed.json().code).toBe("SEND_FIXED");

    getJob(id)!.dispatched = true;
    expect((await change("anna", id, 5)).statusCode).toBe(409);
    expect(setSendMaxAccess).not.toHaveBeenCalled();
  });

  it("sets and clears the other password recipient, only with a valid address", async () => {
    const id = await createJob("oid-anna", "anna@example.test", "a.pdf", "https://vault/#/send/x", 3, "send-9");
    const set = (passwordTo: unknown) =>
      app.inject({ method: "PATCH", url: `/api/jobs/${id}`, headers: as("anna"), payload: { passwordTo } });

    const res = await set("  kollege@example.test ");
    expect(res.statusCode).toBe(200);
    expect(res.json().passwordTo).toBe("kollege@example.test");
    expect(getJob(id)?.passwordTo).toBe("kollege@example.test");
    expect(setSendMaxAccess).not.toHaveBeenCalled(); // only the recipient changed

    expect((await set("a@b.de, c@d.de")).statusCode).toBe(400);
    expect((await set(42)).statusCode).toBe(400);
    expect(getJob(id)?.passwordTo).toBe("kollege@example.test");

    expect((await set("")).json().passwordTo).toBe("");
    expect(getJob(id)?.passwordTo).toBeUndefined();
  });

  it("remembers the Send of a new share", async () => {
    const { body } = await startUpload("anna");
    await putChunk("anna", body.uploadId, 1);
    const { jobId } = (await complete("anna", body.uploadId)).json();
    expect(getJob(jobId)?.sendId).toBe("send-1");
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
