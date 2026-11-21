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
  listFiles: vi.fn(),
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
    appGraphToken: vi.fn(async () => "app-token"),
  };
});
vi.mock("../src/graph.js", async (importOriginal) => ({
  ...(await importOriginal<typeof import("../src/graph.js")>()),
  fetchProfile: vi.fn(async () => ({
    displayName: "Anna Muster", givenName: "Anna", surname: "Muster", jobTitle: "", department: "Vertrieb", company: "",
    email: "anna@example.test", phone: "", mobile: "", office: "",
  })),
  sendMail: vi.fn(async () => undefined),
  sendMailAsApp: vi.fn(async () => undefined),
  fetchContacts: vi.fn(async () => [
    { name: "Jonas Müller-Berg", emails: ["j.mueller@example.org"], phones: [{ number: "0171 1234567", type: "mobile" }] },
    { name: "Erika Beispiel", emails: ["erika@example.test"], phones: [] },
  ]),
  fetchSensitivityLabels: vi.fn(async () => [{ id: "aaaaaaaa-0000-4000-8000-000000000002", name: "Vertraulich", rank: 4, protection: false }]),
  ewsIdFor: vi.fn(async (_token: string, id: string) => (id === "AAMk-gone" ? null : `ews-${id}`)),
  sentMailState: vi.fn(async (_token: string, id: string) => (id === "AAMk-deleted" ? "deleted" : "ok")),
}));
// seven.io is not contacted: the SMS of "Kennwort erneut senden" goes to this double.
vi.mock("../src/sms.js", async (importOriginal) => ({
  ...(await importOriginal<typeof import("../src/sms.js")>()),
  sendSms: vi.fn(async () => "99000001"),
}));
vi.mock("../src/vaultwarden.js", () => ({
  createPasswordSend: vi.fn(async () => ({ url: "https://vault.example.test/#/send/abc", id: "send-1", expiresAt: Date.now() + 5 * 86_400_000 })),
  readSendPassword: vi.fn(async () => "KateGeteFi45"),
  deleteSend: vi.fn(async () => undefined),
  createMessageSend: vi.fn(async () => ({ url: "https://vault.example.test/#/send/msg/key123", id: "send-msg" })),
  MESSAGE_MAX_LENGTH: 1000,
  setSendMaxAccess: vi.fn(async () => undefined),
  setSendExpiry: vi.fn(async () => undefined),
  sendAccessCount: vi.fn(async () => null),
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
const { armJob, createJob, getJob } = await import("../src/jobs.js");
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
    expect(res.json()).toMatchObject({ thresholdMb: 5, attachmentPolicy: "suggest", sendMaxAccessCount: 3, sendMaxAccessLimit: 20 });
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
    ["GET", "/api/shares/1/mail"],
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
    expect(res.json().error).toBe("Nextcloud ist noch nicht verbunden."); // no language header: LANGUAGE
  });

  it("answers in the language of the add-in (X-CloudShare-Language), English for unknown ones", async () => {
    const english = await app.inject({
      method: "POST", url: "/api/uploads", headers: { ...as("carla"), "x-cloudshare-language": "en-US" }, payload: { fileName: "a.pdf", size: 100 },
    });
    expect(english.json()).toMatchObject({ code: "NEXTCLOUD_NOT_CONNECTED", error: "Nextcloud is not connected yet." });
    const french = await app.inject({
      method: "POST", url: "/api/uploads", headers: { ...as("carla"), "x-cloudshare-language": "fr-FR" }, payload: { fileName: "a.pdf", size: 100 },
    });
    expect(french.json().error).toBe("Nextcloud is not connected yet.");
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

  it("keeps one password per mail: further links join the first step, a second password is refused", async () => {
    const upload = async (name: string, body: Record<string, unknown>) => {
      const started = await startUpload("anna", name);
      await putChunk("anna", started.body.uploadId, 1);
      return complete("anna", started.body.uploadId, body);
    };
    // Step 1: two files selected together, then another file with "Gleiches Kennwort wie zuvor".
    const { jobId } = (await upload("a.pdf", { step: "s1", maxAccess: 2 })).json();
    await upload("b.pdf", { step: "s1", jobId, maxAccess: 2 });
    await upload("c.pdf", { step: "s1", jobId, maxAccess: 2 });
    // A second password in the same mail: refused before any link is created.
    const OTHER = "Zr8!mQ4#vT1x_Lp9Wk2s";
    nc.createPublicShare.mockClear();
    const refused = await upload("d.pdf", { step: "s2", jobId, password: OTHER, maxAccess: 5 });
    expect(refused.statusCode).toBe(409);
    expect(refused.json().code).toBe("ONE_PASSWORD");
    expect(nc.createPublicShare).not.toHaveBeenCalled();
    // The same for an upload link and a confidential message, and for a link without a step.
    for (const [url, payload] of [
      ["/api/requests", { jobId, step: "s3", password: OTHER }],
      ["/api/messages", { jobId, step: "s4", password: OTHER, text: "geheim" }],
      ["/api/requests", { jobId, password: OTHER }],
    ] as const) {
      const res = await app.inject({ method: "POST", url, headers: as("anna"), payload });
      expect(res.json().code).toBe("ONE_PASSWORD");
    }

    expect(createPasswordSend.mock.calls.map((c) => [c[1], c[2]])).toEqual([[STRONG, 2]]);
    expect(getJob(jobId)?.separate?.map((p) => [p.step, p.fileNames])).toEqual([["s1", ["a.pdf", "b.pdf", "c.pdf"]]]);
    // An upload link and a message join the step as well.
    const request = await app.inject({ method: "POST", url: "/api/requests", headers: as("anna"), payload: { jobId, step: "s1", password: STRONG } });
    expect(request.statusCode).toBe(200);
    expect(getJob(jobId)?.separate).toHaveLength(1);
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
    // A link with password lasts only as long as its password in Vaultwarden (Sends: at most 31 days).
    const maxDate = new Date(Date.now() + 30 * 86_400_000).toISOString().slice(0, 10);
    expect(res.json().expireDate).toBe(maxDate);
    expect(nc.createPublicShare.mock.calls[0][2]).toBe(maxDate);
    expect(createPasswordSend.mock.calls[0][2]).toBe(20);
    expect(createPasswordSend.mock.calls[0][3]).toBe(30);
    // The Send lives until the link expires, so the password can be sent again.
    expect(Math.round((createPasswordSend.mock.calls[0][4]! - Date.now()) / DAY)).toBe(30);
  });

  it("lets links without password last up to SHARE_MAX_EXPIRE_DAYS", async () => {
    const { body } = await startUpload("anna");
    await putChunk("anna", body.uploadId, 1);
    const res = await complete("anna", body.uploadId, { noPassword: true, expireDays: 1000 });
    expect(res.json().expireDate).toBe(new Date(Date.now() + 90 * 86_400_000).toISOString().slice(0, 10));
  });

  it("ends the recipients' access to the password at the latest with the link", async () => {
    const { body } = await startUpload("anna");
    await putChunk("anna", body.uploadId, 1);
    await complete("anna", body.uploadId, { expireDays: 3, passwordDays: 10 });
    expect(createPasswordSend.mock.calls[0][3]).toBe(3);
    expect(Math.round((createPasswordSend.mock.calls[0][4]! - Date.now()) / DAY)).toBe(3);
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
    expect(folder[0]).toBe("CloudShare-File-Requests");
    expect(permissions).toBe(4);
    expect(label).toBe("Dateianforderung");
    // Watched for received files once the mail is sent (UPLOAD_NOTICE) – a new folder holds nothing yet.
    expect(getJob(res.json().jobId)?.separate).toBeUndefined();
  });

  it("marks the upload link in a step, so the upload notification watches it", async () => {
    const res = await app.inject({ method: "POST", url: "/api/requests", headers: as("anna"), payload: { password: STRONG, step: "req-1" } });
    expect(getJob(res.json().jobId)?.separate?.[0].shares).toEqual([
      expect.objectContaining({ token: "abc", request: true, knownFiles: [] }),
    ]);
  });

  it("keeps the way of each link in its step: the same step by SMS adds the SMS to the password mail", async () => {
    const first = await app.inject({
      method: "POST", url: "/api/requests", headers: as("anna"), payload: { password: STRONG, step: "way-1", passwordTo: "" },
    });
    const { jobId } = first.json();
    expect(getJob(jobId)?.separate?.[0].deliveries).toEqual([{ channel: "recipients" }]);
    await app.inject({
      method: "POST", url: "/api/requests", headers: as("anna"),
      payload: { password: STRONG, step: "way-1", jobId, passwordSms: "0170 1234567" },
    });
    expect(getJob(jobId)?.separate?.[0].deliveries).toEqual([{ channel: "recipients" }, { channel: "sms", to: "+491701234567" }]);
    // A number SMS may not go to is refused before any link is created.
    const refused = await app.inject({
      method: "POST", url: "/api/requests", headers: as("anna"), payload: { password: STRONG, step: "way-1", jobId, passwordSms: "0800 1234567" },
    });
    expect(refused.statusCode).toBe(400);
    expect(getJob(jobId)?.separate).toHaveLength(1);
  });
});

describe("password by SMS: the link of the Send is never sent", () => {
  it("opens the Sends of a job with SMS as long as they live, also those created before the switch", async () => {
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
    // A new mail whose password goes by SMS: its Send lasts as long as it may (one password per mail – no second step here).
    await complete("anna", second.body.uploadId, { password: STRONG, passwordDays: 3, step: "s2", passwordSms: "0170 1234567" });
    expect(createPasswordSend.mock.calls[1][3]).toBe(30);
  });
});

describe("confidential messages", () => {
  const send = (payload: Record<string, unknown>) =>
    app.inject({ method: "POST", url: "/api/messages", headers: as("anna"), payload });
  // The messages created here would show in anna's "Verlauf" in later tests: revoke them.
  afterAll(async () => {
    await app.inject({ method: "DELETE", url: "/api/shares/msg-send-msg", headers: as("anna") });
  });

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

  it("lists the message in \"Verlauf\" without its key, and only its owner revokes it", async () => {
    createMessageSend.mockResolvedValueOnce({ url: "https://vault.example.test/#/send/acc-7/secretkey", id: "send-7" });
    expect((await send({ text: "geheim", password: STRONG, step: "msg7" })).statusCode).toBe(200);
    nc.userInfo.mockResolvedValue({ id: "anna", free: null });
    nc.listPublicShares.mockResolvedValue([]);
    const list = async (who: string) =>
      (await app.inject({ method: "GET", url: "/api/shares", headers: as(who) })).json() as Array<Record<string, unknown>>;

    const entry = (await list("anna")).find((s) => s.id === "msg-send-7")!;
    expect(entry).toMatchObject({ kind: "message", name: "Vertrauliche Nachricht", url: null, accessId: "acc-7", unsent: true, subject: null });
    expect(JSON.stringify(entry)).not.toContain("secretkey"); // CloudShare never keeps the key
    expect((await list("bernd")).some((s) => s.id === "msg-send-7")).toBe(false);

    const deleteSend = vi.mocked(vaultwarden.deleteSend);
    deleteSend.mockClear();
    expect((await app.inject({ method: "DELETE", url: "/api/shares/msg-send-7", headers: as("bernd") })).statusCode).toBe(404);
    const revoked = await app.inject({ method: "DELETE", url: "/api/shares/msg-send-7", headers: as("anna") });
    expect(revoked.json()).toMatchObject({ ok: true, message: true });
    expect(deleteSend).toHaveBeenCalledWith("send-7");
    expect((await list("anna")).some((s) => s.id === "msg-send-7")).toBe(false);
  });

  it("opens the sent mail of an entry: its current EWS ID, only for its owner", async () => {
    const { recordSentMail } = await import("../src/history.js");
    const keepUntil = Date.now() + DAY;
    recordSentMail(["msg-m1"], { owner: "oid-anna", subject: "Zugangsdaten", mailId: "AAMk-1", keepUntil });
    recordSentMail(["msg-m2"], { owner: "oid-anna", subject: "Alt", mailId: "AAMk-gone", keepUntil });
    recordSentMail(["msg-m3"], { owner: "oid-anna", subject: "Freigegebenes Postfach", keepUntil }); // no ID
    recordSentMail(["msg-m4"], { owner: "oid-anna", subject: "Gelöscht", mailId: "AAMk-deleted", keepUntil });
    const mail = (id: string, who = "anna") => app.inject({ method: "GET", url: `/api/shares/${id}/mail`, headers: as(who) });

    const res = await mail("msg-m1");
    expect(res.statusCode).toBe(200);
    expect(res.json()).toEqual({ ewsId: "ews-AAMk-1" }); // translated now, so it follows a moved mail
    expect((await mail("msg-m1", "bernd")).statusCode).toBe(404); // someone else's
    expect((await mail("msg-m2")).statusCode).toBe(404); // deleted in the mailbox
    expect((await mail("msg-m3")).statusCode).toBe(404); // sent from a shared mailbox
    expect((await mail("msg-none")).statusCode).toBe(404); // sent before this version
    // In Deleted Items or – after emptying it – in Recoverable Items: not opened, Outlook would still show it.
    const deleted = await mail("msg-m4");
    expect(deleted.statusCode).toBe(404);
    expect(deleted.json().error).toBe("Die gesendete Nachricht wurde gelöscht.");
  });

  it("tells the add-in the limits", async () => {
    const config = (await app.inject({ method: "GET", url: "/api/config" })).json();
    expect(config).toMatchObject({ messageMaxLength: 1000, sendExpireDays: 5, sendMaxExpireDays: 30 });
  });
});

describe("my shares", () => {
  /** YYYY-MM-DD `offset` days from today. */
  const day = (offset: number) => new Date(Date.now() + offset * DAY).toISOString().slice(0, 10);
  const shares = [
    { id: "1", path: "/CloudShare/2026-11-13_ab/a.pdf", url: "u1", expiration: day(12), sharedAt: Date.now() - 2 * DAY, fileId: 11, isFolder: false, permissions: 1, label: "", size: 4096, hasPassword: true },
    { id: "2", path: "/CloudShare-File-Requests/2026-11-13_cd", url: "u2", expiration: null, sharedAt: Date.now() - DAY, fileId: 12, isFolder: true, permissions: 4, label: "Dateianforderung", size: null, hasPassword: true },
    { id: "3", path: "/Privat/urlaub.jpg", url: "u3", expiration: null, sharedAt: Date.now(), fileId: 13, isFolder: false, permissions: 1, label: "", size: 4096, hasPassword: false },
    { id: "4", path: "/CloudShare/2026-09-01_ef/alt.pdf", url: "u4", expiration: day(-1), sharedAt: Date.now() - 31 * DAY, fileId: 14, isFolder: false, permissions: 1, label: "", size: 1, hasPassword: false },
    { id: "5", path: "/CloudShare/2026-11-13_gh/offen.pdf", url: "u5", expiration: null, sharedAt: Date.now() - 3 * DAY, fileId: 15, isFolder: false, permissions: 1, label: "", size: 1, hasPassword: false },
    { id: "6", path: "/CloudShare/2026-09-01_ij/lang.pdf", url: "u6", expiration: day(30), sharedAt: Date.now() - 60 * DAY, fileId: 16, isFolder: false, permissions: 1, label: "", size: 1, hasPassword: false },
  ];

  it("lists only CloudShare links with their status", async () => {
    nc.listPublicShares.mockResolvedValue(shares);
    nc.publicDownloads.mockResolvedValue({ count: 2, last: "2026-11-13T10:20:00+00:00" });
    nc.countFiles.mockResolvedValue(3);
    const res = await app.inject({ method: "GET", url: "/api/shares", headers: as("anna") });
    expect(res.statusCode).toBe(200);
    // Newest first, only CloudShare folders, every valid link however old ("lang.pdf": 60 days, valid without
    // password), no expired one ("alt.pdf" expired yesterday).
    expect(res.json()).toEqual([
      expect.objectContaining({ id: "2", kind: "request", name: "Dateianforderung", received: 3, downloads: null, protected: true }),
      expect.objectContaining({ id: "1", kind: "file", name: "a.pdf", downloads: { count: 2, last: "2026-11-13T10:20:00+00:00" }, protected: true }),
      expect.objectContaining({ id: "5", kind: "file", name: "offen.pdf", protected: false }),
      expect.objectContaining({ id: "6", kind: "file", name: "lang.pdf", protected: false }),
    ]);
  });

  it("tells an SMS seven.io refused from one it reported as not delivered later", async () => {
    const { recordDelivery: record } = await import("../src/deliveries.js");
    // u1: seven.io refused the SMS (no message ID). u5: accepted, later "not delivered".
    await record(["u1"], { owner: "oid-anna", channel: "sms", to: "+49 170 •••• 67", at: Date.now(), status: "failed" });
    await record(["u5"], { owner: "oid-anna", channel: "sms", to: "+49 170 •••• 68", at: Date.now(), messageId: "881", status: "failed" });
    nc.listPublicShares.mockResolvedValue(shares);
    nc.publicDownloads.mockResolvedValue({ count: 0, last: null });
    nc.countFiles.mockResolvedValue(0);
    const list = (await app.inject({ method: "GET", url: "/api/shares", headers: as("anna") })).json() as Array<{ id: string; password: unknown }>;
    expect(list.find((s) => s.id === "1")?.password).toMatchObject({ channel: "sms", status: "failed", refused: true });
    expect(list.find((s) => s.id === "5")?.password).toMatchObject({ channel: "sms", status: "failed", refused: false });
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
        { ...fileShare("9", "/CloudShare-File-Requests/2026-11-13_cd34ef56"), isFolder: true, permissions: 4 },
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
    nc.listFiles.mockResolvedValue(["schon-da.pdf"]);
    const res = await link("2", { password: STRONG, step: "req-again" });
    expect(res.statusCode).toBe(200);
    expect(res.json().kind).toBe("request");
    expect(nc.createPublicShare.mock.calls[0][3]).toBe(4);
    // Files received through the earlier link are not reported again by the upload notification.
    expect(getJob(res.json().jobId)?.separate?.[0].shares?.[0]).toMatchObject({ request: true, knownFiles: ["schon-da.pdf"] });
  });

  it("marks links whose mail was not sent (yet) and tells when CloudShare revokes them", async () => {
    nc.listPublicShares.mockResolvedValue(shares);
    nc.createPublicShare.mockResolvedValue("https://cloud.example.test/s/unsent1");
    const res = await link("1", { password: STRONG, expireDays: 7, step: "step-unsent", subject: "Entwurf Neubau" });
    const linked = { ...shares[0], id: "21", url: "https://cloud.example.test/s/unsent1", sharedAt: Date.now() };
    nc.listPublicShares.mockResolvedValue([...shares, linked]);
    nc.publicDownloads.mockResolvedValue({ count: 0, last: null });
    nc.countFiles.mockResolvedValue(0);
    const list = async () =>
      (await app.inject({ method: "GET", url: "/api/shares", headers: as("anna") })).json() as Array<{
        id: string;
        unsent: boolean;
        revokeAt: number | null;
      }>;

    const before = await list();
    expect(before.find((s) => s.id === "21")).toMatchObject({ unsent: true, subject: "Entwurf Neubau", mailId: null }); // the draft's subject until it is sent
    // The subject typed later (the pane sends it while the draft is open) replaces it – also if the draft is discarded.
    await app.inject({ method: "PATCH", url: `/api/jobs/${res.json().jobId}`, headers: as("anna"), payload: { subject: "Unterlagen Neubau" } });
    expect((await list()).find((s) => s.id === "21")).toMatchObject({ subject: "Unterlagen Neubau" });
    expect(before.find((s) => s.id === "21")!.revokeAt).toBeGreaterThan(Date.now() + 6 * DAY); // job expiry: 7 days
    expect(before.find((s) => s.id === "1")).toMatchObject({ unsent: false, revokeAt: null });

    // Downloaded already (e.g. sent from an Outlook without CloudShare): still unsent, but not revoked automatically.
    nc.publicDownloads.mockResolvedValue({ count: 1, last: "2026-11-16T09:04:00+00:00" });
    expect((await list()).find((s) => s.id === "21")).toMatchObject({ unsent: true, revokeAt: null });

    // "Send" pressed: the mail is on its way – no longer shown as not sent.
    await armJob(res.json().jobId, { graphToken: "t", tokenExpiresAt: Date.now() + 3_600_000, armedAt: Date.now() });
    expect((await list()).find((s) => s.id === "21")?.unsent).toBe(false);
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

  it("allows SMS only to the countries of SMS_ALLOWED_COUNTRIES and names them", async () => {
    // setup.ts: "Deutschland; AUT, +41"
    expect((await app.inject({ method: "GET", url: "/api/config" })).json()).toMatchObject({
      smsCountryCodes: ["49", "43", "41"],
      smsCountries: "Deutschland, Österreich und Schweiz",
    });
    const id = await createJob("oid-anna", "anna@example.test", "a.pdf", "https://vault/#/send/x", 3, "send-9");
    expect((await patch(id, { passwordSms: "+43 (0)664 1234567" })).statusCode).toBe(200);
    expect(getJob(id)?.passwordSms).toBe("+436641234567");

    const refused = await patch(id, { passwordSms: "+1 202 555 0123" });
    expect(refused.statusCode).toBe(400);
    expect(refused.json().error).toContain("Deutschland, Österreich und Schweiz");
    expect(getJob(id)?.passwordSms).toBe("+436641234567"); // unchanged

    const service = await patch(id, { passwordSms: "0800 1234567" }); // SMS_BLOCKED_NUMBERS (default)
    expect(service.statusCode).toBe(400);
    expect(service.json().error).toContain("0800, 0180 oder 0900");
    expect((await patch(id, { passwordSms: "030 1234567" })).statusCode).toBe(200); // landline is fine
    expect(getJob(id)?.passwordSms).toBe("+49301234567");
    expect((await app.inject({ method: "GET", url: "/api/config" })).json()).toMatchObject({
      smsAllowedNumbers: "",
      smsBlockedNumbers: "^\\+49(800|900|180)",
    });

    // Arming with a number of another country keeps the number the taskpane set.
    await app.inject({ method: "POST", url: `/api/jobs/${id}/arm`, headers: as("anna"), payload: { passwordSms: "0044 7700 900123" } });
    expect(getJob(id)?.passwordSms).toBe("+49301234567");
  });

  it("takes the number and the sender's profile when arming", async () => {
    const id = await createJob("oid-anna", "anna@example.test", "a.pdf", "https://vault/#/send/x", 3, "send-9");
    await app.inject({ method: "POST", url: `/api/jobs/${id}/arm`, headers: as("anna"), payload: { passwordSms: "+49 170 7654321" } });
    expect(getJob(id)).toMatchObject({ passwordSms: "+491707654321", senderName: "Anna Muster", department: "Vertrieb" });
  });

  it("fixes the sender's signature for the password mail when arming (own one, filled from the profile)", async () => {
    const { resetSettings, saveSettings } = await import("../src/settings.js");
    saveSettings("oid-anna", { signature: { format: "text", content: "{displayName} · {department}\n<Team>" } });
    try {
      const id = await createJob("oid-anna", "anna@example.test", "a.pdf", "https://vault/#/send/x", 3, "send-10");
      await app.inject({ method: "POST", url: `/api/jobs/${id}/arm`, headers: as("anna"), payload: {} });
      expect(getJob(id)?.signature).toMatch(/^<p style="[^"]*">Anna Muster · Vertrieb<br>&lt;Team&gt;<\/p>$/);
    } finally {
      resetSettings("oid-anna");
    }
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

describe("showing and sending the password again", () => {
  /** The sent mail went to Erika, Cc to her boss and Bcc to the archive. */
  const DISTRIBUTION = { to: ["erika@example.test"], cc: ["chef@example.test"], bcc: ["archiv@example.test"] };
  const share = { id: "41", path: "/CloudShare/2026-11-13_ab/vertrag.pdf", url: "https://cloud.example.test/s/tokvertrag",
    expiration: "2099-12-01", sharedAt: Date.now(), fileId: 41, isFolder: false, permissions: 1, label: "", size: 1, hasPassword: true };
  const register = async () => {
    const { registerPasswordRef } = await import("../src/passwordRefs.js");
    await registerPasswordRef({
      owner: "oid-anna", ownerUpn: "anna@example.test", sendId: "send-r1", tokens: ["tokvertrag"], fileNames: ["vertrag.pdf"],
      keepUntil: Date.now() + 10 * DAY, subject: "Vertrag", recipients: DISTRIBUTION, sms: "+491701234567",
      maxAccessCount: 3,
    });
  };
  const resend = (user: string, payload: Record<string, unknown>) =>
    app.inject({ method: "POST", url: "/api/shares/41/password/send", headers: as(user), payload });
  const sendMail = async () => vi.mocked((await import("../src/graph.js")).sendMail);
  const sendSms = async () => vi.mocked((await import("../src/sms.js")).sendSms);

  beforeEach(async () => {
    await register();
    nc.listPublicShares.mockResolvedValue([share]);
    nc.publicDownloads.mockResolvedValue({ count: 0, last: null });
    (await sendMail()).mockClear();
    (await sendSms()).mockClear();
  });

  it("shows the password only to its sender and only for its links", async () => {
    const res = await app.inject({ method: "GET", url: "/api/shares/41/password", headers: as("anna") });
    expect(res.statusCode).toBe(200);
    expect(res.json()).toEqual({ password: "KateGeteFi45", recipients: DISTRIBUTION, otherTo: [], sms: "+491701234567" });
    expect(vi.mocked(vaultwarden.readSendPassword)).toHaveBeenCalledWith("send-r1");
    expect((await app.inject({ method: "GET", url: "/api/shares/41/password", headers: as("bernd") })).statusCode).toBe(404);
    expect((await app.inject({ method: "GET", url: "/api/shares/99/password", headers: as("anna") })).statusCode).toBe(404);
    const list = (await app.inject({ method: "GET", url: "/api/shares", headers: as("anna") })).json() as Array<{ passwordAvailable: boolean }>;
    expect(list[0].passwordAvailable).toBe(true);
  });

  it("sends it again to the recipients the sender ticks – each in a mail of their own, only addresses the message went to", async () => {
    const all = [...DISTRIBUTION.to, ...DISTRIBUTION.cc, ...DISTRIBUTION.bcc];
    const res = await resend("anna", { channel: "recipients", to: [all[0], all[all.length - 1]] });
    expect(res.statusCode).toBe(200);
    expect(res.json()).toEqual({ channel: "recipients", to: `${all[0]}, ${all[all.length - 1]}` });
    expect((await sendMail()).mock.calls.map((call) => call[1])).toEqual([
      { to: [all[0]], cc: [], bcc: [] }, { to: [all[all.length - 1]], cc: [], bcc: [] }]);
    (await sendMail()).mockClear();
    const stranger = await resend("anna", { channel: "recipients", to: ["kollege@example.test"] });
    expect(stranger.statusCode).toBe(400);
    expect((await sendMail()).mock.calls).toHaveLength(0);
  });

  it("sends it again to the recipients of the message like the first time – To, Cc and Bcc as they were", async () => {
    createPasswordSend.mockClear();
    const res = await resend("anna", { channel: "recipients" });
    expect(res.statusCode).toBe(200);
    expect(res.json()).toEqual({ channel: "recipients", to: "erika@example.test, chef@example.test" }); // Bcc not shown
    const [, password, maxAccess, days, keepUntil] = createPasswordSend.mock.calls[0];
    expect([password, maxAccess, days]).toEqual(["KateGeteFi45", 3, 5]); // a new Send with the same password
    expect(Math.round((keepUntil! - Date.now()) / DAY)).toBe(10); // deleted together with the links
    expect((await sendMail()).mock.calls).toHaveLength(1); // one mail, as the first time
    const [token, recipients, subject, html] = (await sendMail()).mock.calls[0];
    expect(token).toBe("graph-token");
    expect(recipients).toEqual(DISTRIBUTION);
    expect(subject).toContain("Vertrag");
    expect(html).toContain("https://vault.example.test/#/send/abc");
    expect(html).not.toContain("KateGeteFi45"); // never the password itself
    const { deliveryFor } = await import("../src/deliveries.js");
    expect(deliveryFor("tokvertrag", "oid-anna")).toMatchObject({ channel: "mail", to: "erika@example.test, chef@example.test" });
  });

  it("sends it again to one other address – like \"Kennwort an\" when sending; a list is refused", async () => {
    createPasswordSend.mockClear();
    // Several addresses (an older add-in): refused before anything goes out – several recipients get it with "recipients".
    const list = await resend("anna", { channel: "mail", to: "kollege@example.test; buergerin@example.test" });
    expect(list.statusCode).toBe(400);
    expect(list.json().error).toContain("genau eine gültige E-Mail-Adresse");
    expect(createPasswordSend).not.toHaveBeenCalled();
    expect((await sendMail()).mock.calls).toHaveLength(0);

    const res = await resend("anna", { channel: "mail", to: " kollege@example.test " });
    expect(res.statusCode).toBe(200);
    expect(res.json()).toEqual({ channel: "mail", to: "kollege@example.test" });
    expect(createPasswordSend).toHaveBeenCalledOnce();
    expect((await sendMail()).mock.calls.map((c) => c[1])).toEqual([{ to: ["kollege@example.test"], cc: [], bcc: [] }]);
    // Suggested again next time.
    const shown = (await app.inject({ method: "GET", url: "/api/shares/41/password", headers: as("anna") })).json();
    expect(shown.otherTo).toEqual(["kollege@example.test"]);
  });

  it("sends it again by SMS to a number of the sender's choice", async () => {
    const res = await resend("anna", { channel: "sms", to: "0171 7654321" });
    expect(res.statusCode).toBe(200);
    expect(res.json()).toEqual({ channel: "sms", to: "+49 171 •••• 21" });
    const [number, text] = (await sendSms()).mock.calls[0];
    expect(number).toBe("+491717654321");
    expect(text).toContain("KateGeteFi45");
    expect(text).toContain("\"vertrag.pdf\""); // names the file
    const { deliveryFor } = await import("../src/deliveries.js");
    expect(deliveryFor("tokvertrag", "oid-anna")).toMatchObject({ channel: "sms", messageId: "99000001", status: "submitted" });
  });

  it("refuses invalid targets and other people's links", async () => {
    expect((await resend("anna", { channel: "mail", to: "keine adresse" })).statusCode).toBe(400);
    expect((await resend("anna", { channel: "sms", to: "12" })).statusCode).toBe(400);
    expect((await resend("anna", { channel: "sms", to: "+44 7700 900123" })).statusCode).toBe(400); // not allowed
    expect((await sendSms()).mock.calls).toHaveLength(0);
    expect((await resend("anna", { channel: "fax", to: "x" })).statusCode).toBe(400);
    expect((await resend("bernd", { channel: "mail", to: "bernd@example.test" })).statusCode).toBe(404);
    expect((await sendMail()).mock.calls).toHaveLength(0);
  });

  it("deletes the Send once the last link of the step is revoked", async () => {
    const deleteSend = vi.mocked(vaultwarden.deleteSend);
    deleteSend.mockClear();
    const res = await app.inject({ method: "DELETE", url: "/api/shares/41", headers: as("anna") });
    expect(res.statusCode).toBe(200);
    expect(deleteSend).toHaveBeenCalledWith("send-r1");
    expect((await app.inject({ method: "GET", url: "/api/shares/41/password", headers: as("anna") })).statusCode).toBe(404);
  });

  it("tells the sender – without password – when seven.io reports the SMS as not delivered later", async () => {
    const { recordDelivery } = await import("../src/deliveries.js");
    await recordDelivery(["tokvertrag"], { owner: "oid-anna", channel: "sms", to: "+49 170 •••• 67", at: Date.now(), messageId: "551", status: "submitted" });
    const sendMailAsApp = vi.mocked((await import("../src/graph.js")).sendMailAsApp);
    sendMailAsApp.mockClear();
    const report = { webhook_event: "dlr", data: { msg_id: "551", status: "NOTDELIVERED" } };
    for (let i = 0; i < 2; i++) {
      await app.inject({ method: "POST", url: "/webhooks/seven?token=test-webhook-secret", payload: report });
    }
    expect(sendMailAsApp).toHaveBeenCalledOnce(); // once, not for the repeated report
    const [, owner, to, subject, html] = sendMailAsApp.mock.calls[0];
    expect([owner, to]).toEqual(["oid-anna", "anna@example.test"]);
    expect(subject).toBe("SMS nicht zugestellt – Vertrag");
    expect(html).toContain("wurde nicht zugestellt");
    expect(html).not.toContain("KateGeteFi45");
  });
});

describe("personal settings (/api/settings)", () => {
  const get = (user: string) => app.inject({ method: "GET", url: "/api/settings", headers: as(user) });
  const put = (user: string, payload: unknown) => app.inject({ method: "PUT", url: "/api/settings", headers: as(user), payload: payload as object });

  it("starts with the defaults of the environment, within what the server offers", async () => {
    const res = await get("sina");
    expect(res.statusCode).toBe(200);
    expect(res.json()).toEqual({
      settings: {
        downloadNotice: true, uploadNotice: true, smsNotice: true, proofReminder: true, channels: ["recipients", "mail", "sms"], defaultChannel: "recipients",
        language: "auto", recipientLanguage: "de", signature: null, // recipients: LANGUAGE (de in the tests)
      },
      defaults: expect.objectContaining({ channels: ["recipients", "mail", "sms"], language: "auto" }),
      available: { downloadNotice: true, uploadNotice: true, smsNotice: true, proofReminder: true, channels: ["recipients", "mail", "sms"], languages: ["de", "en"] },
    });
  });

  it("keeps each user's own choices and only those – a reset brings back the defaults", async () => {
    const saved = await put("tom", {
      channels: ["mail", "sms"], defaultChannel: "sms", downloadNotice: false, language: "en",
      signature: { format: "text", content: "Tom\n\nTeam Bau" },
    });
    expect(saved.statusCode).toBe(200);
    expect(saved.json().settings).toMatchObject({
      channels: ["mail", "sms"], defaultChannel: "sms", downloadNotice: false, uploadNotice: true, language: "en",
      signature: { format: "text", content: "Tom\n\nTeam Bau" },
    });
    expect((await get("tom")).json().settings.defaultChannel).toBe("sms");
    expect((await get("sina")).json().settings.defaultChannel).toBe("recipients"); // other users untouched

    // An empty signature is none: the organisation's applies again.
    expect((await put("tom", { signature: { format: "text", content: "  " } })).json().settings.signature).toBeNull();

    const reset = await app.inject({ method: "DELETE", url: "/api/settings", headers: as("tom") });
    expect(reset.json().settings).toMatchObject({ channels: ["recipients", "mail", "sms"], downloadNotice: true, language: "auto" });
  });

  it("refuses invalid settings and needs a sign-in", async () => {
    for (const payload of [
      { channels: [] },
      { channels: ["fax"] },
      { defaultChannel: "fax" },
      { language: "xx" },
      { downloadNotice: "yes" },
      { signature: { format: "pdf", content: "x" } },
      { signature: { format: "html", content: "<p>markup</p>" } }, // plain text only
      { signature: { format: "html", nodes: [{ tag: "b", children: ["x"] }] } },
      { signature: { format: "text", content: "x".repeat(20_001) } },
    ]) {
      const res = await put("tom", payload);
      expect(res.statusCode, JSON.stringify(payload).slice(0, 60)).toBe(400);
      expect(res.json().code).toBe("SETTINGS_INVALID");
    }
    expect((await app.inject({ method: "GET", url: "/api/settings" })).statusCode).toBe(401);
  });

  it("keeps a preselected way among the offered ones", async () => {
    const res = await put("ida", { channels: ["recipients"], defaultChannel: "sms" });
    expect(res.json().settings).toMatchObject({ channels: ["recipients"], defaultChannel: "recipients" });
  });
});

describe("delivery proof (Zustellnachweis)", () => {
  const form = (fields: Record<string, string>) => ({
    headers: { "content-type": "application/x-www-form-urlencoded" },
    payload: new URLSearchParams(fields).toString(),
  });

  it("records the file with its fingerprint and puts the gateway into the mail", async () => {
    nc.createPublicShare.mockResolvedValue("https://cloud.example.test/s/proof-file");
    const { body } = await startUpload("anna", "Angebot.pdf", 100);
    const content = Buffer.from("y".repeat(100));
    await putChunk("anna", body.uploadId, 1, content);
    const res = await complete("anna", body.uploadId, { proof: true, client: { host: "Outlook", platform: "PC" } });
    expect(res.statusCode).toBe(200);
    const { shareUrl, proofId } = res.json();
    expect(shareUrl).toBe("https://cloudshare.example.test/r/s/proof-file");
    expect(proofId).toMatch(/^\d{8}-\d{4}-[2-9A-HJ-NP-Z]{4}$/);

    const { caseById } = await import("../src/proof.js");
    const created = caseById(proofId)!.events.find((e) => e.type === "link.created")!;
    const { createHash } = await import("node:crypto");
    expect(created.data).toMatchObject({
      token: "proof-file", name: "Angebot.pdf", size: 100, protection: "password",
      sha256: createHash("sha256").update(content).digest("hex"),
    });
    expect(caseById(proofId)!.events[0].data).toMatchObject({ client: { host: "Outlook", platform: "PC" } });
  });

  it("without the box ticked: the Nextcloud link as before, no case", async () => {
    nc.createPublicShare.mockResolvedValue("https://cloud.example.test/s/no-proof");
    const { body } = await startUpload("anna");
    await putChunk("anna", body.uploadId, 1);
    const res = await complete("anna", body.uploadId);
    expect(res.json()).toMatchObject({ shareUrl: "https://cloud.example.test/s/no-proof", proofId: null });
  });

  it("lets only a recipient of the sent mail confirm their address with a code, then opens the link", async () => {
    nc.createPublicShare.mockResolvedValue("https://cloud.example.test/s/proof-gw");
    const { body } = await startUpload("anna");
    await putChunk("anna", body.uploadId, 1);
    const { proofId } = (await complete("anna", body.uploadId, { proof: true })).json();
    const sendMailAsApp = vi.mocked((await import("../src/graph.js")).sendMailAsApp);
    sendMailAsApp.mockClear();

    // Before the mail was sent: nobody can get a code yet.
    const early = await app.inject({ method: "GET", url: "/r/s/proof-gw" });
    expect(early.statusCode).toBe(409);
    expect(early.headers["content-security-policy"]).toContain("default-src 'none'");

    const { recordMailSent, caseById } = await import("../src/proof.js");
    recordMailSent(["proof-gw"], { subject: "Unterlagen", messageId: "<m@example.test>", from: "anna@example.test", to: ["erika@example.test"], cc: [], bcc: [] });
    const page = await app.inject({ method: "GET", url: "/r/s/proof-gw" });
    expect(page.statusCode).toBe(200);
    expect(page.body).toContain('action="/r/s/proof-gw/code"');

    // Not a recipient: the same answer, but no mail.
    const stranger = await app.inject({ method: "POST", url: "/r/s/proof-gw/code", ...form({ address: "mallory@example.test" }) });
    expect(stranger.statusCode).toBe(200);
    expect(sendMailAsApp).not.toHaveBeenCalled();

    // A recipient (any spelling): the code comes from the sender's mailbox – with the sender's signature below.
    const { config } = await import("../src/config.js");
    const footer = config.mailFooter;
    config.mailFooter = "Viele Grüße\n{email}";
    const asked = await app.inject({ method: "POST", url: "/r/s/proof-gw/code", ...form({ address: "Erika@Example.test" }) });
    expect(asked.statusCode).toBe(200);
    // The mail goes out after the answer (same timing for every address).
    await vi.waitFor(() => expect(sendMailAsApp).toHaveBeenCalledOnce());
    config.mailFooter = footer;
    expect(asked.body.replaceAll("erika@example.test", "X")).toBe(stranger.body.replaceAll("mallory@example.test", "X"));
    const [, owner, to, subject, html] = sendMailAsApp.mock.calls[0];
    expect([owner, to]).toEqual(["oid-anna", "erika@example.test"]);
    // The code mail explains itself: what was sent, the code, the next steps, what to do if nobody asked for it.
    expect(subject).toContain("Ihr Bestätigungscode für");
    expect(html).toContain("So geht es weiter");
    expect(html).toContain("ignorieren");
    expect(String(html)).toMatch(/ignorieren[\s\S]*Viele Grüße<br>anna@example\.test/);
    const code = />(\d{6})</.exec(String(html))![1];
    // The button in the mail: the code form filled in – the GET alone confirms nothing (mail scanners open links).
    expect(String(html)).toContain(`/r/s/proof-gw/confirm?address=erika%40example.test&amp;code=${code}`);
    const prefilled = await app.inject({ method: "GET", url: `/r/s/proof-gw/confirm?address=erika%40example.test&code=${code}` });
    expect(prefilled.body).toContain(`value="${code}"`);
    expect(caseById(proofId)!.events.some((e) => e.type === "recipient.verified")).toBe(false);

    const wrong =await app.inject({ method: "POST", url: "/r/s/proof-gw/verify", ...form({ address: "erika@example.test", code: code === "000000" ? "111111" : "000000" }) });
    expect(wrong.statusCode).toBe(400);
    // An address that never got a code gets exactly the same answer as a wrong code – nothing to learn by trying.
    const strangerTry = await app.inject({ method: "POST", url: "/r/s/proof-gw/verify", ...form({ address: "mallory@example.test", code: "123456" }) });
    expect(strangerTry.statusCode).toBe(400);
    expect(strangerTry.body.replaceAll("mallory@example.test", "X")).toBe(wrong.body.replaceAll("erika@example.test", "X"));
    expect(wrong.body).toContain("Der Code stimmt nicht.");
    expect(wrong.body).toContain('action="/r/s/proof-gw/verify"');
    // No code (any more) – expired, entered wrongly too often or never asked for: back to the address, filled in, for a
    // new code; from the button in the mail as from the form.
    for (const res of [
      await app.inject({ method: "GET", url: "/r/s/proof-gw/confirm?address=bernd%40example.test&code=123456" }),
      await app.inject({ method: "POST", url: "/r/s/proof-gw/verify", ...form({ address: "bernd@example.test", code: "123456" }) }),
    ]) {
      expect(res.body).toContain('action="/r/s/proof-gw/code"');
      expect(res.body).toContain('value="bernd@example.test"');
      expect(res.body).toContain("Der Code ist abgelaufen oder nicht mehr gültig.");
    }
    const right = await app.inject({
      method: "POST", url: "/r/s/proof-gw/verify", ...form({ address: "erika@example.test", code }),
      headers: { "content-type": "application/x-www-form-urlencoded", "user-agent": "TestBrowser/1.0" },
    });
    expect(right.statusCode).toBe(303);
    expect(right.headers.location).toBe("https://cloud.example.test/s/proof-gw");
    const cookie = String(right.headers["set-cookie"]).split(";")[0];
    expect(String(right.headers["set-cookie"])).toContain("HttpOnly");

    // The same browser comes back: straight on, recorded again. A forged cookie gets the form.
    const again = await app.inject({ method: "GET", url: "/r/s/proof-gw", headers: { cookie } });
    expect(again.statusCode).toBe(303);
    const forged = await app.inject({ method: "GET", url: "/r/s/proof-gw", headers: { cookie: `${cookie.slice(0, -2)}xx` } });
    expect(forged.statusCode).toBe(200);

    const types = caseById(proofId)!.events.map((e) => e.type);
    expect(types).toEqual(expect.arrayContaining(["message.sent", "recipient.codeSent", "recipient.codeFailed", "recipient.verified", "link.opened"]));
    expect(caseById(proofId)!.events.find((e) => e.type === "recipient.verified")!.data).toMatchObject({
      address: "erika@example.test", userAgent: "TestBrowser/1.0",
    });
  });

  it("hands the PDF only to the owner, through a signed address that expires", async () => {
    nc.createPublicShare.mockResolvedValue("https://cloud.example.test/s/proof-pdf");
    const { body } = await startUpload("anna");
    await putChunk("anna", body.uploadId, 1);
    const { proofId } = (await complete("anna", body.uploadId, { proof: true })).json();

    expect((await app.inject({ method: "POST", url: `/api/proofs/${proofId}/link`, headers: as("bernd") })).statusCode).toBe(404);
    const { url } = (await app.inject({ method: "POST", url: `/api/proofs/${proofId}/link`, headers: as("anna") })).json();
    expect(url.startsWith(`https://cloudshare.example.test/proof/${proofId}.pdf?`)).toBe(true);
    const path = url.replace("https://cloudshare.example.test", "");
    const pdf = await app.inject({ method: "GET", url: path });
    expect(pdf.statusCode).toBe(200);
    expect(pdf.headers["content-type"]).toBe("application/pdf");
    expect(pdf.rawPayload.subarray(0, 8).toString("latin1")).toBe("%PDF-1.4");
    expect(pdf.headers["content-disposition"]).toContain(`${proofId}.pdf`);

    const tampered = await app.inject({ method: "GET", url: path.replace(/until=\d+/, "until=99999999999999") });
    expect(tampered.statusCode).toBe(403);
  });
});

describe("sensitivity labels (/api/labels)", () => {
  it("returns the tenant's labels via Graph, kept per user and language", async () => {
    const graph = await import("../src/graph.js");
    const fetchLabels = vi.mocked(graph.fetchSensitivityLabels);
    fetchLabels.mockClear();
    const first = await app.inject({ method: "GET", url: "/api/labels", headers: { ...as("lena"), "x-cloudshare-language": "de" } });
    expect(first.json()).toEqual({ labels: [{ id: "aaaaaaaa-0000-4000-8000-000000000002", name: "Vertraulich", rank: 4, protection: false }] });
    await app.inject({ method: "GET", url: "/api/labels", headers: { ...as("lena"), "x-cloudshare-language": "de" } });
    expect(fetchLabels).toHaveBeenCalledOnce();
    expect(fetchLabels.mock.calls[0][1]).toBe("de");
  });

  it("without the permission: labels null (the pane uses Outlook's catalog)", async () => {
    const graph = await import("../src/graph.js");
    vi.mocked(graph.fetchSensitivityLabels).mockRejectedValueOnce(new Error("Graph sensitivityLabels failed (HTTP 403)"));
    const res = await app.inject({ method: "GET", url: "/api/labels", headers: as("moritz") });
    expect(res.statusCode).toBe(200);
    expect(res.json()).toEqual({ labels: null });
  });
});

describe("feedback (/api/feedback)", () => {
  it("mails a problem report from the user's mailbox to FEEDBACK_MAIL – only known details", async () => {
    const graph = await import("../src/graph.js");
    const sendMail = vi.mocked(graph.sendMail);
    sendMail.mockClear();
    const config = (await app.inject({ method: "GET", url: "/api/config" })).json();
    expect(config.feedback).toBe(true);
    expect(JSON.stringify(config)).not.toContain("cloudshare-support"); // the address stays on the server
    const res = await app.inject({
      method: "POST", url: "/api/feedback", headers: as("anna"),
      payload: { kind: "bug", text: "Der Link landet in der Karte.", diagnostics: { outlook: "Outlook 16.0 (PC)", secret: "nein" } },
    });
    expect(res.statusCode).toBe(200);
    expect(sendMail).toHaveBeenCalledOnce();
    const [, recipients, subject, html, from] = sendMail.mock.calls[0];
    expect(recipients).toEqual({ to: ["cloudshare-support@example.test"], cc: [], bcc: [] });
    expect(subject).toBe("CloudShare: Fehlermeldung von Anna Muster");
    expect(html).toContain("Der Link landet in der Karte.");
    expect(html).toContain("Outlook 16.0 (PC)");
    expect(html).not.toContain("nein");
    expect(from).toBeUndefined(); // from the user's own mailbox
  });

  it("refuses an empty text, and answers 404 without FEEDBACK_MAIL", async () => {
    const empty = await app.inject({ method: "POST", url: "/api/feedback", headers: as("anna"), payload: { kind: "idea", text: " " } });
    expect(empty.statusCode).toBe(400);
    const { config } = await import("../src/config.js");
    const address = config.feedbackMail;
    config.feedbackMail = "";
    try {
      const off = await app.inject({ method: "POST", url: "/api/feedback", headers: as("anna"), payload: { kind: "idea", text: "Hallo" } });
      expect(off.statusCode).toBe(404);
      expect(off.json().code).toBe("FEEDBACK_OFF");
      expect((await app.inject({ method: "GET", url: "/api/config" })).json().feedback).toBe(false);
    } finally {
      config.feedbackMail = address;
    }
  });
});

describe("contact suggestions (/api/contacts)", () => {
  it("matches names and addresses, or the numbers for SMS – Graph asked once per user", async () => {
    const graph = await import("../src/graph.js");
    const fetchContacts = vi.mocked(graph.fetchContacts);
    fetchContacts.mockClear();
    const mail = await app.inject({ method: "GET", url: "/api/contacts?q=m%C3%BCl&kind=mail", headers: as("lena") });
    expect(mail.json()).toEqual({ contacts: [{ name: "Jonas Müller-Berg", value: "j.mueller@example.org" }] });
    const sms = await app.inject({ method: "GET", url: "/api/contacts?q=0171&kind=sms", headers: as("lena") });
    expect(sms.json()).toEqual({ contacts: [{ name: "Jonas Müller-Berg", value: "0171 1234567", type: "mobile" }] });
    expect(fetchContacts).toHaveBeenCalledOnce();
    // The address book page: every address A to Z.
    const book = await app.inject({ method: "GET", url: "/api/contacts?kind=mail&all=1", headers: as("lena") });
    expect(book.json().contacts.map((c: { name: string }) => c.name)).toEqual(["Erika Beispiel", "Jonas Müller-Berg"]);
    expect(fetchContacts).toHaveBeenCalledOnce();
  });

  it("without the permission: contacts null (the pane stops asking)", async () => {
    const graph = await import("../src/graph.js");
    vi.mocked(graph.fetchContacts).mockRejectedValueOnce(new Error("Graph contacts failed (HTTP 403)"));
    const res = await app.inject({ method: "GET", url: "/api/contacts?q=er", headers: as("moritz") });
    expect(res.statusCode).toBe(200);
    expect(res.json()).toEqual({ contacts: null });
  });
});

describe("delivery proof – guessing addresses", () => {
  const form = (fields: Record<string, string>) => ({
    headers: { "content-type": "application/x-www-form-urlencoded" },
    payload: new URLSearchParams(fields).toString(),
  });

  it("locks code requests of a link for an hour after five addresses that are no recipients, and tells the sender", async () => {
    nc.createPublicShare.mockResolvedValue("https://cloud.example.test/s/proof-guess");
    const { body } = await startUpload("anna");
    await putChunk("anna", body.uploadId, 1);
    const { proofId } = (await complete("anna", body.uploadId, { proof: true })).json();
    const { recordMailSent, caseById } = await import("../src/proof.js");
    recordMailSent(["proof-guess"], { subject: "Angebot", messageId: "<g@example.test>", from: "anna@example.test", to: ["erika@example.test"], cc: [], bcc: [] });
    const sendMailAsApp = vi.mocked((await import("../src/graph.js")).sendMailAsApp);
    sendMailAsApp.mockClear();

    for (let i = 1; i <= 5; i++) {
      const res = await app.inject({ method: "POST", url: "/r/s/proof-guess/code", ...form({ address: `rate${i}@example.test` }) });
      expect(res.statusCode).toBe(200);
    }
    // The sender is told once, with the attempts; nobody else got a mail.
    await vi.waitFor(() => expect(sendMailAsApp).toHaveBeenCalledOnce());
    const [, owner, to, subject, html] = sendMailAsApp.mock.calls[0];
    expect([owner, to]).toEqual(["oid-anna", "anna@example.test"]);
    expect(subject).toContain("auffällige Code-Anforderungen");
    expect(html).toContain("rate5@example.test");

    // Locked: even the real recipient gets the same page and no code for now.
    const locked = await app.inject({ method: "POST", url: "/r/s/proof-guess/code", ...form({ address: "erika@example.test" }) });
    expect(locked.statusCode).toBe(200);
    await new Promise((r) => setTimeout(r, 20));
    expect(sendMailAsApp).toHaveBeenCalledOnce();
    expect(caseById(proofId)!.events.map((e) => e.type)).toContain("gateway.locked");
  });
});

describe("delivery proof – password after the confirmation", () => {
  const form = (fields: Record<string, string>) => ({
    headers: { "content-type": "application/x-www-form-urlencoded" },
    payload: new URLSearchParams(fields).toString(),
  });

  it("mails the held-back password only to the confirmed recipient – again on a new confirmation – and shows the next steps", async () => {
    nc.createPublicShare.mockResolvedValue("https://cloud.example.test/s/proof-pw");
    const { body } = await startUpload("anna", "Angebot.pdf");
    await putChunk("anna", body.uploadId, 1);
    const { proofId } = (await complete("anna", body.uploadId, { proof: true })).json();
    const { recordMailSent, caseById } = await import("../src/proof.js");
    recordMailSent(["proof-pw"], { subject: "Angebot", messageId: "<p@example.test>", from: "anna@example.test", to: ["erika@example.test", "bernd@example.test"], cc: [], bcc: [] });
    // What the job watcher stores when it holds the password back (jobs.ts, PasswordRef.deferred).
    const { registerPasswordRef } = await import("../src/passwordRefs.js");
    await registerPasswordRef({
      owner: "oid-anna", ownerUpn: "anna@example.test", sendId: "send-held", tokens: ["proof-pw"], fileNames: ["Angebot.pdf"],
      keepUntil: Date.now() + 7 * DAY, subject: "Angebot", recipients: { to: ["erika@example.test", "bernd@example.test"], cc: [], bcc: [] },
      maxAccessCount: 3, deferred: true, toRecipients: true,
    });
    const sendMailAsApp = vi.mocked((await import("../src/graph.js")).sendMailAsApp);

    const confirm = async (address: string) => {
      sendMailAsApp.mockClear();
      await app.inject({ method: "POST", url: "/r/s/proof-pw/code", remoteAddress: "192.0.2.10", ...form({ address }) });
      await vi.waitFor(() => expect(sendMailAsApp).toHaveBeenCalledOnce());
      const code = />(\d{6})</.exec(String(sendMailAsApp.mock.calls[0][4]))![1];
      sendMailAsApp.mockClear();
      return app.inject({ method: "POST", url: "/r/s/proof-pw/verify", remoteAddress: "192.0.2.10", ...form({ address, code }) });
    };

    const page = await confirm("erika@example.test");
    expect(page.statusCode).toBe(200);
    expect(page.body).toContain("So geht es weiter");
    expect(page.body).toContain("an erika@example.test gesendet");
    expect(page.body).toContain('href="https://cloud.example.test/s/proof-pw"');
    expect(String(page.headers["set-cookie"])).toContain("HttpOnly");
    // The password mail went to her only – not to every recipient of the mail.
    expect(sendMailAsApp).toHaveBeenCalledOnce();
    expect(sendMailAsApp.mock.calls[0][2]).toBe("erika@example.test");
    expect(String(sendMailAsApp.mock.calls[0][4])).toContain("Kennwort anzeigen");
    expect(String(sendMailAsApp.mock.calls[0][4])).toContain("Sie haben eben Ihre E-Mail-Adresse bestätigt");
    // Its button leads back to the gateway, never straight to Nextcloud.
    expect(String(sendMailAsApp.mock.calls[0][4])).toContain('href="https://cloudshare.example.test/r/s/proof-pw"');
    expect(String(sendMailAsApp.mock.calls[0][4])).not.toContain("cloud.example.test/s/proof-pw");

    // Under step 1: the way to a new password link – the address form even for this confirmed browser.
    expect(String(sendMailAsApp.mock.calls[0][4])).toContain('href="https://cloudshare.example.test/r/s/proof-pw/password"');
    expect(String(sendMailAsApp.mock.calls[0][4])).toContain("Neuen Kennwort-Link anfordern");
    const cookie = String(page.headers["set-cookie"]).split(";")[0];
    const askAgain = await app.inject({ method: "GET", url: "/r/s/proof-pw/password", headers: { cookie } });
    expect(askAgain.statusCode).toBe(200);
    expect(askAgain.body).toContain('action="/r/s/proof-pw/code"');
    expect(askAgain.body).toContain("danach erhalten Sie einen neuen Kennwort-Link");

    // Confirming again (the first password link ran out, or another device): a fresh password mail to her.
    const again = await confirm("erika@example.test");
    expect(again.body).toContain("an erika@example.test gesendet");
    expect(sendMailAsApp).toHaveBeenCalledOnce();
    expect(sendMailAsApp.mock.calls[0][2]).toBe("erika@example.test");
    expect(String(sendMailAsApp.mock.calls[0][4])).toContain("Kennwort anzeigen");

    const types = caseById(proofId)!.events.map((e) => e.type);
    expect(types).toEqual(expect.arrayContaining(["recipient.verified", "password.sent"]));
  });

  it("\"Kennwort erneut senden\" never gets round the proof: nothing before a confirmation, then only to confirmed addresses", async () => {
    nc.createPublicShare.mockResolvedValue("https://cloud.example.test/s/proof-again");
    const { body } = await startUpload("anna", "Plan.pdf");
    await putChunk("anna", body.uploadId, 1);
    await complete("anna", body.uploadId, { proof: true });
    const { recordMailSent } = await import("../src/proof.js");
    recordMailSent(["proof-again"], { subject: "Plan", messageId: "<a@example.test>", from: "anna@example.test", to: ["erika@example.test", "bernd@example.test"], cc: [], bcc: [] });
    const { registerPasswordRef } = await import("../src/passwordRefs.js");
    await registerPasswordRef({
      owner: "oid-anna", ownerUpn: "anna@example.test", sendId: "send-again", tokens: ["proof-again"], fileNames: ["Plan.pdf"],
      keepUntil: Date.now() + 7 * DAY, subject: "Plan", recipients: { to: ["erika@example.test", "bernd@example.test"], cc: [], bcc: [] },
      maxAccessCount: 3, deferred: true, toRecipients: true,
    });
    nc.listPublicShares.mockResolvedValue([{ id: "77", path: "/CloudShare/2026-11-20_ab/Plan.pdf", url: "https://cloud.example.test/s/proof-again",
      expiration: "2099-12-01", sharedAt: Date.now(), fileId: 77, isFolder: false, permissions: 1, label: "", size: 1, hasPassword: true }]);
    const graph = await import("../src/graph.js");
    const sendMail = vi.mocked(graph.sendMail);
    const sendMailAsApp = vi.mocked(graph.sendMailAsApp);
    const resend = (payload: Record<string, unknown>) =>
      app.inject({ method: "POST", url: "/api/shares/77/password/send", headers: as("anna"), remoteAddress: "192.0.2.50", payload });

    // Nobody confirmed yet: the pane is told, and no way sends it – not to the recipients, not elsewhere, not by SMS.
    const shown = (await app.inject({ method: "GET", url: "/api/shares/77/password", headers: as("anna"), remoteAddress: "192.0.2.50" })).json();
    expect(shown).toMatchObject({ deferred: true, verified: [] });
    expect(shown).not.toHaveProperty("password"); // never shown with a proof of delivery, not even to the sender
    sendMail.mockClear();
    for (const payload of [{ channel: "recipients" }, { channel: "mail", to: "kollege@example.test" }, { channel: "sms", to: "+491701234567" }]) {
      const refused = await resend(payload);
      expect(refused.statusCode).toBe(409);
      expect(refused.json().code).toBe("PASSWORD_HELD");
    }
    expect(sendMail).not.toHaveBeenCalled();

    // Erika confirms with the code: sent again to the recipients means to her only, Bernd has not confirmed.
    sendMailAsApp.mockClear();
    await app.inject({ method: "POST", url: "/r/s/proof-again/code", remoteAddress: "192.0.2.11", ...form({ address: "erika@example.test" }) });
    await vi.waitFor(() => expect(sendMailAsApp).toHaveBeenCalledOnce());
    const code = />(\d{6})</.exec(String(sendMailAsApp.mock.calls[0][4]))![1];
    await app.inject({ method: "POST", url: "/r/s/proof-again/verify", remoteAddress: "192.0.2.11", ...form({ address: "erika@example.test", code }) });
    expect((await app.inject({ method: "GET", url: "/api/shares/77/password", headers: as("anna"), remoteAddress: "192.0.2.50" })).json().verified).toEqual(["erika@example.test"]);
    sendMail.mockClear();
    // Only confirmed recipients the sender ticks: none ticked, Bernd (not confirmed) or a new address are refused (SMS may go to a corrected number).
    for (const payload of [{ channel: "recipients" }, { channel: "recipients", to: ["bernd@example.test"] },
      { channel: "mail", to: ["kollege@example.test"] }]) {
      const refused = await resend(payload);
      expect(refused.statusCode).toBe(409);
      expect(refused.json().code).toBe("PROOF_RESEND");
    }
    expect(sendMail).not.toHaveBeenCalled();
    const again = await resend({ channel: "recipients", to: ["Erika@example.test"] });
    expect(again.statusCode).toBe(200);
    expect(again.json()).toEqual({ channel: "recipients", to: "erika@example.test" });
    expect(sendMail).toHaveBeenCalledOnce();
    expect(sendMail.mock.calls[0][1]).toEqual({ to: ["erika@example.test"], cc: [], bcc: [] });
  });
});

describe("delivery proof – the list for the search (GET /api/proofs)", () => {
  it("lists the user's own cases with subject, recipients, files and state – nobody else's", async () => {
    nc.createPublicShare.mockResolvedValue("https://cloud.example.test/s/proof-list");
    const { body } = await startUpload("anna", "Liste.pdf");
    await putChunk("anna", body.uploadId, 1);
    const { proofId } = (await complete("anna", body.uploadId, { proof: true })).json();
    const { recordMailSent } = await import("../src/proof.js");
    recordMailSent(["proof-list"], { subject: "Liste", messageId: "<l@example.test>", from: "anna@example.test", to: ["erika@example.test"], cc: [], bcc: ["geheim@example.test"], sentAt: "2026-11-18T10:00:00Z" });

    const list = (await app.inject({ method: "GET", url: "/api/proofs", headers: as("anna") })).json() as Array<Record<string, unknown>>;
    const entry = list.find((p) => p.id === proofId)!;
    expect(entry).toMatchObject({ subject: "Liste", recipients: ["erika@example.test"], files: ["Liste.pdf"], downloads: 0, done: false });
    expect(entry.recipients).not.toContain("geheim@example.test"); // Bcc stays out of the list
    expect((await app.inject({ method: "GET", url: "/api/proofs", headers: as("bernd") })).json()
      .some((p: { id: string }) => p.id === proofId)).toBe(false);
  });

  it("leaves out proofs of mails that were never sent", async () => {
    nc.createPublicShare.mockResolvedValue("https://cloud.example.test/s/proof-draft");
    const { body } = await startUpload("anna", "Entwurf.pdf");
    await putChunk("anna", body.uploadId, 1);
    const { proofId } = (await complete("anna", body.uploadId, { proof: true })).json();
    const list = (await app.inject({ method: "GET", url: "/api/proofs", headers: as("anna") })).json() as Array<{ id: string }>;
    expect(list.some((p) => p.id === proofId)).toBe(false);
  });
});

describe("delivery proof – the pages carry the organisation", () => {
  it("shows the CloudShare logo with the organisation, and website, imprint and privacy links", async () => {
    const { config } = await import("../src/config.js");
    Object.assign(config, { companyName: "Example GmbH", companyWebsiteUrl: "https://www.example.com", imprintUrl: "https://www.example.com/imprint", privacyUrl: "" });
    try {
      const page = await app.inject({ method: "GET", url: "/r/s/unknown-token" });
      expect(page.body).toContain('<img src="/assets/icon-64.png"');
      expect(page.body).toContain('<span class="brand-company">Example GmbH</span>');
      expect(page.body).toContain('<a href="https://www.example.com" target="_blank" rel="noopener noreferrer">example.com</a>');
      expect(page.body).toContain(">Impressum</a>");
      expect(page.body).not.toContain("Datenschutz</a>"); // no PRIVACY_URL: no link
      expect(page.headers["content-security-policy"]).toContain("img-src 'self' data:");
    } finally {
      Object.assign(config, { companyName: "", companyWebsiteUrl: "", imprintUrl: "", privacyUrl: "" });
    }
  });
});

describe("errors – a service that cannot be reached is named", () => {
  const fetchFailed = (code: string, hostname?: string) =>
    Object.assign(new TypeError("fetch failed"), { cause: Object.assign(new Error(`getaddrinfo ${code}`), { code, hostname }) });

  it("names Nextcloud, the password service or the host instead of „Interner Fehler“", async () => {
    const { unreachableService } = await import("../src/app.js");
    expect(unreachableService(fetchFailed("ENOTFOUND", "cloud.example.test"))).toBe("Nextcloud (cloud.example.test) ist nicht erreichbar. Bitte wenden Sie sich an Ihre IT.");
    expect(unreachableService(fetchFailed("ECONNREFUSED", "127.0.0.1"))).toBe("Der Kennwortdienst ist nicht erreichbar. Bitte wenden Sie sich an Ihre IT.");
    expect(unreachableService(fetchFailed("ETIMEDOUT", "graph.microsoft.com"))).toBe("graph.microsoft.com ist nicht erreichbar. Bitte wenden Sie sich an Ihre IT.");
    expect(unreachableService(new Error("something else"))).toBeNull();
  });

  it("answers 502 with code SERVICE_UNREACHABLE", async () => {
    nc.userInfo.mockRejectedValueOnce(fetchFailed("ENOTFOUND", "cloud.example.test"));
    const res = await app.inject({ method: "GET", url: "/api/shares", headers: as("anna") });
    expect(res.statusCode).toBe(502);
    expect(res.json()).toEqual({ error: "Nextcloud (cloud.example.test) ist nicht erreichbar. Bitte wenden Sie sich an Ihre IT.", code: "SERVICE_UNREACHABLE" });
  });
});

describe("delivery proof – state and reminder", () => {
  /** A sent case of anna with one file link ending on `expireDate`. */
  async function sentCase(token: string, expireDate?: string, sentAt = "2026-11-18T10:00:00Z") {
    nc.createPublicShare.mockResolvedValue(`https://cloud.example.test/s/${token}`);
    const { body } = await startUpload("anna", "Bericht.pdf");
    await putChunk("anna", body.uploadId, 1);
    const { proofId } = (await complete("anna", body.uploadId, { proof: true })).json();
    const { recordMailSent, caseById } = await import("../src/proof.js");
    if (expireDate) for (const l of Object.values(caseById(proofId)!.links ?? {})) l.expireDate = expireDate;
    recordMailSent([token], { subject: "Bericht", messageId: `<${token}@example.test>`, from: "anna@example.test", to: ["erika@example.test"], cc: [], bcc: [], sentAt });
    return proofId as string;
  }

  it("lists the state, the end of the links and until when the proof is kept; the seal in \"Verlauf\" gets the state", async () => {
    const proofId = await sentCase("proof-state");
    const entry = ((await app.inject({ method: "GET", url: "/api/proofs", headers: as("anna") })).json() as Array<Record<string, unknown>>)
      .find((p) => p.id === proofId)!;
    expect(entry).toMatchObject({ state: "sent", fetched: 0 });
    expect(entry).toMatchObject({ kind: "file", mailAvailable: false, passwordTo: [] }); // password not sent yet
    // No sent mail known for it (recorded by the job watcher): the subject cannot open one; someone else's case: not found.
    expect((await app.inject({ method: "GET", url: `/api/proofs/${proofId}/mail`, headers: as("anna") })).statusCode).toBe(404);
    expect((await app.inject({ method: "GET", url: `/api/proofs/${proofId}/mail`, headers: as("bernd") })).json().error).toBe("Zustellnachweis nicht gefunden");
    expect(typeof entry.endsOn).toBe("string");
    expect(entry.keptUntil).toBeGreaterThan(Date.now());
  });


  it("reminds when half of the validity has passed – never within the first 24 hours", async () => {
    const { reminderTime } = await import("../src/proofWatcher.js");
    const day = 86_400_000;
    expect(reminderTime(0, 14 * day)).toBe(7 * day); // 14 days: after 7
    expect(reminderTime(0, 2 * day)).toBe(day); // 2 days: after 1
    expect(reminderTime(0, 30 * 3_600_000)).toBe(day); // 30 hours: not before 24
    expect(reminderTime(0, 20 * 3_600_000)).toBeNull(); // ends before 24 hours: none
  });

  it("reminds the sender once when nothing was fetched by half of the validity – not when switched off", async () => {
    const soon = new Date(Date.now() + 86_400_000).toISOString().slice(0, 10);
    const threeDaysAgo = new Date(Date.now() - 3 * 86_400_000).toISOString();
    // Just sent, link ends tomorrow: half of the time has not passed yet – no reminder.
    await sentCase("proof-remind-early", soon, new Date().toISOString());
    const proofId = await sentCase("proof-remind", soon, threeDaysAgo);
    const sendMailAsApp = vi.mocked((await import("../src/graph.js")).sendMailAsApp);
    sendMailAsApp.mockClear();
    const { runProofWatch } = await import("../src/proofWatcher.js");
    const log = { info: vi.fn(), warn: vi.fn() };
    await runProofWatch(log as never);
    const reminders = sendMailAsApp.mock.calls.filter((call) => String(call[3]).startsWith("Noch nicht abgerufen"));
    expect(reminders).toHaveLength(1);
    expect(reminders[0][3]).toBe(`Noch nicht abgerufen – Vorgang ${proofId}`);
    // Once per case.
    await runProofWatch(log as never);
    expect(sendMailAsApp.mock.calls.filter((call) => String(call[3]).startsWith("Noch nicht abgerufen"))).toHaveLength(1);

    // Switched off in the settings: no reminder.
    await app.inject({ method: "PUT", url: "/api/settings", headers: as("anna"), payload: { proofReminder: false } });
    const other = await sentCase("proof-remind-off", soon, threeDaysAgo);
    sendMailAsApp.mockClear();
    await runProofWatch(log as never);
    expect(sendMailAsApp.mock.calls.some((call) => String(call[3]).includes(other))).toBe(false);
    await app.inject({ method: "DELETE", url: "/api/settings", headers: as("anna") });
  });
});

describe("delivery proof – never mixed with links without one in a mail", () => {
  it("refuses a message or an upload link in a mail with a proof, and a proof in a mail without one", async () => {
    nc.createPublicShare.mockResolvedValue("https://cloud.example.test/s/mix-proof");
    const first = await startUpload("anna", "Vertrag.pdf");
    await putChunk("anna", first.body.uploadId, 1);
    const withProof = (await complete("anna", first.body.uploadId, { proof: true, step: "s1" })).json();
    expect(withProof.proofId).toBeTruthy();

    const message = await app.inject({ method: "POST", url: "/api/messages", headers: as("anna"),
      payload: { jobId: withProof.jobId, step: "s1", password: STRONG, text: "geheim" } });
    expect(message.statusCode).toBe(409);
    expect(message.json().code).toBe("PROOF_MIX");
    const request = await app.inject({ method: "POST", url: "/api/requests", headers: as("anna"),
      payload: { jobId: withProof.jobId, step: "s1", password: STRONG } });
    expect(request.json().code).toBe("PROOF_MIX");

    // The other way round: a mail with a plain link takes no file with a proof.
    nc.createPublicShare.mockResolvedValue("https://cloud.example.test/s/mix-plain");
    const plain = await startUpload("anna", "Liste.pdf");
    await putChunk("anna", plain.body.uploadId, 1);
    const plainJob = (await complete("anna", plain.body.uploadId, { step: "p1" })).json().jobId;
    const second = await startUpload("anna", "Plan.pdf");
    await putChunk("anna", second.body.uploadId, 1);
    const refused = await complete("anna", second.body.uploadId, { jobId: plainJob, step: "p1", proof: true });
    expect(refused.statusCode).toBe(409);
    expect(refused.json().error).toContain("ohne Zustellnachweis");
  });

  it("holds back the send of a mail that mixes them anyway (the password would reach them before the confirmation)", async () => {
    nc.createPublicShare.mockResolvedValue("https://cloud.example.test/s/mix-arm");
    const upload = await startUpload("anna", "Bericht.pdf");
    await putChunk("anna", upload.body.uploadId, 1);
    const { jobId } = (await complete("anna", upload.body.uploadId, { proof: true, step: "a1" })).json();
    const { addSeparatePart } = await import("../src/jobs.js");
    await addSeparatePart(jobId, {
      step: "a2", fileNames: ["Alt.pdf"], sendUrl: "https://vault.example.test/#/send/x/y", sendId: "send-x",
      shares: [{ token: "plain-token", name: "Alt.pdf", url: "https://cloud.example.test/s/plain-token", expireDate: "2099-12-31" }],
    } as never);
    const res = await app.inject({ method: "POST", url: `/api/jobs/${jobId}/arm`, headers: as("anna"), payload: {} });
    expect(res.statusCode).toBe(409);
    expect(res.json()).toMatchObject({ code: "PROOF_MIX" });
  });
});
