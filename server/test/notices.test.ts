/**
 * Download confirmation (src/notices.ts): Nextcloud's Activity API, CloudShare's app token and Graph are test
 * doubles; credentials and the encrypted registry are the real modules.
 */
import { beforeEach, describe, expect, it, vi } from "vitest";

const nc = vi.hoisted(() => ({ listPublicShares: vi.fn(), publicDownloads: vi.fn() }));
vi.mock("../src/nextcloud.js", async (importOriginal) => {
  const actual = await importOriginal<typeof import("../src/nextcloud.js")>();
  class FakeNextcloudClient {
    constructor() {
      return nc;
    }
  }
  return { ...actual, NextcloudClient: FakeNextcloudClient };
});
vi.mock("../src/graph.js", () => ({ sendMailAsApp: vi.fn(async () => undefined) }));
vi.mock("../src/auth.js", async (importOriginal) => ({
  ...(await importOriginal<typeof import("../src/auth.js")>()),
  appGraphToken: vi.fn(async () => "app-token"),
}));

const { pendingNotices, registerNotices, runNotices } = await import("../src/notices.js");
const { saveCredentials } = await import("../src/credentials.js");
const graph = await import("../src/graph.js");
const { quietLog } = await import("./helpers.js");
const sendMailAsApp = vi.mocked(graph.sendMailAsApp);

const registeredAt = Date.parse("2026-11-14T10:00:00Z");
const notice = (fileName: string, expireDate = "2026-11-28") => ({
  owner: "oid-anna",
  ownerUpn: "anna@example.test",
  notifyTo: "anna@example.test",
  fileName,
  subject: "Unterlagen",
  recipients: "erika@example.test",
  expireDate,
  registeredAt,
});
const share = (token: string, fileId: number) => ({ id: token, url: `https://cloud.example.test/s/${token}`, fileId, path: "/CloudShare/x/y" });

beforeEach(async () => {
  nc.listPublicShares.mockReset().mockResolvedValue([]);
  nc.publicDownloads.mockReset().mockResolvedValue({ count: 0, last: null });
  sendMailAsApp.mockClear();
  await saveCredentials("oid-anna", { loginName: "anna", appPassword: "pw" });
});

describe("download confirmation", () => {
  it("mails the sender once after the first download, then forgets the link", async () => {
    await registerNotices([["tokA", notice("Angebot.pdf")]]);
    nc.listPublicShares.mockResolvedValue([share("tokA", 11)]);

    await runNotices(quietLog(), registeredAt + 60_000); // not downloaded yet
    expect(sendMailAsApp).not.toHaveBeenCalled();

    nc.publicDownloads.mockResolvedValue({ count: 1, last: "2026-11-14T11:30:00+00:00" });
    await runNotices(quietLog(), registeredAt + 2 * 3_600_000);
    expect(sendMailAsApp).toHaveBeenCalledOnce();
    const [token, user, to, subject, html] = sendMailAsApp.mock.calls[0];
    expect(token).toBe("app-token");
    expect(user).toBe("oid-anna");
    expect(to).toBe("anna@example.test");
    expect(subject).toBe("Heruntergeladen: Angebot.pdf");
    expect(html).toContain("„Angebot.pdf“ aus Ihrer Nachricht „Unterlagen“");
    expect(html).toContain("erika@example.test");
    expect(pendingNotices().some(([t]) => t === "tokA")).toBe(false);

    await runNotices(quietLog(), registeredAt + 3 * 3_600_000);
    expect(sendMailAsApp).toHaveBeenCalledOnce(); // only the first download
  });

  it("ignores downloads before the password went out (older links of the same file)", async () => {
    await registerNotices([["tokB", notice("Alt.pdf")]]);
    nc.listPublicShares.mockResolvedValue([share("tokB", 12)]);
    nc.publicDownloads.mockResolvedValue({ count: 3, last: "2026-11-01T09:00:00+00:00" });
    await runNotices(quietLog(), registeredAt + 3_600_000);
    expect(sendMailAsApp).not.toHaveBeenCalled();
  });

  it("drops links that expired without a download", async () => {
    await registerNotices([["tokC", notice("Nie.pdf", "2026-11-15")]]);
    await runNotices(quietLog(), Date.parse("2026-11-20T00:00:00Z"));
    expect(pendingNotices().some(([t]) => t === "tokC")).toBe(false);
    expect(sendMailAsApp).not.toHaveBeenCalled();
  });
});
