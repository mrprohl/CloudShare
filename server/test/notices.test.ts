/**
 * Download confirmation (src/notices.ts): Nextcloud's Activity API, CloudShare's app token and Graph are test
 * doubles; credentials and the encrypted registry are the real modules.
 */
import { beforeEach, describe, expect, it, vi } from "vitest";

const nc = vi.hoisted(() => ({ listPublicShares: vi.fn(), publicDownloads: vi.fn(), userInfo: vi.fn(), listFiles: vi.fn() }));
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
  nc.userInfo.mockReset().mockResolvedValue({ id: "anna-uid", free: null });
  nc.listFiles.mockReset().mockResolvedValue([]);
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
    // Common layout of the automatic mails: a green card with the file, the line that CloudShare sent it.
    expect(html).toContain("border-left:4px solid #107c10");
    expect(html).toContain("Datei heruntergeladen");
    expect(html).toMatch(/Zum ersten Mal heruntergeladen am .*, \d{2}:\d{2} Uhr · Link gültig bis /);
    expect(html).toContain("automatisch von CloudShare");
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

describe("upload notification (UPLOAD_NOTICE)", () => {
  const upload = (knownFiles: string[] = [], expireDate = "2026-11-28") => ({
    ...notice("Upload-Link für Ihre Dateien", expireDate),
    kind: "upload" as const,
    knownFiles,
  });
  const folder = (token: string) => ({ ...share(token, 30), path: "/CloudShare-File-Requests/2026-11-14_ab12cd34" });

  it("mails the files that arrived since the last mail – one mail per pass, until the link expires", async () => {
    await registerNotices([["tokU", upload()]]);
    nc.listPublicShares.mockResolvedValue([folder("tokU")]);

    await runNotices(quietLog(), registeredAt + 600_000); // nothing uploaded yet
    expect(sendMailAsApp).not.toHaveBeenCalled();

    nc.listFiles.mockResolvedValue(["Ausweis.pdf", "Nachweis.pdf"]);
    await runNotices(quietLog(), registeredAt + 1_200_000);
    expect(sendMailAsApp).toHaveBeenCalledOnce();
    const [token, user, to, subject, html] = sendMailAsApp.mock.calls[0];
    expect([token, user, to]).toEqual(["app-token", "oid-anna", "anna@example.test"]);
    expect(subject).toBe("2 Dateien empfangen: Unterlagen");
    expect(html).toContain("Über den Link zum Hochladen aus Ihrer Nachricht „Unterlagen“ (an erika@example.test)");
    expect(html).toContain("border-left:4px solid #107c10");
    expect(html).toContain("Ausweis.pdf");
    expect(html).toContain("Nachweis.pdf");
    expect(html).toContain("In Nextcloud öffnen");
    expect(html).toContain("/index.php/apps/files/?dir=%2FCloudShare-File-Requests%2F2026-11-14_ab12cd34");
    expect(nc.listFiles).toHaveBeenCalledWith("anna-uid", ["CloudShare-File-Requests", "2026-11-14_ab12cd34"]);

    await runNotices(quietLog(), registeredAt + 1_800_000); // nothing new
    expect(sendMailAsApp).toHaveBeenCalledOnce();

    nc.listFiles.mockResolvedValue(["Ausweis.pdf", "Nachweis.pdf", "Foto.jpg"]);
    await runNotices(quietLog(), registeredAt + 2_400_000);
    expect(sendMailAsApp).toHaveBeenCalledTimes(2);
    expect(sendMailAsApp.mock.calls[1][3]).toBe("Datei empfangen: Unterlagen");
    expect(sendMailAsApp.mock.calls[1][4]).toContain("Foto.jpg");
    expect(sendMailAsApp.mock.calls[1][4]).not.toContain("Ausweis.pdf");
    expect(pendingNotices().some(([t]) => t === "tokU")).toBe(true); // still watched
  });

  it("does not report files that were in the folder before the link went out", async () => {
    await registerNotices([["tokV", upload(["Alt.pdf"])]]);
    nc.listPublicShares.mockResolvedValue([folder("tokV")]);
    nc.listFiles.mockResolvedValue(["Alt.pdf"]);
    await runNotices(quietLog(), registeredAt + 600_000);
    expect(sendMailAsApp).not.toHaveBeenCalled();
  });

  it("names at most ten files and counts the rest", async () => {
    await registerNotices([["tokW", upload()]]);
    nc.listPublicShares.mockResolvedValue([folder("tokW")]);
    nc.listFiles.mockResolvedValue(Array.from({ length: 13 }, (_, i) => `Scan-${i + 1}.pdf`));
    await runNotices(quietLog(), registeredAt + 600_000);
    const html = sendMailAsApp.mock.calls[0][4];
    expect(html).toContain("Scan-10.pdf");
    expect(html).not.toContain("Scan-11.pdf");
    expect(html).toContain("und 3 weitere Dateien");
  });

  it("stops watching when the link expired", async () => {
    await registerNotices([["tokX", upload([], "2026-11-15")]]);
    await runNotices(quietLog(), Date.parse("2026-11-20T00:00:00Z"));
    expect(pendingNotices().some(([t]) => t === "tokX")).toBe(false);
  });
});
