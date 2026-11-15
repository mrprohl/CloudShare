import { beforeEach, describe, expect, it, vi } from "vitest";

const nc = vi.hoisted(() => ({
  userInfo: vi.fn(),
  listPublicShares: vi.fn(),
  deleteShare: vi.fn(),
  deletePath: vi.fn(),
}));
vi.mock("../src/nextcloud.js", async (importOriginal) => {
  const actual = await importOriginal<typeof import("../src/nextcloud.js")>();
  class FakeNextcloudClient {
    constructor() {
      return nc;
    }
  }
  return { ...actual, NextcloudClient: FakeNextcloudClient };
});
vi.mock("../src/vaultwarden.js", () => ({ deleteSend: vi.fn(async () => undefined) }));

const { revokeUnsentLinks } = await import("../src/shares.js");
const { saveCredentials } = await import("../src/credentials.js");
const { HttpError } = await import("../src/auth.js");
const { NOT_CONNECTED } = await import("../src/nextcloud.js");
const vaultwarden = await import("../src/vaultwarden.js");
const { quietLog } = await import("./helpers.js");
const deleteSend = vi.mocked(vaultwarden.deleteSend);

const share = (id: string, path: string, token: string) => ({
  id, path, url: `https://cloud.example.test/s/${token}`, expiration: "2099-01-01", sharedAt: Date.now(), fileId: 1,
  isFolder: false, permissions: 1, label: "", size: 1, hasPassword: true,
});
const links = { owner: "oid-max", ownerUpn: "max@example.test", tokens: ["tokA", "tokB"], sendIds: ["send-1"] };

beforeEach(async () => {
  for (const fn of Object.values(nc)) fn.mockReset();
  nc.userInfo.mockResolvedValue({ id: "uid", free: null });
  deleteSend.mockClear();
  await saveCredentials("oid-max", { loginName: "max", appPassword: "pw" });
});

describe("links of a mail that was never sent", () => {
  it("revokes exactly its CloudShare links, deletes its Sends and moves its uploads to the trash", async () => {
    nc.listPublicShares.mockResolvedValue([
      share("1", "/CloudShare/2026-11-13_ab12cd34/a.pdf", "tokA"),
      share("2", "/CloudShare/2026-11-13_ab12cd34/b.pdf", "tokB"), // same upload, also never sent
      share("3", "/CloudShare/2026-11-14_cd34ef56/c.pdf", "tokC"), // another mail – stays
      share("4", "/Privat/d.pdf", "tokA"), // not CloudShare's – never touched
    ]);
    await revokeUnsentLinks(links, quietLog());
    expect(nc.deleteShare.mock.calls.map((c) => c[0])).toEqual(["1", "2"]);
    expect(deleteSend).toHaveBeenCalledWith("send-1");
    // The second link sees the first one gone: the folder goes to the trash (CLEANUP_AFTER_DAYS=0 in the tests).
    expect(nc.deletePath).toHaveBeenCalledWith("uid", ["CloudShare", "2026-11-13_ab12cd34"]);
  });

  it("keeps a file another link still shares (a link inserted again from \"Meine Freigaben\")", async () => {
    nc.listPublicShares.mockResolvedValue([
      share("1", "/CloudShare/2026-11-13_ab12cd34/a.pdf", "tokA"),
      share("5", "/CloudShare/2026-11-13_ab12cd34/a.pdf", "tokOriginal"), // the link that was sent earlier
    ]);
    await revokeUnsentLinks(links, quietLog());
    expect(nc.deleteShare.mock.calls.map((c) => c[0])).toEqual(["1"]);
    expect(nc.deletePath).not.toHaveBeenCalled();
  });

  it("skips users no longer connected and throws if Nextcloud is unreachable (retried later)", async () => {
    await revokeUnsentLinks({ ...links, owner: "oid-nobody" }, quietLog());
    nc.listPublicShares.mockRejectedValueOnce(new HttpError(401, "x", NOT_CONNECTED));
    await revokeUnsentLinks(links, quietLog());
    expect(nc.deleteShare).not.toHaveBeenCalled();

    nc.listPublicShares.mockRejectedValueOnce(new Error("ECONNREFUSED"));
    await expect(revokeUnsentLinks(links, quietLog())).rejects.toThrow("ECONNREFUSED");
  });
});
