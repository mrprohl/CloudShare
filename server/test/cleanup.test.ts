/**
 * Deleting uploaded files after their link expired (src/cleanup.ts). Nextcloud is a test double; credentials and the
 * encrypted registry use the real modules with the temporary DATA_DIR from setup.ts.
 */
import { beforeEach, describe, expect, it, vi } from "vitest";

const nc = vi.hoisted(() => ({
  userInfo: vi.fn(),
  listPublicShares: vi.fn(),
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

const { cleanupEntries, expireOnRevoke, isUploadFolder, registerForCleanup, runCleanup } = await import("../src/cleanup.js");
const { saveCredentials } = await import("../src/credentials.js");
const { quietLog } = await import("./helpers.js");

const DAY = 86_400_000;
const folder = (id: string) => ["CloudShare", `2026-11-13_${id}`];
const file = (id: string) => [...folder(id), "Bericht.pdf"];
/** Noon UTC of a given day. */
const at = (date: string) => Date.parse(`${date}T12:00:00Z`);

beforeEach(async () => {
  for (const fn of Object.values(nc)) fn.mockReset();
  nc.userInfo.mockResolvedValue({ id: "anna-uid", free: null, used: null, total: null });
  nc.listPublicShares.mockResolvedValue([]);
  nc.deletePath.mockResolvedValue(undefined);
  await saveCredentials("oid-anna", { loginName: "anna", appPassword: "pw" });
});

describe("cleanup of expired uploads", () => {
  it("only accepts upload folders below NEXTCLOUD_FOLDER", () => {
    expect(isUploadFolder(folder("ab12cd34"))).toBe(true);
    expect(isUploadFolder(["CloudShare-File-Requests", "2026-11-13_ab12cd34"])).toBe(false); // received files
    expect(isUploadFolder(["CloudShare"])).toBe(false);
    expect(isUploadFolder(["CloudShare", ".."])).toBe(false);
    expect(isUploadFolder(["Privat", "2026-11-13_ab12cd34"])).toBe(false);
  });

  it("deletes the folder right after the link expired, not before", async () => {
    await registerForCleanup("oid-anna", file("aaaa0001"), "2026-11-27");

    await runCleanup(quietLog(), at("2026-11-27")); // last valid day
    expect(nc.deletePath).not.toHaveBeenCalled();

    await runCleanup(quietLog(), at("2026-11-28"));
    expect(nc.deletePath).toHaveBeenCalledWith("anna-uid", folder("aaaa0001"));
    expect(cleanupEntries().some((e) => e.folder[1].endsWith("aaaa0001"))).toBe(false);
  });

  it("keeps the latest expiry when a file is linked again", async () => {
    await registerForCleanup("oid-anna", file("aaaa0002"), "2026-11-27");
    await registerForCleanup("oid-anna", file("aaaa0002"), "2026-12-24");
    await registerForCleanup("oid-anna", file("aaaa0002"), "2026-12-01"); // earlier – ignored
    expect(cleanupEntries().find((e) => e.folder[1].endsWith("aaaa0002"))?.expireDate).toBe("2026-12-24");
  });

  it("postpones a folder that still has a valid link", async () => {
    await registerForCleanup("oid-anna", file("aaaa0003"), "2026-11-27");
    nc.listPublicShares.mockResolvedValue([
      { id: "1", path: "/CloudShare/2026-11-13_aaaa0003/Bericht.pdf", expiration: "2027-01-15" },
    ]);
    await runCleanup(quietLog(), at("2026-12-01"));
    expect(nc.deletePath).not.toHaveBeenCalled();
    expect(cleanupEntries().find((e) => e.folder[1].endsWith("aaaa0003"))?.expireDate).toBe("2027-01-15");
  });

  it("never registers received files or anything outside the upload folders", async () => {
    await registerForCleanup("oid-anna", ["CloudShare-File-Requests", "2026-11-13_aaaa0004"], "2026-11-27");
    await registerForCleanup("oid-anna", ["Privat", "x.pdf"], "2026-11-27");
    expect(cleanupEntries().some((e) => e.folder[0] !== "CloudShare")).toBe(false);
  });

  it("keeps the folders of a user who is not connected until they connect again – and gives up after 90 days", async () => {
    await registerForCleanup("oid-gone", file("aaaa0005"), "2026-11-01");
    await runCleanup(quietLog(), at("2026-11-20"));
    expect(nc.deletePath).not.toHaveBeenCalledWith(expect.anything(), folder("aaaa0005"));
    expect(cleanupEntries().some((e) => e.owner === "oid-gone")).toBe(true); // "Verbindung trennen" loses nothing

    // Connected again: deleted on the next pass.
    await saveCredentials("oid-gone", { loginName: "gone", appPassword: "pw" });
    await runCleanup(quietLog(), at("2026-11-21"));
    expect(nc.deletePath).toHaveBeenCalledWith("anna-uid", folder("aaaa0005"));
    expect(cleanupEntries().some((e) => e.owner === "oid-gone")).toBe(false);

    // Never connected again: after 90 days the entry is given up (logged), it does not stay forever.
    await registerForCleanup("oid-never", file("aaaa0008"), "2026-11-01");
    await runCleanup(quietLog(), at("2027-01-15"));
    expect(cleanupEntries().some((e) => e.owner === "oid-never")).toBe(true);
    await runCleanup(quietLog(), at("2027-02-15"));
    expect(cleanupEntries().some((e) => e.owner === "oid-never")).toBe(false);
  });

  it("deletes each folder on its own – one Nextcloud refuses does not hold up the others", async () => {
    await registerForCleanup("oid-anna", file("aaaa0009"), "2026-11-01");
    await registerForCleanup("oid-anna", file("aaaa0010"), "2026-11-01");
    nc.deletePath.mockImplementation(async (_uid: string, target: string[]) => {
      if (target[1].endsWith("aaaa0009")) throw new Error("423 Locked");
    });
    await runCleanup(quietLog(), at("2026-11-20"));
    expect(nc.deletePath).toHaveBeenCalledWith("anna-uid", folder("aaaa0010"));
    expect(cleanupEntries().some((e) => e.folder[1].endsWith("aaaa0010"))).toBe(false);
    expect(cleanupEntries().some((e) => e.folder[1].endsWith("aaaa0009"))).toBe(true); // tried again next time
  });

  it("deletes on revoke right away and forgets the folder; keeps it while another link is valid", async () => {
    await registerForCleanup("oid-anna", file("aaaa0007"), "2027-03-01");
    const deleted: string[][] = [];
    const remove = async (target: string[]) => void deleted.push(target);

    expect(await expireOnRevoke("oid-anna", file("aaaa0007"), true, remove)).toBe("kept");
    expect(deleted).toEqual([]);

    expect(await expireOnRevoke("oid-anna", file("aaaa0007"), false, remove)).toBe("deleted");
    expect(deleted).toEqual([folder("aaaa0007")]);
    expect(cleanupEntries().some((e) => e.folder[1].endsWith("aaaa0007"))).toBe(false);

    expect(await expireOnRevoke("oid-anna", ["CloudShare-File-Requests", "x"], false, remove)).toBe("kept");
  });

  it("retries later when Nextcloud is unreachable", async () => {
    await registerForCleanup("oid-anna", file("aaaa0006"), "2026-11-01");
    nc.userInfo.mockRejectedValueOnce(new Error("ECONNREFUSED"));
    await runCleanup(quietLog(), at("2026-11-20"));
    expect(cleanupEntries().some((e) => e.folder[1].endsWith("aaaa0006"))).toBe(true);
    await runCleanup(quietLog(), at("2026-11-20") + DAY);
    expect(nc.deletePath).toHaveBeenCalledWith("anna-uid", folder("aaaa0006"));
  });
});
