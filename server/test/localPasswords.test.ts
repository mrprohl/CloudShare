import { afterEach, describe, expect, it, vi } from "vitest";
import { config } from "../src/config.js";
import { HttpError } from "../src/auth.js";
import {
  createMessageSend,
  createPasswordSend,
  deleteSend,
  isLocalPassword,
  purgeLocalPasswords,
  readSendPassword,
  sendAccessCount,
  setSendExpiry,
  setSendMaxAccess,
} from "../src/vaultwarden.js";
import { mockFetch } from "./helpers.js";

afterEach(() => {
  config.vaultwardenEnabled = true;
  vi.unstubAllGlobals();
});

describe("passwords without Vaultwarden (VAULTWARDEN_ENABLED=false)", () => {
  it("are kept by CloudShare itself until the links expire – no Send, no URL, nothing to bw serve", async () => {
    config.vaultwardenEnabled = false;
    const { fetchMock } = mockFetch();
    const keepUntil = Date.now() + 40 * 86_400_000; // longer than Vaultwarden would allow
    const send = await createPasswordSend("CloudShare: a.pdf", "KateGeteFi45", 3, 5, keepUntil);
    expect(send.url).toBe("");
    expect(isLocalPassword(send.id)).toBe(true);
    expect(send.expiresAt).toBe(keepUntil);
    expect(await readSendPassword(send.id!)).toBe("KateGeteFi45");
    // No Send to change or count.
    await setSendExpiry(send.id!, Date.now());
    await setSendMaxAccess(send.id!, 1);
    expect(await sendAccessCount(send.id!)).toBeNull();
    expect(fetchMock).not.toHaveBeenCalled(); // bw serve was never asked

    await deleteSend(send.id!);
    await expect(readSendPassword(send.id!)).rejects.toBeInstanceOf(HttpError);
  });

  it("are deleted once their links expired – also after Vaultwarden was switched on again", async () => {
    config.vaultwardenEnabled = false;
    const old = await createPasswordSend("CloudShare: alt.pdf", "AltesKennwort1!", 3, 5, Date.now() + 1000);
    const current = await createPasswordSend("CloudShare: neu.pdf", "NeuesKennwort1!", 3, 5, Date.now() + 86_400_000);
    config.vaultwardenEnabled = true;
    expect(await readSendPassword(current.id!)).toBe("NeuesKennwort1!"); // a "local:" ID is read locally whatever the switch
    expect(purgeLocalPasswords(Date.now() + 2000)).toBe(1);
    await expect(readSendPassword(old.id!)).rejects.toBeInstanceOf(HttpError);
    expect(await readSendPassword(current.id!)).toBe("NeuesKennwort1!");
  });

  it("refuse confidential messages – they are Vaultwarden Sends", async () => {
    config.vaultwardenEnabled = false;
    await expect(createMessageSend("Vertrauliche Nachricht", "Text", "KateGeteFi45", new Date(Date.now() + 86_400_000)))
      .rejects.toMatchObject({ status: 409, code: "NO_VAULTWARDEN" });
  });
});
