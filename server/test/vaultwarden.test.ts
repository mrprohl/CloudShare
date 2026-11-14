/**
 * Vaultwarden Sends through `bw serve` (fetch mocked): the recipients' access and the lifetime of the Send are two
 * dates – the Send lives until the links expire, so its sender can still read the password and send it again.
 */
import { afterEach, describe, expect, it, vi } from "vitest";
import { json, mockFetch } from "./helpers.js";

const { createPasswordSend, setSendExpiry } = await import("../src/vaultwarden.js");
const DAY = 86_400_000;
const days = (iso: unknown) => Math.round((Date.parse(String(iso)) - Date.now()) / DAY);

afterEach(() => vi.unstubAllGlobals());

describe("password Sends", () => {
  it("lets the recipients open it for the chosen days and keeps it until the links expire", async () => {
    const { call } = mockFetch(json({ success: true, data: { accessUrl: "https://vault.example.test/#/send/a/b", id: "s1" } }));
    const send = await createPasswordSend("CloudShare: a.pdf", "KateGeteFi45", 3, 5, Date.now() + 14 * DAY);
    const body = JSON.parse(String(call(0).init.body));
    expect(body).toMatchObject({ maxAccessCount: 3, text: { text: "KateGeteFi45", hidden: true } });
    expect(days(body.expirationDate)).toBe(5); // "Kennwort-Link gültig für"
    expect(days(body.deletionDate)).toBe(14); // the links' expiry
    expect(send).toMatchObject({ url: "https://vault.example.test/#/send/a/b", id: "s1" });
    expect(Math.round((send.expiresAt - Date.now()) / DAY)).toBe(5);
  });

  it("never lets the recipients' access outlast the Send", async () => {
    const { call } = mockFetch(json({ success: true, data: { accessUrl: "https://vault.example.test/#/send/a/b", id: "s2" } }));
    await createPasswordSend("CloudShare: a.pdf", "KateGeteFi45", 3, 30, Date.now() + 3 * DAY);
    const body = JSON.parse(String(call(0).init.body));
    expect(days(body.expirationDate)).toBe(3);
    expect(days(body.deletionDate)).toBe(3);
  });

  it("moves only the recipients' access, never past the deletion date", async () => {
    const deletionDate = new Date(Date.now() + 10 * DAY).toISOString();
    const { call } = mockFetch(
      json({ success: true, data: { id: "s3", deletionDate, expirationDate: new Date(Date.now() + DAY).toISOString() } }),
      json({ success: true, data: {} }),
    );
    await setSendExpiry("s3", Date.now() + 30 * DAY);
    const put = JSON.parse(String(call(1).init.body));
    expect(call(1).init.method).toBe("PUT");
    expect(put.deletionDate).toBe(deletionDate);
    expect(days(put.expirationDate)).toBe(10);
  });
});
