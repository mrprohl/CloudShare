import { afterEach, describe, expect, it, vi } from "vitest";
import { mapReportStatus, maskPhone, normalizePhone, parseDeliveryReport, sendSms } from "../src/sms.js";
import { json, mockFetch } from "./helpers.js";

afterEach(() => vi.unstubAllGlobals());

describe("SMS (seven.io)", () => {
  it.each([
    ["0170 1234567", "+491701234567"],
    ["0170/123 45-67", "+491701234567"],
    ["+49 (170) 1234567", "+491701234567"],
    ["0043 664 1234567", "+436641234567"],
  ])("normalizes %j to E.164", (input, expected) => {
    expect(normalizePhone(input, "49")).toBe(expected);
  });

  it.each(["", "12", "abc", "+0123456789", "0170 12", "+49 170 1234567 99999"])("rejects %j", (input) => {
    expect(normalizePhone(input, "49")).toBeNull();
  });

  it("masks numbers for logs and the status line", () => {
    expect(maskPhone("+491701234567")).toBe("+49 170 •••• 67");
    expect(maskPhone("+12025550123")).toBe("+1 202 •••• 23");
  });

  it("sends with the API key from the server and returns the message ID", async () => {
    const { call } = mockFetch(json({ success: "100", messages: [{ id: 77123456789, success: true }] }));
    expect(await sendSms("+491701234567", "Ihr Kennwort: KateGeteFi45")).toBe("77123456789");
    const headers = call(0).init.headers as Record<string, string>;
    expect(headers["X-Api-Key"]).toBeTruthy();
    expect(JSON.parse(String(call(0).init.body))).toMatchObject({ to: "+491701234567", text: "Ihr Kennwort: KateGeteFi45" });
  });

  it("reports an SMS seven.io did not accept", async () => {
    mockFetch(json({ success: "402", messages: [] })); // e.g. no credit
    await expect(sendSms("+491701234567", "x")).rejects.toThrow(/SMS-Versand fehlgeschlagen.*402/);
  });

  it("reads delivery reports in the webhook and the classic format", () => {
    expect(parseDeliveryReport({ webhook_event: "dlr", data: { msg_id: "77", status: "DELIVERED" } }, {})).toEqual({
      messageId: "77",
      status: "DELIVERED",
    });
    expect(parseDeliveryReport({ msg_id: 78, status: "NOTDELIVERED" }, {})).toEqual({ messageId: "78", status: "NOTDELIVERED" });
    expect(parseDeliveryReport(undefined, { msg_id: "79", status: "BUFFERED" })).toEqual({ messageId: "79", status: "BUFFERED" });
    expect(parseDeliveryReport({ webhook_event: "sms_mo", data: { text: "hi" } }, {})).toBeNull();
    expect(mapReportStatus("DELIVERED")).toBe("delivered");
    expect(mapReportStatus("expired")).toBe("failed");
    expect(mapReportStatus("BUFFERED")).toBe("submitted");
    expect(mapReportStatus("SOMETHING")).toBe("unknown");
  });
});
