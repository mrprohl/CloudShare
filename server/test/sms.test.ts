import { afterEach, describe, expect, it, vi } from "vitest";
import { callingCode, countryName, parseCountries, resolveCountry } from "../src/countries.js";
import { config, envFlag, envRegex } from "../src/config.js";
import {
  mapReportStatus,
  maskPhone,
  normalizePhone,
  parseDeliveryReport,
  sendSms,
  smsCountryAllowed,
  smsCountryList,
  smsNumberProblem,
} from "../src/sms.js";
import { json, mockFetch } from "./helpers.js";

afterEach(() => vi.unstubAllGlobals());

describe("SMS (seven.io)", () => {
  it.each([
    ["0170 1234567", "+491701234567"],
    ["0170/123 45-67", "+491701234567"],
    ["+49 (170) 1234567", "+491701234567"],
    ["0043 664 1234567", "+436641234567"],
    ["+49 (0)170 1234567", "+491701234567"],
    ["+49(0)170-1234567", "+491701234567"],
    ["tel:+491701234567", "+491701234567"],
    ["+49 170 – 123 45 67", "+491701234567"],
    ["＋41 79 123 45 67", "+41791234567"],
  ])("normalizes %j to E.164", (input, expected) => {
    expect(normalizePhone(input, "49")).toBe(expected);
  });

  describe("allowed countries (SMS_ALLOWED_COUNTRIES)", () => {
    it.each([
      ["DE", ["DE"]], ["de", ["DE"]], ["DEU", ["DE"]], ["+49", ["DE"]], ["0049", ["DE"]], ["49", ["DE"]],
      ["Deutschland", ["DE"]], ["Germany", ["DE"]], ["  deutschland ", ["DE"]], ["BRD", ["DE"]],
      ["Österreich", ["AT"]], ["Oesterreich", ["AT"]], ["Osterreich", ["AT"]], ["Austria", ["AT"]], ["AUT", ["AT"]],
      ["Schweiz", ["CH"]], ["Switzerland", ["CH"]], ["CHE", ["CH"]], ["+41", ["CH"]],
      ["Vereinigtes Königreich", ["GB"]], ["UK", ["GB"]], ["Großbritannien", ["GB"]], ["Grossbritannien", ["GB"]],
      ["USA", ["US"]], ["+1", ["US"]], ["Niederlande", ["NL"]], ["Holland", ["NL"]], ["Tschechien", ["CZ"]],
      ["Czech Republic", ["CZ"]], ["Türkei", ["TR"]], ["Turkey", ["TR"]], ["Côte d’Ivoire", ["CI"]],
    ])("reads %j", (entry, regions) => {
      expect(resolveCountry(entry)).toEqual(regions);
    });

    it("knows groups and keeps the order without duplicates", () => {
      expect(parseCountries("DACH")).toEqual(["DE", "AT", "CH"]);
      expect(parseCountries("EU")).toHaveLength(27);
      expect(parseCountries("EWR")).toHaveLength(30);
      expect(parseCountries("Deutschland; AUT, +41\nDE, dach")).toEqual(["DE", "AT", "CH"]);
      expect(parseCountries("")).toEqual([]);
      expect(parseCountries(" , ")).toEqual([]);
    });

    it("refuses entries that name no country", () => {
      expect(() => parseCountries("DE, Deutschlnd")).toThrow(/unknown country "Deutschlnd"/);
      expect(() => parseCountries("+999")).toThrow(/unknown country/);
    });

    it("checks a number by its calling code", () => {
      expect(callingCode("+491701234567")).toBe("49");
      expect(callingCode("+12025550123")).toBe("1");
      expect(callingCode("+4917")).toBe("49");
      expect(callingCode("+380501234567")).toBe("380");
      const dach = ["DE", "AT", "CH"];
      expect(smsCountryAllowed("+491701234567", dach)).toBe(true);
      expect(smsCountryAllowed("+436641234567", dach)).toBe(true);
      expect(smsCountryAllowed("+447700900123", dach)).toBe(false);
      expect(smsCountryAllowed("+447700900123", [])).toBe(true); // empty = all countries
      // Countries sharing a calling code cannot be told apart: Canada with "USA" allowed.
      expect(smsCountryAllowed("+14165550123", ["US"])).toBe(true);
    });

    it("allows landline and mobile numbers but no service numbers by default (SMS_BLOCKED_NUMBERS)", () => {
      const rules = { countries: [], allowedNumbers: null, blockedNumbers: config.smsBlockedNumbers };
      expect(smsNumberProblem("+491701234567", rules)).toBeNull(); // mobile
      expect(smsNumberProblem("+49301234567", rules)).toBeNull(); // landline (Berlin)
      expect(smsNumberProblem("+43800123456", rules)).toBeNull(); // only the German service numbers
      for (const number of ["0800 1234567", "0900 1234567", "01801 123456", "01806 123456", "+49 (0)800 1234567"]) {
        expect(smsNumberProblem(normalizePhone(number, "49")!, rules)).toMatch(/0800, 0180 oder 0900/);
      }
    });

    it("takes own rules as regular expressions – e.g. only German mobile numbers (SMS_ALLOWED_NUMBERS)", () => {
      const rules = { countries: [], allowedNumbers: envRegex("X", "", "^\\+491(5|6|7)\\d{8,9}$"), blockedNumbers: null };
      expect(smsNumberProblem("+491701234567", rules)).toBeNull();
      expect(smsNumberProblem("+4915112345678", rules)).toBeNull();
      expect(smsNumberProblem("+49301234567", rules)).toBe("Diese Rufnummer ist für SMS mit dem Kennwort nicht zugelassen.");
      expect(envRegex("SMS_BLOCKED_NUMBERS", "^\\+49", "")).toBeNull(); // set empty = no rule
      expect(() => envRegex("SMS_ALLOWED_NUMBERS", "", "^\\+49(")).toThrow(/SMS_ALLOWED_NUMBERS is no valid regular expression/);
    });

    it("checks the countries before the number rules", () => {
      const rules = { countries: ["DE"], allowedNumbers: null, blockedNumbers: config.smsBlockedNumbers };
      expect(smsNumberProblem("+447700900123", rules)).toBe("SMS mit dem Kennwort dürfen nur an Rufnummern aus Deutschland gehen.");
    });

    it("names the allowed countries in the language of the users", () => {
      expect(smsCountryList(["DE", "AT", "CH"], "de")).toBe("Deutschland, Österreich und Schweiz");
      expect(smsCountryList(["DE", "AT", "CH"], "en")).toBe("Germany, Austria, and Switzerland");
      expect(countryName("GB", "de")).toBe("Vereinigtes Königreich");
    });
  });

  it.each(["", "12", "abc", "+0123456789", "0170 12", "+49 170 1234567 99999"])("rejects %j", (input) => {
    expect(normalizePhone(input, "49")).toBeNull();
  });

  it("SMS_NOTICE switches the note about an undelivered SMS; unset it follows DOWNLOAD_NOTICE as before", () => {
    expect(config.smsNotice).toBe(true); // SMS_NOTICE=true in the tests
    expect(envFlag("SMS_NOTICE", "DOWNLOAD_NOTICE", { SMS_NOTICE: "true" })).toBe(true);
    expect(envFlag("SMS_NOTICE", "DOWNLOAD_NOTICE", { SMS_NOTICE: "false", DOWNLOAD_NOTICE: "true" })).toBe(false);
    expect(envFlag("SMS_NOTICE", "DOWNLOAD_NOTICE", { DOWNLOAD_NOTICE: "true" })).toBe(true); // older setups
    expect(envFlag("SMS_NOTICE", "DOWNLOAD_NOTICE", {})).toBe(false);
    expect(envFlag("UPLOAD_NOTICE", undefined, { DOWNLOAD_NOTICE: "true" })).toBe(false); // no fallback
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
