/**
 * SMS with the password via seven.io (https://docs.seven.io). Only the backend talks to seven.io – the API key never
 * reaches the add-in. Delivery reports arrive at /webhooks/seven (see app.ts) and are kept in deliveries.ts.
 */
import { HttpError } from "./auth.js";
import { config } from "./config.js";
import { callingCode, callingCodeOf, countryName } from "./countries.js";
import { t } from "./i18n.js";

export const smsEnabled = () => Boolean(config.sevenApiKey);

/**
 * Normalises a phone number to E.164 ("+4917012345678"): a "tel:" prefix, the "(0)" of "+49 (0)170 …", spaces,
 * dashes, dots, slashes and brackets removed, "00" → "+", a leading single "0" gets the default country code.
 * Returns null if it is no plausible number. Same rules in the taskpane (normalizePhone in taskpane.ts).
 */
export function normalizePhone(input: string, defaultCountry = config.smsDefaultCountry): string | null {
  let number = input
    .trim()
    .replace(/^tel:/i, "")
    .replace(/＋/g, "+")
    .replace(/(\d)\s*\(0\)/g, "$1")
    .replace(/[\s\-‐‑–—./()]/g, "");
  if (number.startsWith("00")) number = `+${number.slice(2)}`;
  else if (/^0[1-9]/.test(number)) number = `+${defaultCountry}${number.slice(1)}`;
  return /^\+[1-9]\d{7,14}$/.test(number) ? number : null;
}

/** May an SMS go to this number (E.164)? Checks its calling code against SMS_ALLOWED_COUNTRIES (empty = all). */
export function smsCountryAllowed(e164: string, allowed: string[] = config.smsAllowedCountries): boolean {
  if (allowed.length === 0) return true;
  const code = callingCode(e164);
  return code !== null && allowed.some((region) => callingCodeOf(region) === code);
}

/** The allowed countries as text for users: "Deutschland, Österreich und Schweiz"; "" = all countries. */
export function smsCountryList(allowed: string[] = config.smsAllowedCountries, language = config.language): string {
  const names = allowed.map((region) => countryName(region, language));
  try {
    return new Intl.ListFormat([language, "en"], { type: "conjunction" }).format(names);
  } catch {
    return names.join(", ");
  }
}

/** Calling codes of the allowed countries for the taskpane ([] = all). */
export const smsCountryCodes = () => [...new Set(config.smsAllowedCountries.map((r) => callingCodeOf(r)!))];

export interface SmsRules {
  countries: string[];
  allowedNumbers: RegExp | null;
  blockedNumbers: RegExp | null;
}
const configuredRules = (): SmsRules => ({
  countries: config.smsAllowedCountries,
  allowedNumbers: config.smsAllowedNumbers,
  blockedNumbers: config.smsBlockedNumbers,
});

/**
 * Why no SMS may go to this number (E.164), as text for the user, or null: SMS_ALLOWED_COUNTRIES, then
 * SMS_ALLOWED_NUMBERS / SMS_BLOCKED_NUMBERS. Same checks in the taskpane (smsNumberProblem in taskpane.ts).
 */
export function smsNumberProblem(e164: string, rules: SmsRules = configuredRules()): string | null {
  if (!smsCountryAllowed(e164, rules.countries)) return t("server.smsCountryNotAllowed", { countries: smsCountryList(rules.countries) });
  if (rules.allowedNumbers && !rules.allowedNumbers.test(e164)) return t("server.smsNumberNotAllowed");
  if (rules.blockedNumbers?.test(e164)) return t("server.smsNumberBlocked");
  return null;
}

/** A number entered by the user → E.164, or HttpError 400 if it is no valid number or no SMS may go to it. */
export function smsNumber(input: unknown): string {
  const number = typeof input === "string" ? normalizePhone(input) : null;
  if (!number) throw new HttpError(400, t("server.phoneInvalid"));
  const problem = smsNumberProblem(number);
  if (problem) throw new HttpError(400, problem);
  return number;
}

/** Calling code and the rest of an E.164 number ("+4917012345678" → ["49", "17012345678"]). */
function phoneParts(e164: string): [string, string] {
  const digits = e164.slice(1);
  const country = digits.startsWith("1") || digits.startsWith("7") ? digits.slice(0, 1) : digits.slice(0, 2);
  return [country, digits.slice(country.length)];
}

/** "+4917012345678" → "+49 170 •••• 78" for logs and pages a recipient sees (who may not own that phone). */
export function maskPhone(e164: string): string {
  const [country, rest] = phoneParts(e164);
  return `+${country} ${rest.slice(0, 3)} •••• ${rest.slice(-2)}`;
}

/**
 * "+4917012345678" → "+49 170 12345678": the whole number, readable, for the sender – "Verlauf", the delivery proof
 * and the notes to the sender (the sender typed it in; a proof must name where the password went).
 */
export function formatPhone(e164: string): string {
  const [country, rest] = phoneParts(e164);
  return `+${country} ${rest.slice(0, 3)} ${rest.slice(3)}`;
}

/** Sends one SMS; returns seven.io's message ID. Throws HttpError if seven.io does not accept it. */
export async function sendSms(to: string, text: string): Promise<string> {
  const res = await fetch(`${config.sevenUrl}/sms`, {
    method: "POST",
    headers: { "X-Api-Key": config.sevenApiKey, Accept: "application/json", "Content-Type": "application/json" },
    // Delivery reports are requested per message; the callback URL is set in the seven.io dashboard.
    body: JSON.stringify({ to, from: config.sevenSender, text, flash: 0, delay: 0 }),
  });
  const json = (await res.json().catch(() => null)) as { success?: string | number; messages?: Array<{ id?: string | number; success?: boolean }> } | null;
  const id = json?.messages?.[0]?.id;
  if (!res.ok || String(json?.success) !== "100" || id === undefined || id === null) {
    throw new HttpError(502, t("server.smsFailed", { status: res.status, code: String(json?.success ?? "?") }));
  }
  return String(id);
}

export type SmsStatus = "submitted" | "delivered" | "failed" | "unknown";

/** seven.io report status → own status (DELIVERED, NOTDELIVERED, BUFFERED, TRANSMITTED, ACCEPTED, EXPIRED, REJECTED, FAILED). */
export function mapReportStatus(status: string): SmsStatus {
  const s = status.toUpperCase();
  if (s === "DELIVERED") return "delivered";
  if (["NOTDELIVERED", "EXPIRED", "REJECTED", "FAILED"].includes(s)) return "failed";
  if (["BUFFERED", "TRANSMITTED", "ACCEPTED"].includes(s)) return "submitted";
  return "unknown";
}

/**
 * Delivery report from seven.io – the webhook format ({ webhook_event: "dlr", data: { msg_id, status, … } }) or the
 * classic callback with msg_id/status as fields or query parameters. Returns null for anything else.
 */
export function parseDeliveryReport(body: unknown, query: Record<string, unknown>): { messageId: string; status: string } | null {
  const pick = (source: unknown) => {
    if (!source || typeof source !== "object") return null;
    const s = source as Record<string, unknown>;
    const id = s.msg_id ?? s.message_id ?? s.id;
    return (typeof id === "string" || typeof id === "number") && typeof s.status === "string"
      ? { messageId: String(id), status: s.status }
      : null;
  };
  const b = body as { data?: unknown } | undefined;
  return pick(b?.data) ?? pick(body) ?? pick(query);
}
