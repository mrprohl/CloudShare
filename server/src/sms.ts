/**
 * SMS with the password via seven.io (https://docs.seven.io). Only the backend talks to seven.io – the API key never
 * reaches the add-in. Delivery reports arrive at /webhooks/seven (see app.ts) and are kept in deliveries.ts.
 */
import { HttpError } from "./auth.js";
import { config } from "./config.js";

export const smsEnabled = () => Boolean(config.sevenApiKey);

/**
 * Normalises a phone number to E.164 ("+4917012345678"): spaces, dashes, dots, slashes and brackets removed, "00" →
 * "+", a leading single "0" gets the default country code. Returns null if it is no plausible number.
 */
export function normalizePhone(input: string, defaultCountry = config.smsDefaultCountry): string | null {
  let number = input.trim().replace(/[\s\-./()]/g, "");
  if (number.startsWith("00")) number = `+${number.slice(2)}`;
  else if (/^0[1-9]/.test(number)) number = `+${defaultCountry}${number.slice(1)}`;
  return /^\+[1-9]\d{7,14}$/.test(number) ? number : null;
}

/** "+4917012345678" → "+49 170 •••• 78" for status lines and logs. */
export function maskPhone(e164: string): string {
  const digits = e164.slice(1);
  const country = digits.startsWith("1") || digits.startsWith("7") ? digits.slice(0, 1) : digits.slice(0, 2);
  const rest = digits.slice(country.length);
  return `+${country} ${rest.slice(0, 3)} •••• ${rest.slice(-2)}`;
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
    throw new HttpError(502, `SMS-Versand fehlgeschlagen (seven.io, HTTP ${res.status}, Code ${json?.success ?? "?"})`);
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
