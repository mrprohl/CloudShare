/**
 * Personal settings of a user (⚙ in the add-in): notices, the ways to send the password the pane offers and the one
 * it preselects, the language of the add-in and an own signature for the password mails. Stored per Entra oid in the
 * database (store "settings", kept however old) – only the fields the user changed, so a new default in the
 * environment (config.userDefaults, DEFAULT_*) still reaches everyone who left it alone.
 *
 * What the organisation decides stays out of reach: a notice the server has switched off (DOWNLOAD_NOTICE …), SMS
 * without seven.io, and every SMS rule (countries, allowed and blocked numbers) – settings only choose within that.
 */
import { config } from "./config.js";
import { PersistentMap } from "./db.js";
import { HttpError } from "./auth.js";
import { addinLanguages, t } from "./i18n.js";
import type { Signature } from "./signature.js";
import { smsEnabled } from "./sms.js";

export type Channel = "recipients" | "mail" | "sms";
export const CHANNELS: readonly Channel[] = ["recipients", "mail", "sms"];

/** Own signature of the password mails: plain text, profile placeholders like {phone}, empty line = new paragraph. */
export type { Signature } from "./signature.js";

export interface UserSettings {
  downloadNotice: boolean;
  uploadNotice: boolean;
  smsNotice: boolean;
  /** "Zustellnachweis": a mail when nothing was fetched a few days before the link ends (PROOF_REMINDER_DAYS). */
  proofReminder: boolean;
  /** Ways the pane offers for "Kennwort senden an" (at least one). */
  channels: Channel[];
  /** The way preselected for a new mail – one of `channels`. */
  defaultChannel: Channel;
  /** "auto" (as in Outlook) or a language of the add-in. */
  language: string;
  /** Language preselected for what the recipients get (link card, password mail, SMS); changeable per mail. */
  recipientLanguage: string;
  /** Own signature of the password mails; null = the organisation's (MAIL_FOOTER, filled from the profile). */
  signature: Signature | null;
}

/** Longest signature (characters). */
export const SIGNATURE_MAX = 20_000;

const stored = new PersistentMap<Partial<UserSettings>>("settings", { retention: false });

/** What the server offers at all – settings choose within it. */
export function availableSettings() {
  return {
    downloadNotice: config.downloadNotice,
    uploadNotice: config.uploadNotice,
    smsNotice: config.smsNotice && smsEnabled(),
    proofReminder: config.deliveryProof && config.proofReminderDays > 0,
    channels: CHANNELS.filter((c) => c !== "sms" || smsEnabled()),
    languages: addinLanguages(),
  };
}

/** LANGUAGE if the add-in offers it, else English, else the first language. */
const fallbackRecipientLanguage = (languages: string[]) =>
  languages.includes(config.language) ? config.language : languages.includes("en") ? "en" : languages[0];

/** Keeps a channel list within what is available; never empty. */
function channelsWithin(list: unknown, available: Channel[]): Channel[] {
  const wanted = Array.isArray(list) ? available.filter((c) => list.includes(c)) : [];
  return wanted.length > 0 ? wanted : available;
}

/** `values` made consistent with what the server offers (notices off where switched off, channels available …). */
function effective(values: UserSettings): UserSettings {
  const available = availableSettings();
  const channels = channelsWithin(values.channels, available.channels);
  return {
    downloadNotice: available.downloadNotice && values.downloadNotice,
    uploadNotice: available.uploadNotice && values.uploadNotice,
    smsNotice: available.smsNotice && values.smsNotice,
    proofReminder: available.proofReminder && values.proofReminder,
    channels,
    defaultChannel: channels.includes(values.defaultChannel) ? values.defaultChannel : channels[0],
    language: values.language === "auto" || available.languages.includes(values.language) ? values.language : "auto",
    recipientLanguage: available.languages.includes(values.recipientLanguage) ? values.recipientLanguage : fallbackRecipientLanguage(available.languages),
    // Only text signatures count – a formatted one saved by an earlier version counts as none.
    signature: values.signature?.format === "text" && typeof values.signature.content === "string" ? values.signature : null,
  };
}

/** The defaults from the environment (DEFAULT_*), as every new user starts. */
export function defaultSettings(): UserSettings {
  const d = config.userDefaults;
  return effective({
    downloadNotice: d.downloadNotice,
    uploadNotice: d.uploadNotice,
    smsNotice: d.smsNotice,
    proofReminder: d.proofReminder,
    channels: d.channels.filter((c): c is Channel => (CHANNELS as readonly string[]).includes(c)),
    defaultChannel: (CHANNELS as readonly string[]).includes(d.defaultChannel) ? (d.defaultChannel as Channel) : "recipients",
    language: d.language,
    recipientLanguage: config.language,
    signature: null,
  });
}

/** The settings that apply to `oid`: the defaults, the user's own choices on top, within what the server offers. */
export function settingsFor(oid: string): UserSettings {
  return effective({ ...defaultSettings(), ...(stored.get(oid) ?? {}) });
}

const invalid = (what: string) => new HttpError(400, t("server.settingsInvalid", { what }), "SETTINGS_INVALID");

/** Checks a change from the add-in; returns only the fields it contains. */
export function parseSettings(body: unknown): Partial<UserSettings> {
  if (!body || typeof body !== "object") throw invalid("body");
  const input = body as Record<string, unknown>;
  const result: Partial<UserSettings> = {};
  for (const key of ["downloadNotice", "uploadNotice", "smsNotice", "proofReminder"] as const) {
    if (key in input) {
      if (typeof input[key] !== "boolean") throw invalid(key);
      result[key] = input[key] as boolean;
    }
  }
  if ("channels" in input) {
    const list = input.channels;
    if (!Array.isArray(list) || list.length === 0 || !list.every((c) => (CHANNELS as readonly unknown[]).includes(c))) {
      throw invalid("channels");
    }
    result.channels = [...new Set(list as Channel[])];
  }
  if ("defaultChannel" in input) {
    if (!(CHANNELS as readonly unknown[]).includes(input.defaultChannel)) throw invalid("defaultChannel");
    result.defaultChannel = input.defaultChannel as Channel;
  }
  if ("language" in input) {
    const language = input.language;
    if (typeof language !== "string" || (language !== "auto" && !addinLanguages().includes(language))) throw invalid("language");
    result.language = language;
  }
  if ("recipientLanguage" in input) {
    const lang = input.recipientLanguage;
    if (typeof lang !== "string" || !addinLanguages().includes(lang)) throw invalid("recipientLanguage");
    result.recipientLanguage = lang;
  }
  if ("signature" in input) {
    const signature = input.signature;
    if (signature === null) {
      result.signature = null;
    } else {
      // Plain text only – no markup is ever taken over into a mail.
      const s = signature as { format?: unknown; content?: unknown } | undefined;
      if (s?.format !== "text" || typeof s.content !== "string") throw invalid("signature");
      if (s.content.length > SIGNATURE_MAX) throw new HttpError(400, t("server.signatureTooLong", { max: SIGNATURE_MAX }), "SETTINGS_INVALID");
      // An empty signature is no signature: the organisation's applies again.
      result.signature = s.content.trim() ? { format: "text", content: s.content } : null;
    }
  }
  return result;
}

/** Stores the user's changes (on top of earlier ones) and returns the settings that apply now. */
export function saveSettings(oid: string, changes: Partial<UserSettings>): UserSettings {
  stored.set(oid, { ...(stored.get(oid) ?? {}), ...changes });
  return settingsFor(oid);
}

/** Back to the defaults of the environment. */
export function resetSettings(oid: string): UserSettings {
  stored.delete(oid);
  return settingsFor(oid);
}
