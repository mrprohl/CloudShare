import type { Contact } from "./graph.js";

/** One suggestion under "Kennwort an": the contact's name and the address or number to fill in. */
export interface ContactSuggestion {
  name: string;
  value: string;
  /** Kind of number (SMS only). */
  type?: "mobile" | "business" | "home";
}

/** Most suggestions shown under the field. */
export const MAX_SUGGESTIONS = 8;

/** Lower case without accents – "Müller" is found by "mul" and by "mue". */
function forms(text: string): string[] {
  const lower = text.toLowerCase();
  const plain = lower.normalize("NFD").replace(/[̀-ͯ]/g, "");
  const spelled = lower.replace(/ä/g, "ae").replace(/ö/g, "oe").replace(/ü/g, "ue").replace(/ß/g, "ss")
    .normalize("NFD").replace(/[̀-ͯ]/g, "");
  return plain === spelled ? [plain] : [plain, spelled];
}

/** 0 when a word starts with the query, 1 when it is somewhere inside, null when not at all. */
function textScore(text: string, query: string[]): number | null {
  let best: number | null = null;
  for (const form of forms(text)) {
    for (const q of query) {
      const at = form.indexOf(q);
      if (at < 0) continue;
      const score = at === 0 || /[\s.,@_\-'(]/.test(form[at - 1]) ? 0 : 1;
      best = best === null ? score : Math.min(best, score);
    }
  }
  return best;
}

/**
 * The digits of a number with its country code – "0171 …", "0049 171 …" and "+49 (0) 171 …" all become "49171…"; a
 * part without prefix ("171") stays as it is and is found inside.
 */
function digits(text: string, country: string): string {
  const number = text.replace(/(\d)\s*\(0\)/g, "$1").replace(/[^\d+]/g, "");
  if (number.startsWith("+")) return number.replace(/\D/g, "");
  if (number.startsWith("00")) return number.slice(2);
  if (number.startsWith("0")) return country + number.slice(1);
  return number;
}

/**
 * The contacts matching `query` (name, address or – for SMS – the digits of the number), one entry per address or
 * number (`country`: calling code for national numbers, SMS_DEFAULT_COUNTRY), best first: a word starting with the query before a match inside, then by name.
 */
export function matchContacts(contacts: Contact[], query: string, kind: "mail" | "sms", country = "49"): ContactSuggestion[] {
  const q = query.trim();
  if (q.length < 2) return [];
  const queryForms = forms(q);
  const queryDigits = /^[\d\s+()\-./]+$/.test(q) ? digits(q, country) : "";
  const found: (ContactSuggestion & { score: number })[] = [];
  for (const contact of contacts) {
    const byName = textScore(contact.name, queryForms);
    if (kind === "mail") {
      for (const address of contact.emails) {
        const byAddress = textScore(address, queryForms);
        const score = Math.min(byName ?? 9, byAddress ?? 9);
        if (score < 9) found.push({ name: contact.name, value: address, score });
      }
    } else {
      for (const phone of contact.phones) {
        const byNumber = queryDigits.length >= 2 && digits(phone.number, country).includes(queryDigits) ? 0 : null;
        const score = Math.min(byName ?? 9, byNumber ?? 9);
        if (score < 9) found.push({ name: contact.name, value: phone.number.trim(), type: phone.type, score });
      }
    }
  }
  const seen = new Set<string>();
  return found
    .sort((a, b) => a.score - b.score || a.name.localeCompare(b.name))
    .filter((s) => {
      const key = `${s.name}|${s.value.toLowerCase()}`;
      if (seen.has(key)) return false;
      seen.add(key);
      return true;
    })
    .slice(0, MAX_SUGGESTIONS)
    .map(({ score: _score, ...suggestion }) => suggestion);
}

/** Most entries of the address book page – one per address or number. */
export const MAX_BOOK = 2000;

/**
 * The address book page ("Adressbuch" next to the field): every address (`mail`) or number (`sms`) of the contacts,
 * by name from A to Z – the pane filters it while the user types.
 */
export function listContacts(contacts: Contact[], kind: "mail" | "sms"): ContactSuggestion[] {
  const entries: ContactSuggestion[] = [];
  for (const contact of contacts) {
    if (kind === "mail") for (const address of contact.emails) entries.push({ name: contact.name, value: address });
    else for (const phone of contact.phones) entries.push({ name: contact.name, value: phone.number.trim(), type: phone.type });
  }
  const seen = new Set<string>();
  return entries
    .sort((a, b) => a.name.localeCompare(b.name, "de", { sensitivity: "base" }) || a.value.localeCompare(b.value))
    .filter((entry) => {
      const key = `${entry.name}|${entry.value.toLowerCase()}`;
      if (seen.has(key)) return false;
      seen.add(key);
      return true;
    })
    .slice(0, MAX_BOOK);
}
