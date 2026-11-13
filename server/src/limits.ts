/**
 * Turns user input into a whole number within [min, max]; missing or invalid input (NaN, 0, text) yields the
 * fallback. Used for link validity (days) and the number of accesses of the password link.
 */
export function clampInt(value: unknown, fallback: number, min: number, max: number): number {
  const parsed = Math.floor(Number(value ?? fallback));
  const number = Number.isFinite(parsed) && parsed !== 0 ? parsed : fallback;
  return Math.min(Math.max(number, min), max);
}

/**
 * Plausible single e-mail address (e.g. "Kennwort an andere Adresse"): local@domain.tld, no spaces, no list
 * separators or angle brackets, at most 254 characters. Deliverability is Exchange's business.
 */
export function isEmailAddress(value: string): boolean {
  return value.length <= 254 && /^[^\s@,;<>"()]+@[^\s@,;<>"()]+\.[^\s@,;<>"()]{2,}$/.test(value);
}
