/**
 * Password rules and generator for CloudShare shares.
 * The rules must match server/src/password.ts (the server checks again).
 */

export const PASSWORD_MIN_LENGTH = 12;
export const GENERATED_LENGTH = 20;

// Generated passwords avoid easily confused characters (0/O, 1/l/I) and symbols that are awkward to type or
// break when copied from a mail ("'`\ and spaces). 20 characters from 67 symbols ≈ 121 bits of entropy.
const UPPER = "ABCDEFGHJKLMNPQRSTUVWXYZ";
const LOWER = "abcdefghijkmnopqrstuvwxyz";
const DIGITS = "23456789";
const SYMBOLS = "!#%*+-=?@_";

/** Uniform random index in [0, n) from the CSPRNG, without modulo bias (rejection sampling). */
function randomIndex(n: number): number {
  const limit = Math.floor(4294967296 / n) * n;
  const buffer = new Uint32Array(1);
  for (;;) {
    crypto.getRandomValues(buffer);
    if (buffer[0] < limit) return buffer[0] % n;
  }
}

/** Random password with at least one character of every class, shuffled (Fisher–Yates). */
export function generatePassword(length = GENERATED_LENGTH): string {
  const groups = [UPPER, LOWER, DIGITS, SYMBOLS];
  const all = groups.join("");
  const chars = groups.map((g) => g[randomIndex(g.length)]);
  while (chars.length < length) chars.push(all[randomIndex(all.length)]);
  for (let i = chars.length - 1; i > 0; i--) {
    const j = randomIndex(i + 1);
    [chars[i], chars[j]] = [chars[j], chars[i]];
  }
  return chars.join("");
}

export interface PasswordCheck {
  /** At least PASSWORD_MIN_LENGTH characters. */
  length: boolean;
  /** At least 3 of: upper case, lower case, digits, other characters. */
  classes: boolean;
  /** Not a repetition or a simple sequence like "abcdef…"/"123456…", at least 6 different characters. */
  variety: boolean;
  ok: boolean;
}

/** Ascending or descending run of consecutive characters over the whole password (e.g. "123456789012"). */
function isSequence(password: string): boolean {
  const codes = Array.from(password, (c) => c.charCodeAt(0));
  const step = codes[1] - codes[0];
  return Math.abs(step) === 1 && codes.every((c, i) => i === 0 || c - codes[i - 1] === step);
}

export function checkPassword(password: string): PasswordCheck {
  const classes = [/[A-ZÄÖÜ]/, /[a-zäöüß]/, /[0-9]/, /[^A-Za-zÄÖÜäöüß0-9]/].filter((r) => r.test(password)).length;
  const result = {
    length: password.length >= PASSWORD_MIN_LENGTH,
    classes: classes >= 3,
    variety: new Set(password).size >= 6 && !isSequence(password),
  };
  return { ...result, ok: result.length && result.classes && result.variety };
}
