/**
 * Password rules and generator for CloudShare shares.
 * The rules must match server/src/password.ts (the server checks again).
 */

export const PASSWORD_MIN_LENGTH = 12;
export const GENERATED_LENGTH = 12;

// Generated passwords are pronounceable, so they can be read out on the phone: two capitalised "words" of two
// consonant-vowel syllables, one more syllable and two digits, e.g. "FuvoMozeBo28". No special characters, no easily
// confused characters (c/q/x/y/l, 0/1). (16 × 5)^5 × 8^2 ≈ 2^37.6 – enough for links that expire, are only reachable
// with the link itself and are protected by Nextcloud's brute-force throttling.
const CONSONANTS = "bdfghjkmnprstvwz";
const VOWELS = "aeiou";
const DIGITS = "23456789";

/** Uniform random index in [0, n) from the CSPRNG, without modulo bias (rejection sampling). */
function randomIndex(n: number): number {
  const limit = Math.floor(4294967296 / n) * n;
  const buffer = new Uint32Array(1);
  for (;;) {
    crypto.getRandomValues(buffer);
    if (buffer[0] < limit) return buffer[0] % n;
  }
}

const pick = (alphabet: string) => alphabet[randomIndex(alphabet.length)];
const syllable = () => pick(CONSONANTS) + pick(VOWELS);
const capital = (part: string) => part[0].toUpperCase() + part.slice(1);

/**
 * Pronounceable random password "WortWortSi12" (12 characters: upper and lower case, digits). Drawn again in the rare
 * case it would not pass the password rules (e.g. too few different characters).
 */
export function generatePassword(): string {
  for (;;) {
    const password =
      capital(syllable() + syllable()) + capital(syllable() + syllable()) + capital(syllable()) + pick(DIGITS) + pick(DIGITS);
    if (checkPassword(password).ok) return password;
  }
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
