/**
 * Server-side check of share passwords. Same rules as addin/src/shared/password.ts – keep them in sync.
 * Nextcloud's own password policy may be stricter; its message is passed through when the share is created.
 */

const PASSWORD_MIN_LENGTH = 12;

function isSequence(password: string): boolean {
  const codes = Array.from(password, (c) => c.charCodeAt(0));
  const step = codes[1] - codes[0];
  return Math.abs(step) === 1 && codes.every((c, i) => i === 0 || c - codes[i - 1] === step);
}

/** Returns a German message for the user if the password is too weak, otherwise null. */
export function passwordProblem(password: string): string | null {
  if (password.length < PASSWORD_MIN_LENGTH) {
    return `Das Kennwort muss mindestens ${PASSWORD_MIN_LENGTH} Zeichen lang sein.`;
  }
  const classes = [/[A-ZÄÖÜ]/, /[a-zäöüß]/, /[0-9]/, /[^A-Za-zÄÖÜäöüß0-9]/].filter((r) => r.test(password)).length;
  if (classes < 3) {
    return "Das Kennwort muss mindestens drei der folgenden Zeichenarten enthalten: Großbuchstaben, Kleinbuchstaben, Ziffern, Sonderzeichen.";
  }
  if (new Set(password).size < 6 || isSequence(password)) {
    return "Das Kennwort ist zu einfach (Wiederholung oder Zeichenfolge).";
  }
  return null;
}
