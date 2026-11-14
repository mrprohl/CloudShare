/**
 * Server-side check of share passwords. Same rules as addin/src/shared/password.ts – keep them in sync.
 * Nextcloud's own password policy may be stricter; its message is passed through when the share is created.
 */
import { t } from "./i18n.js";

const PASSWORD_MIN_LENGTH = 12;

function isSequence(password: string): boolean {
  const codes = Array.from(password, (c) => c.charCodeAt(0));
  const step = codes[1] - codes[0];
  return Math.abs(step) === 1 && codes.every((c, i) => i === 0 || c - codes[i - 1] === step);
}

/** Returns a German message for the user if the password is too weak, otherwise null. */
export function passwordProblem(password: string): string | null {
  if (password.length < PASSWORD_MIN_LENGTH) {
    return t("server.passwordLength", { min: PASSWORD_MIN_LENGTH });
  }
  const classes = [/[A-ZÄÖÜ]/, /[a-zäöüß]/, /[0-9]/, /[^A-Za-zÄÖÜäöüß0-9]/].filter((r) => r.test(password)).length;
  if (classes < 3) {
    return t("server.passwordClasses");
  }
  if (new Set(password).size < 6 || isSequence(password)) {
    return t("server.passwordSimple");
  }
  return null;
}
