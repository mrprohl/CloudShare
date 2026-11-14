import { t } from "./i18n";

/**
 * Office SSO errors (Office.auth.getAccessToken) are plain objects ({ code, name, message }), not Error instances.
 * This turns them into a readable message with the code, so users can tell support what failed.
 * Codes: https://learn.microsoft.com/office/dev/add-ins/develop/troubleshoot-sso-in-office-add-ins
 * The texts are "sso.<code>" in the language file.
 */
const SSO_CODES = [13000, 13001, 13002, 13003, 13004, 13005, 13006, 13007, 13008, 13010, 13012, 13013];

interface OfficeErrorLike {
  code?: number | string;
  message?: string;
}

/** Readable message for any error, including Office SSO errors. */
export function errorText(err: unknown): string {
  if (err instanceof Error) return err.message;
  if (err && typeof err === "object") {
    const { code, message } = err as OfficeErrorLike;
    const numeric = Number(code);
    if (code !== undefined && SSO_CODES.indexOf(numeric) >= 0) {
      return t("error.withCode", { message: t(`sso.${numeric}`), code: numeric });
    }
    if (message) return code !== undefined ? t("error.withCode", { message, code }) : message;
    if (code !== undefined) return t("error.withCode", { message: t("error.unknown"), code });
    try {
      return JSON.stringify(err);
    } catch {
      // fall through
    }
  }
  return String(err);
}

/** Error thrown by getToken(): keeps the SSO code for callers and shows a readable message. */
export class SsoError extends Error {
  constructor(public readonly ssoCode: number | undefined, message: string) {
    super(message);
    this.name = "SsoError";
  }
}

export function toSsoError(err: unknown): SsoError {
  const code = err && typeof err === "object" ? Number((err as OfficeErrorLike).code) : NaN;
  return new SsoError(Number.isFinite(code) ? code : undefined, errorText(err));
}
