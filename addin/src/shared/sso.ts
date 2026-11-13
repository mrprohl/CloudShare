/**
 * Office SSO errors (Office.auth.getAccessToken) are plain objects ({ code, name, message }), not Error instances.
 * This turns them into a German message with the code, so users can tell support what failed.
 * Codes: https://learn.microsoft.com/office/dev/add-ins/develop/troubleshoot-sso-in-office-add-ins
 */
const SSO_HINTS: Record<number, string> = {
  13000: "Diese Outlook-Version unterstützt die Anmeldung für Add-ins nicht.",
  13001: "Sie sind in Outlook nicht mit Ihrem Microsoft-Konto angemeldet.",
  13002: "Die Anmeldung oder Zustimmung wurde abgebrochen.",
  13003: "Dieses Konto wird nicht unterstützt – bitte mit dem Geschäftskonto anmelden.",
  13004: "Die Anmeldekonfiguration des Add-ins ist ungültig (Manifest). Bitte die IT informieren.",
  13005: "Die Berechtigungen für CloudShare wurden noch nicht bestätigt. Bitte die IT informieren (Administratorzustimmung).",
  13006: "Outlook konnte die Anmeldung nicht abschließen. Bitte Outlook neu starten.",
  13007:
    "Outlook hat kein Anmeldetoken für CloudShare erhalten. Bitte die IT informieren (App-Registrierung: " +
    "Anwendungs-ID-URI, autorisierte Clientanwendungen, Administratorzustimmung).",
  13008: "Eine Anmeldung läuft bereits. Bitte kurz warten und erneut versuchen.",
  13010: "Die Anmeldung ist in diesem Browser blockiert (Sicherheitszonen). Bitte Outlook-Seite und Anmeldeseite derselben Zone zuordnen.",
  13012: "Die Anmeldung wird in dieser Umgebung nicht unterstützt (Add-in muss über das Admin Center bereitgestellt sein).",
  13013: "Zu viele Anmeldeversuche. Bitte kurz warten und erneut versuchen.",
};

interface OfficeErrorLike {
  code?: number | string;
  message?: string;
}

/** Readable German message for any error, including Office SSO errors. */
export function errorText(err: unknown): string {
  if (err instanceof Error) return err.message;
  if (err && typeof err === "object") {
    const { code, message } = err as OfficeErrorLike;
    const numeric = Number(code);
    if (code !== undefined && SSO_HINTS[numeric]) return `${SSO_HINTS[numeric]} (Code ${numeric})`;
    if (message || code !== undefined) return [message || "Unbekannter Fehler", code !== undefined ? `(Code ${code})` : ""].join(" ").trim();
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
