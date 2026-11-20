/**
 * "Feedback senden" (account menu, only with FEEDBACK_MAIL on the server): a bug report or an idea, sent by the server
 * from the user's own mailbox to the organisation's address (POST /api/feedback) – never to anything public. With
 * "Technische Angaben mitsenden" the details below go along; the user sees them before (`#feedback-details`). They
 * hold no mail content: client, versions, languages, window size and the last error messages of the pane.
 */
import { language, t } from "../shared/i18n";
import { api } from "../shared/office";

interface Options {
  token: () => Promise<string>;
  setStatus: (text: string, kind?: "success" | "error") => void;
  errorMessage: (err: unknown) => string;
  /** Back to the view, as "←" does. */
  close: () => void;
}

const $ = <T extends HTMLElement>(id: string) => document.getElementById(id) as T;

/** The last error messages the pane showed (time and text) – the most useful detail of a bug report. */
const errors: string[] = [];

/** Notes an error message of the pane for a later report (the last five). */
export function noteError(text: string): void {
  if (!text) return;
  const time = new Date().toLocaleTimeString("de-DE", { hour: "2-digit", minute: "2-digit", second: "2-digit" });
  errors.push(`${time} ${text}`);
  if (errors.length > 5) errors.shift();
}

/** The highest Mailbox requirement set the client supports (tells the Outlook generation). */
function mailboxApi(): string {
  try {
    for (let minor = 15; minor >= 1; minor--) {
      if (Office.context.requirements.isSetSupported("Mailbox", `1.${minor}`)) return `1.${minor}`;
    }
  } catch {
    // not in Outlook
  }
  return "?";
}

/** The technical details of a report, as label → value (shown to the user before sending). */
export function diagnostics(): Record<string, string> {
  const office = (() => {
    try {
      const d = Office.context.diagnostics;
      return `${d.host} ${d.version} (${d.platform})`;
    } catch {
      return "?";
    }
  })();
  const display = (() => {
    try {
      return Office.context.displayLanguage;
    } catch {
      return "?";
    }
  })();
  return {
    outlook: office,
    mailboxApi: mailboxApi(),
    outlookLanguage: display,
    addinLanguage: language(),
    window: `${window.innerWidth}×${window.innerHeight}`,
    browser: navigator.userAgent,
    time: new Date().toISOString(),
    ...(errors.length ? { errors: errors.join("\n") } : {}),
  };
}

/** Wires the feedback page (panel-feedback in taskpane.html); called once when taskpane.ts starts. */
export function initFeedback(options: Options) {
  const form = $<HTMLFormElement>("feedback-form");
  const text = $<HTMLTextAreaElement>("feedback-text");
  const withDetails = $<HTMLInputElement>("feedback-diagnostics");
  const toggle = $<HTMLButtonElement>("feedback-details-toggle");
  const details = $<HTMLPreElement>("feedback-details");
  const send = $<HTMLButtonElement>("feedback-send");

  const renderDetails = () => {
    details.textContent = Object.entries(diagnostics()).map(([key, value]) => `${t(`feedback.detail.${key}`)}: ${value}`).join("\n");
  };
  const showDetails = (open: boolean) => {
    details.hidden = !open;
    toggle.setAttribute("aria-expanded", String(open));
    toggle.textContent = t(open ? "feedback.hideDetails" : "feedback.showDetails");
    if (open) renderDetails();
  };
  toggle.addEventListener("click", () => showDetails(details.hidden));
  withDetails.addEventListener("change", () => {
    toggle.hidden = !withDetails.checked;
    if (!withDetails.checked) showDetails(false);
  });

  form.addEventListener("submit", async (event) => {
    event.preventDefault();
    const message = text.value.trim();
    if (!message) {
      text.setAttribute("aria-invalid", "true");
      text.focus();
      options.setStatus(t("feedback.empty"), "error");
      return;
    }
    text.removeAttribute("aria-invalid");
    const kind = form.querySelector<HTMLInputElement>('input[name="feedback-kind"]:checked')?.value === "bug" ? "bug" : "idea";
    send.disabled = true;
    try {
      await api("/api/feedback", {
        method: "POST",
        body: JSON.stringify({ kind, text: message, ...(withDetails.checked ? { diagnostics: diagnostics() } : {}) }),
      }, await options.token());
      // Empty for the next time: no text, a suggestion as the kind again.
      text.value = "";
      $<HTMLInputElement>("feedback-kind-idea").checked = true;
      showDetails(false);
      options.setStatus(t("feedback.sent"), "success");
      options.close();
    } catch (err) {
      options.setStatus(t("feedback.failed", { error: options.errorMessage(err) }), "error");
    } finally {
      send.disabled = false;
    }
  });

  /** Shown with the page: the cursor in the text. */
  const open = () => {
    showDetails(false);
    toggle.hidden = !withDetails.checked;
    text.focus();
  };
  return { open };
}
