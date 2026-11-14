/**
 * Texts of the add-in in the language the server is configured for (LANGUAGE, see server/src/i18n.ts). The server
 * replaces the marker below with the texts of locales/<language>.json ("addin" section) when the container starts,
 * so the bundles need no extra request – important for the send handler in classic Outlook's JavaScript-only runtime.
 * The static texts of taskpane.html are filled in the same way (text markers, see server/src/i18n.ts).
 */

interface Bundle {
  lang: string;
  locale: string;
  messages: Record<string, string>;
}

const INJECTED = "%%CLOUDSHARE_I18N%%";

function parse(): Bundle {
  try {
    // Not replaced (unit tests): keys are shown as they are until setMessages() is called.
    if (INJECTED.charAt(0) !== "{") return { lang: "de", locale: "de-DE", messages: {} };
    return JSON.parse(INJECTED) as Bundle;
  } catch {
    return { lang: "de", locale: "de-DE", messages: {} };
  }
}

let bundle = parse();

/** Tests only: use these texts. */
export function setMessages(messages: Record<string, string>, locale = "de-DE"): void {
  bundle = { lang: locale.slice(0, 2), locale, messages };
}

/** BCP 47 locale for dates and numbers ("de-DE", "en-GB"). */
export const locale = () => bundle.locale;

/**
 * Text for `key` with {name} placeholders filled. With a numeric `count`, "key_one" (count 1) or "key_other" is
 * used if the language has it.
 */
export function t(key: string, vars: Record<string, string | number> = {}): string {
  const messages = bundle.messages;
  const count = vars.count;
  let text: string | undefined;
  if (typeof count === "number") text = messages[count === 1 ? `${key}_one` : `${key}_other`];
  if (text === undefined) text = messages[key];
  if (text === undefined) text = key;
  return text.replace(/\{(\w+)\}/g, (match, name: string) =>
    Object.prototype.hasOwnProperty.call(vars, name) ? String(vars[name]) : match,
  );
}
