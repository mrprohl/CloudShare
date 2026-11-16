import { forceLanguage, language } from "../shared/i18n";

/**
 * The taskpane in the language of the Outlook client. taskpane.html comes in LANGUAGE (the server's first guess);
 * the server also writes taskpane.<language>.html for every language (server/src/i18n.ts, localizePublic). When the
 * client has another language, its page replaces the body before taskpane.ts runs – so all element references and
 * handlers are made on the final page. If that page cannot be loaded, everything stays in the page's language
 * (texts of the page and of the code always match).
 */
export async function applyLanguage(): Promise<void> {
  const page = document.documentElement.lang;
  const wanted = language();
  if (!page || wanted === page) {
    if (page) forceLanguage(page);
    return;
  }
  try {
    const url = new URL(location.href);
    url.pathname = url.pathname.replace(/\.html$/, `.${wanted}.html`);
    url.search = "";
    url.hash = "";
    const res = await fetch(url.toString());
    if (!res.ok) throw new Error(`HTTP ${res.status}`);
    const doc = new DOMParser().parseFromString(await res.text(), "text/html");
    if (!doc.body || doc.documentElement.lang !== wanted) throw new Error("unexpected page");
    document.body.replaceWith(document.importNode(doc.body, true));
    document.documentElement.lang = wanted;
    if (doc.title) document.title = doc.title;
    forceLanguage(wanted);
  } catch {
    forceLanguage(page);
  }
}
