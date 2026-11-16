/**
 * Own signature of the link card (settings): pasted from Outlook as formatted text (HTML) or typed as plain text.
 * HTML is cleaned with the browser's parser before it is shown in the pane, stored or inserted into a mail: only
 * harmless formatting stays – no scripts, no event handlers, no styles that load anything, links only to web, mail
 * and phone addresses, images only from https (Outlook's local or embedded images cannot travel with the card).
 */

/** Elements kept (their content too). Anything else is unwrapped – its text stays, the element goes. */
const KEEP = new Set([
  "A", "B", "BR", "DIV", "EM", "FONT", "HR", "I", "IMG", "LI", "OL", "P", "SMALL", "SPAN", "STRONG", "SUB", "SUP",
  "TABLE", "TBODY", "TD", "TH", "THEAD", "TR", "U", "UL",
]);
/** Elements dropped with everything inside. */
const DROP = new Set([
  "SCRIPT", "STYLE", "TEMPLATE", "IFRAME", "OBJECT", "EMBED", "FORM", "INPUT", "BUTTON", "SELECT", "TEXTAREA", "SVG",
  "MATH", "LINK", "META", "TITLE", "HEAD", "NOSCRIPT", "VIDEO", "AUDIO", "CANVAS", "XML", "O:P",
]);
/** Attributes kept per element (besides "style" on every element). */
const ATTRIBUTES: Record<string, string[]> = {
  A: ["href", "title"],
  IMG: ["src", "alt", "width", "height"],
  TABLE: ["width", "cellpadding", "cellspacing", "border", "align"],
  TD: ["width", "colspan", "rowspan", "valign", "align"],
  TH: ["width", "colspan", "rowspan", "valign", "align"],
  FONT: ["color", "face", "size"],
  P: ["align"],
  DIV: ["align"],
};
/** CSS properties kept in style attributes; values that could load or run something are dropped. */
const STYLE_PROPERTIES = new Set([
  "color", "background-color", "font-family", "font-size", "font-weight", "font-style", "text-decoration",
  "line-height", "text-align", "vertical-align", "white-space", "width", "height", "max-width",
  "margin", "margin-top", "margin-right", "margin-bottom", "margin-left",
  "padding", "padding-top", "padding-right", "padding-bottom", "padding-left",
  "border", "border-top", "border-right", "border-bottom", "border-left", "border-collapse", "border-spacing",
]);
const UNSAFE_VALUE = /url\s*\(|expression|javascript:|[<>\\]|@import/i;

function cleanStyle(style: string): string {
  return style
    .split(";")
    .map((declaration) => {
      const colon = declaration.indexOf(":");
      if (colon < 0) return "";
      const property = declaration.slice(0, colon).trim().toLowerCase();
      const value = declaration.slice(colon + 1).trim();
      return STYLE_PROPERTIES.has(property) && value && !UNSAFE_VALUE.test(value) ? `${property}:${value}` : "";
    })
    .filter(Boolean)
    .join(";");
}

const safeHref = (href: string) => /^(https?:|mailto:|tel:)/i.test(href.trim());
const safeSrc = (src: string) => /^https:\/\//i.test(src.trim());

function cleanNode(node: Node, doc: Document): Node[] {
  if (node.nodeType === Node.TEXT_NODE) return [doc.createTextNode(node.textContent ?? "")];
  if (node.nodeType !== Node.ELEMENT_NODE) return []; // comments (Outlook's conditional comments), processing instructions
  const element = node as Element;
  const tag = element.tagName.toUpperCase();
  if (DROP.has(tag) || tag.includes(":")) return [];
  const children = Array.from(element.childNodes).flatMap((child) => cleanNode(child, doc));
  if (!KEEP.has(tag)) return children;
  if (tag === "IMG" && !safeSrc(element.getAttribute("src") ?? "")) return [];

  const copy = doc.createElement(tag.toLowerCase());
  for (const name of ATTRIBUTES[tag] ?? []) {
    const value = element.getAttribute(name);
    if (value === null) continue;
    if (name === "href" && !safeHref(value)) continue;
    copy.setAttribute(name, value);
  }
  const style = cleanStyle(element.getAttribute("style") ?? "");
  if (style) copy.setAttribute("style", style);
  if (tag === "A" && copy.hasAttribute("href")) copy.setAttribute("target", "_blank");
  for (const child of children) copy.appendChild(child);
  return [copy];
}

/** Signature HTML with only harmless formatting left (see the module comment). Empty if nothing visible remains. */
export function sanitizeSignature(html: string): string {
  const source = new DOMParser().parseFromString(html, "text/html");
  const target = document.implementation.createHTMLDocument("");
  const container = target.createElement("div");
  for (const node of Array.from(source.body.childNodes).flatMap((child) => cleanNode(child, target))) container.appendChild(node);
  const cleaned = container.innerHTML.trim();
  return signatureText(cleaned) || container.querySelector("img") ? cleaned : "";
}

/** Text version of signature HTML for plain-text mails: one line per paragraph, row or line break. */
export function signatureText(html: string): string {
  const doc = new DOMParser().parseFromString(html, "text/html");
  const lines: string[] = [];
  let line = "";
  const flush = () => {
    lines.push(line.replace(/\s+/g, " ").trim());
    line = "";
  };
  const walk = (node: Node) => {
    if (node.nodeType === Node.TEXT_NODE) {
      line += node.textContent ?? "";
      return;
    }
    if (node.nodeType !== Node.ELEMENT_NODE) return;
    const tag = (node as Element).tagName.toUpperCase();
    if (tag === "BR") return flush();
    const block = ["P", "DIV", "TR", "LI", "TABLE", "UL", "OL", "HR"].includes(tag);
    if (block && line.trim()) flush();
    node.childNodes.forEach(walk);
    if (tag === "TD" || tag === "TH") line += " ";
    if (block && line.trim()) flush();
  };
  doc.body.childNodes.forEach(walk);
  if (line.trim()) flush();
  return lines
    .filter((l, i, all) => l || (i > 0 && all[i - 1]))
    .join("\n")
    .trim();
}
