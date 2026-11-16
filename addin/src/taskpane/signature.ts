/**
 * Own signature of the password mails (settings): pasted from Outlook as formatted text. The browser inserts the
 * paste into an editable box; the allow-list below cleans it in place right away – only harmless formatting stays: no
 * scripts, no event handlers, no styles that load anything, links only to web, mail and phone addresses, pictures
 * only from https (Outlook's local or embedded pictures cannot travel with a mail). The result is stored as a tree of
 * nodes, never as markup; the server checks it against the same allow-list (server/src/signature.ts) and builds the
 * HTML of the password mail itself.
 */

/** A text, or an element with attributes and children – what /api/settings stores for a formatted signature. */
export type SignatureNode = string | { tag: string; attrs?: Record<string, string>; children?: SignatureNode[] };

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

/**
 * Cleans the content of an element in place (the signature editor right after a paste and before saving): the
 * allow-list applied to the nodes the browser inserted – no markup string is parsed. The caret goes to the end.
 */
export function cleanInPlace(root: HTMLElement): void {
  const doc = root.ownerDocument;
  const nodes = Array.from(root.childNodes).flatMap((child) => cleanNode(child, doc));
  root.replaceChildren(...nodes);
  const selection = doc.getSelection();
  if (selection) {
    selection.selectAllChildren(root);
    selection.collapseToEnd();
  }
}


/** The (cleaned) content of an element as a node tree for the settings. */
export function signatureNodes(root: Element): SignatureNode[] {
  return Array.from(root.childNodes).flatMap((node): SignatureNode[] => {
    if (node.nodeType === Node.TEXT_NODE) return node.textContent ? [node.textContent] : [];
    if (node.nodeType !== Node.ELEMENT_NODE) return [];
    const element = node as Element;
    const attrs: Record<string, string> = {};
    for (const attribute of Array.from(element.attributes)) {
      if (attribute.name !== "target") attrs[attribute.name] = attribute.value;
    }
    const children = signatureNodes(element);
    return [{ tag: element.tagName.toLowerCase(), ...(Object.keys(attrs).length ? { attrs } : {}), ...(children.length ? { children } : {}) }];
  });
}

/**
 * DOM nodes for the editor from a stored tree – built element by element in an inert document (nothing loads or
 * runs there), then copied into `doc` by the allow-list.
 */
export function signatureFromNodes(nodes: SignatureNode[] | undefined, doc: Document): Node[] {
  const inert = doc.implementation.createHTMLDocument("");
  const build = (list: SignatureNode[]): Node[] =>
    list.flatMap((node): Node[] => {
      if (typeof node === "string") return [inert.createTextNode(node)];
      if (!node || typeof node.tag !== "string" || !/^[a-z][a-z0-9]*$/.test(node.tag)) return [];
      const element = inert.createElement(node.tag);
      for (const [name, value] of Object.entries(node.attrs ?? {})) {
        if (/^[a-z][a-z-]*$/.test(name) && typeof value === "string") element.setAttribute(name, value);
      }
      for (const child of build(node.children ?? [])) element.appendChild(child);
      return [element];
    });
  return build(Array.isArray(nodes) ? nodes : []).flatMap((node) => cleanNode(node, doc));
}

/** Whether the editor shows anything (text or a picture). */
export const signatureHasContent = (root: Element) => Boolean(root.textContent?.trim()) || root.querySelector("img") !== null;
