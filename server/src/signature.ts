/**
 * Signature of the password mails, so recipients see who the password comes from: the sender's own one from the
 * settings (text or formatted), else MAIL_FOOTER filled from the sender's Microsoft 365 profile.
 *
 * A formatted signature is pasted from Outlook in the add-in and stored as a tree of nodes (element name, allowed
 * attributes, children, text) – never as markup. The server checks the tree against the same allow-list as the add-in
 * (addin/src/taskpane/signature.ts) and builds the HTML itself, escaping every text and attribute: no script, event
 * handler, style that loads anything, links only to web, mail and phone addresses, pictures only from https.
 */
import { config } from "./config.js";
import type { Profile } from "./graph.js";

/** A text, or an element with attributes and children. */
export type SignatureNode = string | { tag: string; attrs?: Record<string, string>; children?: SignatureNode[] };

const TAGS = new Set([
  "a", "b", "br", "div", "em", "font", "hr", "i", "img", "li", "ol", "p", "small", "span", "strong", "sub", "sup",
  "table", "tbody", "td", "th", "thead", "tr", "u", "ul",
]);
const VOID = new Set(["br", "hr", "img"]);
const ATTRIBUTES: Record<string, string[]> = {
  a: ["href", "title"],
  img: ["src", "alt", "width", "height"],
  table: ["width", "cellpadding", "cellspacing", "border", "align"],
  td: ["width", "colspan", "rowspan", "valign", "align"],
  th: ["width", "colspan", "rowspan", "valign", "align"],
  font: ["color", "face", "size"],
  p: ["align"],
  div: ["align"],
};
const STYLE_PROPERTIES = new Set([
  "color", "background-color", "font-family", "font-size", "font-weight", "font-style", "text-decoration",
  "line-height", "text-align", "vertical-align", "white-space", "width", "height", "max-width",
  "margin", "margin-top", "margin-right", "margin-bottom", "margin-left",
  "padding", "padding-top", "padding-right", "padding-bottom", "padding-left",
  "border", "border-top", "border-right", "border-bottom", "border-left", "border-collapse", "border-spacing",
]);
const UNSAFE_VALUE = /url\s*\(|expression|javascript:|[<>\\"]|@import/i;
const MAX_DEPTH = 24;

export function cleanStyle(style: string): string {
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

/**
 * The tree as sent by the add-in, reduced to what the allow-list permits (unknown elements give way to their
 * children, unknown attributes and unsafe values are dropped). Anything that is not a node at all is dropped.
 */
export function cleanSignatureNodes(input: unknown, depth = 0): SignatureNode[] {
  if (!Array.isArray(input) || depth > MAX_DEPTH) return [];
  return input.flatMap((node): SignatureNode[] => {
    if (typeof node === "string") return [node];
    if (!node || typeof node !== "object") return [];
    const { tag, attrs, children } = node as { tag?: unknown; attrs?: unknown; children?: unknown };
    const name = typeof tag === "string" ? tag.toLowerCase() : "";
    const kids = cleanSignatureNodes(children, depth + 1);
    if (!TAGS.has(name)) return kids;
    const source = attrs && typeof attrs === "object" ? (attrs as Record<string, unknown>) : {};
    const kept: Record<string, string> = {};
    for (const attribute of ATTRIBUTES[name] ?? []) {
      const value = source[attribute];
      if (typeof value !== "string") continue;
      if (attribute === "href" && !safeHref(value)) continue;
      if (attribute === "src" && !safeSrc(value)) continue;
      kept[attribute] = value;
    }
    if (name === "img" && !kept.src) return [];
    const style = typeof source.style === "string" ? cleanStyle(source.style) : "";
    if (style) kept.style = style;
    return [{ tag: name, ...(Object.keys(kept).length ? { attrs: kept } : {}), ...(kids.length && !VOID.has(name) ? { children: kids } : {}) }];
  });
}

const escapeHtml = (s: string) =>
  s.replace(/[&<>"']/g, (c) => ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;", "'": "&#39;" })[c]!);

/** HTML of a cleaned tree – every text and attribute escaped. */
export function renderSignatureNodes(nodes: SignatureNode[]): string {
  return cleanSignatureNodes(nodes)
    .map((node) => {
      if (typeof node === "string") return escapeHtml(node);
      const attrs = Object.entries(node.attrs ?? {})
        .map(([name, value]) => ` ${name}="${escapeHtml(value)}"`)
        .join("");
      const target = node.tag === "a" && node.attrs?.href ? ' target="_blank" rel="noopener"' : "";
      if (VOID.has(node.tag)) return `<${node.tag}${attrs}>`;
      return `<${node.tag}${attrs}${target}>${renderSignatureNodes(node.children ?? [])}</${node.tag}>`;
    })
    .join("");
}

/** Whether a tree shows anything (text or a picture). */
export function signatureHasContent(nodes: SignatureNode[]): boolean {
  return nodes.some((node) => (typeof node === "string" ? node.trim() !== "" : node.tag === "img" || signatureHasContent(node.children ?? [])));
}

const SEPARATOR = " · ";

/**
 * A text template with profile placeholders ({displayName}, {jobTitle}, {phone} …): "\n" starts a new line, a
 * " · " segment whose placeholders are all empty is dropped with its fixed text ("Tel. {phone}" without a number),
 * empty lines are dropped. Same rules as MAIL_FOOTER always had.
 */
export function fillTemplate(template: string, profile: Partial<Profile>): string {
  const values = profile as Record<string, string | undefined>;
  return template
    .split(/\r?\n|\\n/)
    .map((line) =>
      line
        .split(SEPARATOR)
        .map((segment) => {
          let placeholders = 0;
          let filled = 0;
          const text = segment.replace(/\{(\w+)\}/g, (_match, key: string) => {
            placeholders++;
            const value = (values[key] ?? "").trim();
            if (value) filled++;
            return value;
          });
          return placeholders > 0 && filled === 0 ? "" : text.trim();
        })
        .filter(Boolean)
        .join(SEPARATOR),
    )
    .filter(Boolean)
    .join("\n");
}

/** Own signature of the settings (server/src/settings.ts). */
export type Signature = { format: "text"; content: string } | { format: "html"; nodes: SignatureNode[] };

/**
 * HTML of the signature for a password mail: the own formatted one, the own text one or MAIL_FOOTER – texts filled
 * from `profile` (empty values drop their segment). "" if there is none.
 */
export function signatureHtml(own: Signature | null | undefined, profile: Partial<Profile>): string {
  if (own?.format === "html") return renderSignatureNodes(own.nodes);
  const text = fillTemplate(own?.format === "text" ? own.content : config.mailFooter, profile);
  return text ? escapeHtml(text).split("\n").join("<br>") : "";
}
