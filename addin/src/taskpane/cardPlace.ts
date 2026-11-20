/**
 * Where a further link card goes: right after the last CloudShare card already in the mail – wherever the cursor is
 * (the user may have clicked into a card, and a card inserted there would break its layout). Office.js cannot insert
 * after an element, so the pane reads the body, puts the card in with `withCardAfterLast` and writes the body back.
 */

/** The opening tag of a card's frame (mailCard.ts `frame`: width 560) – kept by every Outlook. */
const FRAME = /\b(?:width\s*=\s*["']?560\b|max-width\s*:\s*560px)/i;

/**
 * The position right after the frame of the card holding the last of `markers` (share tokens or URLs of the links
 * inserted into this mail) in `html`, or null when there is none (no card yet, or a body Outlook changed beyond
 * recognition – then the card goes to the cursor as before).
 */
export function afterLastCard(html: string, markers: string[]): number | null {
  let last = -1;
  for (const marker of markers) if (marker) last = Math.max(last, html.lastIndexOf(marker));
  if (last < 0) return null;

  const tags = /<(\/?)table\b[^>]*>/gi;
  const open: { tag: string }[] = [];
  let frame: { tag: string } | null = null;
  for (let match = tags.exec(html); match; match = tags.exec(html)) {
    if (frame === null && match.index > last) {
      // The tables around the marker, outermost first: the card is the outermost one with the card's width – a body
      // laid out in a table of its own (some signatures, templates) stays around it.
      frame = open.find((entry) => FRAME.test(entry.tag)) ?? null;
      if (frame === null) return null;
    }
    if (match[1] === "") open.push({ tag: match[0] });
    else {
      const closed = open.pop();
      if (closed && closed === frame) return match.index + match[0].length;
    }
  }
  return null;
}

/** `html` with `card` right after the last CloudShare card, or null (see afterLastCard). */
export function withCardAfterLast(html: string, markers: string[], card: string): string | null {
  const at = afterLastCard(html, markers);
  return at === null ? null : html.slice(0, at) + card + html.slice(at);
}
