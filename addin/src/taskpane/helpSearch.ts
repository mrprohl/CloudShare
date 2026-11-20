/**
 * Search on the "Hilfe" page: the magnifier in its head folds out a field. While typing, only the sections that
 * contain every word stay (`.search-miss` hides the others – never touching `hidden`, which the pane uses for sections
 * a server switch leaves out), they open, and the hits are marked. Escape or an empty field brings the guide back as it
 * was (the sections open before stay open); closing the page clears the search too (`reset`).
 */

const $ = <T extends HTMLElement>(id: string) => document.getElementById(id) as T;

/** Lower case without accents – "Kennwörter" is found with "kennworter" too. */
const plain = (text: string) => text.toLowerCase().normalize("NFD").replace(/[̀-ͯ]/g, "");

export function initHelpSearch() {
  const toggle = $<HTMLButtonElement>("help-search-toggle");
  const row = $<HTMLDivElement>("help-search-row");
  const input = $<HTMLInputElement>("help-search");
  const none = $<HTMLParagraphElement>("help-search-none");
  const panel = $<HTMLElement>("panel-help");
  const sections = Array.from(panel.querySelectorAll<HTMLDetailsElement>("details.help-section"));
  const intro = panel.querySelector<HTMLElement>(".help-intro");
  /** Which sections were open before the search – restored when it ends. */
  let openBefore: boolean[] | null = null;

  /** Removes the marks of the last search. */
  const unmark = () => {
    for (const mark of Array.from(panel.querySelectorAll("mark.help-hit"))) {
      mark.replaceWith(document.createTextNode(mark.textContent ?? ""));
    }
    for (const section of sections) section.normalize();
  };

  /** Marks every word in the text of `root` (not in the summary's arrow or the icons). */
  const mark = (root: HTMLElement, words: string[]) => {
    const walker = document.createTreeWalker(root, NodeFilter.SHOW_TEXT);
    const nodes: Text[] = [];
    for (let node = walker.nextNode(); node; node = walker.nextNode()) nodes.push(node as Text);
    for (const node of nodes) {
      const text = node.data;
      const lower = plain(text); // same length as the text for umlauts and accents (the marks are dropped again)
      if (lower.length !== text.length) continue; // rare characters that change length: leave unmarked
      const hits: Array<[number, number]> = [];
      for (const word of words) {
        for (let at = lower.indexOf(word); at >= 0; at = lower.indexOf(word, at + word.length)) hits.push([at, at + word.length]);
      }
      if (!hits.length) continue;
      hits.sort((a, b) => a[0] - b[0]);
      const parts: Node[] = [];
      let pos = 0;
      for (const [start, end] of hits) {
        if (start < pos) continue; // overlapping words: the first one wins
        if (start > pos) parts.push(document.createTextNode(text.slice(pos, start)));
        const m = document.createElement("mark");
        m.className = "help-hit";
        m.textContent = text.slice(start, end);
        parts.push(m);
        pos = end;
      }
      if (pos < text.length) parts.push(document.createTextNode(text.slice(pos)));
      node.replaceWith(...parts);
    }
  };

  const search = () => {
    unmark();
    const words = plain(input.value).split(/\s+/).filter(Boolean);
    if (!words.length) {
      for (const section of sections) section.classList.remove("search-miss");
      intro?.classList.remove("search-miss");
      if (openBefore) sections.forEach((section, i) => (section.open = openBefore![i]));
      openBefore = null;
      none.hidden = true;
      return;
    }
    openBefore ??= sections.map((section) => section.open);
    intro?.classList.add("search-miss");
    let found = 0;
    for (const section of sections) {
      const hit = !section.hidden && words.every((word) => plain(section.textContent ?? "").includes(word));
      section.classList.toggle("search-miss", !hit);
      section.open = hit;
      if (hit) {
        found++;
        mark(section, words);
      }
    }
    none.hidden = found > 0;
  };

  const open = (shown: boolean) => {
    row.hidden = !shown;
    toggle.setAttribute("aria-expanded", String(shown));
    if (shown) input.focus();
    else {
      input.value = "";
      search();
    }
  };

  toggle.addEventListener("click", () => open(row.hidden));
  input.addEventListener("input", search);
  input.addEventListener("keydown", (event) => {
    if (event.key !== "Escape") return;
    // Escape closes the search only – not the page (the pane's Escape handler skips handled events).
    event.preventDefault();
    open(false);
    toggle.focus();
  });

  /** The page is closed: no search left over for the next opening. */
  return { reset: () => open(false) };
}
