/**
 * Loading placeholders: grey bars in the shape of what is coming, shimmering like Outlook's while it loads the inbox
 * ("Verlauf", "Zustellnachweise", the links of a mail being read, the settings; the start of the pane has the same in
 * taskpane.html, `#start-view`). The page stands at once and the content takes their place without a jump. Screen
 * readers hear "Wird geladen …" (role status, aria-busy); with reduced motion the bars do not shimmer (taskpane.css).
 */
import { t } from "../shared/i18n";

const box = (className: string) => {
  const element = document.createElement("div");
  element.className = className;
  return element;
};

/** The frame every placeholder shares: busy for assistive technology, the words for screen readers. */
function frame(): HTMLDivElement {
  const wrap = box("skeleton");
  wrap.setAttribute("role", "status");
  wrap.setAttribute("aria-busy", "true");
  const label = document.createElement("span");
  label.className = "visually-hidden";
  label.textContent = t("loading");
  wrap.append(label);
  return wrap;
}

/** `count` entries like a row of "Verlauf": the kind symbol, then the title and two shorter lines. */
export function skeletonRows(count = 3): HTMLDivElement {
  const rows = frame();
  for (let i = 0; i < count; i++) {
    const row = box("skeleton-row");
    const lines = box("skeleton-lines");
    // Slightly different lengths, so it reads as entries and not as a pattern.
    for (const width of [[70, 45, 85], [60, 40, 75], [80, 50, 65]][i % 3]) {
      const line = box("skeleton-line");
      line.style.width = `${width}%`;
      lines.append(line);
    }
    row.append(box("skeleton-icon"), lines);
    rows.append(row);
  }
  return rows;
}

/** `count` fields like a form (the settings): a label and its box. */
export function skeletonForm(count = 3): HTMLDivElement {
  const form = frame();
  for (let i = 0; i < count; i++) {
    const field = box("skeleton-field");
    const label = box("skeleton-line");
    label.style.width = `${[45, 35, 55][i % 3]}%`;
    field.append(label, box("skeleton-box"));
    form.append(field);
  }
  return form;
}
