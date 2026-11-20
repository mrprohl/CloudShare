/**
 * The magnifier in the head of a page ("Hilfe", "Verlauf", "Zustellnachweise"): a click folds out the search field
 * (focus in it), the list filters while typing (`onChange`). Escape in the field, a second click on the magnifier or
 * leaving the page (`reset`) folds it in and clears it, so the whole list is back. Escape closes only the search, not
 * the page – the pane's Escape handler skips handled events.
 */
interface Options {
  toggle: HTMLButtonElement;
  /** The field, or a row around it with more (e.g. "Nichts gefunden"). */
  row: HTMLElement;
  input: HTMLInputElement;
  /** The search changed (typed, or cleared on closing). */
  onChange: () => void;
}

export function initSearchToggle({ toggle, row, input, onChange }: Options) {
  const open = (shown: boolean) => {
    row.hidden = !shown;
    toggle.setAttribute("aria-expanded", String(shown));
    if (shown) {
      input.focus();
      return;
    }
    if (input.value) {
      input.value = "";
      onChange();
    }
  };

  toggle.addEventListener("click", () => open(row.hidden !== false)); // hidden may also be "until-found"
  input.addEventListener("input", onChange);
  input.addEventListener("keydown", (event) => {
    if (event.key !== "Escape") return;
    event.preventDefault();
    open(false);
    toggle.focus();
  });

  /** The page is left: no search left over for the next opening. */
  return { reset: () => open(false) };
}
