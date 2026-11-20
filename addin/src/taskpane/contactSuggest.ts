/**
 * Suggestions from the user's own Outlook contacts under "Kennwort an" (another address, SMS number) – like Outlook's
 * recipient field: from two characters a list opens below the field (ARIA combobox), arrow keys and Enter or a click
 * pick an entry. The server matches (`GET /api/contacts`, Graph with the optional delegated Contacts.Read); when it
 * answers `contacts: null` (not consented) the fields stay plain inputs for the rest of the session.
 */
import { api, getToken } from "../shared/office";
import { t } from "../shared/i18n";

export interface ContactSuggestion {
  name: string;
  value: string;
  type?: "mobile" | "business" | "home";
}

let unavailable = false;
const whenUnavailable: Array<() => void> = [];

/** The server said the contacts cannot be read (no Contacts.Read): suggestions and the address book stop. */
export function contactsUnavailable(): boolean {
  return unavailable;
}

/** Marks the contacts as not readable and tells everyone who asked (the address book buttons hide). */
export function markContactsUnavailable(): void {
  if (unavailable) return;
  unavailable = true;
  for (const listener of whenUnavailable) listener();
}

/** Runs `listener` once the contacts turn out not to be readable. */
export function onContactsUnavailable(listener: () => void): void {
  whenUnavailable.push(listener);
}

/**
 * Turns `input` into a combobox with suggestions of `kind`. `usable` leaves out entries that could not be used (an SMS
 * number of a country that is not allowed). A picked entry is filled in and announced like typed text ("input",
 * "change"), so the field checks and saves it as usual. `kind` may be a function for a field that switches between
 * address and number ("Erneut senden an"); a field for several addresses (`data-multiple`) searches the part after the
 * last comma and adds the picked address to the list.
 */
export function attachContactSuggestions(
  input: HTMLInputElement,
  kindOf: "mail" | "sms" | (() => "mail" | "sms"),
  usable: (value: string, kind: "mail" | "sms") => boolean,
): void {
  const currentKind = () => (typeof kindOf === "function" ? kindOf() : kindOf);
  const multiple = () => currentKind() === "mail" && input.dataset.multiple !== undefined;
  /** What is searched: the whole field, or the address being typed in a list. */
  const typed = () => (multiple() ? (input.value.split(/[,;]/).pop() ?? "") : input.value).trim();
  const list = document.createElement("ul");
  list.id = `${input.id}-suggestions`;
  list.className = "suggestions";
  list.setAttribute("role", "listbox");
  list.setAttribute("aria-label", t("contacts.listLabel"));
  list.hidden = true;
  // After the field – or after its group with the address book button, so the list opens below both.
  (input.closest(".input-group") ?? input).insertAdjacentElement("afterend", list);
  input.setAttribute("role", "combobox");
  input.setAttribute("aria-autocomplete", "list");
  input.setAttribute("aria-expanded", "false");
  input.setAttribute("aria-controls", list.id);

  let shown: ContactSuggestion[] = [];
  let active = -1;
  let timer: ReturnType<typeof setTimeout> | undefined;
  let asked = 0;

  const close = () => {
    list.hidden = true;
    list.replaceChildren();
    shown = [];
    active = -1;
    input.setAttribute("aria-expanded", "false");
    input.removeAttribute("aria-activedescendant");
  };

  const highlight = (index: number) => {
    active = index;
    Array.from(list.children).forEach((item, i) => item.setAttribute("aria-selected", String(i === index)));
    if (index >= 0) {
      input.setAttribute("aria-activedescendant", `${list.id}-${index}`);
      (list.children[index] as HTMLElement).scrollIntoView({ block: "nearest" });
    } else input.removeAttribute("aria-activedescendant");
  };

  const choose = (index: number) => {
    const entry = shown[index];
    if (!entry) return;
    input.value = multiple() ? addToList(input.value, entry.value) : entry.value;
    close();
    input.dispatchEvent(new Event("input", { bubbles: true }));
    input.dispatchEvent(new Event("change", { bubbles: true }));
  };

  const render = (entries: ContactSuggestion[]) => {
    shown = entries;
    list.replaceChildren(...entries.map((entry, i) => {
      const item = document.createElement("li");
      item.id = `${list.id}-${i}`;
      item.className = "suggestion";
      item.setAttribute("role", "option");
      item.setAttribute("aria-selected", "false");
      const name = document.createElement("span");
      name.className = "suggestion-name";
      name.textContent = entry.name;
      const value = document.createElement("span");
      value.className = "suggestion-value";
      value.textContent = entry.type ? `${entry.value} · ${t(`contacts.${entry.type}`)}` : entry.value;
      item.append(name, value);
      // mousedown keeps the focus in the field – a blur would first check the half-typed text.
      item.addEventListener("mousedown", (event) => event.preventDefault());
      item.addEventListener("click", () => choose(i));
      return item;
    }));
    list.hidden = entries.length === 0;
    input.setAttribute("aria-expanded", String(entries.length > 0));
    active = -1;
    input.removeAttribute("aria-activedescendant");
  };

  const search = async () => {
    const query = typed();
    const kind = currentKind();
    if (unavailable || input.hidden || query.length < 2 || document.activeElement !== input) return close();
    const mine = ++asked;
    try {
      const { contacts } = await api<{ contacts: ContactSuggestion[] | null }>(
        `/api/contacts?kind=${kind}&q=${encodeURIComponent(query)}`, {}, await getToken());
      if (mine !== asked || document.activeElement !== input) return;
      if (contacts === null) {
        markContactsUnavailable();
        return close();
      }
      render(contacts.filter((entry) => entry.value !== query && usable(entry.value, kind)));
    } catch {
      if (mine === asked) close(); // no suggestions – typing still works
    }
  };

  input.addEventListener("input", (event) => {
    if (!event.isTrusted) return; // filled in by choose()
    clearTimeout(timer);
    timer = setTimeout(() => void search(), 200);
  });
  input.addEventListener("keydown", (event) => {
    if (list.hidden) {
      if (event.key === "ArrowDown" && typed().length >= 2) {
        event.preventDefault();
        void search();
      }
      return;
    }
    if (event.key === "ArrowDown" || event.key === "ArrowUp") {
      event.preventDefault();
      const step = event.key === "ArrowDown" ? 1 : -1;
      const next = active + step;
      highlight(next >= shown.length ? 0 : next < 0 ? shown.length - 1 : next);
    } else if (event.key === "Enter" && active >= 0) {
      event.preventDefault();
      choose(active);
    } else if (event.key === "Escape") {
      event.preventDefault();
      event.stopPropagation(); // only the list closes, not the page
      close();
    } else if (event.key === "Tab") close();
  });
  input.addEventListener("blur", () => {
    clearTimeout(timer);
    close();
  });
}

/** `list` ("a@x.de, b@y.de" or a half-typed last part) with `address` added – the half-typed part is replaced. */
export function addToList(list: string, address: string): string {
  const parts = list.split(/[,;]/).map((part) => part.trim());
  const last = parts.pop() ?? "";
  // A complete last address stays; a fragment ("mai") gives way to the picked one.
  if (/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(last)) parts.push(last);
  const kept = parts.filter(Boolean).filter((part) => part.toLowerCase() !== address.toLowerCase());
  return [...kept, address].join(", ");
}
