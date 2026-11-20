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
let listCount = 0;

/**
 * Turns `input` into a combobox with suggestions of `kind`. `usable` leaves out entries that could not be used (an SMS
 * number of a country that is not allowed). A picked entry is filled in and announced like typed text ("input",
 * "change"), so the field checks and saves it as usual.
 */
export function attachContactSuggestions(input: HTMLInputElement, kind: "mail" | "sms", usable: (value: string) => boolean): void {
  const list = document.createElement("ul");
  list.id = `contact-suggestions-${++listCount}`;
  list.className = "suggestions";
  list.setAttribute("role", "listbox");
  list.setAttribute("aria-label", t("contacts.listLabel"));
  list.hidden = true;
  input.insertAdjacentElement("afterend", list);
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
    input.value = entry.value;
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
    const query = input.value.trim();
    if (unavailable || query.length < 2 || document.activeElement !== input) return close();
    const mine = ++asked;
    try {
      const { contacts } = await api<{ contacts: ContactSuggestion[] | null }>(
        `/api/contacts?kind=${kind}&q=${encodeURIComponent(query)}`, {}, await getToken());
      if (mine !== asked || document.activeElement !== input) return;
      if (contacts === null) {
        unavailable = true;
        return close();
      }
      render(contacts.filter((entry) => entry.value !== query && usable(entry.value)));
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
      if (event.key === "ArrowDown" && input.value.trim().length >= 2) {
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
