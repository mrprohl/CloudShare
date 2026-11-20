/**
 * "Adressbuch": the button next to "Andere E-Mail-Adresse" / "Rufnummer (SMS)" opens a page with the user's own
 * Outlook contacts A to Z (GET /api/contacts?all=1, optional Contacts.Read), grouped by first letter and filtered while
 * typing. A click takes the address or number into the field – checked and saved like a typed one – and leads back.
 * Buttons in the static page are found at start, those built later (the resend panel) come in by `register`. Without
 * the permission the buttons disappear (onContactsUnavailable).
 */
import { t } from "../shared/i18n";
import { api } from "../shared/office";
import { contactsUnavailable, markContactsUnavailable, onContactsUnavailable, type ContactSuggestion } from "./contactSuggest";

interface Options {
  token: () => Promise<string>;
  setStatus: (text: string, kind?: "info" | "success" | "error") => void;
  errorMessage: (err: unknown) => string;
  /** Shows the page in place of the view, or the view again (null). */
  show: (open: boolean) => void;
  /** Entries that could not be used are left out (an SMS number of a country that is not allowed). */
  usable: (kind: "mail" | "sms", value: string) => boolean;
}

const $ = <T extends HTMLElement>(id: string) => document.getElementById(id) as T;

/** Lower case without accents, for the search. */
const plain = (text: string) => text.toLowerCase().normalize("NFD").replace(/[̀-ͯ]/g, "");

export function initAddressBook(options: Options) {
  const search = $<HTMLInputElement>("contacts-search");
  const list = $<HTMLDivElement>("contacts-list");
  const empty = $<HTMLParagraphElement>("contacts-empty");
  const close = $<HTMLButtonElement>("contacts-close");
  const buttons = Array.from(document.querySelectorAll<HTMLButtonElement>(".book-open"));

  let entries: ContactSuggestion[] = [];
  let target: HTMLInputElement | null = null;
  let kind: "mail" | "sms" = "mail";

  const hideButtons = () => {
    for (const button of buttons) button.hidden = true;
  };
  onContactsUnavailable(hideButtons);
  if (contactsUnavailable()) hideButtons();

  const back = () => {
    options.show(false);
    target?.focus();
  };

  const pick = (entry: ContactSuggestion) => {
    if (!target) return;
    target.value = entry.value;
    options.show(false);
    target.dispatchEvent(new Event("input", { bubbles: true }));
    target.dispatchEvent(new Event("change", { bubbles: true }));
    target.focus();
  };

  /** The entries matching the search, under their first letter. */
  const render = () => {
    const query = plain(search.value.trim());
    const shown = query
      ? entries.filter((e) => plain(e.name).includes(query) || plain(e.value).includes(query)
        || (kind === "sms" && /\d/.test(query) && e.value.replace(/\D/g, "").includes(query.replace(/\D/g, ""))))
      : entries;
    const groups = new Map<string, ContactSuggestion[]>();
    for (const entry of shown) {
      const letter = plain(entry.name).charAt(0).toUpperCase();
      const key = /[A-Z]/.test(letter) ? letter : "#";
      groups.set(key, [...(groups.get(key) ?? []), entry]);
    }
    list.replaceChildren(
      ...[...groups].map(([letter, group]) => {
        const section = document.createElement("div");
        section.className = "contacts-group";
        const heading = document.createElement("h3");
        heading.className = "contacts-letter";
        heading.textContent = letter;
        const items = document.createElement("ul");
        items.className = "contacts-items";
        for (const entry of group) {
          const item = document.createElement("li");
          const button = document.createElement("button");
          button.type = "button";
          button.className = "contact-entry";
          const name = document.createElement("span");
          name.className = "contact-name";
          name.textContent = entry.name;
          const value = document.createElement("span");
          value.className = "contact-value";
          value.textContent = entry.type ? `${entry.value} · ${t(`contacts.${entry.type}`)}` : entry.value;
          button.append(name, value);
          button.addEventListener("click", () => pick(entry));
          item.append(button);
          items.append(item);
        }
        section.append(heading, items);
        return section;
      }),
    );
    empty.hidden = shown.length > 0;
    empty.textContent = shown.length ? "" : t(query ? "contacts.noMatch" : "contacts.none");
  };

  /** Opens the page for `input` (its kind from the field: tel = SMS). */
  async function open(input: HTMLInputElement) {
    target = input;
    kind = input.type === "tel" ? "sms" : "mail";
    search.value = "";
    entries = [];
    list.replaceChildren();
    empty.hidden = false;
    empty.textContent = t("contacts.loading");
    options.show(true);
    search.focus();
    try {
      const { contacts } = await api<{ contacts: ContactSuggestion[] | null }>(`/api/contacts?kind=${kind}&all=1`, {}, await options.token());
      if (contacts === null) {
        markContactsUnavailable();
        back();
        options.setStatus(t("contacts.unavailable"), "info");
        return;
      }
      entries = contacts.filter((entry) => options.usable(kind, entry.value));
      render();
    } catch (err) {
      empty.textContent = t("contacts.failed", { error: options.errorMessage(err) });
    }
  }

  /** Wires a button to the field it fills (`data-target`, looked up on the click – the field may be built later). */
  const register = (button: HTMLButtonElement) => {
    // Rows of "Verlauf" are built anew: buttons no longer in the page are dropped.
    for (let i = buttons.length - 1; i >= 0; i--) if (!buttons[i].isConnected) buttons.splice(i, 1);
    if (!buttons.includes(button)) buttons.push(button);
    if (contactsUnavailable()) button.hidden = true;
    button.addEventListener("click", () => {
      const input = document.getElementById(button.dataset.target!) as HTMLInputElement | null;
      if (input) void open(input);
    });
  };
  buttons.forEach(register);
  search.addEventListener("input", render);
  close.addEventListener("click", back);
  return { back, register };
}
