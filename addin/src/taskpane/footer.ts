/**
 * Footer of the mail card: MAIL_FOOTER is a template with placeholders that are filled from the sender's
 * Microsoft 365 profile, e.g.  "{displayName} · {jobTitle}\n{company} · Tel. {phone} · {email}".
 *
 * Rules:
 *   - "\n" (typed literally in the env file) or a real line break starts a new line.
 *   - Each line is split into segments at " · ". A segment whose placeholders are all empty is dropped,
 *     together with its fixed text ("Tel. {phone}" disappears when there is no phone number).
 *   - Empty lines are dropped. Unknown placeholders count as empty.
 */

export interface Profile {
  displayName: string;
  givenName: string;
  surname: string;
  jobTitle: string;
  department: string;
  company: string;
  email: string;
  phone: string;
  mobile: string;
  office: string;
}

export const EMPTY_PROFILE: Profile = {
  displayName: "",
  givenName: "",
  surname: "",
  jobTitle: "",
  department: "",
  company: "",
  email: "",
  phone: "",
  mobile: "",
  office: "",
};

const SEPARATOR = " · ";

export const hasPlaceholders = (template: string) => /\{\w+\}/.test(template);

export function fillFooter(template: string, profile: Profile): string {
  const values = profile as unknown as Record<string, string>;
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
