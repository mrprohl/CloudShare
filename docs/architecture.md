# CloudShare – architecture

CloudShare is an Outlook add-in plus a small backend. Large files are uploaded to the **sender's own Nextcloud** and
inserted into the mail as a password-protected link; the password reaches the recipients in a **second mail** as a
Vaultwarden Send, and only after the first mail has really been sent. This document describes how the parts fit
together, where state lives, and the Outlook peculiarities the code has to work around. Operation and installation:
[README](../README.md); Microsoft 365 (Entra ID, deployment, Exchange): [microsoft-365.md](microsoft-365.md).

Everything visible to users and recipients (taskpane, Outlook hints, cards, password mail, SMS, download
confirmation) comes from the language files in `locales/` – the add-in in the Outlook client's language, the mails and
SMS in `LANGUAGE` (`en` by default, see [2.8](#28-languages)); code, logs and documentation are English. Labels are quoted as in the English interface
("Insert as link", "History" …); the German texts (the default) are in `locales/de.json` under the same keys.

## Contents

1. [Components](#1-components)
2. [Main flows](#2-main-flows)
3. [State and data](#3-state-and-data)
4. [HTTP API](#4-http-api)
5. [Outlook runtimes and pitfalls](#5-outlook-runtimes-and-pitfalls)
6. [Security](#6-security)
7. [Setup](#7-setup)
8. [User interface](#8-user-interface)
9. [Limits and risks](#9-limits-and-risks)

## 1. Components

```
Outlook (classic on Windows / new Outlook / Outlook on the web)
  ├─ Taskpane            addin/src/taskpane/   connect Nextcloud, pick files, upload, insert cards,
  │                                            "History"; read mode (?view=read): the shares of
  │                                            the mail being read
  └─ Event handlers      addin/src/commands/   OnMessageAttachmentsChanged → info bar
                                               OnMessageRecipientsChanged → same check (size inside / outside)
                                               OnMessageSend (Smart Alert) → attachment policy, arm password mail
        │  HTTPS + Office SSO token (Entra ID, audience api://<host>/<client-id>, scope access_as_user)
        ▼
Caddy (Let's Encrypt, ports 80/443) ──► network "internal"
        ▼
Container "cloudshare" (Node 24, read-only)
  ├─ Fastify backend     server/src/           API, add-in files (texts of all languages filled in at start),
  │     │                                      manifest.xml, well-known file, seven.io webhook
  │     ├─ Nextcloud     app password per user → WebDAV chunked upload, OCS shares, activity, trash
  │     ├─ Graph         On-Behalf-Of token → read "Sent Items", send password mail as the user;
  │     │                app token (optional) → download confirmation
  │     ├─ seven.io      password by SMS (optional)
  │     └─ bw serve      HTTP on 127.0.0.1:8087 → Vaultwarden Sends (service account)
  ├─ Bitwarden CLI "bw serve"
  ├─ /app/locales        built-in language files (locales/ in the repository)
  └─ /config (ro)        optional: templates.json, locales/<language>.json (host: env/config/)
```

| Part | Files | Notes |
|---|---|---|
| Taskpane | `taskpane.ts/.html/.css`, `mailCard.ts`, `settings.ts` | Plain TypeScript, bundled as IIFE (esbuild, ES2017) |
| Event handlers | `commands.ts` → `commands.js` | ES2016 bundle for classic Outlook's JavaScript-only runtime |
| Shared add-in code | `shared/office.ts`, `shared/password.ts`, `shared/sso.ts`, `shared/i18n.ts` | Office.js helpers, draft state, SSO token cache, password rules, `t()` for the texts |
| Languages | `locales/<language>.json`, `server/src/i18n.ts` | Texts of every language; loading, fallback, `t()`, injection into the add-in files ([2.8](#28-languages)) |
| API | `app.ts` (routes), `index.ts` (start: languages, vault unlock, background jobs) | `buildApp()` is used by the API tests |
| Integrations | `auth.ts` (SSO validation, OBO, app token), `nextcloud.ts`, `graph.ts`, `vaultwarden.ts`, `sms.ts` (seven.io) | |
| Texts sent on the sender's behalf | `passwordMail.ts`, `templates.ts` (`templates.json`) | Password mail, SMS, download confirmation |
| Background jobs | `jobs.ts` (password mails and SMS), `cleanup.ts` (expired uploads), `notices.ts` (download confirmation) | Every 20 s / hourly / every 10 min, see below |
| Persistence | `db.ts` (`PersistentMap`), `storage.ts` (sealing) | SQLite `/data/cloudshare.db`, every row AES-256-GCM sealed |

**Why app passwords for Nextcloud?** Nextcloud usually cannot accept Entra ID tokens (it often signs users in through
another IdP, e.g. Keycloak via `user_oidc`). Each user therefore connects once with Nextcloud's Login Flow v2; the
resulting app password is stored encrypted and used as Basic auth. Nextcloud is only ever called server to server,
so it needs no CORS.

**Placeholders.** `docker/entrypoint.sh` copies the add-in files from `/app/public.tpl` to `/app/public` (tmpfs) and
replaces `%%PUBLIC_BASE_URL%%`, `%%PUBLIC_HOST%%` and `%%ENTRA_CLIENT_ID%%`; then the server fills in the texts of
every language ([2.8](#28-languages)). So one image serves every installation and language.

## 2. Main flows

### 2.1 Connect Nextcloud (once per user)

1. Taskpane → "Connect" → `POST /api/nextcloud/connect` starts Login Flow v2 and returns the login URL, which the
   taskpane opens in the browser (usually SSO, no typing).
2. The taskpane polls `POST /api/nextcloud/connect/:id/poll` until the user granted access.
3. The backend stores the app password encrypted under the user's Entra object ID (`oid`). In Nextcloud it appears as
   "CloudShare Outlook" under *Settings → Security* and can be revoked there; a 401 from Nextcloud deletes the stored
   credential and the taskpane shows the connect view again (`NEXTCLOUD_NOT_CONNECTED`).

### 2.2 Share files ("Insert as link")

The "Send" tab starts with a switch for **what** goes into the mail – **Files**, **Message** (a confidential message) or
**Request** (an upload link) – and shows only that part; the one button in the footer names what it inserts ("Insert 2
files as links", "Insert confidential message", "Insert upload link"). Password, the block "Empfänger" ("Kennwort an", for files the switch "Zustellnachweis") and the collapsed
"Gültigkeit" (one-line summary of the times for the chosen kind) apply to all three. "No password" is offered for files
only; a mail already fixed to links without password disables the button for the other two and says why. An option
that does not exist for the chosen kind is not hidden without a word: "Ohne Kennwort" shows
greyed out for "Nachricht"/"Anfordern" (`#mode-none-unavailable` – an own disabled copy, so the
lock logic of the real one stays untouched) with an "i" (`.why-toggle`, `showUnavailable`) that folds out the reason
(`password.noneFilesOnly`); "Zustellnachweis" is simply not shown there. Nextcloud enforcing passwords still hides "Ohne Kennwort" (a rule of the organisation, not of the kind). Files
dragged onto the pane switch it to "Files".

1. The taskpane lists the mail's file attachments (all ticked – whoever opens CloudShare sends through it; the threshold
   only decides when Outlook points to CloudShare; an attachment unticked stays so when the list is redrawn) and files from
   the computer – dropped onto the drop zone (the whole "Send" tab accepts drops) or picked by clicking it (native dialog `window.showOpenFilePicker`, fallback: hidden `<input type="file">`).
2. **One click = one step** with a random `step` ID. All files of the step are uploaded one by one:
   `POST /api/uploads` (quota check via OCS `cloud/user`, 507 → `NEXTCLOUD_QUOTA_EXCEEDED`); a service that cannot be reached (DNS, connection refused, timeout) answers 502 `SERVICE_UNREACHABLE` naming it (`unreachableService`) →
   `PUT /api/uploads/:id/chunks/:n` (10 MiB chunks, Nextcloud chunked upload v2) → `POST /api/uploads/:id/complete`.
   Target: `NEXTCLOUD_FOLDER/<date>_<id>/<file>`. Attachments are removed from the mail only after a successful upload.
3. `complete` creates one public link per file (password, expiry date) and attaches it to the mail's **password job**
   (`attachToJob`): the first file of a step creates a Vaultwarden Send with the step's password, further files of the
   same step join it. The job ID and its parts are described in [3](#3-state-and-data).
4. The taskpane inserts the first card (`mailCard.ts`) at the cursor – falling back to the top of the mail when Outlook
   inserted nothing (new Outlook without a cursor in the body) – followed by an empty paragraph. Every further card of
   an HTML mail goes right after the last CloudShare card, wherever the cursor is (the user may have clicked into a
   card): Office.js cannot insert after an element, so `insertAfterLastCard` reads the body, `cardPlace.ts` finds the
   frame (the table of width 560 around the last token of `DraftState.links`, outside any table the body itself is laid
   out in) and the body is written back with `body.setAsync`; if the card cannot be found it goes to the cursor. Text
   mails always insert at the cursor. It tags the draft with the custom property `cloudshareJob` (job ID) and records the link
   token in the draft state.
5. One password per mail: after the step the taskpane selects "Gleiches Kennwort wie zuvor" (`mode-same`,
   `DraftState.lastStep` = step ID and password) and greys out "Zufällig", "Eigenes", "Ohne Kennwort" and "Kennwort an"
   (`lockSeparate`, `lockChannel`): the next link – files, message or upload link – is sent with the same `step` and
   joins that part (`findPart`/`addFileToPart` on the server: same password, same Send, one password mail or SMS). The
   server refuses a link with a password that would open a second step in the job (`assertOnePassword`, 409
   `ONE_PASSWORD`, before any link is created); mails started with one shared password (no `separate`) are not affected. The password field is
   always `type="text"`; an own password is hidden by CSS (`.masked`, `-webkit-text-security`, `setRevealed`), because
   a real password field made Edge (WebView2) offer to save it as a login after the click.

Variants: **"Request"** (`POST /api/requests`) creates a folder in `NEXTCLOUD_REQUEST_FOLDER` with
a create-only link ("file drop") – same password rules, its own step. **"No password"** (`noPassword: true`, file
links only) creates a link without password and Send; with `DOWNLOAD_NOTICE=true` it joins a job **without password**
(`createNoticeJob`, `PasswordJob.noPassword`, the links in `noticeShares`), which the send handler arms like any job –
the watcher sends nothing and only registers the links for the download confirmation once the mail is in "Sent
Items". Without the confirmation there is no job (`jobId: null`). The answer always says `noPassword: true`. **"History → Insert"**
(`POST /api/shares/:id/link`) creates a new link to an existing upload – in a mail that already has a password it joins its step (same password).

**"Message"** (`POST /api/messages`): a text instead of a file. It is never written to Nextcloud or to
`/data`: `createMessageSend` stores it as a Vaultwarden **text Send protected with the step's password** (Send access
password, `hidden: false`, no access limit, deleted at expiry). The Send's link goes into the mail as a card
("Open message"); the step's password goes out like for files – the job item is "Confidential message", the
password mail shows an envelope badge. Always with a password; at most 1000 characters (Bitwarden's limit for text
Sends). Its lifetime is chosen separately ("Message valid for", `messageDays`): `SEND_EXPIRE_DAYS` (default 5) up to
`SEND_MAX_EXPIRE_DAYS` (at most 30 – Vaultwarden refuses deletion dates more than 31 days ahead).

**"Password link valid for"** (`passwordDays`, same limits, at most the link's validity): how long the recipients can
open the step's password Send – its *expiration date*. It is created with the first link of the step, so it counts
from inserting, not from sending; the date is stored with the step (`sendExpiresAt`) and named in the password mail
("expires on …"). After it the recipients can no longer fetch the password, even if the download link is still valid
– those who opened it in time keep it. The Send itself lives until the links expire (its *deletion date*,
`sendKeepUntil`), so the sender can still see and resend the password (2.10). Vaultwarden accepts deletion dates at
most 31 days ahead, hence **a link with password lasts at most 30 days** (`passwordLinkMaxDays`, also in
`/api/config`); links "Ohne Kennwort" up to `SHARE_MAX_EXPIRE_DAYS`. A confidential message's password Send lives as
long as the message. Messages are listed in "History" (`history.ts`, 2.10) – with password and revoke, never their
key – but get no download confirmation.

### 2.3 Send and password mail

1. On "Send" Outlook runs the Smart Alert `onMessageSendHandler` (manifest `SendMode="SoftBlock"`):
   - Large attachments left in the mail: under `ATTACHMENT_POLICY=suggest` (default) the user is asked once and may
     choose "Send Anyway"; under `enforce` sending is blocked with the button "Open CloudShare"
     (`sendModeOverride: PromptUser`, Mailbox 1.14).
   - If all CloudShare cards were deleted from the mail, the draft is released: no tag, no password mail.
   - Otherwise the job is **armed**: `POST /api/jobs/:id/arm` with an SSO token; the backend exchanges it for a Graph
     token (On-Behalf-Of, `Mail.Read(.Shared)` + `Mail.Send(.Shared)`) and stores it with the job. If arming fails,
     sending is blocked with a message naming the failed step.
2. The watcher (`jobs.ts`, every 20 s) searches "Sent Items" for the mail carrying the `cloudshareJob` tag (MAPI
   property `cecp-<add-in ID>`; `ADDIN_ID` must match the manifest `<Id>`). For a "From" that differs from the user
   (shared mailbox, send on behalf), that mailbox's "Sent Items" are searched as well.
3. Once found, it sends **one password mail per step** (`passwordMail.ts`) to the mail's **actual** recipients – or
   only to the address from "to another email address" – from the same sender. Sent parts are
   remembered, so a retry after an error sends nothing twice. Typical delay: 20–60 s after the main mail.
4. If the mail does not show up within 15 minutes (sending cancelled or failed), the job is disarmed and stays open
   for further links. Jobs expire 7 days after their last link was added.
5. Links are created when they are inserted, not when the mail is sent. Until the watcher has seen the mail sent,
   "History" marks the links of a job as **not sent** (`unsent`, `revokeAt` in `GET /api/shares`, from
   `unsentShares`). A job that expires without "Send" ever having been pressed (`armedOnce: false` – the draft was
   discarded) has its links revoked and its Sends deleted (`revokeUnsentLinks` in `shares.ts`, retried hourly while
   Nextcloud is unreachable). Jobs that were armed once are left alone – the mail may have gone out without the
   watcher seeing it – and so are links without password, which may have been sent from a client without CloudShare.
   A link that was already downloaded (Nextcloud activity) is never revoked automatically either: it reached someone
   (opened by the sender, sent from Outlook without CloudShare, copied). "History" then shows it without a revoke
   date as "Password not sent yet." (`shares.unsentDownloaded`), next to its downloads.

If the add-in cannot run at all, SoftBlock lets the mail through – otherwise a CloudShare outage would stop all mail.
The handlers cache the last configuration in `roamingSettings`, so the attachment policy also applies while the
server is unreachable. Outlook mobile and other mail clients do not run the add-in and are not checked.

### 2.4 Cleanup of uploads

Every file link registers its upload folder with the link's expiry date in the database (store `cleanup`; the latest link
wins) – Nextcloud's housekeeping deletes expired shares, so the date could not be read back later. An hourly pass
(`cleanup.ts`) deletes due folders via WebDAV `DELETE` with the owner's app password, i.e. into the owner's Nextcloud
**trash**:

- due = day after the expiry date + `CLEANUP_AFTER_DAYS` (default `0`; `off` disables the cleanup);
- **"Revoke"** counts as expiry today – with `0` the folder goes to the trash at once;
- folders that still have a valid link are postponed; users who disconnected are skipped;
- only folders matching `NEXTCLOUD_FOLDER/YYYY-MM-DD_xxxxxxxx` are touched – received files
  (`NEXTCLOUD_REQUEST_FOLDER`) are never deleted. Uploads from before the feature existed are not registered.

### 2.5 Password channel: recipients, another address or SMS

"Kennwort an" (a select under "Empfänger") in the taskpane decides **per insertion** where the password goes: every link carries
its way in its request (`passwordTo` – `""` = the recipients – or `passwordSms`, `deliveryFrom` in `app.ts`), and the
step collects the ways of its links (`SeparatePart.deliveries`). A step with links for the recipients and one of the
same password by SMS sends one password mail and one SMS; each way exactly once, also after an error in another one
(`sentWays`). Changing the select only affects the next link. Steps from before this (no `deliveries`) and mails with
one shared password keep the mail's one way (`PasswordJob.passwordTo` / `passwordSms`, `PATCH /api/jobs/:id`, `/arm`).
While typing another address or an SMS number the pane suggests entries of the user's own Outlook contacts
(`contactSuggest.ts`: ARIA combobox, from two characters, `GET /api/contacts`; for SMS only numbers `smsNumberProblem`
accepts); a picked entry is filled in and saved like a typed one. The address book button next to the field
(`.book-open`, `addressBook.ts`) opens the page "Adressbuch" (`#panel-contacts`, `showOverlay("contacts")`): every
entry of that kind A to Z under its first letter (`GET /api/contacts?all=1`), filtered while typing; a click takes it
the same way and leads back to where it was opened (`contactsReturn`). "Erneut senden an" in "Verlauf" has the same
field with both (`addressBook.register`, `attachContactSuggestions` with the kind following the select); it takes several
addresses (`data-multiple`): the suggestions match the part after the last comma and a picked address is added to the list
(`addToList`). Without the permission the buttons disappear (`onContactsUnavailable`). Without the optional delegated `Contacts.Read` the
server answers `contacts: null` and the pane stops asking for the session.
The ways:

- **Recipients of the message** (default) – the password mail to the mail's actual recipients;
- **Another email address** – the password mail only to that address;
- **Phone number (SMS)** (only with `SEVEN_API_KEY`, otherwise not in the list) – one SMS per step via seven.io (`sms.ts`), sent by the watcher after the
  mail appeared in "Sent Items", like the password mail. The number is entered in the pane and normalised to E.164
  (`0170 …` → `+49170…`, `SMS_DEFAULT_COUNTRY`; also `+49 (0)170 …`, `tel:` and typographic dashes).
  `SMS_ALLOWED_COUNTRIES` (`countries.ts`: ISO codes, calling codes, German/English names, groups like `DACH`/`EU`)
  limits the countries by calling code; then `SMS_ALLOWED_NUMBERS` (must match, empty = all) and
  `SMS_BLOCKED_NUMBERS` (must not match, default `^\+49(800|900|180)`: no 0800/0900/0180 numbers) check the E.164
  number as regular expressions – mobile and landline numbers are fine. `/api/config` hands the pane the calling
  codes, the names and both expressions (`smsNumberProblem` in `sms.ts` and `taskpane.ts`); the pane refuses such a
  number as soon as it is entered (and in "Send again"), the server on `PATCH` and resend (400) and the watcher
  before every SMS (then the note of a failed SMS). The password exists in readable form only in the step's Vaultwarden
  Send; the watcher reads it from there right before sending (`readSendPassword`). The Send stays until the links
  expire (2.10). The password-link settings do not apply (the pane hides "Password can be opened" and "Password link
  valid for"): the link of an SMS step's Send is never sent, so its access lasts as long as the Send (`PATCH` sets
  this also for steps inserted before SMS was chosen). The password in the SMS works as long as the link.
  **If the SMS fails, the password goes out no other way** – not to the recipients, not to the sender. The sender
  gets a note without password or link (`smsProblemMailHtml`) and sends it again (2.10). Two cases, told apart in
  `GET /api/shares` by `password.refused` (failed without seven.io message ID):
  - seven.io refuses it (key, credit, `SMS_MAX_PER_DAY`): the watcher sends the note with the armed delegated token;
    the step counts as done (no retry of the SMS on every pass).
  - seven.io reports it later as not delivered: delivery reports arrive at `/webhooks/seven` (secret in the query,
    compared in constant time, idempotent, webhook and classic form callback) and update `deliveries.ts`; a newly
    failed SMS triggers the note with CloudShare's app token (`sendMailAsApp`) – only with `SMS_NOTICE=true`
    (application permission `Mail.Send`; unset it follows `DOWNLOAD_NOTICE`, `envFlag`), otherwise "History"
    shows it only.

How the password of each link went out (channel, masked recipient, SMS status) is shown in "History".

### 2.6 Text templates

`templates.ts` reads the optional `TEMPLATES_FILE` (`/config/templates.json`, mounted from `env/config/`): fields
`sms_template`, `mail_password_subject`, `mail_password_intro`, `download_notice_subject`, `download_notice_body`, each
with a fixed set of placeholders. Resolution per field: `users[<sender UPN>]` → `groups[<department>]` → `default` →
built-in texts of the language (section `templates` of the language file); the department and display name are read
from the sender's profile when arming. `templates.json` holds one set of texts for all languages – after a change of
`LANGUAGE` translate it as well. At start the file is validated: `{password}` outside the SMS, an SMS without `{password}` or with the link / the raw file name
(`{filename}`) drop the
field (built-in text instead); unknown placeholders, non-GSM characters, SMS above 160 characters and a `{password}`
that does not start a line (phones may hyphenate a wrapped password) are warnings.
A broken file never stops CloudShare. Example: `templates.example.json`.

### 2.7 Download confirmation

With `DOWNLOAD_NOTICE=true`, every file link of a sent mail is registered – with password or, via the job without
password of 2.2, without – with the sender, subject, recipients and
expiry (`notices.ts`, store `notices`). Every 10 minutes the Activity API is asked with the owner's
app password; only downloads **after** the password went out count (older links of the same file are ignored). On the
first download the sender gets a mail (`download_notice_*` templates above a green card, layout of 8) – sent with CloudShare's **own app token**
(client credentials, application permission `Mail.Send`), because the sender's delegated token has long expired by
then. Entries are dropped after the download or when the link expired. No IP addresses or data about the person who
downloaded are kept.

**Upload notification** (`UPLOAD_NOTICE=true`, separate switch, same app token): upload links ("Anfordern") carry
`request: true` and `knownFiles` in their `StepShare` (the files already in the folder when the link was created –
none for a new folder, the received ones when an upload folder is linked again) and are registered as
`kind: "upload"`. Each pass lists the folder (`listFiles`, WebDAV `PROPFIND` depth 1) and mails the sender the files
not reported yet – at most one mail per link and pass (green card with up to 10 names and "and N more", button to the
folder in Nextcloud, texts `mail.uploaded*`). The entry stays until the link expired.

### 2.8 Languages

Two kinds of language:

- **The add-in** follows the display language of the Outlook client (`Office.context.displayLanguage`): every
  language in `locales/` and `CUSTOM_LOCALES_DIR` is offered; the exact one (`pt-br`), else the one without region
  (`de-AT` → `de`), else English (`pickLanguage`, the same rule in `addin/src/shared/i18n.ts` and
  `server/src/i18n.ts`). That covers the taskpane, info bar, Smart Alerts, the link card and the error messages the
  server returns to the add-in.
- **`LANGUAGE`** (default `en`) is the first guess of the taskpane page, the preselected language for the recipients
  and the fallback for notes to the sender whose add-in language is unknown.
- **"Language for the recipients"** (only the personal setting `recipientLanguage` – `#settings-recipient-language`;
  "Senden" has no choice per mail, `recipientLanguage()` in `taskpane.ts`): the link card is rendered in it
  (`inLanguage` in `addin/src/shared/i18n.ts`), and every link request carries it as `language`. The server keeps it
  per step (`SeparatePart.language`, the first link of a step decides) and writes the step's password mail and SMS in
  it (`inRecipientLanguage` in `i18n.ts`: sections `mail` and `templates`, dates); sending the password again uses the
  same language (`PasswordRef.language`). The names in the password mail ("Confidential message", upload link) follow
  it too (`tIn`).
- **Notes to the sender** (download and upload confirmation, SMS not sent / not delivered) come in the language of
  their add-in when they pressed "Send": the send handler's `X-CloudShare-Language` is kept with the armed job
  (`ArmedSend.language`, `requestLanguageCode`) and passed on to the notices (`Notice.language`) and the password entry
  (`PasswordRef.senderLanguage`); the notes are rendered in it with `inSenderLanguage`.
- **`templates.json`** may hold texts per language on every level (`"languages": { "en": {…} }`); `templatesFor` uses
  those of the language of the mail being written (`mailLanguage`: the recipients' or the sender's), over the level's
  own.
- **Manifest:** each language's texts are given as Overrides for its `_meta.locale` and the regional variants in
  `_meta.regions` (German: de-DE, de-AT, de-CH, de-LU, de-LI – `manifestLocales`), so Outlook need not guess.

Each language is `locales/<language>.json`. `i18n.ts` merges, per text:

1. English – fallback for texts a language does not have yet;
2. the built-in file of the language (`/app/locales` in the image);
3. an optional own file from `CUSTOM_LOCALES_DIR` (`/config/locales`, mounted from `env/config/locales/`) – adds a
   language or changes single texts of a built-in one without a new image.

An unknown `LANGUAGE` (or one that is not a code like `de`, `en`, `pt-br`) stops the start with the list of available
languages. A language file:

| Part | Content |
|---|---|
| `_meta` | `name`, `locale` (BCP 47, e.g. `de-DE`: date and number format, `Locale` of the manifest Overrides) |
| `addin` | taskpane, info bar, Smart Alerts, link cards, SSO hints (`sso.<code>`) |
| `manifest` | description, ribbon group, tooltip – English as `DefaultValue`, the other languages as `Override` |
| `server` | error messages the taskpane shows (`HttpError`) – in the language of the request, share label of upload links |
| `mail` | password mail (greeting, card, notes), subjects of fallback mails |
| `templates` | built-in texts of [2.6](#26-text-templates) (SMS, password mail subject/intro, download confirmation) |

Keys are flat within a section (`"files.added_one"`); placeholders are `{name}`; `_one` / `_other` are the singular
and plural form – `t("files.added", { count })` picks one. Texts never contain markup: HTML and XML get them escaped,
the bundles as an ASCII-only JSON string.

The add-in gets its texts without an extra request: at startup (`index.ts`, `localizePublic`) the server replaces in
the copies under `/app/public`:

- the string marker `"%%CLOUDSHARE_I18N%%"` in `taskpane.js` / `commands.js` with a JSON string
  `{ fallback: "en", languages: { <language>: { locale, messages } } }` of the `addin` texts of every language (read
  by `addin/src/shared/i18n.ts`, ASCII-escaped) – so the send handler in classic Outlook's JavaScript-only runtime has
  them too and picks the client's language itself;
- `%%t:<section>.<key>%%`, `%%LANG%%`, `%%LOCALE%%` in `taskpane.html` / `commands.html` with the texts of `LANGUAGE`,
  and writes `taskpane.<language>.html` for every language. `taskpane.js` starts with `main.ts`: once Office.js is
  ready, `language.ts` replaces the body with the page of the client's language (if it differs) and only then
  `taskpane.ts` is evaluated (dynamic import inside the one IIFE bundle) – it binds its elements at evaluation. If
  that page cannot be loaded, the pane stays in the page's language, texts of page and code always match;
- `manifest.xml`: each `DefaultValue="%%t:manifest.<key>%%"` becomes the English text with one
  `<Override Locale="<_meta.locale>">` / `<bt:Override …>` per other language; `DefaultLocale` is English's locale.

Every request of the add-in carries `X-CloudShare-Language`; an `onRequest` hook runs the request in an
`AsyncLocalStorage` with that language, so `t("server.…")` (error messages, also the rate limit's) answers in it. All
other sections – mails, templates, SMS – stay in `LANGUAGE`, also for jobs started by a request. Comments in these
files must not contain `%%` – the smoke test fails on any marker left over.

Nextcloud folder names do not depend on the language (`NEXTCLOUD_REQUEST_FOLDER` identifies CloudShare's links).
Outlook caches the add-in files – after adding a language clear the cache ([5](#5-outlook-runtimes-and-pitfalls)); the
ribbon texts change only after the manifest was updated in Microsoft 365.

**Adding a text:** put the key into **every** file in `locales/` with the same placeholders. `server/test/i18n.test.ts`
checks that every language has the keys and placeholders of English, that every key the code uses (`t("…")`,
`%%t:…%%`) exists, Outlook's length limits (Smart Alert button ≤ 20 characters, manifest strings) and that the SMS text
is GSM with `{password}`. Add-in unit tests load `locales/de.json` (`addin/test/setup.ts`); the UI tests serve the
bundles with the German texts and below `/en/` with the English ones for the README pictures
(`addin/scripts/localize.mjs`); the manifest is validated once per language; the smoke
test checks the German texts in the running stack.

### 2.9 Status of sent mails (read mode)

The manifest also adds the CloudShare button to the ribbon of a mail being **read** (`MessageReadCommandSurface`,
pinnable via `SupportsPinning`; legacy `ItemRead` form for older clients). It opens `taskpane.html?view=read`: in place
of "Senden" the links of that mail (`#panel-mail`, "Links in dieser Nachricht" with its own refresh `#refresh-mail`;
shown again when a pinned pane moves to another mail); "Verlauf" and the rest are in the account menu as while
writing. A link revoked in one list disappears from both.

The mail is matched by its body, not by a server-side record: the pane reads the HTML body (`body.getAsync`) and
collects the tokens of Nextcloud links (`/s/<token>`) and the access IDs of message links (`#/send/<accessId>/`); the entries of `GET /api/shares` whose token or `accessId` is among them belong
to this mail. That works for every mail with CloudShare links – sent ones, forwarded ones, also links "Ohne Kennwort"
that never had a job – and needs no new permission (the item is only read). Only the user's own shares can match;
links of other people or other sites in the mail are ignored. Confidential messages are not listed (Vaultwarden
Sends, not Nextcloud shares). While the pane is pinned, `ItemChanged` clears the message bar and reloads the list for the selected mail. Copy
and Revoke work as in the compose pane; Insert is hidden.

### 2.10 Showing and sending the password again

Once a step's mail was sent, the watcher records where its password lives: `passwordRefs.ts`,
store `password-refs` in the database (sealed like the jobs) with the **Send ID** – never the password –, the links' tokens,
file names, subject, the distribution of the sent mail (To, Cc, Bcc), another address or SMS number the password
went to, and the sender. It is kept until the links
expire (`keepUntil`, the Send's deletion date) and only ever returned to its owner.

In "History" and in the read mode (2.9) such links get **Password** (`passwordAvailable` in `GET /api/shares`):

- `GET /api/shares/:id/password` – the server reads the password from the Send with the service account
  (`readSendPassword`; also after the recipients' access ended) and returns it with the suggested targets.
- `POST /api/shares/:id/password/send` – the same ways as the first time (the pane preselects the way it went then):
  `{ channel: "recipients" }` sends the password mail with the **distribution of the sent mail** – To, Cc and Bcc as
  they were (`PasswordRef.recipients`, taken from "Sent Items"; older entries fall back to `mailTo`);
  `{ channel: "mail", to }` sends it to up to 20 other addresses, **each in a mail of its own**, so nobody sees the
  others (remembered in `otherTo`); `{ channel: "sms", to }` sends an SMS to that number (same template and daily
  limit as the first one). A mail gets a **new** Send with the same password (default access and days, deleted with
  the links) and goes out as the user via Graph (delegated, like the first one, also from a shared mailbox); it never
  contains the password itself. "History" shows the new delivery – To and Cc, never Bcc.
- A step whose password is **held back for the delivery proof** (`PasswordRef.deferred`, 2.12) is never sent again
  around it: before any recipient confirmed their address with the code every way answers 409 `PASSWORD_HELD`;
  afterwards `channel: "recipients"` reaches the **confirmed addresses only** (`verifiedAddresses` in
  `proofPassword.ts`: the `recipient.verified` events of the step's links), each in a mail of its own. `GET …/password`
  says `deferred` and `verified`, so the pane shows a note instead of "Senden" before a confirmation and names the
  confirmed recipients afterwards ("Bestätigte Empfänger").

Revoking the last link of a step deletes its Sends at once (`forgetPasswordToken`); otherwise Vaultwarden deletes them
when the links expire. Steps sent before this version have no entry – their password cannot be shown.

**Confidential messages and the sent mail in "History"** (`history.ts`): a message is a Vaultwarden Send, no Nextcloud
share, so CloudShare records it itself when it is created (store `messages`: Send ID, access ID, dates – never the key
from the link or the text) and adds it to its step like a link (`StepShare.message`, token `msg-<Send ID>`). It is then
covered by the step's password entry above – "Password" works as for files – and `DELETE /api/shares/msg-…` deletes its
Send (`revokeMessage`); the watcher deletes the message of a mail never sent with its links. It has no "Copy" or
"Insert again" (no key) and no download state (Vaultwarden does not report openings). When the watcher finds the sent
mail it stores per link the subject and the mail's Graph ID (store `sent-mails`, `recordSentMail`; requested with
`Prefer: IdType="ImmutableId"` so it survives moving the mail; only for the user's own mailbox, not a shared one).
"History" titles each entry with that subject and opens the mail on a click (hovering shows when and to whom it went –
`sentAt`, `sentTo` with To and Cc, never Bcc, kept with the subject) (`openSentMail`: `GET /api/shares/:id/mail`
translates the immutable ID into the current EWS ID with Graph `translateExchangeIds` – Outlook cannot open the
immutable form – then `displayMessageFormAsync`, which reports a failure; a mail in Deleted Items or – after emptying them – in
Recoverable Items counts as deleted, `sentMailState`); until the mail is sent the entry shows the draft's subject (every link request carries `subject`,
`recordDraftSubject`, a subject typed later follows with `PATCH /api/jobs/:id` from the pane's 4-second check,
`syncDraftSubject` – kept even if the draft is then discarded; replaced by the real one when the mail is found); a mail not
sent yet without subject says "(No subject yet)", entries from before "(Subject unknown)".

### 2.11 Personal settings

"Einstellungen" in the account menu (`#settings-open`, bottom left) opens the settings in place of the current tab, like the guide (`settings.ts` in the
add-in, `server/src/settings.ts`, store `settings`). Every change is saved at once (no "Speichern": a box or a list on
`change`, the signature one second after typing stops and on leaving it – one `PUT` at a time, nothing sent when
nothing changed); the page stays open until "←" (`closeSettings`, focus back on the account). Notices and signature
are in sight; the ways for the password, the languages and "Zurücksetzen" (keeps the page open with the defaults in
sight) are folded away under "Weitere Einstellungen" (`#settings-more`):

- **Ways for the password** (`channels`, `defaultChannel`): `applyChannelChoices` hides the other options of "Kennwort
  senden an" (and of "Erneut senden an") but keeps the way a mail already uses; a new mail starts with
  `defaultChannel`. Saving switches a mail without links to the new default if its way is no longer offered.
- **Notices** (`downloadNotice`, `uploadNotice`, `smsNotice`): checked when it matters – `attachNoticeJob` and
  `runNotices` (a notice the sender switched off is dropped), `noteUndeliveredSms`. A refused SMS is always reported:
  the password never went out.
- **Language** (`auto` or a language): kept in `roamingSettings` too, so `i18n.ts` (`preferredLanguage`) has it in
  the taskpane and in the event handlers before any request; loaded from another computer, the pane reloads once
  (`reloadOnce`), saved in the pane it reloads at once. Mails and SMS stay in `LANGUAGE`.
- **Signature of the mails to the recipients** – password mail and the code mail of a proof of delivery
  (`{ format: "text", content }` or null = `MAIL_FOOTER`), so recipients see who it comes from; the link card in the message carries none. **Plain text only** – nothing the user types
  becomes markup (a formatted signature was considered and dropped for safety). `server/src/signature.ts`: an empty
  line separates paragraphs (`signatureParagraphs`, also `\n\n` in the env file), each paragraph is filled from the
  profile like `MAIL_FOOTER` always was (`fillTemplate`), escaped and shown as a `<p>` in the style of the mail text
  (`TEXT_STYLE`, like "Guten Tag"), lines as `<br>`. When the job is armed, `signatureHtml` fixes it with the profile
  of that moment (`PasswordJob.signature`); "send again" builds it anew. At most 20,000 characters; a formatted
  signature saved by an earlier version counts as none.

`settingsFor(oid)` merges `config.userDefaults` (`DEFAULT_*`), the stored fields and what the server offers
(`availableSettings`: notices switched on, SMS with seven.io, the add-in languages) – so a user can never turn on what
the organisation switched off, and the SMS rules (countries, allowed and blocked numbers) apply regardless. Only
changed fields are stored; a new default reaches everyone else.

### 2.12 Delivery proof ("Zustellnachweis")

Optional per mail, for files only (the box is hidden for "Nachricht" and "Anfordern"), never by default: the sender
switches on **"Zustellnachweis"** under "Empfänger" (a switch, `#proof` with `role="switch"`,
offered with `DELIVERY_PROOF`, which follows `DOWNLOAD_NOTICE` when unset – it needs the same application permission
`Mail.Send`). It is evidence that can be checked, not a qualified delivery service in the sense of eIDAS.

- **Case and tracking ID** (`server/src/proof.ts`, store `proofs`, `retention: false`): the first link with the box
  ticked opens a case named like `20261117-1146-7Q2F` (date, time, four characters without look-alikes); the add-in
  keeps it in `DraftState.proofId` and sends it with the next links (`proofRequest`: `proof`, `proofId`, `client` –
  Outlook's host, platform and version), so all links of the mail share one case. The link card, the password mail,
  the download confirmation and "Verlauf" name it.
- **Events, hash-chained**: `case.opened`, `link.created` (name, size, **SHA-256 of the content** – computed while the
  chunks pass through, `UploadState.hash`; protection, expiry, password ways), `message.sent` (subject, **Internet
  Message-ID** for Exchange's message trace, From, To, Cc, Bcc, time – `recordMailSent` from the job watcher),
  `password.sent` / `sms.status` (`recordDelivery`, `updateSmsStatus`), `password.fetched` and `message.opened`
  (access count of the Vaultwarden Send went up), `download` (every download time Nextcloud's activity reports),
  `files.received`, `recipient.codeSent` / `codeFailed` / `verified`, `link.opened`, `link.revoked` / `expired` /
  `gone`. Every event stores `sha256(previous hash + canonical JSON of time, type, data)`; `verifyCase` re-computes
  the chain and the PDF says whether it is intact or from which entry on it was changed.
- **Watcher** (`proofWatcher.ts`, every 10 minutes): downloads, received files, Send access counts, links gone or
  expired; a case is done when all its links ended (or two days after the last expiry) and is then only kept –
  `PROOF_RETENTION_MONTHS` after its last event (calendar months, default 36 = three years, the regular limitation period).
- **Recipient confirms their address** (`proofGateway.ts`): file and upload links of a case go into the mail as
  `https://<host>/r/s/<token>` instead of the Nextcloud link (the path still ends in `/s/<token>`, so the read mode
  and `releaseDraftIfLinksGone` find it). The page asks for the address; only one the mail went to (To, Cc, Bcc – a
  distribution list does not match) gets a six-digit code, sent from the sender's mailbox with the app token (`codeMail`: who sent what with proof of
  delivery, the code large – `MailCard.highlight` –, the next steps depending on where the password will come from, and
  that it can be ignored if nobody asked for it, and the sender's signature below – `PasswordRef.signature`, fixed when the
  mail was sent, else the current one); the page
  answers every address alike – the same text, the same limits (counted for every request), the same timing (the code
  mail is not awaited) and one error text for a wrong, expired or never sent code – so trying does not reveal the
  recipients; an address that is not one is recorded (`recipient.unknownAddress`). Ten minutes, five tries, three codes per
  address in 15 minutes, ten per link an hour. A right code records address, IP address (via Caddy, `trustProxy`) and
  browser, sets a signed cookie for this link until it expires and redirects (303) to Nextcloud – the Nextcloud
  password is still needed. Before the mail was found sent the page says so. Plain HTML forms, no script, own CSP. A header with the CloudShare logo, "CloudShare" and the
  organisation (`COMPANY_NAME`), a footer with `COMPANY_WEBSITE_URL`, `IMPRINT_URL`, `PRIVACY_URL` (`brandHeader`,
  `linksFooter`; CSP `img-src 'self'` for the logo).
  A confidential message has no gateway (its link holds the key CloudShare does not keep); its openings are counted.
- **One mail, one kind** (`assertNoProofMix`, `proofKindOf` in `app.ts`): a password step holds back its password only
  if every link of it goes through the gateway – so a mail never mixes links with and without a proof. Inserting a link
  of the other kind into the mail's job (`jobId`) answers 409 `PROOF_MIX` (files with `proof`, upload links and messages
  never have one); `/arm` refuses a job that mixes them anyway, so the send handler holds the mail back. The pane fixes
  the kind with the first link (`DraftState.proofMail`, `applyProofLock`: the switch fixed, "Nachricht" and
  "Anfordern" disabled with a proof; the reason folds out from an "i" next to the fixed switch, `#proof-lock-toggle` – the hint under
  the switch never changes).
- **Password after the confirmation** (`proofPassword.ts`): the mail to the confirmed address says so ("Sie haben eben Ihre
  E-Mail-Adresse bestätigt …") and its card has two numbered steps (`MailCard.steps`: 1 "Kennwort anzeigen", 2 "Zum Download", each button with its address as plain text below), the download to the gateway – never to Nextcloud: the confirmed browser goes
  straight on, another one or a forwarded mail is asked for the code again, so link and password never travel together
  (`afterConfirmation` in `passwordMailHtml`); a step whose links all go through the gateway
  (`isDeferredStep` in `jobs.ts`) sends no password when the mail is sent (`password.deferred`); its `PasswordRef` is
  `deferred` with `toRecipients` and the signature. After a right code `releasePassword` sends what is still held
  back – a password mail (new Send with the same password, app token, as the mailbox the mail came from) to the
  confirmed address, to other addresses once, the SMS once (`ProofCase.released`) – and the gateway shows the next
  steps with a link on instead of redirecting at once. The link card says so (`card.proofPassword`).
- **Mails to the sender** (`proofMail.ts`, app token, PDF attached via `sendMailAsApp` attachments): after the first
  download (`downloadMailSent`; `notices.ts` skips its download confirmation for links of a case) and when the case is
  done (`finalMailSent`, retried until it went out); `proof.sent` records each. `GET /proof/:id.pdf` first runs
  `refreshCase`, so the PDF has downloads since the last pass.
- **Reminder** (`proofWatcher.ts` `remindIfNothingFetched`, `proofMail.ts` `sendReminder`): a case whose state is still
  "sent" or "confirmed" (`caseSummary` in `proof.ts`: no download, no files received, message not opened) and whose last
  running link has used up half of its validity since the mail was sent (`reminderTime`: 14 days → after 7, 2 days →
  after 1, never within the first 24 hours) mails the sender once (`reminderSent`,
  `proof.reminded`) – unless the personal setting `proofReminder` (`DEFAULT_PROOF_REMINDER`) is off. Recipients are
  never written to.
- **Result first**: the PDF starts with a frame "Ergebnis" (`summaryLines`: sent, address confirmed, downloaded /
  received / opened, link ends or ended, kept until – a ZapfDingbats tick or circle each; `pdf.ts` lines with `box`,
  `mark`, `value`), then the record.
- **Signature** (`pdfSign.ts`, `der.ts`, no dependency): `pdf.ts` leaves an invisible signature field with placeholders
  (`/ByteRange`, `/Contents` of 32768 hex digits), `signPdf` fills them in place: a detached CMS SignedData (SHA-256,
  RSA, the certificate, ESS signing-certificate-v2) over the bytes outside `/Contents`, with an RFC 3161 token from
  `PROOF_TSA_URL` as unsigned attribute (only the hash of the signature leaves the server; without an answer it signs
  without). Identity: `PROOF_SIGN_CERT_FILE` + `PROOF_SIGN_KEY_FILE` (PEM) or a self-made RSA-2048 certificate (ten
  years, `PROOF_SIGN_NAME`/`_ORGANIZATION`) kept in store `signing-identity`. PEM files that cannot be read or whose
  key does not belong to the certificate are named at start (`signingFilesProblem`, log `Proof signing files not
  usable`) and the self-made certificate signs instead. `signedProofPdf` falls back to an unsigned PDF only if signing
  itself fails. The tests check the CMS with OpenSSL (`openssl cms -verify`) and that a changed byte
  breaks it.
- **Guessing** (`proofGateway.ts`): five addresses that are no recipients for a link within an hour lock its code
  requests for an hour (same answer, nothing sent; `gateway.locked`) and mail the sender the attempts once
  (`sendLockNotice`). In memory, like the codes.
- **PDF** (`proofPdf.ts`, `pdf.ts` – a small writer of its own, no dependency, standard fonts in WinAnsi):
  the seal next to the file name in "Verlauf" (`.proof-seal`) asks `POST /api/proofs/:id/link` for an address signed with a key derived from
  `CREDENTIAL_KEY` and valid five minutes, and opens it in the default browser (`openBrowserWindow`) – the language
  of the add-in.

### 2.13 Sensitivity labels

Outlook carries the sensitivity label of an *attachment* over to the message; a file CloudShare shares as a link is no
longer attached, so the add-in does it itself (`addin/src/taskpane/labels.ts`, nothing goes to the server):

- **Reading the label** from the file, only small parts of it (`Blob.slice`): Office Open XML – the ZIP directory, then
  `docMetadata/LabelInfo.xml` (`<clbl:label id enabled siteId removed>`) or the `MSIP_Label_<id>_Enabled` / `_SiteId`
  properties in `docProps/custom.xml` (`DecompressionStream("deflate-raw")`, ZIP64 aware); compound files (encrypted
  Office files and doc/xls/ppt) – FAT, directory and mini stream, the stream `LabelInfo` that Office keeps unencrypted
  next to `EncryptedPackage`, else the `MSIP_Label_…` properties in `DocumentSummaryInformation`; PDF – the XMP
  `MSIP_Label_…` entries in the first and last 2 MB; `.pfile`, `.ppdf` … by their extension (protected, label
  unknown). Result `{ id, siteId, encrypted }`, cached per attachment / file (`labelOf`); attachments above 50 MB are
  read only when inserted.
- **Names and priority** first from the tenant: `GET /api/labels` (Graph, delegated `InformationProtectionPolicy.Read`,
  optional) gives each label's name in the add-in's language and Purview's own priority (`sensitivity` = rank); without
  the permission `Office.context.sensitivityLabelsCatalog` (Mailbox 1.13, its order taken as the rank). A label in
  neither (another organisation) shows as "Vertraulichkeitsbezeichnung" (`labelCatalog` in `labels.ts`).
- **The mail's label** (`item.sensitivityLabel`): after inserting, the highest-ranked label of the files is set if the
  mail's is lower – never lowered (`labelForMail`; `LABEL_POLICY=inherit`), or offered with "Übernehmen" (`suggest`).
- **Encrypted files** (`renderEncrypted`): as soon as a ticked file is known to be encrypted, a note appears right
  below the file list (`#encrypted-field`); `LABEL_ENCRYPTED=warn` – the insert button stays off until "Trotzdem als Link senden"
  is ticked (`insertReady`), `block` – it stays off, `allow` – no note. The note's "×" (`#encrypted-dismiss`) removes the
  encrypted files instead (attachments from the mail, files chosen in the pane from the list); the note follows the list (`renderAttachments` / `renderPicked` call `renderEncrypted`), so a
  file removed from the mail takes its note with it. Large attachments are read only when inserted;
  then the same check runs before any upload.
- No extra permission: the manifest already has `ReadWriteItem`.

### 2.14 Attachments inside and outside the organisation

Which attachments CloudShare steps in for, and how strictly, depends on the recipients (`attachmentRule` in
`addin/src/shared/office.ts`, used by the info bar, the send handler and the preselection in the pane):

- **External** – one recipient in To, Cc or Bcc outside `INTERNAL_DOMAINS` (comma-separated, subdomains included;
  empty = the domain of the user's own address): `ATTACHMENT_THRESHOLD_MB` and `ATTACHMENT_POLICY`. The size is the total
  of all file attachments of the mail, not of each one (`largeAttachments` in `commands.ts`: above it, all of them).
- **Internal** – every recipient inside, and also while none is entered yet: `ATTACHMENT_THRESHOLD_INTERNAL_MB` and
  `ATTACHMENT_POLICY_INTERNAL` (each unset = the external value). A distribution list counts by its own address –
  Outlook does not reveal its members in a draft. Recipients that cannot be read: external (the stricter case).
- **`0` = every attachment**: the info bar and the prompt name no size
  (`notify.any`, `alert.enforceAny`, `alert.suggestAny`).
- `OnMessageRecipientsChanged` (manifest 1.5.0.0, `onRecipientsChangedHandler` = the attachments handler) shows or
  removes the info bar when recipients change; the pane re-preselects only when the size that applies changes
  (`RecipientsChanged`, `shownThreshold`). The send handler checks with the recipients at that moment.

### 2.15 Feedback and problem reports

"Feedback senden" in the account menu (`#feedback-open`, only when `FEEDBACK_MAIL` is set – `/api/config` says
`feedback: true`, never the address) opens a page like the settings (`#panel-feedback`, `addin/src/taskpane/feedback.ts`):
the kind (problem / suggestion), the text and "Technische Angaben mitsenden" (ticked; the details are shown before
they go). `POST /api/feedback` (`server/src/feedback.ts`) mails it from the user's own mailbox to `FEEDBACK_MAIL` in
the layout of every CloudShare mail (red accent for a problem, blue for a suggestion) – a reply goes straight to the
user. Nothing goes anywhere public: no automatic GitHub issue (the repository is public, a report may hold personal
data); the mail asks whoever reads it to remove names, addresses, file names and subjects before making anything
public. The details are a fixed list (`DIAGNOSTICS`, anything else is dropped): Outlook host, version and platform,
the highest Mailbox requirement set, the languages of Outlook and CloudShare, the window size, the browser, the time
and the last five error messages of the pane (`noteError` in `setStatus`); the server adds the add-in version from
the manifest it serves. No mail content, no recipients, no logs.

## 3. State and data

| Where | What | Lifetime |
|---|---|---|
| Draft: `item.sessionData` (`DraftState`, `shared/office.ts`) | `jobId`, `separate` (passwords per step), `noPassword`, `passwordTo` / `passwordSms` (channel), `maxAccess`, `links` (tokens of inserted cards); legacy drafts: `password` | Until the draft is sent or closed |
| Draft: custom property `cloudshareJob` | Job ID; travels with the sent mail into "Sent Items" | With the mail |
| `/data/cloudshare.db`, store `credentials` | Nextcloud app password per `oid` | Until disconnect / 401 |
| … store `password-jobs` | Password jobs: owner, sender name and department, file names, Sends and links per step (`separate`), `passwordTo` / `passwordSms`, armed Graph token | 7 days after the last link |
| … store `cleanup` | Upload folders with expiry date per owner | Until deleted |
| … store `deliveries` | Per link: channel, masked recipient, SMS message ID and status | 120 days |
| … store `password-refs` | Per sent step: Send ID (never the password), link tokens, file names, subject, distribution of the sent mail (To, Cc, Bcc), other address / SMS number, sender | Until the links expire |
| … store `sent-mails` | Per link: subject of the mail it went out with and the Graph ID (immutable) of that mail in the sender's own "Sent Items" – the title in "History", a click opens the mail | Until the link expires |
| … store `messages` | Confidential messages: owner, Send ID, access ID (the link without its key), created, expiry – never key or text | Until the message expires |
| … store `notices` | Links waiting for their first download or for uploads (sender, subject, recipients, expiry, files already reported) | Until downloaded / expired |
| … store `settings` | Personal settings per `oid` (2.11) – only the fields the user changed | Until reset (`retention: false`) |
| Memory (`store.ts`) | Running uploads (12 h), pending Nextcloud logins (20 min) | Lost on restart → single instance only |
| `roamingSettings` | Last server configuration (for the handlers); the language from the settings (`cloudshare-settings`) | Per mailbox |
| `/app/public` (tmpfs) | Add-in files with address, client ID and texts of every language filled in, `taskpane.<language>.html` per language | Rebuilt at every start |
| `/config` (read-only, host `env/config/`) | Optional `templates.json` and own `locales/<language>.json` | Managed by the admin |

Everything in `/data` lives in one SQLite database, `cloudshare.db` (`db.ts`, Node's built-in `node:sqlite`, WAL, file
mode 600): one table `entries (store, key, data, updated)`, one row per entry; every row is sealed AES-256-GCM with
`CREDENTIAL_KEY` (`storage.ts`, AAD = store and key), so the file holds nothing readable. The modules keep their
entries in memory (`PersistentMap`) and write only the row that changed; a row that cannot be decrypted (e.g. the key
changed) is skipped with a warning. **Retention:** entries leave on their own (jobs 7 days after
their last link, password entries and notices when the links expire, delivery states after 120 days, cleanup entries
when the folder was deleted, connections on disconnect); as a safety net a daily pass deletes every row not written
for `DATA_RETENTION_DAYS` (default 365, `off` = never, at least `SHARE_MAX_EXPIRE_DAYS`) and then runs `VACUUM` –
except the connections and the cleanup list (`retention: false`). The upgrade test in CI checks that data written by the published image stays
readable by a new one; the switch from the former JSON files to SQLite took no data over (that one run is skipped).

## 4. HTTP API

All `/api` routes except `config`, `trace` and the well-known file require `Authorization: Bearer <Office SSO token>`.
Errors are `{ error, code? }` with a message in the configured language, shown in the taskpane.

| Route | Purpose |
|---|---|
| `GET /.well-known/microsoft-officeaddins-allowed.json` | Allows SSO/network for `commands.js` in classic Outlook ([5](#5-outlook-runtimes-and-pitfalls)) |
| `GET /api/config` | Public settings for the add-in (threshold, policy, limits, `trace`, `smsEnabled`, message limits) |
| `GET/POST /api/trace` | Diagnostics of the send handler, only with `CLIENT_TRACE=true` (no login) |
| `GET /api/profile` | Sender's Microsoft 365 profile (the server reads it itself for templates and signature) |
| `GET/PUT/DELETE /api/settings` | Personal settings (2.11): `{ settings, defaults, available }`; `PUT` changes single fields, `DELETE` resets |
| `GET /api/labels` | Sensitivity labels published to the user (Graph `/me/security/informationProtection/sensitivityLabels`, delegated `InformationProtectionPolicy.Read`): `{ labels: [{ id, name, rank, protection }] }` in the add-in's language, kept an hour per user; `labels: null` without the permission (2.13) |
| `POST /api/feedback` | "Feedback senden" (2.15): `{ kind: "bug"\|"idea", text, diagnostics? }` → a mail from the user's own mailbox (delegated `Mail.Send`) to `FEEDBACK_MAIL` in `LANGUAGE`; 404 `FEEDBACK_OFF` without it, 400 for an empty or too long text (5,000 characters); the text is not logged |
| `GET /api/contacts` | Suggestions for "Kennwort an" from the user's own Outlook contacts (Graph `/me/contacts`, delegated `Contacts.Read`, optional): `?kind=mail\|sms&q=…` → `{ contacts: [{ name, value, type? }] }` (`all=1`: every entry of that kind A to Z for the address book page, at most 2,000 – `listContacts`; otherwise at most 8; names and addresses, for SMS names and the digits of the numbers – `matchContacts` in `contacts.ts`); the contacts are kept in memory ten minutes per user, never stored, the query is not logged; `contacts: null` without the permission |
| `GET /api/nextcloud/status` | Connected?, storage (free/used/total), `passwordEnforced` |
| `POST /api/nextcloud/connect`, `…/connect/:id/poll`, `DELETE /api/nextcloud/connection` | Login Flow v2, disconnect |
| `POST /api/uploads`, `PUT …/:id/chunks/:n`, `POST …/:id/complete`, `DELETE …/:id` | Chunked upload, share link |
| `POST /api/requests` | Upload link ("Receive files from recipients") |
| `POST /api/messages` | Confidential message: text → password-protected Vaultwarden Send, joins the mail's password job; returns `expiresAt` (exact end, shown with the time in the card) |
| `GET /api/shares`, `DELETE /api/shares/:id`, `POST /api/shares/:id/link` | "History": list (all valid links and confidential messages – `msg-<Send ID>`, `kind: "message"`, deleted in Vaultwarden on revoke –, each with `subject` and `mailId` of its sent mail, newest first, max. 100; loaded with the tab, when the pane regains focus and by the refresh button in the tab bar), revoke, insert again; also the read mode (2.9) |
| `GET /api/shares/:id/password`, `POST /api/shares/:id/password/send` | The password of a sent link (owner only, until the link expires) and sending it again by SMS or password mail (2.10); held back for a delivery proof: 409 `PASSWORD_HELD` before the first confirmation, then to confirmed recipients only |
| `GET /api/shares/:id/mail` | "History": the sent mail of an entry as the EWS ID Outlook opens (`displayMessageFormAsync`) – the immutable Graph ID kept at sending, translated now with `translateExchangeIds` (User.Read), so it follows a moved mail; 404 if none is known or it is gone – also when it lies in Deleted Items or, after emptying them, in Recoverable Items (`sentMailState`, Mail.Read), where Outlook would still open it |
| `PATCH /api/jobs/:id` | Change `maxAccess` (updates the Send), `passwordTo` or `passwordSms` until the password went out; `subject`: the draft's subject changed (the pane checks it every 4 s while open) – "History" shows it for the mail's links until the sent mail is found (`updateDraftSubject`) |
| `POST /api/jobs/:id/arm` | Arm the password mail (send handler); also reads the sender's name and department for the templates |
| `GET /api/proofs` | "Zustellnachweise" (account menu): every case of the user whose mail was sent (drafts never sent are left out), newest first, also finished ones – `{ id, createdAt, subject, sentAt, recipients, files, downloads, confirmed, done }`; the pane searches in it |
| `GET /api/proofs/:id/mail` | "Zustellnachweise": the sent mail of a case as EWS ID (click on the subject, like `GET /api/shares/:id/mail`; shared lookup `sentMailOf`) |
| `POST /api/proofs/:id/link`, `GET /proof/:id.pdf?until=…&lang=…&sig=…` | "Zustellnachweis" (2.12): a signed PDF address valid 5 minutes (owner only), opened in the default browser; the PDF itself needs no login – the HMAC signature (key from `CREDENTIAL_KEY`) binds case, end and language |
| `GET /r/s/:token`, `POST /r/s/:token/code`, `POST /r/s/:token/verify` | Gateway of a link with "Zustellnachweis" (no login): the recipient enters their address, gets a six-digit code (only addresses the mail went to; the page reads the same for any address), confirms it and goes on to Nextcloud; a signed cookie per link lets the same browser come back; every step is recorded |
| `GET /r/s/:token/confirm?address=…&code=…` | The button "Adresse bestätigen" in the code mail: the code form filled in (`confirmFromMail`) – confirms nothing by itself (mail scanners open links), the click posts to `/verify` |
| `GET/POST /webhooks/seven?token=…` | seven.io delivery reports (no login, secret `SEVEN_WEBHOOK_SECRET`) |

Rate limit: `RATE_LIMIT_PER_MINUTE` per user (`oid`, otherwise client address); Nextcloud login, upload links and
"insert again" get 1/60 of it.

## 5. Outlook runtimes and pitfalls

The same add-in runs in two very different environments; most hard-won fixes are about this.

**Classic Outlook on Windows** runs the event handlers in a **JavaScript-only runtime** (no DOM, older engine):

- `commands.js` is bundled as IIFE with target ES2016 and without `Headers`/`Promise.finally`
  (`addin/test/bundle.test.ts` guards this).
- SSO and network requests are only allowed for scripts listed in
  `/.well-known/microsoft-officeaddins-allowed.json` – without it the send dialog hangs at "is taking longer
  than expected".
- Office.js is preloaded, the handlers can be registered right away.
- Outlook caches the script; after an update or a new language clear `%LOCALAPPDATA%\Microsoft\Office\16.0\Wef\`.
- The taskpane runs in WebView2, where the dialog of a file input did not reliably open (script click, `<label for>`
  and a transparent overlay were tried) → the button uses the File System Access API (`window.showOpenFilePicker`)
  first and the file input only as fallback; with `CLIENT_TRACE` every step is logged (`picker:*`). Drag & drop is
  the alternative that needs no dialog at all.

**New Outlook and Outlook on the web** run the handlers in a browser frame (`commands.html`):

- office.js loads the Outlook API (`Office.actions`) only **after** `commands.js` ran. The handlers are therefore
  registered immediately *and* again in `Office.onReady` (`commands-load.test.ts`) – registering only once left the
  send dialog spinning forever.

**Smart Alert limits:** the button text (`cancelLabel`) may have at most 20 characters and the message 500; longer
values make Outlook on the web reject the result silently (`withinLimits()`). Custom buttons and "Send anyway" need
Mailbox 1.14; with 1.12 users are pointed to the ribbon button. The "Send anyway" button itself is Outlook's and
follows Outlook's display language; the message (`alert.suggest`) names it as the language file says.

**Office SSO:** only one `getAccessToken` call may run at a time (13008), so the token is cached and concurrent
callers share one request. Outlook also throttles the calls per add-in (13013, "too many sign-ins"): the token is
kept in `sessionStorage` (this Outlook window only) so a pane loaded anew for every selected mail reuses it, a 13013
is retried after a few seconds, and a token the server refuses (401) is dropped; SSO errors (plain objects with a code) are shown with a hint (language file, `sso.<code>`) and the code
(`shared/sso.ts`). Every step of the send handler has a time limit, so a hanging step ends with a message instead of
"is taking longer than expected".

**Diagnostics:** with `CLIENT_TRACE=true` the send handler reports each step (`start` … `arm:token` … `completed`,
milliseconds, Outlook platform, frame origin, error text) as plain GET requests to `/api/trace`, logged as "Smart
Alert trace". Off by default.

## 6. Security

- **Authentication:** every API call carries an Office SSO token, validated against Entra's signing keys (audience
  `api://<host>/<client-id>`, scope `access_as_user`). Users can only see and change their own uploads, jobs and links
  (tested in `api.test.ts`).
- **Nextcloud app passwords** have full access to the user's Nextcloud account; they are encrypted at rest with
  `CREDENTIAL_KEY` and only used server side.
- **Passwords:** random passwords are pronounceable (`KateGeteFi45`, `47RantoBeKiwu`, `Sumo52DarniPe`: two words of
  two syllables, a short syllable and two digits in random order, 12–15 characters, at least ≈ 42 bits from the
  browser's CSPRNG). That is low for offline attacks but sufficient here: an attacker needs the link
  first, Nextcloud throttles failed attempts on public links, and links expire. Own passwords need ≥ 12 characters and
  three character classes (checked in the taskpane and again by the server). The password exists in readable form
  only in the Vaultwarden Send and briefly in the taskpane. The recipients can open the Send a limited number of
  times for a few days (5 by default); the Send itself lives until the links expire (at most 30 days), so its
  sender can see and resend the password in "History" (2.10) – the server reads it with the service account and
  returns it only to the link's owner. CloudShare stores only the Send ID.
- **Graph:** delegated permissions for everything the user triggers; the armed token is stored encrypted for at most
  15 minutes of use, and the `.Shared` permissions reach only mailboxes the user already has rights to. Only the
  optional download confirmation uses an **application** permission (`Mail.Send`) – restrict it to the CloudShare
  users ([microsoft-365.md → 5](microsoft-365.md#5-optional-application-permission-for-notices-and-the-delivery-proof)), otherwise the app could send as any mailbox of the tenant.
- **Confidential messages** pass the server only in memory; Vaultwarden stores them end-to-end encrypted in the Send
  (the Bitwarden CLI encrypts before upload) and deletes them at expiry. Anyone with link and password can read them
  until then – like a file link.
- **SMS:** the API key and the webhook secret stay on the server; the add-in never talks to seven.io. The SMS carries
  the password in plain text (it is the second channel) and, with `{files}` (the built-in text), the first file name –
  shortened to 32 characters, GSM characters only, "+ N more" – so a recipient of several SMS can tell which password
  belongs to which link; never the link itself. The password is read from the Send only for sending. A
  failed SMS never falls back to another channel on its own. Numbers are stored masked in the status (the full
  number only in the encrypted entry for sending again); `SMS_MAX_PER_DAY` limits costs and abuse, also for SMS sent
  again. Conclude a data processing agreement with seven.io.
- **Language files** are data, not markup: every text is escaped where it is inserted (HTML, XML, JavaScript string),
  so a language file cannot add script or markup to the pane, the manifest or a mail. `LANGUAGE` must be a language
  code (no paths).
- **Container:** read-only, `cap_drop: ALL`, `no-new-privileges`, user 1000, no published ports; `bw serve` listens on
  127.0.0.1 only. Networks: `internal` (Caddy ↔ CloudShare, no outbound), `edge` (Caddy), `egress` (CloudShare).
- **Supply chain:** Trivy scan of every image, `npm audit`, Dependabot, CodeQL, secret scanning; GitHub Actions
  pinned to commit SHAs.

## 7. Setup

### Microsoft 365 (Entra ID, add-in deployment, Exchange Online)

Step by step, with checks, maintenance and error table: **[docs/microsoft-365.md](microsoft-365.md)**. In short:

- App registration "CloudShare", single tenant, access token version 2.
- Application ID URI `api://cloudshare.example.com/<client-id>` (must match the manifest's `<Resource>`), scope
  `access_as_user` (admins and users), authorized client `ea5a67f6-b6f3-4338-b240-c655ddc3cc8e` (Microsoft Office).
- Delegated Graph permissions `openid`, `profile`, `offline_access`, `User.Read`, `Mail.Read`, `Mail.Read.Shared`,
  `Mail.Send`, `Mail.Send.Shared` with **admin consent**, optionally `Contacts.Read` (suggestions under "Kennwort an")
  and `InformationProtectionPolicy.Read` (label names; all of
  them also listed in the manifest's `<WebApplicationInfo><Scopes>`); client secret → `ENTRA_CLIENT_SECRET` (renew before expiry).
- Add-in: admin center → *Integrated apps* → manifest URL `https://cloudshare.example.com/manifest.xml`.
- Only for the download confirmation: application permission `Mail.Send`, **restricted** to the CloudShare users with
  Exchange Online RBAC for applications.

### seven.io (SMS, optional)

- Account with prepaid credit; API key → `SEVEN_API_KEY`, sender (number or up to 11 characters) → `SEVEN_SENDER`.
- Delivery reports: callback URL `https://cloudshare.example.com/webhooks/seven?token=<SEVEN_WEBHOOK_SECRET>` in the
  dashboard (webhook event "dlr").
- Step by step with checks: README → *Password by SMS*.
- Conclude a data processing agreement and document it in the record of processing activities.

### Nextcloud

- No change to the login method (`user_oidc`, LDAP, local) needed. Public links must be allowed; for upload links also
  public upload. If *Enforce password protection* is on, CloudShare hides "No password" (exempt a group there to
  allow it for some users).
- Reached through its public address (`NEXTCLOUD_URL`). Behind a proxy, the request body limit must be ≥ 12 MB, and
  the login flow (`/index.php/login/v2…`) must be reachable.
- Recommended: allow the CloudShare server's address in the brute-force settings, since all requests come from it.

### Vaultwarden

- `SENDS_ALLOWED=true`.
- Service account (with SSO: an IdP user with a verified email, invited in Vaultwarden, master password set at first
  login), API key under *Settings → Security → Keys* → `BW_CLIENTID` / `BW_CLIENTSECRET`, master password →
  `BW_PASSWORD`.

## 8. User interface

- The taskpane looks like part of Outlook: Fluent 2 design tokens (32 px controls, 4 px radius, Segoe UI) in
  `taskpane.css`, no own header, light/dark theme from `Office.context.officeTheme`, Outlook's wording in every
  language ("Nachricht"/"message", "Kennwort"/"password" – translations should use the terms of Outlook in that
  language). Accessibility is tested with axe (WCAG 2 A/AA) in CI.
- Layout and labels must survive longer translations: buttons and labels wrap or are short, the full text is in
  `title`/`aria-label`.
- **Account menu** (`#account-toggle` at the bottom left – the only thing in the bar): its head `.account-menu-head`
  shows the whole account name and the storage – a bar `#storage-meter` with the used share of the quota (red from
  90 %, hidden without a quota) and the text (`storageText`: "37,5 GB von 50 GB belegt" like OneDrive – matching the bar –, or what Nextcloud knows); then the
  entries below.
- **Open Nextcloud** ("Nextcloud öffnen", first entry of the account menu, `#nextcloud-open`): opens `NEXTCLOUD_URL`
  (sent in `/api/config` as `nextcloudUrl`) in the default browser (`openBrowserWindow`, in the browser a new tab) – the
  browser's own Nextcloud session or a sign-in there; the app password of the connection never leaves the server.
- **History** ("Verlauf", `#shares-open`, right above "Zustellnachweise"): a page like the settings (`#panel-shares`,
  `showOverlay("shares")`) – every valid link with its actions, reloaded when opened, by `#refresh-shares` in its head
  and on return to the pane; the search `#shares-search` filters by subject, name, recipients, password address and
  tracking ID (`renderShareList`). The pane has no tabs any more: the view is "Senden" (or, reading, the links of that
  mail – `currentView`), every other thing is a page of the account menu.
- **Proofs of delivery** ("Zustellnachweise", `#proofs-open`, only with `DELIVERY_PROOF`): a page like the settings
  (`#panel-proofs`); loads `GET /api/proofs` (every case of the user whose mail was sent, finished ones too) and
  every entry is built like a row of "Verlauf" (`proofRow`: kind symbol, subject as title – a click opens the sent mail
  via `GET /api/proofs/:id/mail` when `mailAvailable` –, the file with its seal (a click opens the PDF),
  to whom, the tracking ID in bold, then the state with its dot; no action icons – the list is for looking things up);
  filters it while typing over tracking ID, subject, recipients and file names (no other filters); the seal of an entry (`proofSeal`, also in
  "Verlauf" via `proofState` of `/api/shares`) is coloured by `caseSummary(c).state` – grey sent, blue confirmed, green
  fetched – and opens its signed PDF; the sent line says "Nachricht und Kennwort an …" when the password went to exactly
  the recipients (`sameAddresses`, the password line hidden), else "Nachricht an …" and "Kennwort an …"; each entry has the
  line "Gültig bis …" (`endsOn`, `fillValidLine` – a calendar,
  orange from two days before the end; the same line in "Verlauf") or "Aufbewahrt bis …" (`keptUntil`), and the state
  alone on the line below.
- **Help** ("Hilfe" in the account menu of the bottom bar, `#help-toggle`; `#panel-help` in `taskpane.html`, shown in place of the
  view like the history, the settings and the proofs (`showOverlay`, `currentOverlay`), "←" before the title (`.page-back`) or Escape leads back): a user guide for everyday users – short `<details>`
  sections (getting started, sending files, password and channel, confidential message, receiving files, "History",
  FAQ), the most common one open. Texts in the language files (`help.*`); admins put their own support contact into
  `help.support` (`/config/locales`). Plain words, the button names of the pane, no technical terms.
- **Messages** (`#status`, `setStatus`): progress (kind `progress`) and a failed start (with "Erneut versuchen") stay until the next message; all other messages – success, error, grey notes – disappear
  after `SUCCESS_MESSAGE_SECONDS` (default 5, `0` = they stay, served in `/api/config`) – the ring around the tick, "!" or "i" runs
  down clockwise like an OTP timer (CSS animation `status-timer`, the message goes on `animationend`) and stops while
  the pointer or the focus is on the message. A failed start (sign-in, server unreachable) offers "Erneut versuchen" in the
  message (`retryAction`): CloudShare loads anew, as if the pane was opened again.
- Outlook sets the width of the pane; the layout is designed for 320 px (short labels in "History", full text
  in `title`/`aria-label`).
- **Cards** (`mailCard.ts`, inserted by the taskpane) and **every mail CloudShare sends itself** share one look. The
  mails are built by `mailLayout.ts` (`mailDocument`, `mailCard`): text paragraphs, one white card with a coloured
  accent (blue `brand` – password mail, green `success` – download confirmation, red `danger` – SMS not sent / not
  delivered), title, file rows with type badges, at most one button, a grey detail line, then the line "sent
  automatically by CloudShare" (`mail.automatic`). Users of it: `passwordMail.ts` (password mail incl. "send again",
  `smsProblemMailHtml`) and `notices.ts` (download confirmation: the template text above a card with file, time and
  validity). Tables and inline styles only, so they render in classic Outlook, Outlook on the web, Gmail and on phones.
  The link cards share one layout for files, confidential message and upload link: title, a short sentence, the
  button on the left, the details in grey below it (files: one block per file – button, then type badge, name, size).
  Buttons are filled table cells (`bgcolor`, `mso-padding-alt`) in both files, so classic Outlook (Word engine) shows
  them filled instead of blue text in a white box.
  The colours are duplicated in `mailCard.ts` and `mailLayout.ts`; keep them in sync. Plain-text mails get a text
  version of the link card. The password mail ends with the sender's signature (2.11; `MAIL_FOOTER` filled from the
  profile unless the user set an own one); the link card carries none.
- Icons (`addin/assets/icon-*.png`): the rings of the Nextcloud logo, taken unchanged from
  its SVG path, with an upload arrow in the hole of the middle ring. Rendered by `addin/scripts/render-icons.ps1`
  (Windows PowerShell, System.Drawing; drawn at 8x and scaled down). Icon URLs in the
  manifest carry a version parameter because Outlook caches icons by URL – change it together with the icons.

## 9. Limits and risks

1. **Single instance:** running uploads and pending logins live in memory; a restart discards them (users start
   again). Everything in `/data` survives.
2. **Password mail needs "Sent Items":** with saving of sent mail switched off, no password mail is sent. Files always
   go to the sender's own Nextcloud, also when sending from a shared mailbox.
3. **Outlook's attachment limit:** attachments already in the mail are limited by Outlook/Exchange (≈ 35 MB) and read
   into memory; larger files come from the computer via the file picker or drag & drop.
4. **Requirement sets:** Smart Alerts need Mailbox 1.12, custom buttons and "Send anyway" 1.14. Under `enforce`, a
   CloudShare outage means large attachments cannot be sent until it is back.
5. **Bitwarden CLI and Vaultwarden versions must match** (`BW_CLI_VERSION` in the `Dockerfile`).
6. **Trademark:** the icon uses the Nextcloud mark – fine for internal use; check Nextcloud's trademark guidelines
   before distributing CloudShare to third parties.
7. **Languages:** the add-in follows each user's Outlook language, what the recipients get the language chosen per
   mail, the notes to the sender the language of their add-in (2.8). Only mails sent before this version fall back
   to `LANGUAGE`, and a regional variant not listed in `_meta.regions` is left to Outlook (the pane maps it to its
   language anyway).
8. **Possible later:** OIDC token exchange (e.g. Keycloak) instead of app passwords would remove the connect step –
   not built on purpose: Keycloak supports exchanging an external (Entra) token only in its legacy preview feature, and
   the background tasks would still need a stored offline token, so it would add risk without a security gain.
