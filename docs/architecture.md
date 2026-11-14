# CloudShare – architecture

CloudShare is an Outlook add-in plus a small backend. Large files are uploaded to the **sender's own Nextcloud** and
inserted into the mail as a password-protected link; the password reaches the recipients in a **second mail** as a
Vaultwarden Send, and only after the first mail has really been sent. This document describes how the parts fit
together, where state lives, and the Outlook peculiarities the code has to work around. Operation and installation:
[README](../README.md); Microsoft 365 (Entra ID, deployment, Exchange): [microsoft-365.md](microsoft-365.md).

Everything visible to users and recipients (taskpane, Outlook hints, cards, password mail, SMS, download
confirmation) comes from the language file picked by `LANGUAGE` (`locales/de.json` by default, see
[2.8](#28-languages)); code, logs and documentation are English. Labels are quoted as in the English interface
("Insert as link", "My shares" …); the German texts (the default) are in `locales/de.json` under the same keys.

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
  │                                            "My shares"
  └─ Event handlers      addin/src/commands/   OnMessageAttachmentsChanged → info bar
                                               OnMessageSend (Smart Alert) → attachment policy, arm password mail
        │  HTTPS + Office SSO token (Entra ID, audience api://<host>/<client-id>, scope access_as_user)
        ▼
Caddy (Let's Encrypt, ports 80/443) ──► network "internal"
        ▼
Container "cloudshare" (Node 22, read-only)
  ├─ Fastify backend     server/src/           API, add-in files (texts of LANGUAGE filled in at start),
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
| Taskpane | `taskpane.ts/.html/.css`, `mailCard.ts`, `footer.ts` | Plain TypeScript, bundled as IIFE (esbuild, ES2017) |
| Event handlers | `commands.ts` → `commands.js` | ES2016 bundle for classic Outlook's JavaScript-only runtime |
| Shared add-in code | `shared/office.ts`, `shared/password.ts`, `shared/sso.ts`, `shared/i18n.ts` | Office.js helpers, draft state, SSO token cache, password rules, `t()` for the texts |
| Languages | `locales/<language>.json`, `server/src/i18n.ts` | Texts of every language; loading, fallback, `t()`, injection into the add-in files ([2.8](#28-languages)) |
| API | `app.ts` (routes), `index.ts` (start: languages, vault unlock, background jobs) | `buildApp()` is used by the API tests |
| Integrations | `auth.ts` (SSO validation, OBO, app token), `nextcloud.ts`, `graph.ts`, `vaultwarden.ts`, `sms.ts` (seven.io) | |
| Texts sent on the sender's behalf | `passwordMail.ts`, `templates.ts` (`templates.json`) | Password mail, SMS, download confirmation |
| Background jobs | `jobs.ts` (password mails and SMS), `cleanup.ts` (expired uploads), `notices.ts` (download confirmation) | Every 20 s / hourly / every 10 min, see below |
| Persistence | `credentials.ts`, `storage.ts`, `deliveries.ts` | AES-256-GCM files in `/data` |

**Why app passwords for Nextcloud?** Nextcloud usually cannot accept Entra ID tokens (it often signs users in through
another IdP, e.g. Keycloak via `user_oidc`). Each user therefore connects once with Nextcloud's Login Flow v2; the
resulting app password is stored encrypted and used as Basic auth. Nextcloud is only ever called server to server,
so it needs no CORS.

**Placeholders.** `docker/entrypoint.sh` copies the add-in files from `/app/public.tpl` to `/app/public` (tmpfs) and
replaces `%%PUBLIC_BASE_URL%%`, `%%PUBLIC_HOST%%` and `%%ENTRA_CLIENT_ID%%`; then the server fills in the texts of
`LANGUAGE` ([2.8](#28-languages)). So one image serves every installation and language.

## 2. Main flows

### 2.1 Connect Nextcloud (once per user)

1. Taskpane → "Connect" → `POST /api/nextcloud/connect` starts Login Flow v2 and returns the login URL, which the
   taskpane opens in the browser (usually SSO, no typing).
2. The taskpane polls `POST /api/nextcloud/connect/:id/poll` until the user granted access.
3. The backend stores the app password encrypted under the user's Entra object ID (`oid`). In Nextcloud it appears as
   "CloudShare Outlook" under *Settings → Security* and can be revoked there; a 401 from Nextcloud deletes the stored
   credential and the taskpane shows the connect view again (`NEXTCLOUD_NOT_CONNECTED`).

### 2.2 Share files ("Insert as link")

1. The taskpane lists the mail's file attachments (those above `ATTACHMENT_THRESHOLD_MB` preselected) and files from
   the computer – dropped onto the drop zone (the whole "Send" tab accepts drops) or picked by clicking it (native dialog `window.showOpenFilePicker`, fallback: hidden `<input type="file">`).
2. **One click = one step** with a random `step` ID. All files of the step are uploaded one by one:
   `POST /api/uploads` (quota check via OCS `cloud/user`, 507 → `NEXTCLOUD_QUOTA_EXCEEDED`) →
   `PUT /api/uploads/:id/chunks/:n` (10 MiB chunks, Nextcloud chunked upload v2) → `POST /api/uploads/:id/complete`.
   Target: `NEXTCLOUD_FOLDER/<date>_<id>/<file>`. Attachments are removed from the mail only after a successful upload.
3. `complete` creates one public link per file (password, expiry date) and attaches it to the mail's **password job**
   (`attachToJob`): the first file of a step creates a Vaultwarden Send with the step's password, further files of the
   same step join it. The job ID and its parts are described in [3](#3-state-and-data).
4. The taskpane inserts a card (`mailCard.ts`) at the cursor – falling back to the top of the mail when Outlook
   inserted nothing (new Outlook without a cursor in the body) – followed by an empty paragraph, so the next card does
   not land inside this one. It tags the draft with the custom property `cloudshareJob` (job ID) and records the link
   token in the draft state.
5. After the step the taskpane offers a new random password (or clears an own one) for the next step.

Variants: **"Insert upload link"** (`POST /api/requests`) creates a folder in `NEXTCLOUD_REQUEST_FOLDER` with
a create-only link ("file drop") – same password rules, its own step. **"No password"** (`noPassword: true`, file
links only) creates a link without password and Send; with `DOWNLOAD_NOTICE=true` it joins a job **without password**
(`createNoticeJob`, `PasswordJob.noPassword`, the links in `noticeShares`), which the send handler arms like any job –
the watcher sends nothing and only registers the links for the download confirmation once the mail is in "Sent
Items". Without the confirmation there is no job (`jobId: null`). The answer always says `noPassword: true`. **"My shares → Insert"**
(`POST /api/shares/:id/link`) creates a new link to an existing upload with a new password, as a new step.

**"Confidential message"** (`POST /api/messages`): a text instead of a file. It is never written to Nextcloud or to
`/data`: `createMessageSend` stores it as a Vaultwarden **text Send protected with the step's password** (Send access
password, `hidden: false`, no access limit, deleted at expiry). The Send's link goes into the mail as a card
("Open message"); the step's password goes out like for files – the job item is "Confidential message", the
password mail shows an envelope badge. Always with a password; at most 1000 characters (Bitwarden's limit for text
Sends). Its lifetime is chosen separately ("Message valid for", `messageDays`): `SEND_EXPIRE_DAYS` (default 5) up to
`SEND_MAX_EXPIRE_DAYS` (at most 30 – Vaultwarden refuses deletion dates more than 31 days ahead).

**"Password link valid for"** (`passwordDays`, same limits): lifetime of the step's password Send. It is created with
the first link of the step, so it counts from inserting, not from sending; the date is stored with the step
(`sendExpiresAt`) and named in the password mail ("expires on …"). After it the password can no longer be fetched,
even if the download link is still valid – recipients who opened it in time keep the password. Messages are not listed in "My shares" (they are no Nextcloud shares) and get no
download confirmation.

### 2.3 Send and password mail

1. On "Send" Outlook runs the Smart Alert `onMessageSendHandler` (manifest `SendMode="SoftBlock"`):
   - Large attachments left in the mail: under `ATTACHMENT_POLICY=enforce` sending is blocked with the button
     "Open CloudShare"; under `suggest` the user is asked once and may choose "Send Anyway"
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
   for further links. Jobs expire after 7 days.

If the add-in cannot run at all, SoftBlock lets the mail through – otherwise a CloudShare outage would stop all mail.
The handlers cache the last configuration in `roamingSettings`, so the attachment policy also applies while the
server is unreachable. Outlook mobile and other mail clients do not run the add-in and are not checked.

### 2.4 Cleanup of uploads

Every file link registers its upload folder with the link's expiry date in `/data/cleanup.json` (the latest link
wins) – Nextcloud's housekeeping deletes expired shares, so the date could not be read back later. An hourly pass
(`cleanup.ts`) deletes due folders via WebDAV `DELETE` with the owner's app password, i.e. into the owner's Nextcloud
**trash**:

- due = day after the expiry date + `CLEANUP_AFTER_DAYS` (default `0`; `off` disables the cleanup);
- **"Revoke"** counts as expiry today – with `0` the folder goes to the trash at once;
- folders that still have a valid link are postponed; users who disconnected are skipped;
- only folders matching `NEXTCLOUD_FOLDER/YYYY-MM-DD_xxxxxxxx` are touched – received files
  (`NEXTCLOUD_REQUEST_FOLDER`) are never deleted. Uploads from before the feature existed are not registered.

### 2.5 Password channel: recipients, another address or SMS

"Deliver the password" in the taskpane decides per mail where the password goes (`DraftState.passwordTo` /
`passwordSms`, stored on the job via `PATCH /api/jobs/:id` and passed again with `/arm`):

- **to the recipients of the message** (default) – the password mail to the mail's actual recipients;
- **to another email address** – the password mail only to that address;
- **per SMS** (only with `SEVEN_API_KEY`) – one SMS per step via seven.io (`sms.ts`), sent by the watcher after the
  mail appeared in "Sent Items", like the password mail. The number is entered in the pane and normalised to E.164
  (`0170 …` → `+49170…`, `SMS_DEFAULT_COUNTRY`). The password exists in readable form only in the step's Vaultwarden
  Send; the watcher reads it from there right before sending (`readSendPassword`) and deletes the Send afterwards.
  So the password-link settings do not apply (the pane hides "Password can be opened" and "Password link valid for"):
  the Send of an SMS step always gets the maximum lifetime (`SEND_MAX_EXPIRE_DAYS`), also when SMS is chosen after
  links were inserted (`PATCH` extends the Sends of steps not sent yet), so a draft can wait that long. The password
  in the SMS works as long as the link.
  If seven.io refuses the SMS (or `SMS_MAX_PER_DAY` is reached), the password mail goes **to the sender** instead –
  never automatically to the recipients by another channel. Delivery reports arrive at `/webhooks/seven` (secret in
  the query, compared in constant time, idempotent, webhook and classic form callback) and update `deliveries.ts`.

How the password of each link went out (channel, masked recipient, SMS status) is shown in "My shares".

### 2.6 Text templates

`templates.ts` reads the optional `TEMPLATES_FILE` (`/config/templates.json`, mounted from `env/config/`): fields
`sms_template`, `mail_password_subject`, `mail_password_intro`, `download_notice_subject`, `download_notice_body`, each
with a fixed set of placeholders. Resolution per field: `users[<sender UPN>]` → `groups[<department>]` → `default` →
built-in texts of the language (section `templates` of the language file); the department and display name are read
from the sender's profile when arming. `templates.json` holds one set of texts for all languages – after a change of
`LANGUAGE` translate it as well. At start the file is validated: `{password}` outside the SMS, an SMS without `{password}` or with the link / a file name drop the
field (built-in text instead); unknown placeholders, non-GSM characters and SMS above 160 characters are warnings.
A broken file never stops CloudShare. Example: `templates.example.json`.

### 2.7 Download confirmation

With `DOWNLOAD_NOTICE=true`, every file link of a sent mail is registered – with password or, via the job without
password of 2.2, without – with the sender, subject, recipients and
expiry (`notices.ts`, `/data/download-notices.json`). Every 10 minutes the Activity API is asked with the owner's
app password; only downloads **after** the password went out count (older links of the same file are ignored). On the
first download the sender gets a mail (`download_notice_*` templates) – sent with CloudShare's **own app token**
(client credentials, application permission `Mail.Send`), because the sender's delegated token has long expired by
then. Entries are dropped after the download or when the link expired. No IP addresses or data about the person who
downloaded are kept.

### 2.8 Languages

`LANGUAGE` (default `de`; `de` and `en` are included) picks `locales/<language>.json`. `i18n.ts` merges, per text:

1. English – fallback for texts a language does not have yet;
2. the built-in file of the language (`/app/locales` in the image);
3. an optional own file from `CUSTOM_LOCALES_DIR` (`/config/locales`, mounted from `env/config/locales/`) – adds a
   language or changes single texts of a built-in one without a new image.

An unknown language (or one that is not a code like `de`, `en`, `pt-br`) stops the start with the list of available
languages. A language file:

| Part | Content |
|---|---|
| `_meta` | `name`, `locale` (BCP 47, e.g. `de-DE`: date and number format, manifest `DefaultLocale`) |
| `addin` | taskpane, info bar, Smart Alerts, link cards, SSO hints (`sso.<code>`) |
| `manifest` | description, ribbon group, tooltip |
| `server` | error messages the taskpane shows (`HttpError`), share label of upload links |
| `mail` | password mail (greeting, card, notes), subjects of fallback mails |
| `templates` | built-in texts of [2.6](#26-text-templates) (SMS, password mail subject/intro, download confirmation) |

Keys are flat within a section (`"files.added_one"`); placeholders are `{name}`; `_one` / `_other` are the singular
and plural form – `t("files.added", { count })` picks one. Texts never contain markup: HTML and XML get them escaped,
the bundles as an ASCII-only JSON string.

The add-in gets its texts without an extra request: at startup (`index.ts`, `localizePublic`) the server replaces in
the copies under `/app/public` the string marker `"%%CLOUDSHARE_I18N%%"` in `taskpane.js` / `commands.js` with a JSON
string of the `addin` texts (read by `addin/src/shared/i18n.ts`, ASCII-escaped), and `%%t:<section>.<key>%%`,
`%%LANG%%`, `%%LOCALE%%` in `taskpane.html`, `commands.html` and `manifest.xml`. So the send handler in classic
Outlook's JavaScript-only runtime has the texts too. Comments in these files must not contain `%%` – the smoke test
fails on any marker left over.

One language per installation, the same for all users; Nextcloud folder names do not depend on it
(`NEXTCLOUD_REQUEST_FOLDER` identifies CloudShare's links). Outlook caches the add-in files – after a change of
`LANGUAGE` clear the cache ([5](#5-outlook-runtimes-and-pitfalls)).

**Adding a text:** put the key into **every** file in `locales/` with the same placeholders. `server/test/i18n.test.ts`
checks that every language has the keys and placeholders of English, that every key the code uses (`t("…")`,
`%%t:…%%`) exists, Outlook's length limits (Smart Alert button ≤ 20 characters, manifest strings) and that the SMS text
is GSM with `{password}`. Add-in unit tests load `locales/de.json` (`addin/test/setup.ts`); the UI tests serve the
bundles with the German texts (`addin/scripts/localize.mjs`); the manifest is validated once per language; the smoke
test checks the German texts in the running stack.

## 3. State and data

| Where | What | Lifetime |
|---|---|---|
| Draft: `item.sessionData` (`DraftState`, `shared/office.ts`) | `jobId`, `separate` (passwords per step), `noPassword`, `passwordTo` / `passwordSms` (channel), `maxAccess`, `links` (tokens of inserted cards); legacy drafts: `password` | Until the draft is sent or closed |
| Draft: custom property `cloudshareJob` | Job ID; travels with the sent mail into "Sent Items" | With the mail |
| `/data/nextcloud-credentials.json` | Nextcloud app password per `oid` | Until disconnect / 401 |
| `/data/password-jobs.json` | Password jobs: owner, sender name and department, file names, Sends and links per step (`separate`), `passwordTo` / `passwordSms`, armed Graph token | 7 days |
| `/data/cleanup.json` | Upload folders with expiry date per owner | Until deleted |
| `/data/deliveries.json` | Per link: channel, masked recipient, SMS message ID and status | 120 days |
| `/data/download-notices.json` | Links waiting for their first download (sender, subject, recipients, expiry) | Until downloaded / expired |
| Memory (`store.ts`) | Running uploads (12 h), pending Nextcloud logins (20 min) | Lost on restart → single instance only |
| `roamingSettings` | Last server configuration (for the handlers) | Per mailbox |
| `/app/public` (tmpfs) | Add-in files with address, client ID and texts of `LANGUAGE` filled in | Rebuilt at every start |
| `/config` (read-only, host `env/config/`) | Optional `templates.json` and own `locales/<language>.json` | Managed by the admin |

All files in `/data` are AES-256-GCM encrypted with `CREDENTIAL_KEY` (`storage.ts`, atomic writes). The upgrade test
in CI checks that data written by the published image stays readable by a new one.

## 4. HTTP API

All `/api` routes except `config`, `trace` and the well-known file require `Authorization: Bearer <Office SSO token>`.
Errors are `{ error, code? }` with a message in the configured language, shown in the taskpane.

| Route | Purpose |
|---|---|
| `GET /.well-known/microsoft-officeaddins-allowed.json` | Allows SSO/network for `commands.js` in classic Outlook ([5](#5-outlook-runtimes-and-pitfalls)) |
| `GET /api/config` | Public settings for the add-in (threshold, policy, limits, footer, `trace`, `smsEnabled`, message limits) |
| `GET/POST /api/trace` | Diagnostics of the send handler, only with `CLIENT_TRACE=true` (no login) |
| `GET /api/profile` | Sender's Microsoft 365 profile for `MAIL_FOOTER` |
| `GET /api/nextcloud/status` | Connected?, storage (free/used/total), `passwordEnforced` |
| `POST /api/nextcloud/connect`, `…/connect/:id/poll`, `DELETE /api/nextcloud/connection` | Login Flow v2, disconnect |
| `POST /api/uploads`, `PUT …/:id/chunks/:n`, `POST …/:id/complete`, `DELETE …/:id` | Chunked upload, share link |
| `POST /api/requests` | Upload link ("Receive files from recipients") |
| `POST /api/messages` | Confidential message: text → password-protected Vaultwarden Send, joins the mail's password job |
| `GET /api/shares`, `DELETE /api/shares/:id`, `POST /api/shares/:id/link` | "My shares": list (30 days, max. 100), revoke, insert again |
| `PATCH /api/jobs/:id` | Change `maxAccess` (updates the Send), `passwordTo` or `passwordSms` until the password went out |
| `POST /api/jobs/:id/arm` | Arm the password mail (send handler); also reads the sender's name and department for the templates |
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
- Outlook caches the script; after an update or a change of `LANGUAGE` clear `%LOCALAPPDATA%\Microsoft\Office\16.0\Wef\`.
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
callers share one request; SSO errors (plain objects with a code) are shown with a hint (language file, `sso.<code>`) and the code
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
- **Passwords:** random passwords are pronounceable (`KateGeteFi45`: consonant-vowel syllables, two digits, ≈ 37.6
  bits from the browser's CSPRNG). That is low for offline attacks but sufficient here: an attacker needs the link
  first, Nextcloud throttles failed attempts on public links, and links expire. Own passwords need ≥ 12 characters and
  three character classes (checked in the taskpane and again by the server). The password exists in readable form
  only in the Vaultwarden Send (limited accesses, 5 days by default, at most 30) and briefly in the taskpane.
- **Graph:** delegated permissions for everything the user triggers; the armed token is stored encrypted for at most
  15 minutes of use, and the `.Shared` permissions reach only mailboxes the user already has rights to. Only the
  optional download confirmation uses an **application** permission (`Mail.Send`) – restrict it to the CloudShare
  users ([microsoft-365.md → 5](microsoft-365.md#5-optional-application-permission-for-the-download-confirmation)), otherwise the app could send as any mailbox of the tenant.
- **Confidential messages** pass the server only in memory; Vaultwarden stores them end-to-end encrypted in the Send
  (the Bitwarden CLI encrypts before upload) and deletes them at expiry. Anyone with link and password can read them
  until then – like a file link.
- **SMS:** the API key and the webhook secret stay on the server; the add-in never talks to seven.io. The SMS carries
  the password in plain text (it is the second channel); the password is read from the Send only for sending and
  the Send is deleted afterwards. Numbers are stored masked in the status; `SMS_MAX_PER_DAY` limits costs and abuse.
  Conclude a data processing agreement with seven.io.
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
  `Mail.Send`, `Mail.Send.Shared` with **admin consent**; client secret → `ENTRA_CLIENT_SECRET` (renew before expiry).
- Add-in: admin center → *Integrated apps* → manifest URL `https://cloudshare.example.com/manifest.xml`.
- Only for the download confirmation: application permission `Mail.Send`, **restricted** to the CloudShare users with
  Exchange Online RBAC for applications.

### seven.io (SMS, optional)

- Account with prepaid credit; API key → `SEVEN_API_KEY`, sender (number or up to 11 characters) → `SEVEN_SENDER`.
- Delivery reports: callback URL `https://cloudshare.example.com/webhooks/seven?token=<SEVEN_WEBHOOK_SECRET>` in the
  dashboard (webhook event "dlr").
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
- Outlook sets the width of the pane; the layout is designed for 320 px (short labels in "My shares", full text
  in `title`/`aria-label`).
- **Cards** (`mailCard.ts`) and the **password mail** (`passwordMail.ts`) share one look: white card, blue accent,
  coloured file type badges, one button – tables and inline styles only, so they render in classic Outlook, Outlook on
  the web, Gmail and on phones. Colours are duplicated in both files; keep them in sync. Plain-text mails get a text
  version. `MAIL_FOOTER` (optional) is a template filled from the sender's profile (`footer.ts`).
- Icons (`addin/assets/icon-*.png`): the Nextcloud mark as bold rings with an upload arrow; icon URLs carry a version
  parameter because Outlook caches icons by URL.

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
7. **One language per installation:** `LANGUAGE` applies to all users; the add-in does not follow each user's Outlook
   language. Recipients get the password mail and the SMS in the same language. `templates.json` is not per
   language.
8. **Possible later:** OIDC token exchange (e.g. Keycloak) instead of app passwords would remove the connect step;
   a language per user (Outlook display language) would need the texts of all languages in the bundles.
