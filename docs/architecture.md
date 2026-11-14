# CloudShare – architecture

CloudShare is an Outlook add-in plus a small backend. Large files are uploaded to the **sender's own Nextcloud** and
inserted into the mail as a password-protected link; the password reaches the recipients in a **second mail** as a
Vaultwarden Send, and only after the first mail has really been sent. This document describes how the parts fit
together, where state lives, and the Outlook peculiarities the code has to work around. Operation and installation:
[README](../README.md).

Everything visible to users and recipients is German (taskpane, Outlook hints, cards, password mail); code, logs and
documentation are English.

## Contents

1. [Components](#1-components)
2. [Main flows](#2-main-flows)
3. [State and data](#3-state-and-data)
4. [HTTP API](#4-http-api)
5. [Outlook runtimes and pitfalls](#5-outlook-runtimes-and-pitfalls)
6. [Security](#6-security)
7. [Setup](#setup)
8. [User interface](#8-user-interface)
9. [Limits and risks](#9-limits-and-risks)

## 1. Components

```
Outlook (classic on Windows / new Outlook / Outlook on the web)
  ├─ Taskpane            addin/src/taskpane/   connect Nextcloud, pick files, upload, insert cards,
  │                                            "Meine Freigaben"
  └─ Event handlers      addin/src/commands/   OnMessageAttachmentsChanged → info bar
                                               OnMessageSend (Smart Alert) → attachment policy, arm password mail
        │  HTTPS + Office SSO token (Entra ID, audience api://<host>/<client-id>, scope access_as_user)
        ▼
Caddy (Let's Encrypt, ports 80/443) ──► network "internal"
        ▼
Container "cloudshare" (Node 22, read-only)
  ├─ Fastify backend     server/src/           API, add-in files, manifest.xml, well-known file
  │     ├─ Nextcloud     app password per user → WebDAV chunked upload, OCS shares, trash
  │     ├─ Graph         On-Behalf-Of token → read "Sent Items", send password mail as the user
  │     └─ bw serve      HTTP on 127.0.0.1:8087 → Vaultwarden Sends (service account)
  └─ Bitwarden CLI "bw serve"
```

| Part | Files | Notes |
|---|---|---|
| Taskpane | `taskpane.ts/.html/.css`, `mailCard.ts`, `footer.ts` | Plain TypeScript, bundled as IIFE (esbuild, ES2017) |
| Event handlers | `commands.ts` → `commands.js` | ES2016 bundle for classic Outlook's JavaScript-only runtime |
| Shared add-in code | `shared/office.ts`, `shared/password.ts`, `shared/sso.ts` | Office.js helpers, draft state, SSO token cache, password rules |
| API | `app.ts` (routes), `index.ts` (start, vault unlock, background jobs) | `buildApp()` is used by the API tests |
| Integrations | `auth.ts` (SSO validation, OBO), `nextcloud.ts`, `graph.ts`, `vaultwarden.ts` | |
| Background jobs | `jobs.ts` (password mails), `cleanup.ts` (expired uploads) | Hourly / every 20 s, see below |
| Persistence | `credentials.ts`, `storage.ts` | AES-256-GCM files in `/data` |

**Why app passwords for Nextcloud?** Nextcloud usually cannot accept Entra ID tokens (it often signs users in through
another IdP, e.g. Keycloak via `user_oidc`). Each user therefore connects once with Nextcloud's Login Flow v2; the
resulting app password is stored encrypted and used as Basic auth. Nextcloud is only ever called server to server,
so it needs no CORS.

**Placeholders.** `%%PUBLIC_BASE_URL%%`, `%%PUBLIC_HOST%%` and `%%ENTRA_CLIENT_ID%%` in the add-in files and the
manifest are replaced by `docker/entrypoint.sh` at container start, so one image serves every installation.

## 2. Main flows

### 2.1 Connect Nextcloud (once per user)

1. Taskpane → "Verbinden" → `POST /api/nextcloud/connect` starts Login Flow v2 and returns the login URL, which the
   taskpane opens in the browser (usually SSO, no typing).
2. The taskpane polls `POST /api/nextcloud/connect/:id/poll` until the user granted access.
3. The backend stores the app password encrypted under the user's Entra object ID (`oid`). In Nextcloud it appears as
   "CloudShare Outlook" under *Settings → Security* and can be revoked there; a 401 from Nextcloud deletes the stored
   credential and the taskpane shows the connect view again (`NEXTCLOUD_NOT_CONNECTED`).

### 2.2 Share files ("Als Link einfügen")

1. The taskpane lists the mail's file attachments (those above `ATTACHMENT_THRESHOLD_MB` preselected) and files from
   the computer – dropped onto the drop zone (the whole "Senden" tab accepts drops) or picked by clicking it (native dialog `window.showOpenFilePicker`, fallback: hidden `<input type="file">`).
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

Variants: **"Link zum Hochladen einfügen"** (`POST /api/requests`) creates a folder in `NEXTCLOUD_REQUEST_FOLDER` with
a create-only link ("file drop") – same password rules, its own step. **"Ohne Kennwort"** (`noPassword: true`, file
links only) creates a link without password, Send or job. **"Meine Freigaben → Einfügen"**
(`POST /api/shares/:id/link`) creates a new link to an existing upload with a new password, as a new step.

### 2.3 Send and password mail

1. On "Senden" Outlook runs the Smart Alert `onMessageSendHandler` (manifest `SendMode="SoftBlock"`):
   - Large attachments left in the mail: under `ATTACHMENT_POLICY=enforce` sending is blocked with the button
     "CloudShare öffnen"; under `suggest` the user is asked once and may choose "Trotzdem senden"
     (`sendModeOverride: PromptUser`, Mailbox 1.14).
   - If all CloudShare cards were deleted from the mail, the draft is released: no tag, no password mail.
   - Otherwise the job is **armed**: `POST /api/jobs/:id/arm` with an SSO token; the backend exchanges it for a Graph
     token (On-Behalf-Of, `Mail.Read(.Shared)` + `Mail.Send(.Shared)`) and stores it with the job. If arming fails,
     sending is blocked with a message naming the failed step.
2. The watcher (`jobs.ts`, every 20 s) searches "Sent Items" for the mail carrying the `cloudshareJob` tag (MAPI
   property `cecp-<add-in ID>`; `ADDIN_ID` must match the manifest `<Id>`). For a "From" that differs from the user
   (shared mailbox, send on behalf), that mailbox's "Sent Items" are searched as well.
3. Once found, it sends **one password mail per step** (`passwordMail.ts`) to the mail's **actual** recipients – or
   only to the address from "Kennwort an eine andere E-Mail-Adresse senden" – from the same sender. Sent parts are
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
- **"Zurückziehen"** counts as expiry today – with `0` the folder goes to the trash at once;
- folders that still have a valid link are postponed; users who disconnected are skipped;
- only folders matching `NEXTCLOUD_FOLDER/YYYY-MM-DD_xxxxxxxx` are touched – received files
  (`NEXTCLOUD_REQUEST_FOLDER`) are never deleted. Uploads from before the feature existed are not registered.

### 2.5 Password channel: recipients, another address or SMS

"Kennwort übermitteln" in the taskpane decides per mail where the password goes (`DraftState.passwordTo` /
`passwordSms`, stored on the job via `PATCH /api/jobs/:id` and passed again with `/arm`):

- **an die Empfänger der Nachricht** (default) – the password mail to the mail's actual recipients;
- **an eine andere E-Mail-Adresse** – the password mail only to that address;
- **per SMS** (only with `SEVEN_API_KEY`) – one SMS per step via seven.io (`sms.ts`), sent by the watcher after the
  mail appeared in "Sent Items", like the password mail. The number is entered in the pane and normalised to E.164
  (`0170 …` → `+49170…`, `SMS_DEFAULT_COUNTRY`). The password exists in readable form only in the step's Vaultwarden
  Send; the watcher reads it from there right before sending (`readSendPassword`) and deletes the Send afterwards.
  If seven.io refuses the SMS (or `SMS_MAX_PER_DAY` is reached), the password mail goes **to the sender** instead –
  never automatically to the recipients by another channel. Delivery reports arrive at `/webhooks/seven` (secret in
  the query, compared in constant time, idempotent, webhook and classic form callback) and update `deliveries.ts`.

How the password of each link went out (channel, masked recipient, SMS status) is shown in "Meine Freigaben".

### 2.6 Text templates

`templates.ts` reads the optional `TEMPLATES_FILE` (`/config/templates.json`, mounted from `env/templates/`): fields
`sms_template`, `mail_password_subject`, `mail_password_intro`, `download_notice_subject`, `download_notice_body`, each
with a fixed set of placeholders. Resolution per field: `users[<sender UPN>]` → `groups[<department>]` → `default` →
built-in German texts; the department and display name are read from the sender's profile when arming. At start the
file is validated: `{password}` outside the SMS, an SMS without `{password}` or with the link / a file name drop the
field (built-in text instead); unknown placeholders, non-GSM characters and SMS above 160 characters are warnings.
A broken file never stops CloudShare. Example: `templates.example.json`.

### 2.7 Download confirmation

With `DOWNLOAD_NOTICE=true`, every file link of a sent mail is registered with the sender, subject, recipients and
expiry (`notices.ts`, `/data/download-notices.json`). Every 10 minutes the Activity API is asked with the owner's
app password; only downloads **after** the password went out count (older links of the same file are ignored). On the
first download the sender gets a mail (`download_notice_*` templates) – sent with CloudShare's **own app token**
(client credentials, application permission `Mail.Send`), because the sender's delegated token has long expired by
then. Entries are dropped after the download or when the link expired. No IP addresses or data about the person who
downloaded are kept.


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

All files in `/data` are AES-256-GCM encrypted with `CREDENTIAL_KEY` (`storage.ts`, atomic writes). The upgrade test
in CI checks that data written by the published image stays readable by a new one.

## 4. HTTP API

All `/api` routes except `config`, `trace` and the well-known file require `Authorization: Bearer <Office SSO token>`.
Errors are `{ error, code? }` with a German message shown in the taskpane.

| Route | Purpose |
|---|---|
| `GET /.well-known/microsoft-officeaddins-allowed.json` | Allows SSO/network for `commands.js` in classic Outlook ([5](#5-outlook-runtimes-and-pitfalls)) |
| `GET /api/config` | Public settings for the add-in (threshold, policy, limits, footer, `trace`, `smsEnabled`) |
| `GET/POST /api/trace` | Diagnostics of the send handler, only with `CLIENT_TRACE=true` (no login) |
| `GET /api/profile` | Sender's Microsoft 365 profile for `MAIL_FOOTER` |
| `GET /api/nextcloud/status` | Connected?, storage (free/used/total), `passwordEnforced` |
| `POST /api/nextcloud/connect`, `…/connect/:id/poll`, `DELETE /api/nextcloud/connection` | Login Flow v2, disconnect |
| `POST /api/uploads`, `PUT …/:id/chunks/:n`, `POST …/:id/complete`, `DELETE …/:id` | Chunked upload, share link |
| `POST /api/requests` | Upload link ("Dateien von Empfängern erhalten") |
| `GET /api/shares`, `DELETE /api/shares/:id`, `POST /api/shares/:id/link` | "Meine Freigaben": list (30 days, max. 100), revoke, insert again |
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
  `/.well-known/microsoft-officeaddins-allowed.json` – without it the send dialog hangs at "dauert länger als
  erwartet".
- Office.js is preloaded, the handlers can be registered right away.
- Outlook caches the script; after an update clear `%LOCALAPPDATA%\Microsoft\Office\16.0\Wef\`.
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
Mailbox 1.14; with 1.12 users are pointed to the ribbon button.

**Office SSO:** only one `getAccessToken` call may run at a time (13008), so the token is cached and concurrent
callers share one request; SSO errors (plain objects with a code) are shown with a German hint and the code
(`shared/sso.ts`). Every step of the send handler has a time limit, so a hanging step ends with a message instead of
"dauert länger als erwartet".

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
  only in the Vaultwarden Send (limited accesses, 7 days) and briefly in the taskpane.
- **Graph:** delegated permissions for everything the user triggers; the armed token is stored encrypted for at most
  15 minutes of use, and the `.Shared` permissions reach only mailboxes the user already has rights to. Only the
  optional download confirmation uses an **application** permission (`Mail.Send`) – restrict it to the CloudShare
  users (see *Setup*), otherwise the app could send as any mailbox of the tenant.
- **SMS:** the API key and the webhook secret stay on the server; the add-in never talks to seven.io. The SMS carries
  the password in plain text (it is the second channel); the password is read from the Send only for sending and
  the Send is deleted afterwards. Numbers are stored masked in the status; `SMS_MAX_PER_DAY` limits costs and abuse.
  Conclude a data processing agreement with seven.io.
- **Container:** read-only, `cap_drop: ALL`, `no-new-privileges`, user 1000, no published ports; `bw serve` listens on
  127.0.0.1 only. Networks: `internal` (Caddy ↔ CloudShare, no outbound), `edge` (Caddy), `egress` (CloudShare).
- **Supply chain:** Trivy scan of every image, `npm audit`, Dependabot, CodeQL, secret scanning; GitHub Actions
  pinned to commit SHAs.

## Setup

### Entra ID (app registration "CloudShare")

- Application ID URI `api://cloudshare.example.com/<client-id>`, scope `access_as_user` (admins and users).
- Authorized client application `ea5a67f6-b6f3-4338-b240-c655ddc3cc8e` (all Microsoft Office endpoints) for
  `access_as_user`.
- Manifest: `requestedAccessTokenVersion` (old editor: `accessTokenAcceptedVersion`) = `2`.
- Client secret → `ENTRA_CLIENT_SECRET`.
- Delegated Graph permissions: `openid`, `profile`, `offline_access`, `User.Read`, `Mail.Read`, `Mail.Read.Shared`,
  `Mail.Send`, `Mail.Send.Shared` – **grant admin consent**.
- **Only for the download confirmation** (`DOWNLOAD_NOTICE=true`): application permission Graph `Mail.Send` with admin
  consent, **restricted to the CloudShare users** with Exchange Online RBAC for applications, e.g.:
  ```powershell
  New-ManagementScope -Name "CloudShare users" -RecipientRestrictionFilter "MemberOfGroup -eq '<DN of the group>'"
  New-ServicePrincipal -AppId <client-id> -ObjectId <enterprise-app-object-id> -DisplayName CloudShare
  New-ManagementRoleAssignment -App <client-id> -Role "Application Mail.Send" -CustomResourceScope "CloudShare users"
  ```
  Then remove the tenant-wide consent for `Mail.Send` in Entra, so only the scoped assignment applies.

### seven.io (SMS, optional)

- Account with prepaid credit; API key → `SEVEN_API_KEY`, sender (number or up to 11 characters) → `SEVEN_SENDER`.
- Delivery reports: callback URL `https://cloudshare.example.com/webhooks/seven?token=<SEVEN_WEBHOOK_SECRET>` in the
  dashboard (webhook event "dlr").
- Conclude a data processing agreement and document it in the record of processing activities.

### Nextcloud

- No change to the login method (`user_oidc`, LDAP, local) needed. Public links must be allowed; for upload links also
  public upload. If *Enforce password protection* is on, CloudShare hides "Ohne Kennwort" (exempt a group there to
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
  `taskpane.css`, no own header, light/dark theme from `Office.context.officeTheme`, Outlook's wording ("Nachricht",
  "Kennwort"). Accessibility is tested with axe (WCAG 2 A/AA) in CI.
- Outlook sets the width of the pane; the layout is designed for 320 px (short labels in "Meine Freigaben", full text
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
7. **Possible later:** OIDC token exchange (e.g. Keycloak) instead of app passwords would remove the connect step.
