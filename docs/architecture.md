# Architecture: CloudShare

Outlook add-in (classic and new Outlook, Exchange Online) that sends large files not as attachments but as a
password-protected Nextcloud share link. The password is sent automatically to the same recipients in a second
mail as a Vaultwarden Send with a limited number of accesses.

The add-in UI, the Outlook hints and the password mail are German, because users and recipients are German-speaking.

## Components

```
Outlook (classic / new)
  ├─ Taskpane    (taskpane.html)  – connect Nextcloud, pick files, upload, insert links
  └─ Event handlers (commands.js) – OnMessageAttachmentsChanged: info bar
                                    OnMessageSend: trigger password mail, prompt on large attachments
        │  Office SSO token (Entra ID, scope access_as_user)
        ▼
Caddy in the stack (ports 80/443, Let's Encrypt) ─► network internal
        ▼
Docker container "cloudshare"
  ├─ Node/Fastify backend  (server/)  – also serves the add-in files + manifest.xml
  │     ├─ app password per user (Login Flow v2, AES-GCM-encrypted in /data)
  │     │     → WebDAV chunked upload + OCS share in the user's account
  │     ├─ OBO → Graph token → /me/sendMail (password mail on behalf of the user)
  │     └─ HTTP → bw serve (127.0.0.1:8087)
  └─ Bitwarden CLI "bw serve"  – service account in Vaultwarden, creates text Sends
```

Nextcloud usually cannot accept Entra ID tokens (e.g. when it logs users in through Keycloak via `user_oidc`). Therefore the
backend does not use Entra tokens for Nextcloud but one app password per user.

## Flow

1. **Once per user:** taskpane → "Mit Nextcloud verbinden" → "Verbinden". The backend starts the Nextcloud Login Flow v2 and
   opens the Nextcloud login page in the browser (with SSO usually without input). After "Allow access" the
   backend receives an app password and stores it encrypted under the user's Entra object ID (`oid`). In Nextcloud
   it appears under *Settings → Security* as "CloudShare Outlook" and can be revoked there (the add-in then asks to
   reconnect).
2. **Attachment > threshold (`ATTACHMENT_THRESHOLD_MB`, default 1 MB):** the `OnMessageAttachmentsChanged` event
   shows an info bar above the mail with the button "Mit CloudShare senden". Outlook does **not** allow add-ins to
   open the taskpane themselves – a user click is always required.
3. The taskpane lists the mail's file attachments (large ones preselected) and additionally accepts files from the
   computer (without Outlook's size limit). Attachments are read via `getAttachmentContentAsync`, uploaded and
   removed from the mail **only after success**.
4. Upload in 10 MiB chunks through the backend using Nextcloud chunked upload v2 (`/remote.php/dav/uploads/…`),
   target `Outlook-Anhaenge/<date>_<id>/<file>`. No fixed maximum: the free Nextcloud quota is checked up front
   (technical limit 10,000 chunks ≈ 100 GB).
5. One public link per file (password + expiry date). The number of times the password link (Vaultwarden Send) can be opened is set per mail in the taskpane ("Kennwort abrufbar … mal", default `SEND_MAX_ACCESS_COUNT` = 3, limited to `SEND_MAX_ACCESS_LIMIT` = 20 by the server). It stays changeable until sending: afterwards `PATCH /api/jobs/:id` updates the existing Send via `bw serve` (`GET`/`PUT /object/send/:id`, the job keeps the Send ID; jobs from before this version cannot be changed, `SEND_FIXED`); the password mail states the current value ("bis zu 3 Abrufe"). The user chooses a **random password** (default, 20 characters from 67 symbols ≈ 121 bits, CSPRNG with rejection sampling, all four character classes) or an **own password** (min. 12 characters, at least 3 of 4 character classes, no repetition/simple sequence; live check in the taskpane, same rules enforced by the server in `server/src/password.ts`; Nextcloud's password policy applies on top). **All files of one mail share one password, one Vaultwarden
   Send and one password mail.** Job ID and password live in `item.sessionData`; the password field is locked after
   the first file. If the user deletes all inserted links from the mail again, the draft is released (the taskpane
   checks on focus and every 4 s while locked, the send handler before arming): password options unlock, the
   `cloudshareJob` tag is removed and no password mail goes out. Drafts from before link tracking are never released.
6. On send the Smart Alert runs (`OnMessageSend`, manifest `SendMode="SoftBlock"`). Behaviour depends on
   `ATTACHMENT_POLICY` (env, default `enforce`; delivered to the add-in via `/api/config`, so switching needs no
   redeployment of the add-in):
   - **`enforce`** – large attachments remaining → sending is **blocked**, there is no "Send anyway".
     The alert offers "CloudShare öffnen" (opens the taskpane,
     Mailbox 1.14+); older clients are told to use the ribbon button.
   - **`suggest`** – the password mail is armed first (see below), then a prompt with "CloudShare öffnen" or
     "Send anyway" (SoftBlock relaxed via `sendModeOverride: PromptUser`, Mailbox 1.14+; older clients are not
     prompted). "Send anyway" sends the attachments normally.
   - Pending job → the handler only **arms** it (`/api/jobs/:id/arm`): the backend stores a delegated Graph token
     (Mail.Read + Mail.Send). If arming fails, the mail is blocked.
7. **Password mail only after the mail was really sent** (`server/src/jobs.ts`): the taskpane tags the draft with
   the job ID as an Office.js custom property (`cloudshareJob`, stored as MAPI property `cecp-<add-in ID>`, travels
   into "Sent Items"). A watcher checks the user's "Sent Items" every 20 s via Graph; once the tagged mail is there,
   it sends the password mail to that mail's **actual** recipients and subject. For mails sent from a shared mailbox or on behalf of someone (the add-in reports a "From" that differs from the user), the shared mailbox's "Sent Items" are searched as well and the password mail is sent from the same address (send as / on behalf, using the user's Exchange rights). If the mail does not appear within
   15 minutes (prompt cancelled, sending failed, draft discarded), the job is disarmed and stays open – further
   files join it, so there is only ever one password mail, and it always arrives after the main mail (typically
   20–60 s later). Jobs (incl. the armed token) are stored encrypted in `DATA_DIR/password-jobs.json`.
   - The handlers cache the last known configuration in `roamingSettings`, so the policy also applies while the
     server is unreachable. If the add-in cannot run at all, SoftBlock lets the mail through – otherwise a server
     outage would stop all outgoing mail.
   - Outlook mobile and other mail clients do not run the add-in and are not checked (deliberately no Exchange
     transport rule).

## Setup

### Entra ID (app registration "CloudShare")
- Application ID URI: `api://cloudshare.example.com/<client-id>`, scope `access_as_user` (admins and users).
- Authorized client application: `ea5a67f6-b6f3-4338-b240-c655ddc3cc8e` (all Microsoft Office endpoints) for
  `access_as_user`.
- Manifest: `requestedAccessTokenVersion` (old editor: `accessTokenAcceptedVersion`) = `2`.
- Create a client secret → `ENTRA_CLIENT_SECRET`.
- Delegated API permissions: Graph `openid`, `profile`, `offline_access`, `Mail.Read`, `Mail.Read.Shared`,
  `Mail.Send`, `Mail.Send.Shared`. **Grant admin consent.** (`Mail.Read` lets the backend detect the sent mail in
  "Sent Items"; the `.Shared` variants cover shared mailboxes / sending on behalf – they only reach mailboxes the
  user already has access and send rights to in Exchange.)

### Nextcloud
- No change to the login method (`user_oidc`, LDAP, local) needed. Public links must be allowed (for "Dateien
  anfordern" also public upload); password policy/expiry as desired.
- CloudShare reaches Nextcloud through its public address (`NEXTCLOUD_URL`). If Nextcloud runs behind a proxy, its
  request body limit must be ≥ 12 MB (chunk size). No CORS needed, as only the backend talks
  to Nextcloud (server to server). The login flow (`/index.php/login/v2…`) must be reachable.

### Vaultwarden
- `SENDS_ALLOWED=true`.
- Create the service account (with SSO: an IdP user with a verified email, invitation in Vaultwarden, first login sets
  the master password), create an API key (Settings → Security → Keys) → `BW_CLIENTID`/`BW_CLIENTSECRET`,
  master password → `BW_PASSWORD`.

### Deployment
- On the Docker host as described in the [README](../README.md): `/opt/cloudshare`, `./setup.sh` (writes `env/`,
  generates
  `CREDENTIAL_KEY` – keep it safe), Portainer stack. GitHub Actions builds the image `ghcr.io/mrprohl/cloudshare`
  on every push to `main`; updates by re-pull. Include the volume `cloudshare_data` (encrypted app passwords) in
  backups.
- Microsoft 365 admin center → Integrated apps → Upload custom apps → manifest URL
  `https://cloudshare.example.com/manifest.xml`. Rollout can take up to 24 h.

## Open issues / risks
1. The **Bitwarden CLI version** must match the Vaultwarden version (`ARG BW_CLI_VERSION` in the `Dockerfile`,
   tested with Vaultwarden 1.37.5).
2. Running uploads and pending logins are kept in memory → single instance only; a restart discards them (the user
   starts them again). App passwords and password jobs are persistent.
3. The password mail relies on a sent copy in "Sent Items" – of the user or, when sending from a shared mailbox / on
   behalf of someone ("From" differs from the user), of that mailbox. With saving to "Sent Items" disabled no
   password mail is sent. Files always go to the sender's own Nextcloud, also for shared mailboxes. The armed Graph token must still be valid (≥ 1 h), which covers the 15-minute window.
4. Attachments already in Outlook are subject to the Outlook/Exchange limit (approx. 35 MB) and are loaded into
   memory completely during transfer. Pick larger files via "Dateien vom Computer hinzufügen".
5. Smart Alerts require Mailbox requirement set 1.12 (current Microsoft 365 / new Outlook); the custom
   "CloudShare öffnen" button and the "Send anyway" option under `suggest` need 1.14. Outlook limits the button
   text to 20 characters and the message to 500 (`withinLimits` in `commands.ts`); longer values make Outlook on
   the web reject the result and leave the send dialog spinning.
   Under `enforce`, a server outage means large attachments cannot be sent at all until CloudShare is back.
   Classic Outlook on Windows runs the handlers from `commands.js` in a JavaScript-only runtime, which only allows
   SSO and network requests for scripts listed in `/.well-known/microsoft-officeaddins-allowed.json` (served by the
   backend). If the host name or a proxy in front of CloudShare does not deliver that file, the send dialog hangs at
   "takes longer than expected". Classic Outlook caches the script: after an update, clear
   `%LOCALAPPDATA%\Microsoft\Office\16.0\Wef\` or restart Outlook.
6. App passwords have full access to the user's Nextcloud account; protection depends on `CREDENTIAL_KEY` and the
   hardening of the container.
7. Optional later: OIDC token exchange (e.g. Keycloak) instead of app passwords (no connect step, but a preview feature in
   Keycloak).
8. The icon uses the Nextcloud mark (trademark of Nextcloud GmbH) – fine for internal use; check Nextcloud's trademark guidelines before distributing CloudShare to third parties.

## Look and feel

The taskpane is meant to look like part of Outlook, not like an external tool:
- Microsoft Fluent 2 design tokens (colours, 32 px controls, 4 px radius, Segoe UI type ramp) in
  `addin/src/taskpane/taskpane.css`; no own header – Outlook shows the add-in name above the pane.
- Light/dark theme follows `Office.context.officeTheme` (exact background colour), with `prefers-color-scheme` as
  fallback; updated on `OfficeThemeChanged`.
- Outlook's German terminology in all visible texts: "Nachricht" (not "Mail"), "Kennwort" (not "Passwort").
- Ribbon/info bar icon (`addin/assets/icon-*.png`): the Nextcloud mark as fine Fluent-style line rings in Nextcloud blue (centres from the official SVG, radii so the rings touch), with a grey upload arrow in the big ring.

## Mail card, upload links and "Meine Freigaben"

- **Mail card** (`addin/src/taskpane/mailCard.ts`): the taskpane inserts a card instead of a plain list – title,
  one row per file with a coloured file type badge, name, size and a "Herunterladen" button, then expiry and the
  password note, optionally a footer (`MAIL_FOOTER`, a template whose placeholders `{displayName} {givenName} {surname} {jobTitle} {department} {company} {email} {phone} {mobile} {office}` are filled from the sender's Microsoft 365 profile via `GET /api/profile` – Graph `/me`, scope `User.Read`; `
` = new line; " · "-separated segments whose placeholders are all empty are dropped, see `addin/src/taskpane/footer.ts`). Tables and inline styles only, so it works in classic
  Outlook (degrades to a bordered box), Outlook on the web, Gmail and on phones. Text mails get a plain-text
  version.
- **"Ohne Kennwort"** (third password option, always available, file links only): `/complete` with `noPassword: true`
  creates the public share without password and no Vaultwarden Send or password job (`jobId: null`); the card omits
  the password note and the send handler has nothing to arm. The choice is fixed for the whole mail after the first
  link (`DraftState.noPassword`); upload links always keep a password. If Nextcloud enforces passwords for link
  shares (capability `files_sharing.public.password.enforced`, evaluated per user incl. group exceptions, reported
  as `passwordEnforced` by `/api/nextcloud/status`), the option is hidden.
- **"Für jede Datei ein eigenes Kennwort"** (checkbox, default off): every link gets its own password, its own Send and
  its own password mail ("Kennwort für „Datei“: Betreff"). Still one job per mail (`separate: true` on
  `/complete`, `/api/requests`, `/api/shares/:id/link`; `PasswordJob.separate` lists the parts), so tagging and arming
  stay unchanged; the watcher sends one mail per part and remembers sent parts (no duplicates after a retry). The mode is
  fixed by the first link of the mail (the server refuses mixing, 409); the password field stays open – random
  passwords are regenerated per file, an own password is cleared after each step; "Kennwort abrufbar" applies to the
  next file.
- **"Kennwort an eine andere E-Mail-Adresse senden"** (checkbox + one address per mail, `DraftState.passwordTo`): the
  password mail goes *only* to that address, not to the recipients of the mail. The taskpane stores it on the job
  (`PATCH /api/jobs/:id` with `passwordTo`, "" = back to the recipients; also right after the first link), the send
  handler passes it again with `/arm` as a fallback. Validated server-side (`isEmailAddress`, a single address);
  kept when the links are deleted from the mail.
- **Upload links ("Dateien von Empfängern erhalten", button "Link zum Hochladen einfügen")**: `POST /api/requests` creates `NEXTCLOUD_REQUEST_FOLDER/<date>_<id>` and a
  create-only public share (permissions 4, "file drop", label "Dateianforderung") with the same password/expiry
  rules. It joins the draft's password job, so files and upload link share one password mail. Nextcloud must
  allow public uploads (Settings → Sharing).
- **"Meine Freigaben"** (second tab): `GET /api/shares` lists the user's public links below the two CloudShare
  folders from the last 30 days (newest first, at most 100) with expiry, password status (Nextcloud reports only
  whether a link has a password, not whether it was random or own) and status – downloads from the activity app (type `public_links`) for files,
  number of received files (WebDAV PROPFIND) for upload links. "Link kopieren", "Als Link einfügen" and "Zurückziehen"
  (`DELETE /api/shares/:id`, two-step confirmation) – the share ends, the file stays in the user's Nextcloud. "Als Link einfügen" (`POST /api/shares/:id/link`) creates a
  *new* link to the same file or upload folder with the current mail’s settings (password + password mail via the draft’s job, or
  "Ohne Kennwort" for files) – the original password cannot be recovered (Nextcloud stores a hash), the old link stays valid.
