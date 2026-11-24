# CloudShare – connecting Microsoft 365

Everything CloudShare needs in Microsoft 365, step by step: the app registration in Entra ID, the deployment of the
add-in and the Exchange Online settings. Menu names are given in English with the German portal wording in
brackets. The examples use `cloudshare.example.com` as CloudShare's address (`PUBLIC_BASE_URL`); replace it with
yours everywhere.

Roles needed: *Application Administrator* (or *Cloud Application Administrator*) for the app registration, *Global
Administrator* or *Privileged Role Administrator* for the admin consent, *Exchange Administrator* for Exchange
settings, and *Global Administrator* or *Exchange Administrator* for deploying the add-in in the admin center.

## Contents

1. [How CloudShare signs in](#1-how-cloudshare-signs-in)
2. [App registration in Entra ID](#2-app-registration-in-entra-id)
3. [Deploying the add-in](#3-deploying-the-add-in)
4. [Exchange Online](#4-exchange-online)
5. [Optional: application permission for notices and the delivery proof](#5-optional-application-permission-for-notices-and-the-delivery-proof)
6. [Checking the setup](#6-checking-the-setup)
7. [Maintenance](#7-maintenance)
8. [Errors and their causes](#8-errors-and-their-causes)

## 1. How CloudShare signs in

```
Outlook ──Office SSO──► token for "api://cloudshare.example.com/<client-id>" (scope access_as_user)
   │
   └─► CloudShare server ──validates the token (tenant, audience, scope)
                         └─On-Behalf-Of (client secret)──► Graph token of the user
                                                             Mail.Read(.Shared): find the sent mail
                                                             Mail.Send(.Shared): send the password mail
                                                             User.Read: name, department, signature
```

- Users never sign in to CloudShare themselves: Outlook hands the add-in a token for the app registration
  (Office single sign-on). This only works if Office is an **authorized client** of the registration and the
  **Application ID URI** matches the manifest.
- The server exchanges that token for a Graph token of the same user (**On-Behalf-Of**, needs the **client
  secret**). There is no interactive consent in the add-in, so the delegated permissions need **admin consent**.
- One registration serves classic Outlook, new Outlook and Outlook on the web.

## 2. App registration in Entra ID

[Microsoft Entra admin center](https://entra.microsoft.com) → *Identity → Applications → App registrations*
(*App-Registrierungen*).

### 2.1 Register the app

1. **New registration** (*Neue Registrierung*):
   - Name: `CloudShare Server` – not just `CloudShare`: deploying the add-in (step 3) creates a second registration
     named `CloudShare`, and the two are easy to mix up
   - Supported account types: **Accounts in this organizational directory only** (single tenant)
   - Redirect URI: leave empty
2. **Register**. On the *Overview* page copy:
   - *Application (client) ID* → `ENTRA_CLIENT_ID`
   - *Directory (tenant) ID* → `ENTRA_TENANT_ID`

### 2.2 Token version

*Manifest* → set the access token version to **2**:

- Microsoft Graph App Manifest (new view): `"api": { "requestedAccessTokenVersion": 2 }`
- AAD Graph App Manifest (old view): `"accessTokenAcceptedVersion": 2`

**Save.** CloudShare accepts v1 and v2 tokens, but Office SSO is documented and tested with v2.

### 2.3 Expose the API (`access_as_user`)

*Expose an API* (*API verfügbar machen*):

1. **Application ID URI** → *Add* (*Hinzufügen*) → `api://cloudshare.example.com/<client-id>`
   - exactly the host of `PUBLIC_BASE_URL` (no `https://`, no port, no trailing slash), followed by the client ID;
   - it must match `<Resource>` in the manifest – CloudShare writes it from `PUBLIC_BASE_URL` and `ENTRA_CLIENT_ID`.
     If the address of CloudShare changes, change the URI as well.
2. **Add a scope** (*Bereich hinzufügen*):

   | Field | Value |
   |---|---|
   | Scope name | `access_as_user` |
   | Who can consent? | **Admins and users** |
   | Admin consent display name | `Office can act as the user` |
   | Admin consent description | `Enables Office to call CloudShare with the same rights as the current user.` |
   | User consent display name / description | same texts |
   | State | Enabled |

3. **Add a client application** (*Clientanwendung hinzufügen*) → client ID
   **`ea5a67f6-b6f3-4338-b240-c655ddc3cc8e`** (Microsoft Office – covers Outlook on Windows, Mac, the web and new
   Outlook) → tick `api://cloudshare.example.com/<client-id>/access_as_user` → *Add application*.

### 2.4 API permissions (delegated) and admin consent

*API permissions* (*API-Berechtigungen*) → *Add a permission* → *Microsoft Graph* → **Delegated permissions**:

| Permission | Used for |
|---|---|
| `openid`, `profile`, `offline_access` | sign-in basics of the token exchange |
| `User.Read` | sender's name, department and contact data (signature of the password mail, text templates) |
| `Mail.Read` | find the sent mail in the sender's "Sent Items" before the password goes out |
| `Mail.Read.Shared` | the same for shared mailboxes the user sends from |
| `Mail.Send` | send the password mail as the sender |
| `Mail.Send.Shared` | send it from the shared mailbox the mail came from |
| `Contacts.Read` | *optional:* suggestions from the user's own Outlook contacts (default contacts folder) while typing another address or an SMS number under *Password to*, and the address book page next to that field. Read on demand, kept in memory ten minutes, never stored. Without it the fields are plain inputs |
| `InformationProtectionPolicy.Read` | *optional:* the names and the priority of the sensitivity labels published to the user (a shared file's label in the pane, which label is higher). Without it the pane uses Outlook's own label catalog – names may then be missing in some Outlook versions |

Then **Grant admin consent for \<organisation\>** (*Administratorzustimmung für … erteilen*). Every line must show a
green tick ("Granted"). Delegated permissions never exceed the user's own rights: `.Shared` only reaches mailboxes
the user already has access and send rights to.

### 2.5 Client secret

*Certificates & secrets* (*Zertifikate & Geheimnisse*) → *Client secrets* → **New client secret**:

- Description: e.g. `CloudShare server 2026`; expiry: up to 24 months.
- Copy the **Value** (not the *Secret ID*) right away – it is shown only once → `ENTRA_CLIENT_SECRET`.
- Put the expiry date into your calendar (see [7. Maintenance](#7-maintenance)): when the secret expires, the pane
  still opens but no password mail can be sent.

### 2.6 Optional: limit CloudShare to a group

Entra → *Enterprise applications* (*Unternehmensanwendungen*) → *CloudShare* → *Properties* → **Assignment required?**
(*Zuweisung erforderlich?*) = Yes, then *Users and groups* → add the group. Users outside the group get a sign-in error
in the pane. Deploy the add-in (step 3) to the same group.

### 2.7 Conditional Access

Policies for "All cloud apps" also apply to CloudShare (the On-Behalf-Of exchange). Requirements the Outlook session
already meets (MFA, compliant device) are passed on. Policies that demand an interactive step only for CloudShare
(e.g. sign-in frequency or terms of use for this app alone) make the exchange fail – exclude CloudShare from them or
scope them to Outlook as well.

## 3. Deploying the add-in

CloudShare must be running and reachable at `https://cloudshare.example.com` first – Microsoft reads the manifest from
there. **Deploy only after `ENTRA_CLIENT_ID` is set** (step 2) and the manifest shows it:

```bash
curl -s https://cloudshare.example.com/manifest.xml | grep -E '<Version>|<Id>|<Resource>'
```

`<Id>` under `<WebApplicationInfo>` and `<Resource>` must contain the client ID of 2.1. Microsoft 365 keeps the
manifest as it was when it read it – a later change of the client ID only reaches Outlook with a higher manifest
version (*Update*) or by removing the app and deploying it again.

[Microsoft 365 admin center](https://admin.microsoft.com) → *Settings → Integrated apps* (*Einstellungen → Integrierte
Apps*) → **Upload custom apps** (*Benutzerdefinierte Apps hochladen*):

1. App type: **Office Add-in**; *Provide link to manifest file*: `https://cloudshare.example.com/manifest.xml` →
   **Validate**.
2. **Assign users**: *Just me* or a test group first, later *Specific users/groups* or *Entire organization*.
3. **Accept permissions**: the wizard lists the Graph permissions of the manifest (`<WebApplicationInfo>`); accepting
   them grants the admin consent of 2.4 if it is still missing.
4. **Finish deployment.** Outlook shows the **CloudShare** button in the compose window after a few hours, at the
   latest after 24 hours (classic Outlook: tab *Message* (*Nachricht*); new Outlook and the web: the ribbon or the
   *Apps* menu of the compose window). When a mail is **read**, the button is there as well (classic Outlook: tab
   *Home* (*Start*) and the reading window; new Outlook and the web: the ribbon or the *Apps* menu of the mail) – it
   opens "History" with the links of that mail.

Alternative with a clearer error message on failure (Exchange Online PowerShell, `Connect-ExchangeOnline` first):

```powershell
New-App -OrganizationApp -Url https://cloudshare.example.com/manifest.xml -ProvidedTo Everyone -DefaultStateForUser Enabled
```

Notes:

- **A second app registration "CloudShare" appears.** Deploying the add-in in *Integrated apps* makes Microsoft 365
  create an app registration of its own, named after the add-in and without a client secret (*App registrations* →
  *All applications*, created on the day of the upload). It belongs to the deployment, not to CloudShare:
  - do **not** put its ID into `ENTRA_CLIENT_ID` – CloudShare uses the registration of step 2 (the one with the
    client secret, `CloudShare Server`);
  - do **not** change or delete it while the add-in is deployed. Removing the add-in and deploying it again creates
    a new one; the old one can then be deleted.
- The manifest's description and ribbon texts are English with an `Override` per further language, so the ribbon
  follows the Outlook language. Microsoft 365 reads the manifest with a delay – the pane itself and its texts change
  immediately after a stack update.
- The send check (Smart Alert) is part of the manifest; nothing else has to be enabled per user. Users cannot remove
  a centrally deployed add-in that is set to *fixed*.
- **New manifest version** (e.g. 1.1 added the button for mails being read): Microsoft 365 picks up the changed
  manifest from the URL on its own, which can take a few days. Check *Integrated apps → CloudShare* for the version;
  if the new button does not show up, remove the app there and deploy it again with the same manifest URL. Classic
  Outlook additionally keeps its cache ([manual](manual.md#operation) → *Operation*).

## 4. Exchange Online

- **Mailboxes in Exchange Online.** Clients: classic Outlook (Microsoft 365 Apps, current channel or monthly
  enterprise channel), new Outlook, Outlook on the web. Outlook mobile and Outlook for Mac do not run the send check.
  Smart Alerts need requirement set Mailbox 1.12, the "Open CloudShare" button in the alert and "Send anyway" (policy
  `suggest`) need 1.14 – older builds are pointed to the ribbon button.
- **"Sent Items" must be kept** (default). The password mail goes out only after the mail appeared there; with saving
  switched off, no password mail is sent.
- **Shared mailboxes:** by default the sent copy lands in the sender's own "Sent Items"; CloudShare searches both the
  user's and the shared mailbox. Optional, for a copy in the shared mailbox as well:
  ```powershell
  Set-Mailbox -Identity info@example.com -MessageCopyForSentAsEnabled $true -MessageCopyForSendOnBehalfEnabled $true
  ```
  The user needs *Send As* or *Send on Behalf* on the shared mailbox (as for sending without CloudShare).
- **Classic Outlook** runs the send check in a JavaScript-only runtime; it may only use sign-in and the network
  because CloudShare serves `https://cloudshare.example.com/.well-known/microsoft-officeaddins-allowed.json` – nothing
  to configure, but proxies must not block that path.

## 5. Optional: application permission for notices and the delivery proof

For `DOWNLOAD_NOTICE=true`, `UPLOAD_NOTICE=true`, `SMS_NOTICE=true` and `DELIVERY_PROOF=true` (unset it follows
`DOWNLOAD_NOTICE`): these mails are sent later, when the user's delegated token has long expired, so CloudShare uses
its own **application** permission `Mail.Send`:

- **Notices to the sender** – download confirmation, files received through an upload link, password SMS reported
  as not delivered (without the permission only "History" shows these); with the delivery proof also the signed PDF
  after the first download and at the end, and the note when code requests were locked. Sent as the user who shared
  the link, to that user.
- **Delivery proof to the recipients** – the confirmation code, and the password mail after a recipient confirmed
  their address. The code goes out as the user who shared the link; the password mail as the mailbox the original
  mail came from (the user's own, or a shared mailbox they sent from), to the recipient who confirmed (or the other
  password address). Only `Mail.Send` is needed: nothing is read with this permission –
subject and recipients were taken from *Sent Items* with the delegated token when the mail was sent, downloads and
uploads come from Nextcloud. There is no application version of `Mail.Send.Shared`, and none is needed.

**All delegated permissions of step 2 stay** – including the delegated `Mail.Send`, `Mail.Send.Shared`, `Mail.Read`
and `Mail.Read.Shared`. They serve the password mail: right after sending, CloudShare finds the mail in *Sent Items*
and sends the password with the user's own token, as the user or the shared mailbox (also "send again" in
"History"). The application permission cannot replace them; it is only used days later for the hints above.

| `Mail.Send` | Token | When | For |
|---|---|---|---|
| delegated | the user's, when "Send" is pressed | seconds to minutes later | password mail, sending it again |
| application | CloudShare's own | later, the user's token has expired | download confirmation, files received, SMS not delivered; delivery proof: code, password after the confirmation, PDF, lock note |

The permission belongs to the registration CloudShare signs in with: **`CloudShare Server`** from step 2 (its client
ID is `ENTRA_CLIENT_ID`) – not the registration "CloudShare" that *Integrated apps* creates. Choose **one** of two ways:

### Way A (recommended): only the CloudShare users – Exchange RBAC for applications

Exchange grants `Mail.Send` to the app, limited to the members of one group. **No** application permission is entered
in Entra. The notices are sent from the mailbox of the user who shared the link, so the group holds the **users**.
With the delivery proof, the password mail after a confirmation comes from the mailbox the original mail was sent
from: if users send mails with a delivery proof **from shared mailboxes**, add those shared mailboxes to the group
too – otherwise the password mail fails with `ErrorAccessDenied` (the recipient's page then says the password could
not be sent).

1. A mail-enabled security group (or distribution group) with the CloudShare users, e.g. "CloudShare users".
2. Note two IDs of `CloudShare Server`: the *Application (client) ID* (= `ENTRA_CLIENT_ID`) and the *Object ID* of the
   **enterprise application** (Entra → *Enterprise applications* – German *Unternehmensanwendungen* → CloudShare
   Server → *Overview*) – not the object ID of the app registration.
3. In Exchange Online PowerShell (role *Organization Management*) – fill in the first three lines, run the rest as is:
   ```powershell
   $appId    = "<ENTRA_CLIENT_ID>"
   $objectId = "<object ID of the enterprise application>"
   $group    = "CloudShare users"

   Connect-ExchangeOnline
   # Makes the app known to Exchange.
   New-ServicePrincipal -AppId $appId -ObjectId $objectId -DisplayName "CloudShare Server"
   # The scope: every member of the group. This is what limits the permission – without it Mail.Send would apply to
   # all mailboxes (then way B is the simpler way to get the same).
   $dn = (Get-Group $group).DistinguishedName
   New-ManagementScope -Name "CloudShare users" -RecipientRestrictionFilter "MemberOfGroup -eq '$dn'"
   # Mail.Send for the app, within that scope.
   New-ManagementRoleAssignment -App $appId -Role "Application Mail.Send" -CustomResourceScope "CloudShare users"
   ```
   New CloudShare users only need to be added to the group.
4. Make sure the app registration has **no** `Mail.Send` of type *Application* (*API permissions* – German
   *API-Berechtigungen*: a row `Mail.Send · Application / Anwendung · Send mail as any user`). If there is one, remove
   it (*…* → *Remove permission* / *Berechtigung entfernen*): Exchange adds an Entra grant to the RBAC scope, so with
   it the app could still send as any mailbox. The **delegated** `Mail.Send` / `Mail.Send.Shared` stay – the password
   mail needs them.
5. Check (changes can take up to two hours to apply):
   ```powershell
   Test-ServicePrincipalAuthorization -Identity $appId -Resource <mailbox of a CloudShare user>  # InScope: True
   Test-ServicePrincipalAuthorization -Identity $appId -Resource <any other mailbox>             # InScope: False
   ```

### Way B (simple): all mailboxes – application permission in Entra

App registration `CloudShare Server` → *API permissions* → *Add a permission* → *Microsoft Graph* → **Application
permissions** → `Mail.Send` → *Add permissions* → *Grant admin consent*. The row then reads
`Mail.Send · Application / Anwendung · Send mail as any user · Granted`. The app may then send as **any** mailbox of the
tenant – CloudShare only does so for the user who shared a link, but the permission itself is not limited. Switch to
way A later by doing its steps 1–5 (step 4 removes this row again).

### Then

Set `DOWNLOAD_NOTICE=true`, `UPLOAD_NOTICE=true`, `SMS_NOTICE=true` and/or `DELIVERY_PROOF=true` in
`env/cloudshare.env` and update the stack.
Without the permission these mails fail with a warning in the log (`Note about an undelivered SMS not sent`, `… check
failed, retrying later`); everything else keeps working.

## 6. Checking the setup

1. `curl -s https://cloudshare.example.com/api/config` returns JSON; `https://cloudshare.example.com/manifest.xml`
   shows your address and client ID.
2. In Outlook, write a new mail → **CloudShare** → the pane opens without asking for a password → *Connect*
   connects Nextcloud.
3. Attach files larger than the threshold together (default 5 MB) → the info bar appears; open CloudShare and click *Insert file as
   link* → a card is inserted.
4. Send the mail to yourself or a colleague → the password mail follows within about a minute. The server log shows
   `Password mail sent` (`docker logs cloudshare 2>&1 | grep "Password mail"`).
5. Open the sent mail in *Sent Items* → **CloudShare** → the pane lists its link under *In this message*.
6. With `DELIVERY_PROOF`: switch on *Proof of delivery* (below the password), send a file to a second address of yours →
   no password mail arrives with the mail. Open the link, confirm the address with the code → the password mail
   arrives, *To the download* leads on. In *History* the seal next to the file name opens the PDF; Adobe Reader shows
   it as signed and unchanged (signer "unknown" with the certificate CloudShare made itself).
7. With `InformationProtectionPolicy.Read` (2.4): attach a labelled Office file → its label name shows under it in
   the pane.
8. With `Contacts.Read` (2.4): *Password to* → *Another email address*, type two letters of a contact → the contact
   drops down below the field; Enter fills in the address.

## 7. Maintenance

| What | When | How |
|---|---|---|
| Client secret | before it expires (2.5) | new secret → `ENTRA_CLIENT_SECRET` in `env/cloudshare.env` → update the stack → delete the old secret |
| CloudShare's address changes | – | Application ID URI (2.3), DNS, `PUBLIC_BASE_URL`; then deploy the add-in again with the new manifest URL |
| Switching to another app registration | – | set it up completely (2.1–2.5, and 2.6 / 2.7 / 5 if used) → `ENTRA_CLIENT_ID` and `ENTRA_CLIENT_SECRET` → update the stack → check the manifest (3) → remove the add-in in *Integrated apps* and deploy it again → only then delete the old registration. Users keep their Nextcloud connection (stored per user, not per app). Until Outlook has the new manifest, sign-in in the pane fails – plan a quiet time |
| New Graph permission in a CloudShare release | per release notes | add it (2.4), grant admin consent |
| Users leave the CloudShare group | – | with 2.6 / 5 the access ends automatically |
| Certificate that signs the proof PDF | own one before it expires; the self-made one lasts ten years | PEM files under `env/config` → `PROOF_SIGN_CERT_FILE`, `PROOF_SIGN_KEY_FILE` → update the stack; PDFs signed before stay valid |

## 8. Errors and their causes

Messages are quoted in English (the default, `LANGUAGE=en`), the German text (`LANGUAGE=de`) in brackets.

| Message in Outlook / log | Cause | Fix |
|---|---|---|
| Pane: sign-in error **13007** ("did not get a sign-in token" / „kein Anmeldetoken“) | Application ID URI does not match the manifest, Office not authorized, or consent missing | 2.3 (URI exactly `api://<host>/<client-id>`, client `ea5a67f6-…`), 2.4 admin consent |
| Pane: **13005** (permissions not confirmed) | admin consent missing | 2.4 |
| Pane: **13001 / 13006** | Outlook not signed in with the work account / sign-in not finished | sign in to Outlook again, restart Outlook |
| Pane: "Invalid or expired token" („Ungültiges oder abgelaufenes Token“) | `ENTRA_TENANT_ID` / `ENTRA_CLIENT_ID` do not belong to this registration | compare with the *Overview* page (2.1) |
| Pane: "Scope access_as_user missing" („Scope access_as_user fehlt“) | scope named differently or not authorized for Office | 2.3 |
| "Token for … not available (check admin consent)" („Token für … nicht erhältlich (Administratorzustimmung prüfen)“) when sending, `AADSTS65001` in the log | delegated permission without admin consent | 2.4 |
| same with `AADSTS7000215` / `AADSTS7000222` | client secret wrong or expired | 2.5, 7 |
| same with `AADSTS50076` / `AADSTS50079` | a Conditional Access policy demands an interactive step | 2.7 |
| Sign-in error `AADSTS50105` | user not assigned while *Assignment required* is on | 2.6 |
| No password mail, log `Graph sendMail failed (HTTP 403)` for a shared mailbox | no *Send As* / *Send on Behalf* right | 4 |
| Download confirmation fails with `ErrorAccessDenied` | application permission missing or RBAC scope does not contain the sender | 5 |
| Delivery proof: recipients get no code, or the page says the password could not be sent; log `Graph sendMail (app) failed (HTTP 403)` | application `Mail.Send` missing, or the sender's mailbox (or the shared mailbox the mail came from) is outside the RBAC scope | 5 |
| Label names missing in the pane ("Sensitivity label"), log `Sensitivity labels not readable via Graph` | `InformationProtectionPolicy.Read` not added or no admin consent – Outlook's own catalog is used instead | 2.4 |
| No suggestions under *Password to*, log `Contacts not readable via Graph` | `Contacts.Read` not added or no consent – the fields stay plain inputs | 2.4 |
| Start log `Proof signing files not usable` (the PDFs carry the self-made certificate instead of your own) | `PROOF_SIGN_CERT_FILE` / `PROOF_SIGN_KEY_FILE` set but the file is missing (`ENOENT`, e.g. the example lines taken over without the files), not readable, or the key does not belong to the certificate | put the PEM files into `env/config` (= `/config`) or remove both lines to use the self-made certificate; then `docker compose up -d --force-recreate cloudshare` |
| Proof PDF not signed, log `Delivery proof PDF not signed` | signing itself failed (the log names the error) | check the log; the PDF is made anew, signed, when it is opened again |
| Pane: "Nextcloud (…) cannot be reached" („Nextcloud (…) ist nicht erreichbar“), "The password service cannot be reached" or "… cannot be reached"; log `fetch failed` with `ENOTFOUND` / `ECONNREFUSED` / `ETIMEDOUT` | `NEXTCLOUD_URL` (or another address in `env/cloudshare.env`) wrong – e.g. still the example value –, DNS, or a firewall between the container and that host; the password service is `bw serve` inside the container (Vaultwarden login failed at start) | correct `env/cloudshare.env`, then `docker compose up -d --force-recreate cloudshare`; check the start log |
| Pane: sign-in error right after a new deployment, `ENTRA_CLIENT_ID` shows the ID of a registration **without** client secret | the registration Microsoft 365 created for the deployment was taken instead of the one of step 2 | use the client ID of 2.1 (the registration with the secret); see the notes of 3 |
| The CloudShare button does not appear | deployment still in progress (up to 24 h), user not assigned, or an old Outlook cache | admin center → *Integrated apps*; classic Outlook: clear `%LOCALAPPDATA%\Microsoft\Office\16.0\Wef\` |
