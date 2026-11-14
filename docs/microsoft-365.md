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
5. [Optional: application permission for the download confirmation](#5-optional-application-permission-for-the-download-confirmation)
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
                                                             User.Read: name, department, footer
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
   - Name: `CloudShare`
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
| `User.Read` | sender's name, department and contact data (footer of the card, text templates) |
| `Mail.Read` | find the sent mail in the sender's "Sent Items" before the password goes out |
| `Mail.Read.Shared` | the same for shared mailboxes the user sends from |
| `Mail.Send` | send the password mail as the sender |
| `Mail.Send.Shared` | send it from the shared mailbox the mail came from |

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
there.

[Microsoft 365 admin center](https://admin.microsoft.com) → *Settings → Integrated apps* (*Einstellungen → Integrierte
Apps*) → **Upload custom apps** (*Benutzerdefinierte Apps hochladen*):

1. App type: **Office Add-in**; *Provide link to manifest file*: `https://cloudshare.example.com/manifest.xml` →
   **Validate**.
2. **Assign users**: *Just me* or a test group first, later *Specific users/groups* or *Entire organization*.
3. **Accept permissions**: the wizard lists the Graph permissions of the manifest (`<WebApplicationInfo>`); accepting
   them grants the admin consent of 2.4 if it is still missing.
4. **Finish deployment.** Outlook shows the **CloudShare** button in the compose window after a few hours, at the
   latest after 24 hours (classic Outlook: tab *Message* (*Nachricht*); new Outlook and the web: the ribbon or the
   *Apps* menu of the compose window).

Alternative with a clearer error message on failure (Exchange Online PowerShell, `Connect-ExchangeOnline` first):

```powershell
New-App -OrganizationApp -Url https://cloudshare.example.com/manifest.xml -ProvidedTo Everyone -DefaultStateForUser Enabled
```

Notes:

- The manifest's name, description and ribbon tooltip come from the configured language (`LANGUAGE`). Microsoft
  365 reads the manifest with a delay – the pane itself and its texts change immediately after a stack update.
- The send check (Smart Alert) is part of the manifest; nothing else has to be enabled per user. Users cannot remove
  a centrally deployed add-in that is set to *fixed*.

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

## 5. Optional: application permission for the download confirmation

Only for `DOWNLOAD_NOTICE=true`: the confirmation mail is sent days later, when the user's delegated token has long
expired, so CloudShare uses its own **application** permission.

1. App registration → *API permissions* → *Add a permission* → *Microsoft Graph* → **Application permissions** →
   `Mail.Send` → *Grant admin consent*.
2. Restrict it to the CloudShare users with **Exchange Online RBAC for applications** – otherwise the app could send
   as any mailbox of the tenant:
   ```powershell
   Connect-ExchangeOnline
   New-ManagementScope -Name "CloudShare users" -RecipientRestrictionFilter "MemberOfGroup -eq '<distinguished name of the group>'"
   New-ServicePrincipal -AppId <client-id> -ObjectId <object ID of the enterprise application> -DisplayName CloudShare
   New-ManagementRoleAssignment -App <client-id> -Role "Application Mail.Send" -CustomResourceScope "CloudShare users"
   ```
   The object ID is that of the *enterprise application* (Entra → *Enterprise applications* → CloudShare), not of
   the app registration.
3. Remove the tenant-wide `Mail.Send` **application** permission from the app registration again (step 1), so only
   the scoped RBAC assignment applies. Check with
   `Test-ServicePrincipalAuthorization -Identity <client-id> -Resource <mailbox>`.
4. Set `DOWNLOAD_NOTICE=true` in `env/cloudshare.env` and update the stack.

## 6. Checking the setup

1. `curl -s https://cloudshare.example.com/api/config` returns JSON; `https://cloudshare.example.com/manifest.xml`
   shows your address and client ID.
2. In Outlook, write a new mail → **CloudShare** → the pane opens without asking for a password → *Connect*
   connects Nextcloud.
3. Attach a file above the threshold → the info bar appears; click *Insert as link* → a card is inserted.
4. Send the mail to yourself or a colleague → the password mail follows within about a minute. The server log shows
   `Password mail sent` (`docker logs cloudshare 2>&1 | grep "Password mail"`).

## 7. Maintenance

| What | When | How |
|---|---|---|
| Client secret | before it expires (2.5) | new secret → `ENTRA_CLIENT_SECRET` in `env/cloudshare.env` → update the stack → delete the old secret |
| CloudShare's address changes | – | Application ID URI (2.3), DNS, `PUBLIC_BASE_URL`; then deploy the add-in again with the new manifest URL |
| New Graph permission in a CloudShare release | per release notes | add it (2.4), grant admin consent |
| Users leave the CloudShare group | – | with 2.6 / 5 the access ends automatically |

## 8. Errors and their causes

Messages are quoted in English (`LANGUAGE=en`), the German text of the default installation in brackets.

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
| The CloudShare button does not appear | deployment still in progress (up to 24 h), user not assigned, or an old Outlook cache | admin center → *Integrated apps*; classic Outlook: clear `%LOCALAPPDATA%\Microsoft\Office\16.0\Wef\` |
