# Guide: installing CloudShare (step by step)

From nothing to the first link sent from Outlook. Every step ends with a **checkpoint** – go on only when it holds.
Details behind the links: [docs/microsoft-365.md](docs/microsoft-365.md) (Entra ID, add-in, Exchange),
[docs/manual.md](docs/manual.md) (settings, operation, troubleshooting), [docs/architecture.md](docs/architecture.md)
(how it works). The examples use `cloudshare.example.com`, `cloud.example.com` and `vault.example.com`.

| Step | What | Where |
|---|---|---|
| 1 | Prepare: decisions, accounts, DNS name | – |
| 2 | Register the app in Entra ID | Entra admin center |
| 3 | Nextcloud (and Vaultwarden) | their admin interfaces |
| 4 | Install on the Docker host with `setup.sh` | server |
| 5 | Start and check | server |
| 6 | Deploy the add-in | Microsoft 365 admin center |
| 7 | First test in Outlook | Outlook |
| 8 | Optional features | server, Entra, seven.io |

---

## Step 1 – Prepare

Have ready (do not paste secrets into chats or tickets):

- [ ] a **Docker host** with Docker Compose ≥ 2.20, ports **80 and 443** reachable from the internet (443 also UDP)
- [ ] a **DNS name** for CloudShare (e.g. `cloudshare.example.com`) pointing to that host
- [ ] an account allowed to **register apps in Entra ID**, grant admin consent and **deploy add-ins** in Microsoft 365
- [ ] **Nextcloud** with public links allowed (any login method – users connect once with the Nextcloud login flow)
- [ ] decision **Vaultwarden yes/no**: with it the password mail carries a link to the password that runs out and opens
      only a few times, and "Message" (confidential message) is offered; without it the password mail carries the
      password itself ([manual → Without Vaultwarden](docs/manual.md#without-vaultwarden))
- [ ] with Vaultwarden: an own **service account** (e.g. `cloudshare@example.com`) and Sends allowed
- [ ] decisions on the optional features (step 8): password by **SMS** (seven.io account), **download confirmation**,
      **upload notification**, **proof of delivery** – the last three need an application permission in Entra
- [ ] the answers for `setup.sh` (step 4): language of mails and SMS (`en`, `de`), from which attachment size CloudShare
      steps in (default 5 MB), how long links last (default 14, at most 90 days), a signature template for the mails to
      recipients, an address for feedback from the add-in (optional)

**Checkpoint:** everything ticked, the DNS name resolves to the host.

---

## Step 2 – Register the app in Entra ID

Follow [docs/microsoft-365.md → 2. App registration](docs/microsoft-365.md#2-app-registration-in-entra-id):

- [ ] app registration (e.g. "CloudShare Server"), single tenant, access token version 2
- [ ] application ID URI `api://cloudshare.example.com/<client-id>`, scope `access_as_user`, Microsoft Office as
      authorized client
- [ ] delegated Graph permissions with **admin consent** (optional `Contacts.Read` for address suggestions and
      `InformationProtectionPolicy.Read` for label names)
- [ ] client secret – note its **value** and the expiry date

Note: tenant ID, client ID, client secret.

**Checkpoint:** the permissions show "Granted for <organisation>".

---

## Step 3 – Nextcloud and Vaultwarden

Nextcloud ([architecture → Setup](docs/architecture.md#7-setup)):

- [ ] *Administration → Sharing*: public links allowed, for upload links also public upload
- [ ] behind a proxy: request body ≥ 12 MB, the login flow `/index.php/login/v2…` reachable
- [ ] recommended: allow the CloudShare server's address in the brute-force settings

Vaultwarden (only if used):

- [ ] `SENDS_ALLOWED=true`
- [ ] log in once as the service account, *Settings → Security → Keys* → API key: note `client_id` (`user.…`) and
      `client_secret`, and its master password

**Checkpoint:** a test user can create a public link in Nextcloud; the Vaultwarden service account can log in.

---

## Step 4 – Install on the Docker host

```bash
git clone https://github.com/mrprohl/CloudShare.git /opt/cloudshare
cd /opt/cloudshare && ./setup.sh
```

`setup.sh` asks for the public address, language, attachment size, link validity, signature, feedback address, the
Entra values of step 2, Nextcloud (address and folders), "Use Vaultwarden" and – with it – its address, API key,
master password and the password link settings. It writes `env/` (directory 700, files 600) and generates
`CREDENTIAL_KEY`.

```bash
grep CREDENTIAL_KEY env/cloudshare.env    # store this key in your password manager
```

Further settings (all optional, list in [manual → Configuration](docs/manual.md#configuration)) go into
`env/cloudshare.env`; `cloudshare.env.example` is the reference.

**Checkpoint:** `ls -l env/` shows `versions.env` and `cloudshare.env`, both `-rw-------`; the key is stored.

---

## Step 5 – Start and check

```bash
cd /opt/cloudshare && docker compose up -d   # or a Portainer/Dockhand stack with docker-compose.yml
```

Caddy fetches the certificate on the first start. The image `ghcr.io/mrprohl/cloudshare` comes from GitHub – nothing
is compiled on the server.

```bash
docker ps --filter name=cloudshare                  # cloudshare "healthy", cloudshare-caddy "Up"
docker logs cloudshare 2>&1 | grep -i vault         # "Vaultwarden vault unlocked" (or "Vaultwarden off …")
curl -s https://cloudshare.example.com/api/config   # JSON
```

Already running a reverse proxy? See the end of [manual → Installation](docs/manual.md#installation).

**Checkpoint:** both containers run, `https://cloudshare.example.com/manifest.xml` shows your address and client ID.

---

## Step 6 – Deploy the add-in

Only now – Microsoft 365 keeps the manifest as it read it. Microsoft 365 admin center → *Settings → Integrated apps →
Upload custom apps* → manifest URL `https://cloudshare.example.com/manifest.xml`; start with your own account, then
roll out ([docs/microsoft-365.md → 3](docs/microsoft-365.md#3-deploying-the-add-in)).

- [ ] the upload creates a **second app registration "CloudShare"** without a secret – leave it, never use its ID for
      `ENTRA_CLIENT_ID`
- [ ] Exchange Online: settings of [docs/microsoft-365.md → 4](docs/microsoft-365.md#4-exchange-online) checked

**Checkpoint:** the add-in is listed under *Integrated apps*. Outlook may take up to 24 hours to show the button.

---

## Step 7 – First test in Outlook

The checks of [docs/microsoft-365.md → 6](docs/microsoft-365.md#6-checking-the-setup), in short:

- [ ] new mail → **CloudShare** → the pane opens without asking for a password → *Connect* connects Nextcloud
- [ ] attach a file, *Insert file as link* → a card is inserted
- [ ] send it to yourself → the password mail follows within about a minute
      (`docker logs cloudshare 2>&1 | grep "Password mail"`)
- [ ] the sent mail in *Sent Items* → **CloudShare** → its link is listed

**Checkpoint:** link and password arrived; the download works with the password.

---

## Step 8 – Optional features

| Feature | Needs | How |
|---|---|---|
| Password by SMS | seven.io account | [manual → Password by SMS](docs/manual.md#password-by-sms-optional) |
| Download confirmation, upload notification | application permission `Mail.Send` (restricted to the CloudShare users) | [docs/microsoft-365.md → 5](docs/microsoft-365.md#5-optional-application-permission-for-notices-and-the-delivery-proof), `DOWNLOAD_NOTICE`, `UPLOAD_NOTICE` |
| Proof of delivery | the same permission | [manual → Proof of delivery](docs/manual.md#proof-of-delivery-optional), `DELIVERY_PROOF` |
| Text templates per department | `env/config/templates.json` | [manual → Text templates](docs/manual.md#text-templates-optional) |
| Sensitivity labels | Microsoft Purview | [manual → Sensitivity labels](docs/manual.md#sensitivity-labels) |
| Own languages or texts | `env/config/locales/` | [manual → Languages](docs/manual.md#languages) |

After a change to `env/cloudshare.env`: `docker compose up -d` (or update the stack).

**Checkpoint:** every feature you switched on passes its check in the linked section.

---

## Afterwards

- **Back up** the volume `cloudshare_data` and `env/` – without `CREDENTIAL_KEY` the stored connections cannot be read.
- **Update** by re-pulling the image ([manual → Operation](docs/manual.md#operation)); renew the client secret
  before it expires.
- Something does not work: [manual → Troubleshooting](docs/manual.md#troubleshooting) and
  [docs/microsoft-365.md → 8](docs/microsoft-365.md#8-errors-and-their-causes).
