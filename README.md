# CloudShare – hardened, large Outlook attachments as Nextcloud links

Outlook add-in (classic and new Outlook, Exchange Online) that sends large files as a password-protected Nextcloud
link instead of an attachment. Recipients get the password automatically in a second mail as a Vaultwarden Send.
How it works and how Entra ID, Nextcloud and Vaultwarden are set up: [docs/architecture.md](docs/architecture.md).

## What it looks like in Outlook

![CloudShare in a new-Outlook compose window: "CloudShare" in the ribbon, the inserted card with download buttons and profile footer in the mail, and the CloudShare taskpane on the right](docs/images/outlook-compose.png)

*Composing a mail: the ribbon button **CloudShare** opens the taskpane. Files are uploaded to the sender's
Nextcloud and inserted as a card – one row per file with a download button, expiry date and a footer filled from
the sender's Microsoft 365 profile (`MAIL_FOOTER`). The password follows automatically in a second mail after
sending. The taskpane offers a random or own password, link validity and **Dateien anfordern** (upload link).*

<img src="docs/images/outlook-my-shares.png" alt="CloudShare taskpane, tab 'Meine Freigaben' with download status, received uploads, expired links and revoke confirmation" width="376" align="right">

*Tab **Meine Freigaben**: the user's CloudShare links with download status ("2× heruntergeladen"), received files
for upload links, expired links, **Link kopieren** and **Zurückziehen** (with confirmation).*

*The images show example data ("Max Mustermann", example.com). They are rendered from
[docs/mockups](docs/mockups) with the real taskpane styles and icon:*
```powershell
& "C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe" --headless=new --hide-scrollbars `
  --window-size=1440,1010 --screenshot="docs\images\outlook-compose.png" "file:///<repo>/docs/mockups/compose-send.html"
& "C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe" --headless=new --hide-scrollbars `
  --window-size=376,640 --screenshot="docs\images\outlook-my-shares.png" "file:///<repo>/docs/mockups/shares-pane.html"
```

<br clear="right">

## Requirements

- Microsoft 365 / Exchange Online, classic Outlook (Windows) or new Outlook / Outlook on the web
- Nextcloud with public link sharing (any login method – users connect once via Login Flow v2)
- Vaultwarden with Sends enabled and a service account for CloudShare
- A Docker host with Docker Compose ≥ 2.20 and a **central Caddy** acting as reverse proxy for several stacks
  (Let's Encrypt). Nextcloud can run behind the same Caddy – CloudShare then reaches it internally.

## Deployment

The examples use `cloudshare.example.com` (CloudShare), `cloud.example.com` (Nextcloud) and
`vault.example.com` (Vaultwarden) – replace them with your own names.

```
Outlook add-in ─HTTPS─► central Caddy (Let's Encrypt) ─cloudshare-proxy─► cloudshare:3000
cloudshare ─cloudshare-proxy─► central Caddy ─► Nextcloud (same host, via extra_hosts – optional)
cloudshare ─egress─► Entra ID/Graph (password mail) · Vaultwarden (Send)
```

## Layout `/opt/cloudshare` (Git clone)

| Path | Content |
|---|---|
| `docker-compose.yml` | stub for Portainer/Dockhand (`include` + variables) |
| `compose.yml` | CloudShare container |
| `caddy/cloudshare.caddy` | site block for the central Caddy (mounted read-only at the same path), written by `setup.sh` from `caddy/cloudshare.caddy.example` |
| `env/versions.env` | `CLOUDSHARE_VERSION` (`latest`), `NEXTCLOUD_HOST`, `PROXY_IP` |
| `env/cloudshare.env` | Entra ID, Nextcloud, Vaultwarden service account, `CREDENTIAL_KEY` – **secrets** |
| volume `cloudshare_data` | users' encrypted Nextcloud app passwords – **back up** |
| `setup.sh` | new installation: creates the network `cloudshare-proxy`, writes `env/` and the Caddy site block |

GitHub builds the image `ghcr.io/mrprohl/cloudshare` on every push to `main` (`.github/workflows/image.yml`) –
nothing is built on the server. `env/` and `caddy/cloudshare.caddy` are not in Git.

Hardening: read-only, `cap_drop: ALL`, `no-new-privileges`, limits, log rotation, no `ports:`, runs as user 1000.
Networks: `cloudshare-proxy` (internal, default 172.30.13.0/24, central Caddy = 172.30.13.2) and `egress`
(outbound). `bw serve` listens on 127.0.0.1 inside the container only. The site block sets no `X-Frame-Options`,
because new Outlook and Outlook on the web show the taskpane in an iframe.

## New installation

### 1. Clone and set up
```bash
git clone https://github.com/mrprohl/CloudShare.git /opt/cloudshare
cd /opt/cloudshare && ./setup.sh     # creates cloudshare-proxy, asks for Entra, Nextcloud, Vaultwarden
grep CREDENTIAL_KEY env/cloudshare.env   # store in your password manager
```
The network defaults to 172.30.13.0/24; if that is taken, run `PROXY_NET=10.99.0 ./setup.sh` (first three octets
of a free /24).

### 2. Central Caddy (`/opt/caddy/compose.yml`)
Add the network and the site directory:
```yaml
    networks:
      # …
      cloudshare-proxy: { ipv4_address: 172.30.13.2 }
    volumes:
      # …
      - /opt/cloudshare/caddy:/opt/cloudshare/caddy:ro

networks:
  # …
  cloudshare-proxy:
    external: true            # created by cloudshare/setup.sh
```
If the central `Caddyfile` imports the services one by one, add `import /opt/cloudshare/caddy/*.caddy`. Then
recreate the Caddy stack (new network and mount need a redeploy, a reload is not enough).

### 3. DNS and deployment
- DNS: `cloudshare.example.com` → public address of the Docker host (A/AAAA or CNAME). Ports 80/443 must reach the
  central Caddy (directly or forwarded by an upstream proxy); Caddy fetches the certificate on the first request.
- Portainer/Dockhand: stack "cloudshare" → contents of `/opt/cloudshare/docker-compose.yml` → Deploy
  (or `cd /opt/cloudshare && docker compose up -d`).
- Vaultwarden: `SENDS_ALLOWED=true`.
- Microsoft 365 admin center → *Integrated apps* → *Upload custom apps* → manifest URL
  `https://cloudshare.example.com/manifest.xml` (first only for your own account). Alternative with a clear error
  message: Exchange Online PowerShell,
  `New-App -OrganizationApp -Url https://cloudshare.example.com/manifest.xml -ProvidedTo Everyone -DefaultStateForUser Enabled`.
## Tests

Backend and add-in have automated tests (vitest) in `server/test/` and `addin/test/`:

- **API** (`api.test.ts`, via `fastify.inject`): login required, users cannot touch each other's uploads, jobs or
  links, only CloudShare links can be revoked, file names/sizes/chunk numbers are validated, limits are enforced.
- **Outlook event handlers** (`commands.test.ts`, Office.js mock): `enforce`/`suggest`, opt-out, arming before the
  prompt, Mailbox 1.12 vs 1.14, shared mailboxes, server unreachable (cached configuration), info bar.
- Password rules and generator, encryption, Graph, Nextcloud client, password-job flow, footer template, mail card.

Network calls are mocked. Before every image build GitHub runs: type check and tests for both packages,
`npm audit` (runtime dependencies, high/critical), manifest validation (`npm run validate:manifest`), ShellCheck for
the shell scripts, and a Trivy scan of the image (critical/high with an available fix) before it is pushed. Locally:
```bash
cd server && npm install && npm test
cd addin && npm install && npm test
```

## Updating

- **Application:** stack "cloudshare" → "Update the stack" with **re-pull**. To pin a state:
  `CLOUDSHARE_VERSION=sha-<commit>` in `env/versions.env`.
- **compose.yml from the repository:** `cd /opt/cloudshare && git pull`, then update the stack.
  `caddy/cloudshare.caddy` is local – if `caddy/cloudshare.caddy.example` changed, carry the change over and run
  `docker exec caddy caddy reload --config /etc/caddy/Caddyfile`.
- **Settings:** edit `env/cloudshare.env` → update the stack.

## Checks

```bash
docker ps --filter name=cloudshare                         # "healthy"
docker logs cloudshare 2>&1 | grep -i vault                # "Vaultwarden vault unlocked"
docker exec cloudshare wget -qO- http://127.0.0.1:3000/api/config
curl -s https://cloudshare.example.com/api/config         # from outside: same JSON
```
Note for Nextcloud: all requests from CloudShare arrive from one internal address (cloudshare-proxy). If Nextcloud's
brute-force protection ever throttles them, allow the cloudshare-proxy subnet (default `172.30.13.0/24`) in the "Brute-force settings" app.
