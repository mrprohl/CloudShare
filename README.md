# CloudShare

Outlook add-in for Microsoft 365 / Exchange Online: large and confidential files go to the sender's own **Nextcloud**
and into the mail as a password-protected link. The password follows on its own once the mail has really been sent –
in a second mail or by text message, as a link from the own **Vaultwarden** or (without Vaultwarden) right in the
mail. Optional: confidential messages, upload links and a **proof of delivery** with a signed PDF.

![A compose window in new Outlook: the CloudShare button in the ribbon, a download card with three files in the mail, and the CloudShare pane on the right](docs/images/outlook-compose.png)

## Structure

One container (Fastify backend and the add-in) behind its own Caddy, deployed with Docker Compose to `/opt/cloudshare`:

| Folder | Content |
|---|---|
| [`addin/`](addin/) | Outlook add-in: manifest, taskpane, event handlers (Office.js, TypeScript) |
| [`server/`](server/) | backend: Microsoft Graph, Nextcloud, Vaultwarden, SMS, proof of delivery |
| [`locales/`](locales/) | texts of the add-in, the mails and the SMS (German, English) |
| [`caddy/`](caddy/), [`docker/`](docker/) | reverse proxy and container start |
| [`ci/`](ci/) | checks of the stack: setup, smoke, upgrade and load test |
| [`docs/`](docs/) | documentation, pictures and their mockups |

## Documentation

- [docs/manual.md](docs/manual.md) – what senders and recipients see, installation, settings, operation, troubleshooting
- [docs/microsoft-365.md](docs/microsoft-365.md) – Microsoft 365 step by step: Entra app registration, add-in deployment, Exchange
- [docs/architecture.md](docs/architecture.md) – components, data flows, state, HTTP API, security

## Quick start

```bash
git clone https://github.com/mrprohl/CloudShare.git /opt/cloudshare
cd /opt/cloudshare && ./setup.sh          # asks for everything, writes env/
docker compose up -d                      # or as a Portainer/Dockhand stack from docker-compose.yml
```

Before that: the Entra app registration ([docs/microsoft-365.md](docs/microsoft-365.md)); afterwards deploy the add-in
in the Microsoft 365 admin center – both step by step in the [manual](docs/manual.md#installation).

## Not in this repository

Secrets and everything generated on the server stay out: `env/` (written by `setup.sh`), the data volume and own
texts in `env/config/`. The documentation uses example names only (`cloudshare.example.com`, `cloud.example.com`,
`vault.example.com`).
