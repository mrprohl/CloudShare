# Third-party software

CloudShare itself is under the [MIT License](LICENSE). The image `ghcr.io/mrprohl/cloudshare` also contains the
software below, each under its own license – unchanged, as published by its authors. Its license texts are in the
image next to each package (`/app/node_modules/<package>/LICENSE`, `/opt/bw/node_modules/@bitwarden/cli/`) and at the
links below.

## In the image

| Software | Used for | License |
|---|---|---|
| [Bitwarden CLI](https://github.com/bitwarden/clients/tree/main/apps/cli) (`@bitwarden/cli`, version `BW_CLI_VERSION` in the `Dockerfile`) | `bw serve` – Vaultwarden Sends (not started with `VAULTWARDEN_ENABLED=false`) | GPL-3.0; parts of its repository under the Bitwarden License v1.0 – see its [LICENSE.txt](https://github.com/bitwarden/clients/blob/main/LICENSE.txt) |
| [Fastify](https://github.com/fastify/fastify), [@fastify/rate-limit](https://github.com/fastify/fastify-rate-limit), [@fastify/static](https://github.com/fastify/fastify-static) | web server | MIT |
| [@azure/msal-node](https://github.com/AzureAD/microsoft-authentication-library-for-js) | Microsoft sign-in (On-Behalf-Of, app token) | MIT |
| [jose](https://github.com/panva/jose) | checking the Office SSO token | MIT |
| [Node.js](https://github.com/nodejs/node) (base image `node:24-alpine`) | runtime | MIT (with the licenses of its bundled parts, see its [LICENSE](https://github.com/nodejs/node/blob/main/LICENSE)) |
| [tini](https://github.com/krallin/tini) | init process | MIT |
| Alpine Linux packages of the base image | operating system | their own free licenses |

Their own dependencies come with them, each with its license file in `node_modules`.

**Bitwarden CLI and the GPL:** CloudShare runs the CLI as a separate program and talks to it over HTTP on
`127.0.0.1` – it does not link it. Whoever passes the image on passes the CLI on as well and has to keep to the GPL-3.0
for it: its source code is in the [bitwarden/clients](https://github.com/bitwarden/clients) repository (tag
`cli-v<version>`). "Bitwarden" is a trademark of Bitwarden Inc.; CloudShare is not related to Bitwarden or
Vaultwarden.

## Not in the image

Run next to CloudShare or used by it, each under its own license: [Caddy](https://github.com/caddyserver/caddy)
(own container, Apache-2.0), Nextcloud, Vaultwarden, Microsoft Office.js (loaded by Outlook from Microsoft), seven.io
(service). The build and test tools of `addin/` and `server/` (devDependencies) are not part of the image.
