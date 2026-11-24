## What and why

<!-- One or two sentences; the commit messages can hold the details. -->

## Documentation

Tick what applies, or say why it is not needed. The CI checks part of it (server/test/docs.test.ts, README pictures in
the UI tests); the rest is up to us.

- [ ] **Help in the pane** (`help.*` in `locales/*.json`, sections in `taskpane.html`) describes the changed behaviour – in plain words, every language
- [ ] **INSTALL.md** – steps and checks still match setup.sh and the linked guides
- [ ] **docs/manual.md** – features, settings table (`env`), Personal settings, Languages, troubleshooting
- [ ] **docs/architecture.md** – flows, state and data, HTTP API table, security
- [ ] **docs/microsoft-365.md** – permissions, deployment, manifest version
- [ ] **CLAUDE.md** – conventions and the module map
- [ ] **cloudshare.env.example** / **setup.sh** – new or changed settings
- [ ] **README pictures of the pane** (`docs/images/taskpane-*.png`) – the UI tests fail if they are outdated; take them from the artifact `readme-pictures`
- [ ] **Mockup pictures** (`docs/images/outlook-compose.png`, `mail-received.png`, `password-mail.png`, `sender-notices.png`) – re-render from `docs/mockups` when the card, the password mail or a notice changed
- [ ] **Manifest** – new `<Version>` if `manifest.xml` changed
