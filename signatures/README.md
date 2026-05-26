# Shipide Gmail signatures

This folder contains the dark Shipide Gmail signature template and deployment script for Google Workspace mailboxes.

## Files

- `shipide-signature-template.html`: deployable dark signature template.
- `users.json`: per-user data for `louis@shipide.com` and `noah@shipide.com`.
- `deploy-gmail-signatures.js`: renders and deploys signatures with the Gmail API.
- `group-11182.svg`: corner artwork used by the signature preview/template.
- `shipide-signature-preview.html`: visual preview page.

## Required Google setup

You need a Google Cloud service account with domain-wide delegation enabled, then authorize the service account client ID in Google Workspace Admin Console with this scope:

```text
https://www.googleapis.com/auth/gmail.settings.basic
```

Download the service account JSON key locally. Do not commit it.

## Required public assets

Before deploying, these URLs must be reachable publicly:

```text
https://shipide.com/assets/shipide-logo-white.svg
https://shipide.com/signatures/group-11182.svg
```

The current local files are:

```text
/Users/louis/shipide/portal/assets/shipide-logo-white.svg
/Users/louis/shipide/portal/signatures/group-11182.svg
```

## Render locally

```bash
node signatures/deploy-gmail-signatures.js --dry-run
```

This writes rendered HTML files to `signatures/dist/`.

## Deploy

```bash
node signatures/deploy-gmail-signatures.js --credentials /path/to/service-account.json
```

The script impersonates each user in `users.json` and updates their primary Gmail send-as signature.
