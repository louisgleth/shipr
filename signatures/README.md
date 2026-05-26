# Shipide Gmail signatures

This folder contains the dark Shipide Gmail signature template and deployment script for Google Workspace mailboxes.

## Files

- `shipide-signature-template.html`: deployable dark signature template.
- `users.json`: per-user data for `louis@shipide.com` and `noah@shipide.com`.
- `deploy-gmail-signatures.js`: renders and deploys signatures with the Gmail API.
- `group-11182.png`: corner artwork used by the signature preview/template.
- `shipide-signature-preview.html`: visual preview page.

## Required Google setup

You need a Google Cloud service account with domain-wide delegation enabled, then authorize the service account client ID in Google Workspace Admin Console with this scope:

```text
https://www.googleapis.com/auth/gmail.settings.basic
```

Service account JSON keys are not required if your organization blocks key creation. Use the keyless `signJwt` flow below.

## Required public assets

Before deploying, these URLs must be reachable publicly:

```text
https://portal.shipide.com/assets/shipide-logo-white.png
https://portal.shipide.com/signatures/group-11182.png
```

The current local files are:

```text
/Users/louis/shipide/portal/assets/shipide-logo-white.png
/Users/louis/shipide/portal/signatures/group-11182.png
```

## Render locally

```bash
node signatures/deploy-gmail-signatures.js --dry-run
```

This writes rendered HTML files to `signatures/dist/`.

## Deploy with keyless service account auth

If service account key creation is disabled, keep that policy enabled and use IAM Credentials signing instead.

Required one-time setup:

1. Enable the IAM Service Account Credentials API in the Google Cloud project.
2. Grant your admin user `roles/iam.serviceAccountTokenCreator` on the service account.
3. Install/login with Google Cloud CLI:

```bash
gcloud auth login
gcloud config set project YOUR_PROJECT_ID
```

Deploy:

```bash
node signatures/deploy-gmail-signatures.js --service-account-email SERVICE_ACCOUNT_EMAIL
```

Example:

```bash
node signatures/deploy-gmail-signatures.js --service-account-email gmail-signatures@YOUR_PROJECT_ID.iam.gserviceaccount.com
```

## Deploy with a JSON key

```bash
node signatures/deploy-gmail-signatures.js --credentials /path/to/service-account.json
```

The script impersonates each user in `users.json` and updates their primary Gmail send-as signature.
