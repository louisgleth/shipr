#!/usr/bin/env bash
set -euo pipefail

ROOT_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")/.." && pwd)"
cd "$ROOT_DIR"

if ! git rev-parse --is-inside-work-tree >/dev/null 2>&1; then
  echo "Error: this command must run inside the git repo."
  exit 1
fi

COMMIT_MSG="${*:-Live deploy $(date '+%Y-%m-%d %H:%M')}"

if [[ -n "$(git status --porcelain)" ]]; then
  echo "Committing local changes..."
  git add -A
  git commit -m "$COMMIT_MSG"
else
  echo "No local changes to commit."
fi

echo "Pushing frontend to origin/main..."
git push origin main

echo "Deploying Worker..."
npx wrangler deploy --config wrangler.shopify-api.toml

echo ""
echo "Done. Frontend push + Worker deploy completed."
