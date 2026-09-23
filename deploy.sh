#!/usr/bin/env bash
# One-command deploy: pull latest from git, push to Apps Script, activate.
#
# Usage:  ./deploy.sh                # deploys with auto-timestamp label
#         ./deploy.sh "my notes"     # deploys with custom label
#
# Requires clasp installed + logged in (`clasp login` once).
# .clasp.json at the repo root tells clasp where to send the code.

set -euo pipefail

BRANCH="$(git symbolic-ref --short HEAD 2>/dev/null || echo detached)"
LABEL="${1:-Auto-deploy $(date '+%Y-%m-%d %H:%M') from ${BRANCH}}"

echo "→ Pulling latest from git (${BRANCH})…"
git pull --ff-only

echo "→ Pushing code to Apps Script…"
clasp push -f

echo "→ Creating + activating deployment: ${LABEL}"
clasp deploy --description "${LABEL}"

echo ""
echo "✅ Done. Hard-refresh the dashboard to see the new code."
