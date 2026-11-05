#!/usr/bin/env bash
# One-command deploy: pull latest from git, push to Apps Script, update
# the live deployment.
#
# Usage:  ./deploy.sh                # deploys with auto-timestamp label
#         ./deploy.sh "my notes"     # deploys with custom label
#
# Requires clasp installed + logged in (`clasp login` once).
# .clasp.json at the repo root tells clasp where to send the code.
#
# DEPLOYMENT_ID is the active web-app deployment. Passing it to
# `clasp deploy` UPDATES that deployment instead of creating a new one,
# which keeps the /exec URL stable and avoids the 20-deployment cap.

set -euo pipefail

DEPLOYMENT_ID="AKfycbz4ySRzlSamI9bkgnEjlQKh1ypawUX5I25D8I9j1Uy4aSJNpa7tRDrYPgczVcpJypFLGw"

BRANCH="$(git symbolic-ref --short HEAD 2>/dev/null || echo detached)"
LABEL="${1:-Auto-deploy $(date '+%Y-%m-%d %H:%M') from ${BRANCH}}"

echo "→ Pulling latest from git (${BRANCH})…"
git pull --ff-only

echo "→ Pushing code to Apps Script…"
clasp push -f

echo "→ Updating deployment: ${LABEL}"
clasp deploy --deploymentId "${DEPLOYMENT_ID}" --description "${LABEL}"

echo ""
echo "✅ Done. Hard-refresh the dashboard to see the new code."
