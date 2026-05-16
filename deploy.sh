#!/usr/bin/env bash
set -euo pipefail

PROJECT_ID="${PROJECT_ID:-visual-report-sa}"

echo "→ Deploying Visual Report to Firebase Hosting (Spark — free)"
firebase use "$PROJECT_ID"
firebase deploy --only hosting --project "$PROJECT_ID"
echo ""
echo "✓ Done: https://${PROJECT_ID}.web.app"
