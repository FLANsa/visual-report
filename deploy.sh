#!/usr/bin/env bash
set -euo pipefail

PROJECT_ID="${PROJECT_ID:-visual-report-sa}"

echo "→ نشر Visual Report على Firebase Hosting (Spark — مجاني)"
firebase use "$PROJECT_ID"
firebase deploy --only hosting --project "$PROJECT_ID"
echo ""
echo "✓ https://${PROJECT_ID}.web.app"
