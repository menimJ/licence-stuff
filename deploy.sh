#!/usr/bin/env bash

set -euo pipefail

# ==========================================================
# CRFFN Licensing - Apps Script Deployment
# ==========================================================

DEPLOYMENT_ID="AKfycbxnpEmnzekbkbSdTiQKixvGSiJvvN0ScdN6PH2KiFWLMdYi6FF2NgX2aaw968nCu2YXHw"

TIMESTAMP="$(date '+%Y-%m-%d %H:%M:%S')"

echo
echo "======================================================"
echo " CRFFN Licensing Deployment"
echo "======================================================"
echo

read -r -p "What changed in this deployment? " CHANGE_DESCRIPTION

if [[ -z "$CHANGE_DESCRIPTION" ]]; then
  echo
  echo "ERROR: Deployment description cannot be empty."
  exit 1
fi

DESCRIPTION="${CHANGE_DESCRIPTION} - ${TIMESTAMP}"
COMMIT_MESSAGE="Deploy: ${DESCRIPTION}"

echo
echo "Deployment ID:"
echo "$DEPLOYMENT_ID"
echo
echo "Change description:"
echo "$CHANGE_DESCRIPTION"
echo
echo "Commit message:"
echo "$COMMIT_MESSAGE"
echo

# ==========================================================
# 1. Git status
# ==========================================================

echo "------------------------------------------------------"
echo "1. Git working tree status"
echo "------------------------------------------------------"

git status --short

echo
echo "------------------------------------------------------"
echo "2. Git change summary"
echo "------------------------------------------------------"

git diff --stat

echo
echo "------------------------------------------------------"
echo "3. Files clasp will push"
echo "------------------------------------------------------"

clasp status

echo
read -r -p "Commit and deploy these changes? [y/N]: " CONFIRM

if [[ ! "$CONFIRM" =~ ^[Yy]$ ]]; then
  echo
  echo "Deployment cancelled."
  exit 0
fi

# ==========================================================
# 2. Git commit
# ==========================================================

echo
echo "------------------------------------------------------"
echo "4. Creating Git commit"
echo "------------------------------------------------------"

git add -A

if git diff --cached --quiet; then
  echo
  echo "No new Git changes to commit."
  echo "Continuing with the latest existing commit."
else
  echo
  echo "Files being committed:"
  git status --short

  echo
  echo "Commit message:"
  echo "$COMMIT_MESSAGE"
  echo

  git commit -m "$COMMIT_MESSAGE"

  echo
  echo "Git commit created successfully."
fi

# ==========================================================
# 3. Show commit being deployed
# ==========================================================

echo
echo "------------------------------------------------------"
echo "5. Git commit being deployed"
echo "------------------------------------------------------"

git log -1 --oneline

# ==========================================================
# 4. Push source
# ==========================================================

echo
echo "------------------------------------------------------"
echo "6. Pushing source to Apps Script"
echo "------------------------------------------------------"

clasp push

# ==========================================================
# 5. Create Apps Script version
# ==========================================================

echo
echo "------------------------------------------------------"
echo "7. Creating immutable Apps Script version"
echo "------------------------------------------------------"

VERSION_OUTPUT="$(clasp version "$DESCRIPTION")"

echo "$VERSION_OUTPUT"

VERSION_NUMBER="$(
  echo "$VERSION_OUTPUT" |
  grep -Eo '[0-9]+' |
  tail -1
)"

if [[ -z "$VERSION_NUMBER" ]]; then
  echo
  echo "ERROR: Could not determine newly-created version number."
  echo "The Git commit exists, but the deployment was NOT updated."
  exit 1
fi

echo
echo "Created Apps Script version: $VERSION_NUMBER"

# ==========================================================
# 6. Update existing deployment
# ==========================================================

echo
echo "------------------------------------------------------"
echo "8. Updating EXISTING deployment"
echo "------------------------------------------------------"

clasp deploy \
  --deploymentId "$DEPLOYMENT_ID" \
  --versionNumber "$VERSION_NUMBER" \
  --description "$DESCRIPTION"

# ==========================================================
# 7. Verify deployment
# ==========================================================

echo
echo "------------------------------------------------------"
echo "9. Deployment status"
echo "------------------------------------------------------"

clasp deployments

# ==========================================================
# Final summary
# ==========================================================

echo
echo "======================================================"
echo " CRFFN Licensing Deployment Complete"
echo "======================================================"
echo
echo "Change:"
echo "$CHANGE_DESCRIPTION"
echo
echo "Git commit:"
git log -1 --oneline
echo
echo "Apps Script version:"
echo "$VERSION_NUMBER"
echo
echo "Deployment ID:"
echo "$DEPLOYMENT_ID"
echo
echo "Web App URL:"
echo "https://script.google.com/macros/s/$DEPLOYMENT_ID/exec"
echo
echo "Admin URL:"
echo "https://script.google.com/macros/s/$DEPLOYMENT_ID/exec?view=admin"
echo
echo "======================================================"