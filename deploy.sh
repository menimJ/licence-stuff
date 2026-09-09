#!/usr/bin/env bash

set -euo pipefail

# ==========================================================
# CRFFN Licensing - Apps Script Deployment
# ==========================================================

DEPLOYMENT_ID="AKfycbz19WiC1nq8ieMvjf9AraQ83mlv_wVCJNQRzWlJMdk8fVLOEFOlPSxH1hKgGuZetfm2eg"

CLASP_USER="crffn-new-owner"

TIMESTAMP="$(date '+%Y-%m-%d %H:%M:%S')"

echo
echo "======================================================"
echo " CRFFN Licensing Deployment"
echo "======================================================"
echo

# ==========================================================
# 0. Confirm clasp account
# ==========================================================

echo "------------------------------------------------------"
echo "0. Authorized clasp account"
echo "------------------------------------------------------"

clasp show-authorized-user \
  --user "$CLASP_USER"

echo

# ==========================================================
# 1. Check Git working tree
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

if [[ -n "$(git status --porcelain)" ]]; then
  HAS_CHANGES=true
else
  HAS_CHANGES=false
fi

# ==========================================================
# 2. Build deployment description
# ==========================================================

if [[ "$HAS_CHANGES" == true ]]; then

  echo
  echo "Local changes detected."
  echo

  read -r -p "What changed in this deployment? " CHANGE_DESCRIPTION

  if [[ -z "$CHANGE_DESCRIPTION" ]]; then
    echo
    echo "ERROR: Deployment description cannot be empty."
    exit 1
  fi

  DESCRIPTION="${CHANGE_DESCRIPTION} - ${TIMESTAMP}"
  COMMIT_MESSAGE="Deploy: ${DESCRIPTION}"

else

  echo
  echo "No local Git changes detected."
  echo "No new Git commit is required."
  echo

  CHANGE_DESCRIPTION="Redeploy existing source"
  DESCRIPTION="${CHANGE_DESCRIPTION} - ${TIMESTAMP}"
  COMMIT_MESSAGE=""

fi

echo
echo "Deployment ID:"
echo "$DEPLOYMENT_ID"

echo
echo "Clasp profile:"
echo "$CLASP_USER"

echo
echo "Deployment description:"
echo "$DESCRIPTION"
echo

# ==========================================================
# 3. Show files clasp will push
# ==========================================================

echo "------------------------------------------------------"
echo "3. Files clasp will push"
echo "------------------------------------------------------"

clasp show-file-status \
  --user "$CLASP_USER"

echo

if [[ "$HAS_CHANGES" == true ]]; then
  read -r -p "Commit and deploy these changes? [y/N]: " CONFIRM
else
  read -r -p "No Git changes. Deploy the current source anyway? [y/N]: " CONFIRM
fi

if [[ ! "$CONFIRM" =~ ^[Yy]$ ]]; then
  echo
  echo "Deployment cancelled."
  exit 0
fi

# ==========================================================
# 4. Commit local changes
# ==========================================================

if [[ "$HAS_CHANGES" == true ]]; then

  echo
  echo "------------------------------------------------------"
  echo "4. Creating Git commit"
  echo "------------------------------------------------------"

  git add -A

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

else

  echo
  echo "------------------------------------------------------"
  echo "4. Git commit"
  echo "------------------------------------------------------"
  echo
  echo "No changes detected."
  echo "Skipping Git commit."

fi

# ==========================================================
# 5. Show commit being deployed
# ==========================================================

echo
echo "------------------------------------------------------"
echo "5. Git commit being deployed"
echo "------------------------------------------------------"

git log -1 --oneline

# ==========================================================
# 6. Push source to Apps Script
# ==========================================================

echo
echo "------------------------------------------------------"
echo "6. Pushing source to Apps Script"
echo "------------------------------------------------------"

clasp push \
  --user "$CLASP_USER"

# ==========================================================
# 7. Create immutable Apps Script version
# ==========================================================

echo
echo "------------------------------------------------------"
echo "7. Creating immutable Apps Script version"
echo "------------------------------------------------------"

VERSION_OUTPUT="$(
  clasp create-version \
    "$DESCRIPTION" \
    --user "$CLASP_USER"
)"

echo "$VERSION_OUTPUT"

VERSION_NUMBER="$(
  echo "$VERSION_OUTPUT" |
  grep -Eo '[0-9]+' |
  tail -1
)"

if [[ -z "$VERSION_NUMBER" ]]; then
  echo
  echo "ERROR: Could not determine newly-created version number."
  echo "The live deployment was NOT updated."
  exit 1
fi

echo
echo "Created Apps Script version:"
echo "$VERSION_NUMBER"

# ==========================================================
# 8. Update EXISTING deployment
# ==========================================================

echo
echo "------------------------------------------------------"
echo "8. Updating EXISTING deployment"
echo "------------------------------------------------------"

clasp create-deployment \
  --deploymentId "$DEPLOYMENT_ID" \
  --versionNumber "$VERSION_NUMBER" \
  --description "$DESCRIPTION" \
  --user "$CLASP_USER"

# ==========================================================
# 9. Verify deployment
# ==========================================================

echo
echo "------------------------------------------------------"
echo "9. Deployment status"
echo "------------------------------------------------------"

clasp list-deployments \
  --user "$CLASP_USER"

# ==========================================================
# 10. Push Git commit to GitHub
# ==========================================================

echo
echo "------------------------------------------------------"
echo "10. GitHub push"
echo "------------------------------------------------------"

if [[ "$HAS_CHANGES" == true ]]; then

  git push origin main

  echo
  echo "GitHub updated successfully."

else

  echo
  echo "No new Git commit."
  echo "Skipping GitHub push."

fi

# ==========================================================
# 11. Final result
# ==========================================================

WEB_APP_URL="https://script.google.com/macros/s/${DEPLOYMENT_ID}/exec"

echo
echo "======================================================"
echo " DEPLOYMENT COMPLETE"
echo "======================================================"

echo
echo "Apps Script version:"
echo "$VERSION_NUMBER"

echo
echo "Deployment ID:"
echo "$DEPLOYMENT_ID"

echo
echo "Production URL:"
echo "$WEB_APP_URL"

echo
echo "Existing Admin Portal:"
echo "${WEB_APP_URL}?view=admin"

echo
echo "New Admin Dashboard:"
echo "${WEB_APP_URL}?view=dashboard"

echo
echo "Clasp profile used:"
echo "$CLASP_USER"

echo
echo "======================================================"
echo