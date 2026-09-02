#!/usr/bin/env bash

set -euo pipefail

# ==========================================================
# CRFFN Licensing - Apps Script Deployment
# ==========================================================

# IMPORTANT:
# Set this to the deployment ID from the CURRENT production /exec URL.
DEPLOYMENT_ID="AKfycbxnpEmnzekbkbSdTiQKixvGSiJvvN0ScdN6PH2KiFWLMdYi6FF2NgX2aaw968nCu2YXHw"

TIMESTAMP="$(date '+%Y-%m-%d %H:%M:%S')"
DESCRIPTION="CRFFN deployment - ${TIMESTAMP}"

echo
echo "======================================================"
echo " CRFFN Licensing Deployment"
echo "======================================================"
echo
echo "Deployment ID:"
echo "$DEPLOYMENT_ID"
echo
echo "Description:"
echo "$DESCRIPTION"
echo

echo "------------------------------------------------------"
echo "1. Files clasp will push"
echo "------------------------------------------------------"
clasp status

echo
read -r -p "Continue and PUSH these files? [y/N]: " CONFIRM

if [[ ! "$CONFIRM" =~ ^[Yy]$ ]]; then
  echo "Deployment cancelled."
  exit 0
fi

echo
echo "------------------------------------------------------"
echo "2. Pushing source to Apps Script"
echo "------------------------------------------------------"
clasp push

echo
echo "------------------------------------------------------"
echo "3. Creating immutable Apps Script version"
echo "------------------------------------------------------"

VERSION_OUTPUT="$(clasp version "$DESCRIPTION")"

echo "$VERSION_OUTPUT"

# clasp normally outputs something such as:
# Created version 106.

VERSION_NUMBER="$(echo "$VERSION_OUTPUT" | grep -Eo '[0-9]+' | tail -1)"

if [[ -z "$VERSION_NUMBER" ]]; then
  echo
  echo "ERROR: Could not determine newly-created version number."
  echo "Deployment has NOT been updated."
  exit 1
fi

echo
echo "Created version: $VERSION_NUMBER"

echo
echo "------------------------------------------------------"
echo "4. Updating EXISTING deployment"
echo "------------------------------------------------------"

clasp deploy \
  --deploymentId "$DEPLOYMENT_ID" \
  --versionNumber "$VERSION_NUMBER" \
  --description "$DESCRIPTION"

echo
echo "------------------------------------------------------"
echo "5. Deployment status"
echo "------------------------------------------------------"

clasp deployments

echo
echo "======================================================"
echo " Deployment complete"
echo "======================================================"
echo
echo "Deployment ID remains:"
echo "$DEPLOYMENT_ID"
echo
echo "Version deployed:"
echo "$VERSION_NUMBER"
echo
echo "The existing /exec URL remains unchanged."
echo