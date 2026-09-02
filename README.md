# CRFFN Licensing System

Google Apps Script-based licensing system for CRFFN.

The system handles:

- applicant submission
- application ID generation
- applicant portal access
- payment proof upload
- supporting document upload
- admin review
- licence generation
- stamped licence upload
- licence release
- automated emails
- scheduled/background processing
- test/reset utilities

---

## Technology

The project uses:

- Google Apps Script
- Google Sheets
- Google Forms
- Google Drive
- MailApp
- HTML Service
- clasp
- Git

The Apps Script project is developed locally and synchronized with Google Apps Script using `clasp`.

---

## Project Structure

Main Apps Script files:

```text
AdminAuth.js
AdminDocumentViewer.html
AdminPortal.js
AdminPortalPage.html
ApplicantPortal.html
ApplicantSupportingDocuments.html
appsscript.json
Code.js
Email.js
GoLiveReset.js
Licence.js
LicenceDownload.html
Payment.js
PerformanceCaches.js
PerformanceTests.js
Portal.js
StampedLicenceWorkflow.js
SupportingDocuments.js
SystemJobs.js
```

Local-only development files:

```text
.clasp.json
.claspignore
.git/
.gitignore
deploy.sh
README.md
```

These should not be pushed into Apps Script.

---

## Requirements

Install Node.js and clasp.

```bash
node --version
npm --version
```

Install clasp:

```bash
npm install -g @google/clasp
```

Check version:

```bash
clasp --version
```

---

## Google Apps Script Authentication

Login:

```bash
clasp login
```

Check the currently authorized Google account:

```bash
clasp show-authorized-user
```

A second Google account can be saved with a label:

```bash
clasp login --user crffn-owner
```

Check that account:

```bash
clasp show-authorized-user --user crffn-owner
```

The Google account used by clasp must have edit access to the Apps Script project.

---

## Connecting to the Apps Script Project

The local project connection is stored in:

```text
.clasp.json
```

The Script ID can be found in:

```text
Apps Script
→ Project Settings
→ IDs
→ Script ID
```

Clone an existing Apps Script project:

```bash
clasp clone SCRIPT_ID
```

---

## URLs and IDs Used by the System

The project uses several different Google URLs and IDs. They are not interchangeable.

### 1. Google Apps Script Editor URL

Example format:

```text
https://script.google.com/home/projects/SCRIPT_ID/edit
```

The value between `/projects/` and `/edit` is the **Script ID**.

Example:

```text
https://script.google.com/home/projects/1ABCDEF123456789/edit
```

Script ID:

```text
1ABCDEF123456789
```

This Script ID is what `clasp clone` uses:

```bash
clasp clone 1ABCDEF123456789
```

---

### 2. Deployed Web App URL

A deployed Apps Script web app normally looks like:

```text
https://script.google.com/macros/s/DEPLOYMENT_ID/exec
```

Example:

```text
https://script.google.com/macros/s/AKfycbxxxxxxxxxxxxxxxx/exec
```

The value between `/s/` and `/exec` is the **Deployment ID**.

Example:

```text
AKfycbxxxxxxxxxxxxxxxx
```

This is the ID used when redeploying to the same web-app URL:

```bash
clasp deploy \
  --deploymentId AKfycbxxxxxxxxxxxxxxxx \
  --versionNumber 106 \
  --description "CRFFN release 106"
```

Updating the same deployment ID keeps the same `/exec` URL.

Creating a new deployment usually creates a new deployment ID and therefore a new `/exec` URL.

---

### 3. HEAD / Development Deployment

`clasp deployments` may show a deployment such as:

```text
AKfycbxxxx @HEAD
```

`@HEAD` points to the current project source rather than a fixed immutable version.

Do not assume `@HEAD` is the production deployment.

For normal production-style releases, use the deployment ID tied to the actual working `/exec` URL and deploy a numbered version to it.

---

### 4. Google Form URL

Public Google Form URL example:

```text
https://docs.google.com/forms/d/e/FORM_PUBLIC_ID/viewform
```

or, depending on form type:

```text
https://docs.google.com/forms/d/FORM_ID/edit
```

Transferring ownership of the existing Form does not normally change its Form ID or public Form URL.

Do not create a replacement Form unless necessary.

---

### 5. Google Sheet URL

Example:

```text
https://docs.google.com/spreadsheets/d/SPREADSHEET_ID/edit
```

The value between `/d/` and `/edit` is the **Spreadsheet ID**.

Example:

```text
https://docs.google.com/spreadsheets/d/1SheetABC123/edit
```

Spreadsheet ID:

```text
1SheetABC123
```

If ownership of the existing Sheet is transferred, the Spreadsheet ID and URL remain the same.

---

### 6. Google Drive Folder URL

Example:

```text
https://drive.google.com/drive/folders/FOLDER_ID
```

The value after `/folders/` is the **Folder ID**.

Example:

```text
https://drive.google.com/drive/folders/1FolderABC123
```

Folder ID:

```text
1FolderABC123
```

The system should continue using the same folder IDs after ownership transfer.

Changing ownership of an existing folder does not create a new folder ID.

---

### 7. Google Drive File URL

Example:

```text
https://drive.google.com/file/d/FILE_ID/view
```

The value between `/d/` and `/view` is the **File ID**.

For Google Docs:

```text
https://docs.google.com/document/d/DOCUMENT_ID/edit
```

For Google Sheets:

```text
https://docs.google.com/spreadsheets/d/SPREADSHEET_ID/edit
```

For Google Forms:

```text
https://docs.google.com/forms/d/FORM_ID/edit
```

The same underlying file/document ID remains valid after ownership transfer.

---

### 8. Applicant Portal URLs

Applicant portal URLs are served by the deployed Apps Script web app.

Conceptually:

```text
WEB_APP_EXEC_URL?ref=APPLICATION_ID&token=SECURE_TOKEN
```

Example shape:

```text
https://script.google.com/macros/s/DEPLOYMENT_ID/exec?ref=APP-0017&token=...
```

The exact query-string structure must match the current code.

Do not manually invent applicant portal links.

Use the existing helper/functions in the codebase that build the portal URL.

If the deployment ID changes, every generated portal URL based on the old `/exec` URL must be reviewed.

---

### 9. Admin Portal URLs

The admin portal is also served by the Apps Script web app.

Its URL is based on the same deployed `/exec` endpoint and whatever route/query parameter the current code uses.

Conceptually:

```text
https://script.google.com/macros/s/DEPLOYMENT_ID/exec?...admin route/query...
```

Do not hardcode a new admin portal URL without checking the current routing code.

---

### 10. Licence Download URLs

Licence download links may also use the Apps Script deployment URL plus application/licence identifiers.

The existing code should remain the source of truth for how download URLs are built.

When changing deployment IDs, verify:

- applicant portal links
- admin portal links
- payment links
- licence download links
- email links
- any Script/Document Properties containing the web-app URL

---

## Finding URL References in the Code

To search the local project for Apps Script or deployment URL references:

```bash
grep -Rni \
  --exclude-dir=.git \
  --exclude=.clasp.json \
  -E "script\.google\.com|/exec|WEB_APP|BASE_URL|PORTAL_URL|DEPLOYMENT" .
```

To search for Drive, Sheet, Form, and Docs URLs:

```bash
grep -Rni \
  --exclude-dir=.git \
  --exclude=.clasp.json \
  -E "drive\.google\.com|docs\.google\.com|spreadsheets|forms|document" .
```

Before changing any deployment ID, audit all matches.

---

## Checking Local Files

Before every push:

```bash
clasp status
```

Only Apps Script-compatible files not excluded by `.claspignore` should appear under `Tracked files`.

Local-only files should appear under `Untracked files`.

---

## Adding New Apps Script Files

There is no `clasp add` command.

Create the file normally:

```bash
touch NewFeature.js
```

or:

```bash
touch NewPage.html
```

Then:

```bash
clasp status
```

If the file is not ignored, clasp automatically recognizes it.

Push with:

```bash
clasp push
```

---

## Git vs clasp

Git and clasp have different jobs.

### Git

Git stores local version history:

```bash
git status
git add .
git commit -m "Describe the change"
```

### clasp

clasp synchronizes Apps Script code:

```bash
clasp status
clasp push
```

A file does not need to be committed to Git before clasp can push it.

Recommended workflow:

```text
Edit code
↓
git diff
↓
git status
↓
clasp status
↓
clasp push
↓
test
↓
git add / commit
↓
deploy
```

---

## .claspignore

Recommended local-only entries:

```text
deploy.sh
.git/**
.gitignore
.clasp.json
.claspignore
README.md
```

After changing `.claspignore`, always run:

```bash
clasp status
```

and confirm all intended Apps Script source files remain tracked.

---

## Push vs Deployment

These are separate operations.

### Push

```bash
clasp push
```

This updates the source in the Apps Script project.

It does not automatically update a versioned web-app deployment.

### Version

Create an immutable Apps Script version:

```bash
clasp version "Description"
```

Example output:

```text
Created version 106.
```

### Deployment

List deployments:

```bash
clasp deployments
```

Redeploy an existing deployment ID:

```bash
clasp deploy \
  --deploymentId DEPLOYMENT_ID \
  --versionNumber VERSION_NUMBER \
  --description "Release description"
```

This is the preferred way to keep the same `/exec` URL.

---

## Local Deployment Script

`deploy.sh` can automate:

```text
clasp status
↓
confirmation
↓
clasp push
↓
create version
↓
redeploy same deployment ID
↓
show deployment status
```

Run:

```bash
./deploy.sh
```

`deploy.sh` should remain local to the repository and should not be pushed to Apps Script.

---

## Ownership and Drive IDs

The system uses Google Drive IDs for important resources such as:

- supporting documents folder
- payment proof folder
- draft licence folder
- stamped licence folder
- licence template
- generated documents
- other configured Drive resources

If ownership of the same existing file/folder is transferred:

```text
same Drive object
+ new owner
= same file/folder ID
```

Therefore, the application should continue using the same configured IDs.

Only replace an ID when a completely new Drive object has been created.

---

## Installable Triggers

Current triggers include:

```text
onFormSubmit
processSystemJobs
cleanupExpiredPaymentProofs
```

Installable triggers execute under the Google account that created them.

They are not automatically transferred when Sheet ownership changes.

To move trigger execution to a new account:

1. Sign into the new account.
2. Open the Sheet.
3. Open `Extensions → Apps Script`.
4. Open `Triggers`.
5. Recreate the required triggers.
6. Authorize the new account.
7. Test them.
8. Remove the old account's triggers.

Avoid leaving duplicate `onFormSubmit` triggers active unnecessarily.

---

## Email Identity

The project uses Apps Script mail functionality such as:

```javascript
MailApp.sendEmail(...)
```

The account under whose authorization the function runs can affect the sender identity.

Important execution sources include:

- `onFormSubmit`
- `processSystemJobs`
- other time-driven triggers
- direct web-app requests

After moving ownership or triggers, test each email-producing workflow.

---

## Application Workflow

```text
Applicant submits Google Form
        ↓
onFormSubmit
        ↓
Application ID generated
        ↓
Applicant token generated
        ↓
Confirmation email
        ↓
Applicant portal
        ↓
Payment proof
        ↓
Supporting documents
        ↓
Admin review
        ↓
Final approval
        ↓
Unstamped licence generation
        ↓
Stamped licence upload
        ↓
Licence release
        ↓
Applicant licence download
```

---

## Main Modules

### Applicant and Portal

```text
ApplicantPortal.html
ApplicantSupportingDocuments.html
Portal.js
SupportingDocuments.js
```

### Admin

```text
AdminAuth.js
AdminDocumentViewer.html
AdminPortal.js
AdminPortalPage.html
```

### Email

```text
Email.js
```

### Payment

```text
Payment.js
```

### Licence

```text
Licence.js
LicenceDownload.html
StampedLicenceWorkflow.js
```

### Background Jobs

```text
SystemJobs.js
```

### Reset / Go-Live Utilities

```text
GoLiveReset.js
```

---

## Before Pushing

Run:

```bash
git status
git diff
clasp status
```

Confirm:

- expected files changed
- no unrelated files changed
- required Apps Script files remain tracked
- local files remain untracked

Then:

```bash
clasp push
```

---

## Before Deployment

Recommended sequence:

```bash
git status
git diff
clasp status
clasp push
clasp version "Release description"
clasp deploy --deploymentId DEPLOYMENT_ID --versionNumber VERSION_NUMBER
clasp deployments
```

Or use:

```bash
./deploy.sh
```

when the deployment script is configured with the correct deployment ID.

---

## After Deployment

Perform a full smoke test:

1. Open the Google Form.
2. Submit a test application.
3. Confirm application row creation.
4. Confirm Application ID generation.
5. Confirm applicant email.
6. Open applicant portal.
7. Upload payment proof.
8. Upload supporting documents.
9. Open admin portal.
10. Review payment/documents.
11. Final approve application.
12. Generate licence.
13. Upload stamped licence.
14. Release licence.
15. Verify applicant can download licence.
16. Verify System Jobs runs.
17. Verify uploaded/generated files are stored in the expected Drive folders.
18. Verify email sender identity.
19. Verify all portal/download URLs use the intended deployment.

---

## Useful Commands

```bash
clasp --version
clasp show-authorized-user
clasp status
clasp pull
clasp push
clasp version "Description"
clasp deployments
clasp open
```

Redeploy:

```bash
clasp deploy \
  --deploymentId DEPLOYMENT_ID \
  --versionNumber VERSION_NUMBER \
  --description "Description"
```

---

## Safety Rules

- Do not commit Google credentials or secrets.
- Do not commit `.clasp.json` unless intentionally required.
- Always check `clasp status` before pushing.
- Preserve the existing deployment ID where possible.
- Do not create a new deployment unnecessarily.
- Keep Drive file/folder IDs stable where possible.
- Do not replace existing folders merely to change ownership.
- Do not leave duplicate installable triggers active.
- Do not remove the old account's access until the ownership handover has been tested.
- Search for all `/exec` URL references before changing deployment ID.
- Test the full applicant-to-licence flow before go-live.

---

## Current Status

The project is currently sandbox/pre-production.

This is the correct stage to test:

- ownership transfer
- new Google account authorization
- trigger ownership
- email sender identity
- Drive ownership
- deployment execution identity
- stable Form/Sheet/Drive URLs
- applicant/admin portal URLs
- licence download URLs

before opening the system to real applicants.
