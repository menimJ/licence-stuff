/**
 * ============================================================
 * GO-LIVE RESET + COUNTER INITIALISATION
 * ============================================================
 *
 * PURPOSE
 * -------
 * One-time production cutover utility.
 *
 * It:
 * 1. Reads every current test row before clearing it.
 * 2. Moves files referenced by those rows to Google Drive Trash.
 * 3. Clears all application rows from the response sheet but keeps row 1.
 * 4. Removes test applicant memory from Document Properties.
 * 5. Sets the Application ID counter and S/N counter to the last official
 *    numbers already used outside this system.
 * 6. Writes the configured counter values to the Setup sheet for visibility.
 *
 * IMPORTANT
 * ---------
 * - It DOES NOT delete the Google Form.
 * - It DOES NOT delete the response sheet headers.
 * - It DOES NOT delete the licence template.
 * - It DOES NOT delete the configured Drive folders themselves.
 * - It DOES NOT permanently delete files; files are moved to Drive Trash.
 *
 * Existing project dependencies:
 * - PROP_KEYS.APP_COUNTER
 * - PROP_KEYS.SN_COUNTER
 * - getResponseSheet_(spreadsheet)
 * - SETUP_SHEET_NAME (optional; falls back to "Setup")
 * - ADMIN_AUTH_CONFIG
 * - ADMIN_USER_HEADERS
 * - ensureAdminAuthSheet_()
 * - findAdminUserByEmail_()
 */



const CRFFN_GO_LIVE_RESET_VERSION =
  '2026-08-31-v4-clear-only';


/**
 * Confirms that the V4 reset engine is loaded in the project.
 *
 * Run this from the Apps Script editor if you suspect an older
 * duplicate reset file is still active.
 */
function verifyGoLiveResetVersion() {
  const result = {
    ok: true,
    version:
      CRFFN_GO_LIVE_RESET_VERSION,
    rowDeletionStrategy:
      'clearContent only'
  };

  console.log(
    JSON.stringify(
      result
    )
  );

  try {
    SpreadsheetApp
      .getUi()
      .alert(
        'CRFFN Reset Engine',
        [
          'Loaded version: ' +
            CRFFN_GO_LIVE_RESET_VERSION,
          '',
          'Row cleanup strategy: clearContent only',
          '',
          'This V4 file contains no deleteRow() or deleteRows() calls.'
        ].join('\n'),
        SpreadsheetApp
          .getUi()
          .ButtonSet
          .OK
      );
  } catch (error) {
    // The return value and execution log are sufficient outside a Sheet UI.
  }

  return result;
}


/**
 * Existing menu entry point.
 */
function goLiveResetAndSetStartingNumbers() {
  return crffnResetV4_runGoLiveResetAndSetStartingNumbers_();
}


/**
 * Unique V4 entry point for direct testing from the Apps Script editor.
 * Use this if an older duplicate goLiveResetAndSetStartingNumbers()
 * function may still exist elsewhere in the project.
 */
function runGoLiveResetV4() {
  return crffnResetV4_runGoLiveResetAndSetStartingNumbers_();
}


/**
 * Existing menu entry point for resetting one test application.
 */
function showResetTestApplicationPrompt() {
  return crffnResetV4_showResetTestApplicationPrompt_();
}


/**
 * Unique V4 single-reset prompt for direct testing.
 */
function runResetTestApplicationPromptV4() {
  return crffnResetV4_showResetTestApplicationPrompt_();
}


/**
 * Preserves programmatic compatibility with callers that use
 * resetTestApplication(applicationId).
 */
function resetTestApplication(
  applicationId
) {
  return crffnResetV4_resetTestApplication_(
    applicationId
  );
}



/**
 * Authorises destructive utilities launched directly from the
 * bound Google Spreadsheet.
 *
 * IMPORTANT:
 * Admin Portal actions use browser session-token authentication.
 * Spreadsheet menu actions do not have that browser session token,
 * so they must authenticate the Google account executing the Sheet action.
 */
function crffnResetV4_requireSpreadsheetSuperAdminAccess_() {
  const email =
    String(
      Session
        .getActiveUser()
        .getEmail() ||
      Session
        .getEffectiveUser()
        .getEmail() ||
      ''
    )
      .trim()
      .toLowerCase();

  if (!email) {
    throw new Error(
      'Your Google account could not be identified. Open the CRFFN spreadsheet using the authorised Super Admin Google account and try again.'
    );
  }

  const spreadsheet =
    SpreadsheetApp
      .getActiveSpreadsheet();

  if (!spreadsheet) {
    throw new Error(
      'The CRFFN spreadsheet could not be identified.'
    );
  }

  if (
    typeof ADMIN_AUTH_CONFIG ===
      'undefined' ||
    typeof ADMIN_USER_HEADERS ===
      'undefined' ||
    typeof ensureAdminAuthSheet_ !==
      'function' ||
    typeof findAdminUserByEmail_ !==
      'function'
  ) {
    throw new Error(
      'The CRFFN administrator authentication system is not available. Confirm that AdminAuth.gs is present and saved.'
    );
  }

  const usersSheet =
    ensureAdminAuthSheet_(
      spreadsheet,
      ADMIN_AUTH_CONFIG.USERS_SHEET,
      ADMIN_USER_HEADERS
    );

  const admin =
    findAdminUserByEmail_(
      usersSheet,
      email
    );

  if (!admin) {
    throw new Error(
      'Your Google account is not registered as a CRFFN administrator.'
    );
  }

  const status =
    String(
      admin.values &&
      admin.values.Status ||
      ''
    ).trim();

  const role =
    String(
      admin.values &&
      admin.values.Role ||
      ''
    ).trim();

  if (status !== 'Active') {
    throw new Error(
      'Your CRFFN administrator account is not active.'
    );
  }

  if (role !== 'Super Admin') {
    throw new Error(
      'Only a Super Admin can run this reset utility.'
    );
  }

  return {
    email: email,
    role: role
  };
}


/**
 * Main function.
 *
 * Run this manually once when you are ready to go live,
 * or expose it through the Form Tools menu.
 */
function crffnResetV4_runGoLiveResetAndSetStartingNumbers_() {
  console.log('CRFFN reset engine: ' + CRFFN_GO_LIVE_RESET_VERSION);

  crffnResetV4_requireSpreadsheetSuperAdminAccess_();

  const ui =
    SpreadsheetApp.getUi();

  /*
   * ----------------------------------------------------------
   * STEP 1: LAST OFFICIAL APPLICATION NUMBER
   * ----------------------------------------------------------
   *
   * Example:
   * Existing official last number: APP-1250
   * Enter: 1250
   *
   * The next system-generated number becomes APP-1251.
   */
  const applicationResponse =
    ui.prompt(
      'Go Live — Application Number',
      [
        'Enter the LAST official application number already used.',
        '',
        'Example:',
        'If the last existing application is APP-1250, enter 1250.',
        '',
        'The first new application generated by this system will then be APP-1251.'
      ].join('\n'),
      ui.ButtonSet.OK_CANCEL
    );

  if (
    applicationResponse
      .getSelectedButton() !==
    ui.Button.OK
  ) {
    return;
  }

  const lastApplicationNumber =
    crffnResetV4_parsePositiveWholeNumber_(
      applicationResponse
        .getResponseText()
    );

  if (
    lastApplicationNumber === null
  ) {
    ui.alert(
      'Go Live cancelled. Enter a valid whole application number.'
    );
    return;
  }


  /*
   * ----------------------------------------------------------
   * STEP 2: LAST OFFICIAL SERIAL NUMBER
   * ----------------------------------------------------------
   */
  const serialResponse =
    ui.prompt(
      'Go Live — Serial Number',
      [
        'Enter the LAST official S/N already used.',
        '',
        'Example:',
        'If the last S/N is 850, enter 850.',
        '',
        'The first new S/N generated by this system will then be 851.'
      ].join('\n'),
      ui.ButtonSet.OK_CANCEL
    );

  if (
    serialResponse
      .getSelectedButton() !==
    ui.Button.OK
  ) {
    return;
  }

  const lastSerialNumber =
    crffnResetV4_parsePositiveWholeNumber_(
      serialResponse
        .getResponseText()
    );

  if (
    lastSerialNumber === null
  ) {
    ui.alert(
      'Go Live cancelled. Enter a valid whole serial number.'
    );
    return;
  }


  /*
   * ----------------------------------------------------------
   * STEP 3: FINAL DESTRUCTIVE CONFIRMATION
   * ----------------------------------------------------------
   */
  const expectedConfirmation =
    'GO LIVE ' +
    String(
      lastApplicationNumber
    );

  const confirmationResponse =
    ui.prompt(
      'FINAL GO-LIVE CONFIRMATION',
      [
        'This will CLEAR all current test application records and move their uploaded/generated files to Google Drive Trash.',
        '',
        'The Google Form, sheet headers, templates and folders will remain.',
        '',
        'Last Application Number: ' +
          lastApplicationNumber,
        'Next Application Number: ' +
          (
            lastApplicationNumber + 1
          ),
        '',
        'Last S/N: ' +
          lastSerialNumber,
        'Next S/N: ' +
          (
            lastSerialNumber + 1
          ),
        '',
        'Type exactly:',
        expectedConfirmation
      ].join('\n'),
      ui.ButtonSet.OK_CANCEL
    );

  if (
    confirmationResponse
      .getSelectedButton() !==
    ui.Button.OK
  ) {
    return;
  }

  const enteredConfirmation =
    String(
      confirmationResponse
        .getResponseText() ||
      ''
    ).trim();

  if (
    enteredConfirmation !==
    expectedConfirmation
  ) {
    ui.alert(
      'Go Live cancelled because the confirmation text did not match.'
    );
    return;
  }


  /*
   * ----------------------------------------------------------
   * STEP 4: LOCK SYSTEM DURING RESET
   * ----------------------------------------------------------
   */
  const lock =
    LockService.getDocumentLock();

  lock.waitLock(
    30000
  );

  try {
    const spreadsheet =
      SpreadsheetApp
        .getActiveSpreadsheet();

    if (!spreadsheet) {
      throw new Error(
        'The active spreadsheet could not be found.'
      );
    }

    const sheet =
      getResponseSheet_(
        spreadsheet
      );

    if (
      !sheet ||
      typeof sheet.getLastRow !==
        'function'
    ) {
      throw new Error(
        'The application response sheet could not be found.'
      );
    }

    const resetApplicationIds =
      crffnResetV4_collectResetApplicationIds_(
        sheet
      );

    /*
     * First clear files while the sheet still contains
     * all file IDs and URLs.
     */
    const fileCleanupResult =
      crffnResetV4_trashFilesReferencedByTestRows_(
        sheet
      );

    /*
     * Clear response rows but keep:
     * - response sheet
     * - row 1 headers
     * - formatting
     * - form destination
     */
    const clearedRows =
      crffnResetV4_clearAllApplicationRows_(
        sheet
      );

    /*
     * Remove per-applicant memory so a test email is treated
     * as completely new after go-live.
     */
    const removedPropertyCount =
      crffnResetV4_clearApplicantMemoryProperties_();

    const clearedSystemJobs =
      crffnResetV4_clearAllResetSystemJobs_(
        spreadsheet
      );

    const cacheInvalidationActions =
      crffnResetV4_invalidateResetCachesForApplications_(
        resetApplicationIds
      );

    /*
     * Initialise production counters.
     */
    const properties =
      PropertiesService
        .getDocumentProperties();

    properties.setProperty(
      PROP_KEYS.APP_COUNTER,
      String(
        lastApplicationNumber
      )
    );

    properties.setProperty(
      PROP_KEYS.SN_COUNTER,
      String(
        lastSerialNumber
      )
    );

    /*
     * Make the values visible in the Setup sheet.
     */
    crffnResetV4_writeGoLiveCountersToSetup_(
      spreadsheet,
      lastApplicationNumber,
      lastSerialNumber
    );

    SpreadsheetApp.flush();

    ui.alert(
      'GO LIVE SETUP COMPLETE',
      [
        'Test application rows cleared: ' +
          clearedRows,
        '',
        'Drive files moved to Trash: ' +
          fileCleanupResult.trashed,
        'Drive file references skipped/not found: ' +
          fileCleanupResult.skipped,
        '',
        'Applicant memory properties removed: ' +
          removedPropertyCount,
        'System Jobs cleared: ' +
          clearedSystemJobs,
        'Cache invalidation actions: ' +
          cacheInvalidationActions,
        '',
        'Last Application Counter: ' +
          lastApplicationNumber,
        'Next Application: APP-' +
          crffnResetV4_padApplicationNumber_(
            lastApplicationNumber + 1
          ),
        '',
        'Last S/N Counter: ' +
          lastSerialNumber,
        'Next S/N: ' +
          (
            lastSerialNumber + 1
          ),
        '',
        'The form, headers, templates and folders were preserved.',
        '',
        'Reset Engine: ' +
          CRFFN_GO_LIVE_RESET_VERSION
      ].join('\n'),
      ui.ButtonSet.OK
    );
  } finally {
    lock.releaseLock();
  }
}


/**
 * Moves all application-related files referenced by
 * current test rows to Google Drive Trash.
 *
 * It does NOT empty entire folders.
 * That prevents accidental removal of templates or other
 * non-test files.
 */
function crffnResetV4_trashFilesReferencedByTestRows_(
  sheet
) {
  const lastRow =
    sheet.getLastRow();

  const lastColumn =
    sheet.getLastColumn();

  if (
    lastRow < 2 ||
    lastColumn < 1
  ) {
    return {
      trashed: 0,
      skipped: 0
    };
  }

  const headers =
    sheet
      .getRange(
        1,
        1,
        1,
        lastColumn
      )
      .getDisplayValues()[0]
      .map(function(header) {
        return String(
          header || ''
        ).trim();
      });

  const rows =
    sheet
      .getRange(
        2,
        1,
        lastRow - 1,
        lastColumn
      )
      .getDisplayValues();

  /*
   * File-ID columns known in the current application.
   */
  const fileIdHeaders = [
    'Payment Proof File ID',

    'CAC Document File ID',
    'Passport Photograph File ID',
    'Educational Certificates File ID',
    'Proof of Experience File ID',
    'Means of Identification File ID',

    'Stamped Licence File ID'
  ];

  /*
   * URL columns that may point to Drive files even where
   * a separate File ID was not stored.
   */
  const driveUrlHeaders = [
    'Application PDF URL',
    'Receipt PDF URL',

    'CAC Document URL',
    'Passport Photograph URL',
    'Educational Certificates URL',
    'Proof of Experience URL',
    'Means of Identification URL',

    'Licence Document URL',
    'Licence PDF URL',
    'Stamped Licence PDF URL'
  ];

  const headerMap = {};

  headers.forEach(
    function(
      header,
      index
    ) {
      if (header) {
        headerMap[header] =
          index;
      }
    }
  );

  const fileIds =
    new Set();

  rows.forEach(
    function(row) {
      fileIdHeaders
        .forEach(
          function(header) {
            const index =
              headerMap[header];

            if (
              index === undefined
            ) {
              return;
            }

            const fileId =
              String(
                row[index] || ''
              ).trim();

            if (fileId) {
              fileIds.add(
                fileId
              );
            }
          }
        );

      driveUrlHeaders
        .forEach(
          function(header) {
            const index =
              headerMap[header];

            if (
              index === undefined
            ) {
              return;
            }

            const url =
              String(
                row[index] || ''
              ).trim();

            const fileId =
              crffnResetV4_extractGoogleDriveFileId_(
                url
              );

            if (fileId) {
              fileIds.add(
                fileId
              );
            }
          }
        );
    }
  );

  let trashed =
    0;

  let skipped =
    0;

  fileIds.forEach(
    function(fileId) {
      try {
        const file =
          DriveApp
            .getFileById(
              fileId
            );

        if (
          !file.isTrashed()
        ) {
          file.setTrashed(
            true
          );
        }

        trashed +=
          1;
      } catch (error) {
        skipped +=
          1;

        console.warn(
          'Could not move Drive file to Trash: ' +
          fileId +
          ' — ' +
          error.message
        );
      }
    }
  );

  return {
    trashed:
      trashed,
    skipped:
      skipped
  };
}


/**
 * Clears all application response values below row 1.
 *
 * Uses clearContent() instead of deleting the response sheet
 * or destroying its structure.
 */
function crffnResetV4_clearAllApplicationRows_(
  sheet
) {
  const rowsWithData =
    Math.max(
      sheet.getLastRow() - 1,
      0
    );

  const maxRows =
    Number(
      sheet.getMaxRows() || 0
    );

  const maxColumns =
    Number(
      sheet.getMaxColumns() || 0
    );

  /*
   * IMPORTANT:
   * NEVER delete physical rows from the Google Form response sheet.
   *
   * Google Sheets rejects attempts to delete every non-frozen row.
   * Clearing contents below the header is sufficient: after the reset,
   * getLastRow() returns 1 and the next Google Form response can use
   * the next available response row.
   */
  if (
    maxRows > 1 &&
    maxColumns > 0
  ) {
    sheet
      .getRange(
        2,
        1,
        maxRows - 1,
        maxColumns
      )
      .clearContent();
  }

  SpreadsheetApp.flush();

  return rowsWithData;
}


/**
 * Removes persistent applicant test memory.
 *
 * Preserves:
 * - FORM_ID
 * - WEB_APP_URL
 * - folder configuration
 * - other application configuration
 */
function crffnResetV4_clearApplicantMemoryProperties_() {
  const properties =
    PropertiesService
      .getDocumentProperties();

  const allProperties =
    properties.getProperties();

  const prefixes = [
    'EMAIL_TO_APP::',
    'EMAIL_TO_TOKEN::',
    'EMAIL_SUBMISSION_COUNT::',
    'SUBMISSION_COUNT::',
    'APP_TO_TOKEN::',
    'APP_TO_EMAIL::',
    'APP_TO_PORTAL::',
    'APPLICATION_TOKEN::',
    'APPLICATION_EMAIL::',
    'APPLICATION_PORTAL_URL::',
    'TOKEN_TO_APP::',
    'PORTAL_TOKEN::',
    'SECURE_TOKEN::'
  ];

  let removed =
    0;

  Object.keys(
    allProperties
  ).forEach(
    function(key) {
      const shouldRemove =
        prefixes.some(
          function(prefix) {
            return key.indexOf(
              prefix
            ) === 0;
          }
        );

      if (!shouldRemove) {
        return;
      }

      properties.deleteProperty(
        key
      );

      removed +=
        1;
    }
  );

  return removed;
}


/**
 * Extracts a Drive file ID from common Google Drive URLs.
 */
function crffnResetV4_extractGoogleDriveFileId_(
  value
) {
  const text =
    String(
      value || ''
    ).trim();

  if (!text) {
    return '';
  }

  /*
   * /d/FILE_ID/
   */
  let match =
    text.match(
      /\/d\/([A-Za-z0-9_-]{20,})/
    );

  if (
    match &&
    match[1]
  ) {
    return match[1];
  }

  /*
   * ?id=FILE_ID
   */
  match =
    text.match(
      /[?&]id=([A-Za-z0-9_-]{20,})/
    );

  if (
    match &&
    match[1]
  ) {
    return match[1];
  }

  /*
   * Sometimes a raw Drive ID may be stored instead of a URL.
   */
  if (
    /^[A-Za-z0-9_-]{20,}$/
      .test(
        text
      )
  ) {
    return text;
  }

  return '';
}


/**
 * Returns all Application IDs currently present in the response sheet.
 * Used only for targeted cache/job cleanup before rows are removed.
 */
function crffnResetV4_collectResetApplicationIds_(
  sheet
) {
  const lastRow =
    sheet.getLastRow();

  const lastColumn =
    sheet.getLastColumn();

  if (
    lastRow < 2 ||
    lastColumn < 1
  ) {
    return [];
  }

  const headers =
    sheet
      .getRange(
        1,
        1,
        1,
        lastColumn
      )
      .getDisplayValues()[0]
      .map(function(header) {
        return String(header || '').trim();
      });

  const applicationIndex =
    headers.indexOf(
      'Application ID'
    );

  if (applicationIndex < 0) {
    return [];
  }

  const values =
    sheet
      .getRange(
        2,
        applicationIndex + 1,
        lastRow - 1,
        1
      )
      .getDisplayValues();

  const ids =
    {};

  values.forEach(function(row) {
    const value =
      String(row[0] || '').trim();

    if (value) {
      ids[value] = true;
    }
  });

  return Object.keys(ids);
}


/**
 * Best-effort invalidation for every cache layer currently used by
 * the licensing project. Missing optional helpers are ignored so the
 * reset remains compatible with deployments that do not use a layer.
 */
function crffnResetV4_invalidateResetApplicationCaches_(
  applicationId
) {
  const cleanApplicationId =
    String(applicationId || '').trim();

  if (!cleanApplicationId) {
    return 0;
  }

  let invalidated =
    0;

  const optionalInvalidators = [
    'invalidateAdminDocumentReviewCache_',
    'invalidateAdminApplicationCache_',
    'invalidateApplicantPortalCache_',
    'invalidateApplicantApplicationCache_',
    'invalidateApplicationCache_',
    'invalidatePortalCache_',
    'invalidatePaymentCache_',
    'invalidateSupportingDocumentsCache_'
  ];

  optionalInvalidators.forEach(function(name) {
    try {
      const fn =
        globalThis &&
        typeof globalThis[name] === 'function'
          ? globalThis[name]
          : null;

      if (fn) {
        fn(cleanApplicationId);
        invalidated += 1;
      }
    } catch (error) {
      console.warn(
        'Cache invalidation warning for ' +
        cleanApplicationId +
        ' via ' +
        name +
        ': ' +
        error.message
      );
    }
  });

  try {
    CacheService
      .getScriptCache()
      .removeAll([
        'CRFFN:ADMIN_DOC_REVIEW:' +
          cleanApplicationId +
          ':summary',
        'CRFFN:ADMIN_DOC_REVIEW:' +
          cleanApplicationId +
          ':render'
      ]);

    invalidated += 1;
  } catch (error) {
    console.warn(
      'Direct cache cleanup warning for ' +
      cleanApplicationId +
      ': ' +
      error.message
    );
  }

  return invalidated;
}


function crffnResetV4_invalidateResetCachesForApplications_(
  applicationIds
) {
  let count =
    0;

  (applicationIds || [])
    .forEach(function(applicationId) {
      count +=
        crffnResetV4_invalidateResetApplicationCaches_(
          applicationId
        );
    });

  return count;
}


/**
 * Finds the background-job sheet without depending on SystemJobs.gs.
 */
function crffnResetV4_getResetSystemJobsSheet_(
  spreadsheet
) {
  if (!spreadsheet) {
    return null;
  }

  const possibleNames = [
    typeof SYSTEM_JOBS_SHEET_NAME !== 'undefined'
      ? SYSTEM_JOBS_SHEET_NAME
      : '',
    'System Jobs',
    'SystemJobs'
  ].filter(Boolean);

  for (
    let index = 0;
    index < possibleNames.length;
    index++
  ) {
    const sheet =
      spreadsheet.getSheetByName(
        possibleNames[index]
      );

    if (sheet) {
      return sheet;
    }
  }

  return null;
}


function crffnResetV4_clearAllResetSystemJobs_(
  spreadsheet
) {
  const sheet =
    crffnResetV4_getResetSystemJobsSheet_(
      spreadsheet
    );

  if (!sheet) {
    return 0;
  }

  const rows =
    Math.max(
      sheet.getLastRow() - 1,
      0
    );

  const lastColumn =
    Number(
      sheet.getLastColumn() || 0
    );

  if (
    rows > 0 &&
    lastColumn > 0
  ) {
    sheet
      .getRange(
        2,
        1,
        rows,
        lastColumn
      )
      .clearContent();
  }

  return rows;
}


function crffnResetV4_clearResetSystemJobsForApplication_(
  spreadsheet,
  applicationId
) {
  const sheet =
    crffnResetV4_getResetSystemJobsSheet_(
      spreadsheet
    );

  const cleanApplicationId =
    String(applicationId || '').trim();

  if (
    !sheet ||
    !cleanApplicationId ||
    sheet.getLastRow() < 2
  ) {
    return 0;
  }

  const lastRow =
    sheet.getLastRow();

  const lastColumn =
    sheet.getLastColumn();

  if (lastColumn < 1) {
    return 0;
  }

  const values =
    sheet
      .getRange(
        2,
        1,
        lastRow - 1,
        lastColumn
      )
      .getDisplayValues();

  const matchingRows =
    [];

  values.forEach(function(row, index) {
    const belongsToApplication =
      row.some(function(cell) {
        return (
          String(cell || '').trim() ===
          cleanApplicationId
        );
      });

    if (belongsToApplication) {
      matchingRows.push(
        index + 2
      );
    }
  });

  crffnResetV4_clearSpecificRows_(
    sheet,
    matchingRows
  );

  return matchingRows.length;
}


/**
 * Clears selected rows without physically deleting Sheet rows.
 *
 * Adjacent row numbers are grouped into a single clearContent() call.
 */
function crffnResetV4_clearSpecificRows_(
  sheet,
  rowNumbers
) {
  if (!sheet) {
    return 0;
  }

  const maxRows =
    Number(
      sheet.getMaxRows() || 0
    );

  const maxColumns =
    Number(
      sheet.getMaxColumns() || 0
    );

  if (
    maxRows < 2 ||
    maxColumns < 1
  ) {
    return 0;
  }

  const rows =
    Array.from(
      new Set(
        (rowNumbers || [])
          .map(function(rowNumber) {
            return Number(rowNumber);
          })
          .filter(function(rowNumber) {
            return (
              Number.isInteger(rowNumber) &&
              rowNumber > 1 &&
              rowNumber <= maxRows
            );
          })
      )
    )
      .sort(function(a, b) {
        return a - b;
      });

  if (!rows.length) {
    return 0;
  }

  let cleared =
    0;

  let rangeStart =
    rows[0];

  let rangeEnd =
    rows[0];

  function clearCurrentRange_() {
    const count =
      rangeEnd -
      rangeStart +
      1;

    sheet
      .getRange(
        rangeStart,
        1,
        count,
        maxColumns
      )
      .clearContent();

    cleared +=
      count;
  }

  for (
    let index = 1;
    index < rows.length;
    index++
  ) {
    const current =
      rows[index];

    if (
      current ===
      rangeEnd + 1
    ) {
      rangeEnd =
        current;

      continue;
    }

    clearCurrentRange_();

    rangeStart =
      current;

    rangeEnd =
      current;
  }

  clearCurrentRange_();

  return cleared;
}


/**
 * Writes a visible go-live block to the Setup sheet.
 *
 * Document Properties remain the operational source of truth.
 */
function crffnResetV4_writeGoLiveCountersToSetup_(
  spreadsheet,
  lastApplicationNumber,
  lastSerialNumber
) {
  const setupSheetName =
    typeof SETUP_SHEET_NAME !==
      'undefined'
      ? SETUP_SHEET_NAME
      : 'Setup';

  let sheet =
    spreadsheet.getSheetByName(
      setupSheetName
    );

  if (!sheet) {
    sheet =
      spreadsheet.insertSheet(
        setupSheetName
      );
  }

  /*
   * Locate an existing GO LIVE block so repeated runs update
   * the same area instead of endlessly appending.
   */
  const marker =
    'GO LIVE COUNTERS';

  const lastRow =
    Math.max(
      sheet.getLastRow(),
      1
    );

  const values =
    sheet
      .getRange(
        1,
        1,
        lastRow,
        1
      )
      .getDisplayValues();

  let startRow =
    0;

  for (
    let index = 0;
    index < values.length;
    index++
  ) {
    if (
      String(
        values[index][0] || ''
      ).trim() ===
      marker
    ) {
      startRow =
        index + 1;
      break;
    }
  }

  if (!startRow) {
    startRow =
      Math.max(
        sheet.getLastRow() + 2,
        1
      );
  }

  const activeEmail =
    Session
      .getActiveUser()
      .getEmail() ||
    Session
      .getEffectiveUser()
      .getEmail() ||
    '';

  const rows = [
    [
      marker,
      ''
    ],
    [
      'Current Application Counter',
      lastApplicationNumber
    ],
    [
      'Next Application Number',
      'APP-' +
        crffnResetV4_padApplicationNumber_(
          lastApplicationNumber + 1
        )
    ],
    [
      'Current S/N Counter',
      lastSerialNumber
    ],
    [
      'Next S/N',
      lastSerialNumber + 1
    ],
    [
      'Go Live Initialised At',
      new Date()
    ],
    [
      'Go Live Initialised By',
      activeEmail
    ]
  ];

  sheet
    .getRange(
      startRow,
      1,
      rows.length,
      2
    )
    .clearContent()
    .setValues(
      rows
    );

  sheet
    .getRange(
      startRow,
      1,
      1,
      2
    )
    .setFontWeight(
      'bold'
    );

  sheet
    .getRange(
      startRow,
      1,
      rows.length,
      2
    )
    .setWrap(
      true
    );

  sheet.autoResizeColumns(
    1,
    2
  );
}


/**
 * Parses a zero-or-greater whole number.
 */
function crffnResetV4_parsePositiveWholeNumber_(
  value
) {
  const cleanValue =
    String(
      value || ''
    )
      .trim()
      .replace(
        /,/g,
        ''
      );

  if (
    !/^\d+$/
      .test(
        cleanValue
      )
  ) {
    return null;
  }

  const number =
    Number(
      cleanValue
    );

  if (
    !Number.isSafeInteger(
      number
    ) ||
    number < 0
  ) {
    return null;
  }

  return number;
}


/**
 * Mirrors the existing APP-0001 style.
 *
 * If the number becomes longer than 4 digits,
 * JavaScript keeps the full number.
 */
function crffnResetV4_padApplicationNumber_(
  number
) {
  return String(
    number
  ).padStart(
    4,
    '0'
  );
}


/**
 * ============================================================
 * RESET ONE TEST APPLICATION
 * ============================================================
 *
 * Menu:
 * Form Tools -> Reset Test Application
 */
function crffnResetV4_showResetTestApplicationPrompt_() {
  console.log('CRFFN reset engine: ' + CRFFN_GO_LIVE_RESET_VERSION);

  crffnResetV4_requireSpreadsheetSuperAdminAccess_();

  const ui =
    SpreadsheetApp.getUi();

  const response =
    ui.prompt(
      'Reset Test Application',
      [
        'Enter the Application ID you want to reset.',
        '',
        'Example:',
        'APP-0001'
      ].join('\n'),
      ui.ButtonSet.OK_CANCEL
    );

  if (
    response.getSelectedButton() !==
    ui.Button.OK
  ) {
    return;
  }

  const applicationId =
    String(
      response.getResponseText() || ''
    ).trim();

  if (!applicationId) {
    ui.alert(
      'Application ID is required.'
    );
    return;
  }

  const confirmation =
    ui.alert(
      'Confirm Test Application Reset',
      [
        'Application ID:',
        applicationId,
        '',
        'This will remove all Sheet rows for this test Application ID and move known related uploaded/generated files to Drive Trash.',
        '',
        'The global application-number sequence will NOT be reset.',
        '',
        'Continue?'
      ].join('\n'),
      ui.ButtonSet.YES_NO
    );

  if (
    confirmation !==
    ui.Button.YES
  ) {
    return;
  }

  try {
    const result =
      crffnResetV4_resetTestApplication_(
        applicationId
      );

    ui.alert(
      'Reset Complete',
      result &&
      result.message
        ? result.message
        : (
          'Test application ' +
          applicationId +
          ' was reset successfully.'
        ),
      ui.ButtonSet.OK
    );
  } catch (error) {
    ui.alert(
      'Reset Failed',
      error &&
      error.message
        ? error.message
        : String(error),
      ui.ButtonSet.OK
    );
  }
}


function crffnResetV4_resetTestApplication_(
  applicationId
) {
  const cleanApplicationId =
    String(
      applicationId || ''
    ).trim();

  if (!cleanApplicationId) {
    throw new Error(
      'Application ID is required.'
    );
  }

  const lock =
    LockService.getDocumentLock();

  lock.waitLock(
    30000
  );

  try {
    const spreadsheet =
      SpreadsheetApp
        .getActiveSpreadsheet();

    const sheet =
      getResponseSheet_(
        spreadsheet
      );

    const lastRow =
      sheet.getLastRow();

    const lastColumn =
      sheet.getLastColumn();

    if (
      lastRow < 2 ||
      lastColumn < 1
    ) {
      throw new Error(
        'There are no application records to reset.'
      );
    }

    const headers =
      sheet
        .getRange(
          1,
          1,
          1,
          lastColumn
        )
        .getDisplayValues()[0]
        .map(function(header) {
          return String(
            header || ''
          ).trim();
        });

    const headerMap =
      getHeaderMap_(
        headers
      );

    const applicationIdColumn =
      headerMap[
        'Application ID'
      ];

    if (!applicationIdColumn) {
      throw new Error(
        'Application ID column was not found.'
      );
    }

    const values =
      sheet
        .getRange(
          2,
          1,
          lastRow - 1,
          lastColumn
        )
        .getValues();

    const matchingRows =
      [];

    const fileIdsToTrash =
      {};

    const applicantEmails =
      {};

    const secureTokens =
      {};

    const fileIdHeaders = [
      'Payment Proof File ID',
      'CAC Document File ID',
      'Passport Photograph File ID',
      'Educational Certificates File ID',
      'Proof of Experience File ID',
      'Means of Identification File ID',
      'Stamped Licence File ID'
    ];

    const fileUrlHeaders = [
      'Payment Proof URL',
      'Payment Receipt URL',
      'Receipt PDF URL',
      'CAC Document URL',
      'Passport Photograph URL',
      'Educational Certificates URL',
      'Proof of Experience URL',
      'Means of Identification URL',
      'Licence Document URL',
      'Licence PDF URL',
      'Stamped Licence PDF URL'
    ];

    values.forEach(
      function(row, index) {
        const storedApplicationId =
          String(
            row[
              applicationIdColumn - 1
            ] || ''
          ).trim();

        if (
          storedApplicationId !==
          cleanApplicationId
        ) {
          return;
        }

        matchingRows.push(
          index + 2
        );

        [
          'Email Address',
          'Email',
          'Applicant Email',
          'Company Email'
        ].forEach(
          function(header) {
            const column =
              headerMap[
                header
              ];

            if (!column) {
              return;
            }

            const email =
              String(
                row[
                  column - 1
                ] || ''
              )
                .trim()
                .toLowerCase();

            if (email) {
              applicantEmails[
                email
              ] =
                true;
            }
          }
        );

        const tokenColumn =
          headerMap[
            'Secure Token'
          ];

        if (tokenColumn) {
          const token =
            String(
              row[
                tokenColumn - 1
              ] || ''
            ).trim();

          if (token) {
            secureTokens[
              token
            ] =
              true;
          }
        }

        fileIdHeaders.forEach(
          function(header) {
            const column =
              headerMap[
                header
              ];

            if (!column) {
              return;
            }

            const fileId =
              String(
                row[
                  column - 1
                ] || ''
              ).trim();

            if (fileId) {
              fileIdsToTrash[
                fileId
              ] =
                true;
            }
          }
        );

        fileUrlHeaders.forEach(
          function(header) {
            const column =
              headerMap[
                header
              ];

            if (!column) {
              return;
            }

            const fileId =
              crffnResetV4_extractResetDriveFileId_(
                row[
                  column - 1
                ]
              );

            if (fileId) {
              fileIdsToTrash[
                fileId
              ] =
                true;
            }
          }
        );
      }
    );

    if (!matchingRows.length) {
      throw new Error(
        'No application was found with Application ID ' +
        cleanApplicationId +
        '.'
      );
    }

    let filesTrashed =
      0;

    let fileCleanupErrors =
      0;

    Object.keys(
      fileIdsToTrash
    ).forEach(
      function(fileId) {
        try {
          const file =
            DriveApp.getFileById(
              fileId
            );

          if (
            !file.isTrashed()
          ) {
            file.setTrashed(
              true
            );
          }

          filesTrashed +=
            1;
        } catch (error) {
          fileCleanupErrors +=
            1;

          console.warn(
            'Could not trash test file ' +
            fileId +
            ': ' +
            error.message
          );
        }
      }
    );

    crffnResetV4_clearSpecificRows_(
      sheet,
      matchingRows
    );

    SpreadsheetApp.flush();

    const clearedSystemJobs =
      crffnResetV4_clearResetSystemJobsForApplication_(
        spreadsheet,
        cleanApplicationId
      );

    const cacheInvalidationActions =
      crffnResetV4_invalidateResetApplicationCaches_(
        cleanApplicationId
      );

    const properties =
      PropertiesService
        .getDocumentProperties();

    Object.keys(
      applicantEmails
    ).forEach(
      function(email) {
        [
          'EMAIL_TO_APP::',
          'EMAIL_TO_TOKEN::',
          'EMAIL_SUBMISSION_COUNT::',
          'SUBMISSION_COUNT::'
        ].forEach(
          function(prefix) {
            properties.deleteProperty(
              prefix +
              email
            );
          }
        );
      }
    );

    [
      'APP_TO_TOKEN::',
      'APP_TO_EMAIL::',
      'APP_TO_PORTAL::',
      'APPLICATION_TOKEN::',
      'APPLICATION_EMAIL::',
      'APPLICATION_PORTAL_URL::'
    ].forEach(
      function(prefix) {
        properties.deleteProperty(
          prefix +
          cleanApplicationId
        );
      }
    );

    Object.keys(
      secureTokens
    ).forEach(
      function(token) {
        [
          'TOKEN_TO_APP::',
          'PORTAL_TOKEN::',
          'SECURE_TOKEN::'
        ].forEach(
          function(prefix) {
            properties.deleteProperty(
              prefix +
              token
            );
          }
        );
      }
    );

    return {
      ok: true,

      applicationId:
        cleanApplicationId,

      rowsDeleted:
        matchingRows.length,

      filesTrashed:
        filesTrashed,

      fileCleanupErrors:
        fileCleanupErrors,

      systemJobsCleared:
        clearedSystemJobs,

      cacheInvalidationActions:
        cacheInvalidationActions,

      message: [
        'Test application reset completed.',
        '',
        'Application ID: ' +
          cleanApplicationId,
        'Rows removed: ' +
          matchingRows.length,
        'Files moved to Trash: ' +
          filesTrashed,
        'System Jobs cleared: ' +
          clearedSystemJobs,
        'Cache invalidation actions: ' +
          cacheInvalidationActions,
        fileCleanupErrors
          ? (
            'Drive cleanup warnings: ' +
            fileCleanupErrors
          )
          : '',
        'Reset Engine: ' +
          CRFFN_GO_LIVE_RESET_VERSION
      ]
        .filter(Boolean)
        .join('\n')
    };
  } finally {
    lock.releaseLock();
  }
}


function crffnResetV4_extractResetDriveFileId_(
  url
) {
  const value =
    String(
      url || ''
    ).trim();

  if (!value) {
    return '';
  }

  const patterns = [
    /\/file\/d\/([a-zA-Z0-9_-]+)/,
    /\/document\/d\/([a-zA-Z0-9_-]+)/,
    /\/spreadsheets\/d\/([a-zA-Z0-9_-]+)/,
    /[?&]id=([a-zA-Z0-9_-]+)/
  ];

  for (
    let index = 0;
    index <
      patterns.length;
    index++
  ) {
    const match =
      value.match(
        patterns[
          index
        ]
      );

    if (match) {
      return match[1];
    }
  }

  return '';
}
function freshStartSystem() {
  const ss =
    SpreadsheetApp.getActiveSpreadsheet();

  if (!ss) {
    throw new Error(
      'Could not open the active spreadsheet.'
    );
  }

  const lock =
    LockService.getDocumentLock();

  lock.waitLock(30000);

  try {
    const responseSheet =
      getResponseSheet_(ss);

    /*
     * 1. Trash files referenced by old test rows.
     *    Keeps folders/templates themselves.
     */
    const fileCleanup =
      crffnResetV4_trashFilesReferencedByTestRows_(
        responseSheet
      );

    /*
     * 2. Clear application data only.
     *    Row 1 / headers remain.
     */
    const clearedRows =
      crffnResetV4_clearAllApplicationRows_(
        responseSheet
      );

    /*
     * 3. Clear old applicant memory.
     */
    const clearedApplicantProperties =
      crffnResetV4_clearApplicantMemoryProperties_();

    /*
     * 4. Clear queued system jobs.
     */
    const clearedJobs =
      crffnResetV4_clearAllResetSystemJobs_(
        ss
      );

    /*
     * 5. Reset numbering.
     *
     * Next application:
     * APP-0001
     *
     * Next S/N:
     * 1
     */
    const documentProperties =
      PropertiesService
        .getDocumentProperties();

    documentProperties.setProperty(
      PROP_KEYS.APP_COUNTER,
      '0'
    );

    documentProperties.setProperty(
      PROP_KEYS.SN_COUNTER,
      '0'
    );

    /*
     * 6. Clear Admin Users data,
     *    but preserve headers.
     */
    const adminUsers =
      ss.getSheetByName(
        ADMIN_AUTH_CONFIG.USERS_SHEET
      );

    if (
      adminUsers &&
      adminUsers.getLastRow() > 1
    ) {
      adminUsers
        .getRange(
          2,
          1,
          adminUsers.getLastRow() - 1,
          adminUsers.getLastColumn()
        )
        .clearContent();
    }

    /*
     * 7. Clear Admin Sessions data,
     *    but preserve headers.
     */
    const adminSessions =
      ss.getSheetByName(
        ADMIN_AUTH_CONFIG.SESSIONS_SHEET
      );

    if (
      adminSessions &&
      adminSessions.getLastRow() > 1
    ) {
      adminSessions
        .getRange(
          2,
          1,
          adminSessions.getLastRow() - 1,
          adminSessions.getLastColumn()
        )
        .clearContent();
    }

    /*
     * 8. Reset admin-auth secrets only.
     *
     * Do NOT touch ZeptoMail,
     * WEB_APP_URL, etc.
     */
    const scriptProperties =
      PropertiesService
        .getScriptProperties();

    scriptProperties.deleteProperty(
      ADMIN_AUTH_CONFIG.PROP_PEPPER
    );

    scriptProperties.deleteProperty(
      ADMIN_AUTH_CONFIG.PROP_DEFAULT_PASSWORD
    );

    /*
     * 9. Create fresh admin authentication.
     */
    const adminSetup =
      setupAdminAuthentication();

    SpreadsheetApp.flush();

    /*
     * 10. Refresh Setup tab.
     */
    refreshSetupTab();

    Logger.log(
      'FRESH START COMPLETE'
    );

    Logger.log(
      'Rows cleared: ' +
      clearedRows
    );

    Logger.log(
      'Files moved to Trash: ' +
      fileCleanup.trashed
    );

    Logger.log(
      'System Jobs cleared: ' +
      clearedJobs
    );

    Logger.log(
      'Applicant properties cleared: ' +
      clearedApplicantProperties
    );

    Logger.log(
      'Next Application: APP-0001'
    );

    Logger.log(
      'Next S/N: 1'
    );

    Logger.log(
      'TEMPORARY ADMIN PASSWORD: ' +
      adminSetup.defaultPassword
    );

    return {
      ok: true,
      nextApplication: 'APP-0001',
      nextSerialNumber: 1,
      temporaryAdminPassword:
        adminSetup.defaultPassword
    };
  } finally {
    lock.releaseLock();
  }
}