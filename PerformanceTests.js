/**
 * ============================================================
 * CRFFN PERFORMANCE / SMOKE TESTS
 * ============================================================
 *
 * These tests DO NOT submit forms, upload files, approve stages,
 * reset passwords, or modify applicant records.
 *
 * Run them manually from Apps Script after replacing the files.
 */


/**
 * General dependency test.
 */
function testCRFFNPerformanceDependencies() {
  const requiredFunctions = [
    'crffnGetCachedJson_',
    'invalidateApplicantPerformanceCache_',
    'getApplicantPortalData_',
    'getApplicantSupportingDocuments',
    'getApplicantPaymentData',
    'validateAdminSession_'
  ];

  const missing =
    requiredFunctions.filter(
      function(name) {
        return (
          typeof this[name] !==
          'function'
        );
      }
    );

  if (missing.length) {
    throw new Error(
      'Missing required function(s): ' +
      missing.join(', ')
    );
  }

  console.log(
    'PASS: Performance dependencies are available.'
  );

  return {
    ok: true,
    message:
      'Performance dependencies are available.'
  };
}


/**
 * Applicant read-speed test.
 *
 * Replace the sample values before running.
 * This makes two read calls. The second call should normally
 * report performanceCache = HIT and be materially faster.
 */
function testApplicantPortalCachePerformance() {
  const applicationId =
    'REPLACE_WITH_APPLICATION_ID';

  const secureToken =
    'REPLACE_WITH_SECURE_TOKEN';

  if (
    applicationId.indexOf(
      'REPLACE_'
    ) === 0 ||
    secureToken.indexOf(
      'REPLACE_'
    ) === 0
  ) {
    throw new Error(
      'Edit testApplicantPortalCachePerformance() and enter a real test Application ID and secure token first.'
    );
  }

  invalidateApplicantPerformanceCache_(
    applicationId,
    secureToken
  );

  invalidateApplicantCacheByApplicationId_(
    applicationId
  );

  const firstStart =
    Date.now();

  const first =
    getApplicantPortalData_(
      applicationId,
      secureToken
    );

  const firstMs =
    Date.now() -
    firstStart;

  const secondStart =
    Date.now();

  const second =
    getApplicantPortalData_(
      applicationId,
      secureToken
    );

  const secondMs =
    Date.now() -
    secondStart;

  const result = {
    ok:
      Boolean(
        first &&
        first.ok &&
        second &&
        second.ok
      ),
    firstCallMs:
      firstMs,
    firstCache:
      first &&
      first.performanceCache,
    secondCallMs:
      secondMs,
    secondCache:
      second &&
      second.performanceCache,
    improvementMs:
      firstMs -
      secondMs
  };

  console.log(
    JSON.stringify(
      result,
      null,
      2
    )
  );

  return result;
}


/**
 * Supporting-document read cache test.
 */
function testSupportingDocumentsCachePerformance() {
  const applicationId =
    'REPLACE_WITH_APPLICATION_ID';

  const secureToken =
    'REPLACE_WITH_SECURE_TOKEN';

  if (
    applicationId.indexOf(
      'REPLACE_'
    ) === 0 ||
    secureToken.indexOf(
      'REPLACE_'
    ) === 0
  ) {
    throw new Error(
      'Enter a real test Application ID and secure token before running this test.'
    );
  }

  invalidateApplicantPerformanceCache_(
    applicationId,
    secureToken
  );

  const start1 =
    Date.now();

  const first =
    getApplicantSupportingDocuments(
      applicationId,
      secureToken
    );

  const ms1 =
    Date.now() -
    start1;

  const start2 =
    Date.now();

  const second =
    getApplicantSupportingDocuments(
      applicationId,
      secureToken
    );

  const ms2 =
    Date.now() -
    start2;

  const result = {
    ok:
      Boolean(
        first &&
        first.ok &&
        second &&
        second.ok
      ),
    firstCallMs:
      ms1,
    firstCache:
      first &&
      first.performanceCache,
    secondCallMs:
      ms2,
    secondCache:
      second &&
      second.performanceCache
  };

  console.log(
    JSON.stringify(
      result,
      null,
      2
    )
  );

  return result;
}


/**
 * Payment read cache test.
 */
function testPaymentCachePerformance() {
  const applicationId =
    'REPLACE_WITH_APPLICATION_ID';

  const secureToken =
    'REPLACE_WITH_SECURE_TOKEN';

  if (
    applicationId.indexOf(
      'REPLACE_'
    ) === 0 ||
    secureToken.indexOf(
      'REPLACE_'
    ) === 0
  ) {
    throw new Error(
      'Enter a real test Application ID and secure token before running this test.'
    );
  }

  invalidateApplicantPerformanceCache_(
    applicationId,
    secureToken
  );

  const start1 =
    Date.now();

  const first =
    getApplicantPaymentData(
      applicationId,
      secureToken
    );

  const ms1 =
    Date.now() -
    start1;

  const start2 =
    Date.now();

  const second =
    getApplicantPaymentData(
      applicationId,
      secureToken
    );

  const ms2 =
    Date.now() -
    start2;

  const result = {
    ok:
      Boolean(
        first &&
        first.ok &&
        second &&
        second.ok
      ),
    firstCallMs:
      ms1,
    firstCache:
      first &&
      first.performanceCache,
    secondCallMs:
      ms2,
    secondCache:
      second &&
      second.performanceCache
  };

  console.log(
    JSON.stringify(
      result,
      null,
      2
    )
  );

  return result;
}


/**
 * Read-only sheet benchmark. Useful for seeing how much of the
 * delay comes from Google Sheets itself.
 */
function benchmarkCRFFNResponseSheetRead() {
  const ss =
    SpreadsheetApp.getActiveSpreadsheet();

  const start =
    Date.now();

  const sheet =
    getResponseSheet_(
      ss
    );

  const lastRow =
    sheet.getLastRow();

  const lastColumn =
    sheet.getLastColumn();

  let rowsRead =
    0;

  if (
    lastRow >= 2 &&
    lastColumn >= 1
  ) {
    const rows =
      Math.min(
        lastRow - 1,
        100
      );

    sheet
      .getRange(
        lastRow - rows + 1,
        1,
        rows,
        lastColumn
      )
      .getDisplayValues();

    rowsRead =
      rows;
  }

  const elapsed =
    Date.now() -
    start;

  const result = {
    ok: true,
    rowsRead:
      rowsRead,
    columns:
      lastColumn,
    elapsedMs:
      elapsed
  };

  console.log(
    JSON.stringify(
      result,
      null,
      2
    )
  );

  return result;
}
