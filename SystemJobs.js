/**
 * ============================================================
 * CRFFN SYSTEM JOB QUEUE
 * ============================================================
 *
 * Slow follow-up work (primarily email) runs here.
 *
 * IMPORTANT:
 * - AdminPortal.gs decides and commits business state first.
 * - This queue does NOT decide whether a stage is approved/rejected.
 * - The Applications sheet remains the source of truth.
 * - Jobs are retried safely and duplicate active/completed jobs are blocked.
 *
 * Run setupSystemJobs() ONCE from the Apps Script editor.
 */

const SYSTEM_JOBS_CONFIG = Object.freeze({
  SHEET_NAME: 'System Jobs',
  HANDLER: 'processSystemJobs',
  MAX_JOBS_PER_RUN: 10,
  MAX_ATTEMPTS: 3,
  STALE_PROCESSING_MINUTES: 10,
  RETRY_MINUTES: [1, 5, 15],
  HEADERS: [
    'Job ID',
    'Job Key',
    'Job Type',
    'Application ID',
    'Payload JSON',
    'Status',
    'Attempts',
    'Created At',
    'Next Attempt At',
    'Processing Started At',
    'Last Error',
    'Completed At'
  ]
});


function setupSystemJobs() {
  ensureSystemJobsSheet_();
  installSystemJobsTrigger_();

  return {
    ok: true,
    message: 'System Jobs queue is ready and the one-minute trigger is installed.'
  };
}


function queueSystemJob_(options) {
  const input = options || {};
  const jobType = String(input.jobType || '').trim().toUpperCase();
  const applicationId = String(input.applicationId || '').trim();
  const jobKey = String(input.jobKey || '').trim();
  const payload =
    input.payload && typeof input.payload === 'object'
      ? input.payload
      : {};

  if (!jobType) {
    throw new Error('System job type is required.');
  }

  if (!applicationId) {
    throw new Error('Application ID is required.');
  }

  if (!jobKey) {
    throw new Error('System job key is required.');
  }

  const lock = LockService.getScriptLock();

  if (!lock.tryLock(5000)) {
    throw new Error('The background-job queue is busy. Please try again.');
  }

  try {
    const sheet = ensureSystemJobsSheet_();
    const table = readSystemJobsTable_(sheet);

    const duplicate = table.rows.find(function(row) {
      return (
        String(row['Job Key'] || '').trim() === jobKey &&
        ['Pending', 'Processing', 'Completed'].includes(
          String(row['Status'] || '').trim()
        )
      );
    });

    if (duplicate) {
      return {
        ok: true,
        duplicate: true,
        jobId: String(duplicate['Job ID'] || ''),
        status: String(duplicate['Status'] || '')
      };
    }

    const now = new Date();
    const jobId = 'JOB-' + Utilities.getUuid();

    sheet
      .getRange(sheet.getLastRow() + 1, 1, 1, SYSTEM_JOBS_CONFIG.HEADERS.length)
      .setValues([[
        jobId,
        jobKey,
        jobType,
        applicationId,
        JSON.stringify(payload),
        'Pending',
        0,
        now,
        now,
        '',
        '',
        ''
      ]]);

    return {
      ok: true,
      duplicate: false,
      jobId: jobId,
      status: 'Pending'
    };
  } finally {
    lock.releaseLock();
  }
}


function processSystemJobs() {
  const sheet =
    ensureSystemJobsSheet_();

  /*
   * Claim a small batch under a short lock, then release the lock
   * BEFORE sending email. Slow email must never block admin writes.
   */
  const lock =
    LockService.getScriptLock();

  if (!lock.tryLock(3000)) {
    return;
  }

  let claimedJobs =
    [];

  try {
    recoverStaleSystemJobs_(
      sheet
    );

    const table =
      readSystemJobsTable_(
        sheet
      );

    const now =
      Date.now();

    claimedJobs =
      table.rows
        .filter(function(row) {
          if (
            String(
              row['Status'] || ''
            ).trim() !==
            'Pending'
          ) {
            return false;
          }

          const nextAttemptAt =
            row[
              'Next Attempt At'
            ];

          return !(
            nextAttemptAt instanceof Date &&
            nextAttemptAt.getTime() >
              now
          );
        })
        .slice(
          0,
          SYSTEM_JOBS_CONFIG
            .MAX_JOBS_PER_RUN
        );

    claimedJobs.forEach(
      function(job) {
        const attempts =
          Number(
            job['Attempts'] || 0
          ) + 1;

        updateSystemJobRow_(
          sheet,
          table.headerMap,
          job.__rowNumber,
          {
            'Status':
              'Processing',
            'Attempts':
              attempts,
            'Processing Started At':
              new Date(),
            'Last Error':
              ''
          }
        );

        job.__claimedAttempts =
          attempts;

        job.__headerMap =
          table.headerMap;
      }
    );
  } finally {
    lock.releaseLock();
  }

  /*
   * Email / other slow work happens with NO script lock held.
   */
  claimedJobs.forEach(
    function(job) {
      processClaimedSystemJob_(
        sheet,
        job
      );
    }
  );
}


function processClaimedSystemJob_(
  sheet,
  job
) {
  const rowNumber =
    Number(
      job.__rowNumber
    );

  const attempts =
    Number(
      job.__claimedAttempts ||
      job['Attempts'] ||
      1
    );

  const headerMap =
    job.__headerMap ||
    readSystemJobsTable_(
      sheet
    ).headerMap;

  const jobType =
    String(
      job['Job Type'] || ''
    )
      .trim()
      .toUpperCase();

  const applicationId =
    String(
      job['Application ID'] || ''
    ).trim();

  try {
    const payload =
      parseSystemJobPayload_(
        job['Payload JSON']
      );

    executeSystemJob_(
      jobType,
      applicationId,
      payload
    );

    updateSystemJobRow_(
      sheet,
      headerMap,
      rowNumber,
      {
        'Status':
          'Completed',
        'Completed At':
          new Date(),
        'Processing Started At':
          '',
        'Next Attempt At':
          '',
        'Last Error':
          ''
      }
    );
  } catch (error) {
    const exhausted =
      attempts >=
      SYSTEM_JOBS_CONFIG
        .MAX_ATTEMPTS;

    const retryMinutes =
      SYSTEM_JOBS_CONFIG
        .RETRY_MINUTES[
          Math.min(
            attempts - 1,
            SYSTEM_JOBS_CONFIG
              .RETRY_MINUTES.length - 1
          )
        ] ||
      15;

    updateSystemJobRow_(
      sheet,
      headerMap,
      rowNumber,
      {
        'Status':
          exhausted
            ? 'Failed'
            : 'Pending',
        'Processing Started At':
          '',
        'Next Attempt At':
          exhausted
            ? ''
            : new Date(
                Date.now() +
                retryMinutes *
                  60 *
                  1000
              ),
        'Last Error':
          String(
            error &&
            error.message
              ? error.message
              : error
          )
      }
    );
  }
}


function executeSystemJob_(jobType, applicationId, payload) {
  switch (jobType) {
    case 'PAYMENT_INVITATION_EMAIL':
      sendCrffnPaymentInvitationEmail_(payload);
      markSystemJobApplicationAudit_(applicationId, {
        'Payment Invitation Sent At': new Date()
      });
      return;

    case 'DOCUMENT_CORRECTION_EMAIL':
      sendSupportingDocumentsCorrectionSummaryEmail_(payload);
      markSystemJobApplicationAudit_(applicationId, {
        'Document Correction Email Sent At': new Date(),
        'Document Correction Email Sent By': 'System Job'
      });
      return;

    case 'ADMIN_REVIEW_CORRECTION_EMAIL':
      sendAdminReviewCorrectionEmail_(payload);
      return;

    default:
      throw new Error('Unsupported system job type: ' + jobType);
  }
}


function markSystemJobApplicationAudit_(applicationId, changes) {
  const sheet =
    getResponseSheet_(
      SpreadsheetApp.getActiveSpreadsheet()
    );

  const record =
    findAdminApplicationRecord_(
      sheet,
      String(applicationId || '').trim()
    );

  Object.keys(changes || {}).forEach(function(header) {
    if (!record.headerMap[header]) {
      console.warn('Optional audit column missing: ' + header);
      return;
    }

    sheet
      .getRange(record.rowNumber, record.headerMap[header])
      .setValue(changes[header]);

    record.rowValues[record.headerMap[header] - 1] = changes[header];
  });

  invalidateAdminApplicationCacheSafe_(applicationId);

  if (typeof invalidateApplicantCacheByApplicationId_ === 'function') {
    invalidateApplicantCacheByApplicationId_(applicationId);
  }
}


function retryFailedSystemJobs() {
  const sheet = ensureSystemJobsSheet_();
  const table = readSystemJobsTable_(sheet);
  let resetCount = 0;

  table.rows.forEach(function(job) {
    if (String(job['Status'] || '').trim() !== 'Failed') {
      return;
    }

    updateSystemJobRow_(sheet, table.headerMap, job.__rowNumber, {
      'Status': 'Pending',
      'Attempts': 0,
      'Next Attempt At': new Date(),
      'Processing Started At': '',
      'Last Error': ''
    });

    resetCount++;
  });

  return {
    ok: true,
    resetCount: resetCount
  };
}


function getSystemJobsSummary() {
  const sheet = ensureSystemJobsSheet_();
  const table = readSystemJobsTable_(sheet);
  const summary = {
    Pending: 0,
    Processing: 0,
    Completed: 0,
    Failed: 0
  };

  table.rows.forEach(function(job) {
    const status = String(job['Status'] || '').trim();

    if (Object.prototype.hasOwnProperty.call(summary, status)) {
      summary[status]++;
    }
  });

  return {
    ok: true,
    total: table.rows.length,
    summary: summary
  };
}


function ensureSystemJobsSheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(SYSTEM_JOBS_CONFIG.SHEET_NAME);

  if (!sheet) {
    sheet = ss.insertSheet(SYSTEM_JOBS_CONFIG.SHEET_NAME);
  }

  if (sheet.getMaxColumns() < SYSTEM_JOBS_CONFIG.HEADERS.length) {
    sheet.insertColumnsAfter(
      sheet.getMaxColumns(),
      SYSTEM_JOBS_CONFIG.HEADERS.length - sheet.getMaxColumns()
    );
  }

  const existing =
    sheet.getRange(1, 1, 1, SYSTEM_JOBS_CONFIG.HEADERS.length)
      .getDisplayValues()[0];

  const mismatch =
    SYSTEM_JOBS_CONFIG.HEADERS.some(function(header, index) {
      return String(existing[index] || '').trim() !== header;
    });

  if (sheet.getLastRow() === 0 || mismatch) {
    sheet
      .getRange(1, 1, 1, SYSTEM_JOBS_CONFIG.HEADERS.length)
      .setValues([SYSTEM_JOBS_CONFIG.HEADERS])
      .setFontWeight('bold');

    sheet.setFrozenRows(1);
  }

  return sheet;
}


function readSystemJobsTable_(sheet) {
  const lastRow = sheet.getLastRow();
  const lastColumn = SYSTEM_JOBS_CONFIG.HEADERS.length;
  const headers =
    sheet.getRange(1, 1, 1, lastColumn)
      .getDisplayValues()[0]
      .map(function(value) {
        return String(value || '').trim();
      });

  const headerMap = {};
  headers.forEach(function(header, index) {
    if (header) {
      headerMap[header] = index + 1;
    }
  });

  if (lastRow < 2) {
    return {
      headers: headers,
      headerMap: headerMap,
      rows: []
    };
  }

  const values =
    sheet.getRange(2, 1, lastRow - 1, lastColumn)
      .getValues();

  const rows =
    values.map(function(row, index) {
      const item = {
        __rowNumber: index + 2
      };

      headers.forEach(function(header, columnIndex) {
        item[header] = row[columnIndex];
      });

      return item;
    });

  return {
    headers: headers,
    headerMap: headerMap,
    rows: rows
  };
}


function updateSystemJobRow_(sheet, headerMap, rowNumber, changes) {
  const lastColumn = SYSTEM_JOBS_CONFIG.HEADERS.length;
  const row =
    sheet.getRange(rowNumber, 1, 1, lastColumn)
      .getValues()[0];

  Object.keys(changes || {}).forEach(function(header) {
    const column = headerMap[header];

    if (!column) {
      throw new Error('System Jobs column is missing: ' + header);
    }

    row[column - 1] = changes[header];
  });

  sheet
    .getRange(rowNumber, 1, 1, lastColumn)
    .setValues([row]);
}


function recoverStaleSystemJobs_(sheet) {
  const table = readSystemJobsTable_(sheet);
  const cutoff =
    Date.now() -
    SYSTEM_JOBS_CONFIG.STALE_PROCESSING_MINUTES * 60 * 1000;

  table.rows.forEach(function(job) {
    if (String(job['Status'] || '').trim() !== 'Processing') {
      return;
    }

    const startedAt = job['Processing Started At'];

    if (
      startedAt instanceof Date &&
      startedAt.getTime() > cutoff
    ) {
      return;
    }

    updateSystemJobRow_(sheet, table.headerMap, job.__rowNumber, {
      'Status': 'Pending',
      'Processing Started At': '',
      'Next Attempt At': new Date(),
      'Last Error': 'Recovered from stale Processing state.'
    });
  });
}


function parseSystemJobPayload_(value) {
  const text = String(value || '').trim();

  if (!text) {
    return {};
  }

  try {
    const parsed = JSON.parse(text);

    return parsed && typeof parsed === 'object'
      ? parsed
      : {};
  } catch (error) {
    throw new Error('System job payload is invalid JSON.');
  }
}


function installSystemJobsTrigger_() {
  const handler = SYSTEM_JOBS_CONFIG.HANDLER;

  ScriptApp.getProjectTriggers()
    .filter(function(trigger) {
      return trigger.getHandlerFunction() === handler;
    })
    .forEach(function(trigger) {
      ScriptApp.deleteTrigger(trigger);
    });

  ScriptApp.newTrigger(handler)
    .timeBased()
    .everyMinutes(1)
    .create();
}


function invalidateAdminApplicationCacheSafe_(applicationId) {
  try {
    if (typeof invalidateAdminApplicationCache_ === 'function') {
      invalidateAdminApplicationCache_(applicationId);
    }
  } catch (error) {
    console.warn('Admin cache invalidation warning: ' + error.message);
  }
}
