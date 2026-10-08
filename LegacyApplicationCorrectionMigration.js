/**
 * One-time, editor-only migration. Trailing underscore prevents web-app RPC use.
 * Default is read-only. Explicit { dryRun: false } is reserved for a separately
 * approved production run; no triggers, menus, deployment or automatic execution.
 */
function runLegacyCorrectionMigrationDryRun() {
  return migrateLegacyApplicationCorrections_({ dryRun: true });
}
function migrateLegacyApplicationCorrections_(options) {
  const dryRun = !options || options.dryRun !== false;
  const report = { dryRun: dryRun, startedAt: new Date().toISOString(), entries: [], counts: {} };
  const sheet = getResponseSheet_(SpreadsheetApp.getActiveSpreadsheet());
  const snapshot = readLegacyCorrectionMigrationRecords_(sheet);
  // Scan only active correction candidates, but authority considers ALL rows.
  snapshot.forEach(function(record) {
    if (String(getAdminRecordValue_(record, 'Application Information Status') || '').trim().toLowerCase() !== 'correction required') { return; }
    const id = String(getAdminRecordValue_(record, 'Application ID') || '').trim();
    const token = String(getAdminRecordValue_(record, 'Secure Token') || '').trim();
    let entry = describeLegacyCorrectionMigrationRecord_(record);
    let scriptLock;
    let documentLock;
    try {
      if (!id || !token) {
        entry.migrationStatus = 'SKIPPED_MISSING_CREDENTIALS';
      } else {
        if (!dryRun) {
          scriptLock = LockService.getScriptLock();
          scriptLock.waitLock(30000);
          documentLock = LockService.getDocumentLock();
          if (documentLock) { documentLock.waitLock(30000); }
        }
        // Re-read every candidate; never apply a cached dry-run preview.
        const current = findApplicantApplicationRecordByCredentials_(id, token);
        if (current.record.rowNumber !== record.rowNumber) {
          entry.migrationStatus = 'SKIPPED_NOT_AUTHORITATIVE';
        } else {
          entry = describeLegacyCorrectionMigrationRecord_(current.record);
          const ambiguity = readLegacyCorrectionMigrationRecords_(current.sheet).some(function(other) {
            return String(getAdminRecordValue_(other, 'Application ID') || '').trim() === id &&
              String(getAdminRecordValue_(other, 'Secure Token') || '').trim() !== token;
          });
          // One migrated row per ID. Different-token histories need review: the
          // pair selector alone cannot establish which identity should migrate.
          if (ambiguity) {
            entry.migrationStatus = 'SKIPPED_NOT_AUTHORITATIVE';
            entry.detail = 'Multiple or missing tokens for this Application ID; authority requires review.';
          } else {
            entry = prepareLegacyCorrectionMigration_(current.record);
            if (!dryRun && entry.migrationStatus === 'SAFE_TO_MIGRATE') {
              entry.backup = { applicationId: id, rowNumber: current.record.rowNumber,
                originalJson: getAdminRecordValue_(current.record, 'Application Correction Fields JSON'),
                newJson: entry.generatedJsonPreview, timestamp: new Date().toISOString() };
              // Fail closed if the pre-write audit cannot be retained in the log.
              console.log(JSON.stringify({ migration: 'legacy-application-corrections', phase: 'BEFORE_WRITE', backup: entry.backup }));
              current.sheet.getRange(current.record.rowNumber,
                current.record.headerMap['Application Correction Fields JSON'], 1, 1)
                .setValues([[entry.generatedJsonPreview]]);
              entry.migrationStatus = 'MIGRATED';
              // A cache failure cannot turn a completed write into a retry request.
              try { invalidateApplicantCacheByApplicationId_(id); }
              catch (error) { entry.cacheWarning = error.message; }
            }
          }
        }
      }
    } catch (error) {
      entry.migrationStatus = 'ERROR';
      entry.detail = error.message;
    } finally {
      if (documentLock && documentLock.hasLock()) { documentLock.releaseLock(); }
      if (scriptLock && scriptLock.hasLock()) { scriptLock.releaseLock(); }
    }
    report.entries.push(entry);
    report.counts[entry.migrationStatus] = (report.counts[entry.migrationStatus] || 0) + 1;
    console.log(JSON.stringify({ migration: 'legacy-application-corrections', dryRun: dryRun, entry: entry }));
  });
  report.completedAt = new Date().toISOString();
  return report;
}

function readLegacyCorrectionMigrationRecords_(sheet) {
  const lastColumn = sheet.getLastColumn();
  const headers = sheet.getRange(1, 1, 1, lastColumn).getDisplayValues()[0].map(function(header) { return String(header || '').trim(); });
  ['Application ID', 'Secure Token', 'Application Information Status', 'Application Information Review Notes', 'Application Correction Fields JSON'].forEach(function(header) {
    if (headers.filter(function(value) { return value === header; }).length !== 1) {
      throw new Error('Migration requires exactly one existing column: ' + header);
    }
  });
  if (sheet.getLastRow() < 2) { return []; }
  const range = sheet.getRange(2, 1, sheet.getLastRow() - 1, lastColumn);
  const values = range.getValues();
  const formulas = range.getFormulas();
  const headerMap = getHeaderMap_(headers);
  return values.map(function(row, index) {
    // A formula yielding an empty string is not an empty cell to overwrite.
    formulas[index].forEach(function(formula, column) { if (formula) { row[column] = formula; } });
    return { rowNumber: index + 2, headers: headers, headerMap: headerMap, rowValues: row };
  });
}

function describeLegacyCorrectionMigrationRecord_(record) {
  const value = function(header) { return getAdminRecordValue_(record, header); };
  return { applicationId: String(value('Application ID') || '').trim(), rowNumber: record.rowNumber,
    applicantName: value('Full Name') || '', submissionCount: value('Submission Count'),
    recordStatus: value('Record Status'), informationStatus: value('Application Information Status'),
    jsonBlank: String(value('Application Correction Fields JSON') == null ? '' : value('Application Correction Fields JSON')).trim() === '',
    correctionNote: value('Application Information Review Notes') || '', resolvedTargets: [], migrationStatus: '' };
}

function prepareLegacyCorrectionMigration_(record) {
  const entry = describeLegacyCorrectionMigrationRecord_(record);
  if (String(entry.informationStatus || '').trim().toLowerCase() !== 'correction required') {
    entry.migrationStatus = 'SKIPPED_STATUS_CHANGED';
    return entry;
  }
  if (!entry.jsonBlank) {
    const registry = getApplicationCorrectionRegistry_();
    const seen = {};
    let valid = false;
    try {
      const items = JSON.parse(String(getAdminRecordValue_(record, 'Application Correction Fields JSON')));
      valid = Array.isArray(items) && items.length > 0 && items.every(function(item) {
        if (!item || !Object.prototype.hasOwnProperty.call(registry, item.targetCode) || seen[item.targetCode] ||
            (item.status && ['pending', 'submitted', 'resolved'].indexOf(item.status) === -1)) { return false; }
        seen[item.targetCode] = true;
        return true;
      });
    } catch (error) { /* Never overwrite nonblank content, including malformed JSON. */ }
    entry.migrationStatus = valid ? 'SKIPPED_ALREADY_STRUCTURED' : 'SKIPPED_MALFORMED_JSON';
    if (!valid) { entry.detail = 'MALFORMED_EXISTING_JSON'; }
    return entry;
  }
  if (!String(entry.correctionNote).trim()) {
    entry.migrationStatus = 'SKIPPED_NO_CORRECTION_NOTE';
    return entry;
  }
  let items;
  try { items = getLegacyApplicationCorrectionItems_(record); }
  catch (error) {
    entry.migrationStatus = error.correctionCode === 'NO_EXISTING_FIELD' ? 'SKIPPED_NO_EXISTING_FIELD' : 'SKIPPED_UNSAFE_CORRECTION';
    entry.detail = error.message;
    return entry;
  }
  // Identity, aliases, issue codes, field availability and shared requestId were
  // already resolved by the strict legacy parser. Never infer submitted values.
  const preview = items.map(function(item) {
    const result = Object.assign({}, item, { reason: item.issueText });
    const marker = ' ' + item.issueText + ':';
    const offset = item.reason.toLowerCase().indexOf(marker.toLowerCase());
    if (item.issueCode === 'other' && offset !== -1) {
      let details = item.reason.slice(offset + marker.length).trim();
      // Some historical notes repeated the generated target/issue prefix after
      // the first prefix. Remove only those leading generated prefixes; preserve
      // all actual administrator instruction text and the original audit note.
      const registryField = getApplicationCorrectionRegistry_()[item.targetCode];
      let removedPrefix = true;
      while (removedPrefix) {
        removedPrefix = false;
        (registryField ? registryField.aliases : []).some(function(alias) {
          const repeatedPrefix = alias + ' ' + item.issueText + ':';
          if (details.slice(0, repeatedPrefix.length).toLowerCase() === repeatedPrefix.toLowerCase()) {
            details = details.slice(repeatedPrefix.length).trim();
            removedPrefix = true;
            return true;
          }
          return false;
        });
      }
      result.customDetails = details;
      result.reason = result.customDetails;
    }
    return result;
  });
  entry.resolvedTargets = preview.map(function(item) { return item.targetCode; });
  entry.generatedJsonPreview = JSON.stringify(preview);
  if (entry.generatedJsonPreview.length > 45000) {
    entry.migrationStatus = 'SKIPPED_UNSAFE_CORRECTION';
    entry.detail = 'Generated correction JSON exceeds the existing correction size limit.';
    return entry;
  }
  entry.migrationStatus = 'SAFE_TO_MIGRATE';
  return entry;
}
