/** Stable Application Information fields. Labels are presentation, never identity. */
function getApplicationCorrectionRegistry_() {
  const fields = [
    ['full_name', 'Full Name', 'text', true, ['Applicant Name']],
    ['gender', 'Gender', 'select', true, [], ['Male', 'Female']],
    ['date_of_birth', 'Date of Birth', 'date', true],
    ['nationality', 'Nationality', 'select', true],
    ['state_of_origin', 'State of Origin — Nigerian Applicants Only', 'select', false, ['State of Origin']],
    ['residential_address', 'Residential Address', 'textarea', true],
    ['phone_number', 'Phone Number', 'text', true, ['Company Contact Number', 'Phone', 'Contact Number']],
    ['email_address', 'Email Address', 'email', true, ['Email']],
    ['means_of_identification', 'Means of Identification', 'select', true, [], ['National ID', 'International Passport', 'Driver’s License', 'Voter’s Card']],
    ['id_number', 'ID Number', 'text', true],
    ['expiry_date', 'Expiry Date', 'date', false],
    ['area_of_practice', 'Area of Practice', 'multiselect', true, [], ['Clearing & Forwarding', 'Haulage/Transportation Services', 'Warehousing', 'Courier Services', 'Cold Chain', 'Other']],
    ['other_area_of_practice', 'If Other, please specify', 'text', false],
    ['company_name', 'Company Name', 'text', true, ['Business Name']],
    ['company_rc_number', 'Company RC Number', 'text', true, ['RC Number', 'CAC Registration Number']],
    ['company_tin', 'Company TIN', 'text', true, ['Company Tax Identification Number (TIN)', 'Tax Identification Number', 'TIN']],
    ['crffn_membership_number', 'CRFFN Corporate Membership Number', 'text', true, ['CRFFN Membership Number', 'CRFFN Membership No', 'CRFFN Registration Number', 'CRFFN Registration No', 'CRFFN Reg Number', 'CRFFN Number', 'Registration Number']],
    ['company_address', 'Company Address', 'textarea', true, ['Company Office Address', 'Office Address', 'Address']],
    ['position_held', 'Position Held', 'text', true],
    ['other', 'Other Application Information', 'textarea', true]
  ];
  const registry = {};
  fields.forEach(function(field) {
    registry[field[0]] = { key: field[0], label: field[1], type: field[2], required: field[3], aliases: [field[1]].concat(field[4] || []), choices: field[5] || [] };
  });
  return registry;
}

function getApplicationCorrectionTargets_() {
  const registry = getApplicationCorrectionRegistry_();
  return Object.keys(registry).map(function(key) { return [key, registry[key].label]; });
}

function getApplicationCorrectionLabels_() {
  const labels = {};
  getApplicationCorrectionTargets_().forEach(function(field) { labels[field[0]] = field[1]; });
  return Object.freeze(labels);
}

function normalizeApplicationFieldLabel_(value) {
  return String(value || '').toLowerCase().replace(/[^a-z0-9]/g, '');
}

function resolveApplicationCorrectionKey_(value) {
  const registry = getApplicationCorrectionRegistry_();
  if (Object.prototype.hasOwnProperty.call(registry, value)) { return value; }
  const label = normalizeApplicationFieldLabel_(value);
  return Object.keys(registry).find(function(key) {
    return registry[key].aliases.some(function(alias) { return normalizeApplicationFieldLabel_(alias) === label; });
  }) || '';
}

function getApplicationCorrectionHeader_(record, key) {
  const field = getApplicationCorrectionRegistry_()[key];
  if (!field || key === 'other') { return ''; }
  for (let index = 0; index < field.aliases.length; index++) {
    const alias = field.aliases[index];
    if (record.headerMap[alias]) { return alias; }
  }
  // Preserve the existing membership reader's normalized legacy-header support.
  return record.headers.find(function(header) {
    const label = normalizeApplicationFieldLabel_(header);
    return field.aliases.some(function(alias) { return normalizeApplicationFieldLabel_(alias) === label; }) ||
      (key === 'crffn_membership_number' && label.indexOf('crffn') !== -1 &&
        (label.indexOf('membership') !== -1 || label.indexOf('registration') !== -1 || label.indexOf('regnumber') !== -1));
  }) || '';
}

function getApplicationCorrectionValue_(record, key) {
  const field = getApplicationCorrectionRegistry_()[key];
  if (!field || key === 'other') { return ''; }
  const headers = field.aliases.concat([getApplicationCorrectionHeader_(record, key)]);
  for (let index = 0; index < headers.length; index++) {
    const value = getAdminRecordValue_(record, headers[index]);
    if (value instanceof Date) { return Utilities.formatDate(value, Session.getScriptTimeZone(), 'yyyy-MM-dd'); }
    if (value != null && String(value).trim()) { return String(value).trim(); }
  }
  return '';
}

function getApplicationCorrectionChoices_(field) {
  if (field.key === 'nationality') { return getCountries_(); }
  if (field.key === 'state_of_origin') { return getNigerianStates_(); }
  return field.choices;
}

function buildApplicationCorrectionRequest_(record, items, details) {
  if (!Array.isArray(items) || !items.length) { throw new Error('Select at least one correction field.'); }
  const registry = getApplicationCorrectionRegistry_();
  const seen = {};
  const requestId = Utilities.getUuid();
  return items.map(function(item) {
    const key = item && item.targetCode;
    if (!Object.prototype.hasOwnProperty.call(registry, key) || seen[key]) { throw new Error('Unknown or duplicate correction target.'); }
    if (key !== 'other' && !getApplicationCorrectionHeader_(record, key)) {
      throw new Error('The application has no existing column for ' + registry[key].label + '. No correction request was saved.');
    }
    seen[key] = true;
    return { targetCode: key, targetLabel: registry[key].label, issueCode: item.issueCode,
      issueText: item.issueText, reason: String(details || item.issueText || '').trim(),
      requestId: requestId, status: 'pending' };
  });
}

// Only recognise complete sentences produced by the legacy Admin correction formatter.
// Never search the explanatory text for field names or accept a partially matched request.
function getLegacyApplicationCorrectionItems_(record) {
  const fail = function(code) {
    const error = new Error('This correction requires administrator review because its requested fields cannot be safely identified.');
    if (code) { error.correctionCode = code; }
    throw error;
  };
  if (String(getAdminRecordValue_(record, 'Application Information Status') || '').trim().toLowerCase() !== 'correction required' ||
      !record.headerMap['Application Correction Fields JSON']) { return fail(); }
  const registry = getApplicationCorrectionRegistry_();
  const notes = String(getAdminRecordValue_(record, 'Application Information Review Notes') || '').trim();
  if (!notes) { return fail(); }
  const seen = {};
  const items = notes.split(/\r?\n/).filter(function(line) { return line.trim(); }).map(function(line) {
    const text = line.trim().replace(/\s+/g, ' ');
    const matches = [];
    Object.keys(registry).forEach(function(key) {
      // A generic explanation is not an identifiable application field.
      if (key === 'other') { return; }
      Object.keys(ADMIN_TARGET_ISSUES.information).forEach(function(issueCode) {
        const issueText = ADMIN_TARGET_ISSUES.information[issueCode];
        registry[key].aliases.forEach(function(alias) {
          const prefix = (alias + ' ' + issueText).replace(/\s+/g, ' ');
          if (text.slice(0, prefix.length).toLowerCase() !== prefix.toLowerCase()) { return; }
          const suffix = text.slice(prefix.length);
          // The old formatter appended custom details only for the "other" issue.
          if (!/^[.!?]?$/.test(suffix) && !(issueCode === 'other' && /^:\s+\S/.test(suffix))) { return; }
          if (!matches.some(function(item) { return item.targetCode === key && item.issueCode === issueCode; })) {
            matches.push({ targetCode: key, targetLabel: registry[key].label, issueCode: issueCode,
              issueText: issueText, reason: text, status: 'pending' });
          }
        });
      });
    });
    if (matches.length !== 1 || seen[matches[0].targetCode]) { return fail(); }
    if (!getApplicationCorrectionHeader_(record, matches[0].targetCode)) { return fail('NO_EXISTING_FIELD'); }
    seen[matches[0].targetCode] = true;
    return matches[0];
  });
  const requestId = Utilities.getUuid();
  return items.map(function(item) { return Object.assign(item, { requestId: requestId }); });
}

function getPendingApplicationCorrectionItems_(record) {
  let items;
  try { items = JSON.parse(String(getAdminRecordValue_(record, 'Application Correction Fields JSON') || '')); }
  catch (error) { return getLegacyApplicationCorrectionItems_(record); }
  const registry = getApplicationCorrectionRegistry_();
  const seen = {};
  // Do not resurrect a submitted/resolved structured request from stale review notes.
  if (Array.isArray(items) && items.some(function(item) { return item && item.status && item.status !== 'pending'; })) {
    throw new Error('The correction request is invalid or has already been submitted. Please contact the administrator.');
  }
  const valid = Array.isArray(items) && items.length && items.every(function(item) {
    if (!item || !Object.prototype.hasOwnProperty.call(registry, item.targetCode) || seen[item.targetCode]) { return false; }
    seen[item.targetCode] = true;
    return true;
  });
  if (!valid) { return getLegacyApplicationCorrectionItems_(record); }
  items.forEach(function(item) {
    if (item.targetCode !== 'other' && !getApplicationCorrectionHeader_(record, item.targetCode)) {
      throw new Error('The requested field is unavailable: ' + registry[item.targetCode].label + '. Please contact the administrator.');
    }
  });
  return items;
}

/**
 * Applicant identity is the pair (Application ID, Secure Token), never email.
 * The last physical response row matching BOTH is authoritative, including when
 * historical resubmissions reused both credentials. Never search older rows for
 * a preferred status or correction JSON. Callers retain their existing locks.
 */
function findApplicantApplicationRecordByCredentials_(applicationId, secureToken) {
  const id = String(applicationId || '').trim();
  const token = String(secureToken || '').trim();
  const fail = function(code, message) {
    const error = new Error(message);
    error.errorCode = code;
    throw error;
  };
  if (!id || !token) { return fail('MISSING_CREDENTIALS', 'The secure application link is incomplete.'); }
  const sheet = getResponseSheet_(SpreadsheetApp.getActiveSpreadsheet());
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) { return fail('APPLICATION_NOT_FOUND', 'No applications were found.'); }
  const lastColumn = sheet.getLastColumn();
  const headers = sheet.getRange(1, 1, 1, lastColumn).getDisplayValues()[0].map(function(header) { return String(header || '').trim(); });
  const headerMap = getHeaderMap_(headers);
  ['Application ID', 'Secure Token'].forEach(function(header) {
    if (!headerMap[header]) { throw new Error('Required column is missing: ' + header); }
  });
  const rows = sheet.getRange(2, 1, lastRow - 1, lastColumn).getDisplayValues();
  for (let index = rows.length - 1; index >= 0; index--) {
    if (String(rows[index][headerMap['Application ID'] - 1] || '').trim() !== id ||
        String(rows[index][headerMap['Secure Token'] - 1] || '').trim() !== token) { continue; }
    const rowNumber = index + 2;
    const range = sheet.getRange(rowNumber, 1, 1, lastColumn);
    const rowValues = range.getValues()[0];
    range.getFormulas()[0].forEach(function(formula, column) {
      if (formula) { rowValues[column] = formula; }
    });
    return { sheet: sheet, displayRowValues: rows[index],
      record: { rowNumber: rowNumber, headers: headers, headerMap: headerMap, rowValues: rowValues } };
  }
  return fail('INVALID_ACCESS', 'The secure application link is invalid or no longer active.');
}

function requireApplicantCorrectionRecord_(applicationId, secureToken) {
  const context = findApplicantApplicationRecordByCredentials_(applicationId, secureToken);
  const record = context.record;
  if (String(getAdminRecordValue_(record, 'Application Information Status') || '').trim().toLowerCase() !== 'correction required') {
    throw new Error('There is no active application information correction request.');
  }
  if (['generated', 'printed', 'stamped', 'uploaded', 'released'].indexOf(String(getAdminRecordValue_(record, 'Licence Status') || '').trim().toLowerCase()) !== -1) {
    throw new Error('Application information cannot be changed after licence processing has started.');
  }
  return context;
}

// Presentation only: never replace the administrator's stored audit/history text.
function getApplicantApplicationCorrectionReason_(item) {
  const customDetails = typeof item.customDetails === 'string' ? item.customDetails.trim() : '';
  const reason = typeof item.reason === 'string' ? item.reason.trim() : '';
  return customDetails || reason || item.issueText || '';
}

function getApplicantApplicationCorrectionReasons_(record, resolvedItems) {
  if (String(getAdminRecordValue_(record, 'Application Information Status') || '').trim().toLowerCase() !== 'correction required') { return []; }
  let items;
  try { items = resolvedItems || getPendingApplicationCorrectionItems_(record); }
  catch (error) {
    // The editor retains its existing administrator-review error for unsafe requests.
    return [];
  }
  const registry = getApplicationCorrectionRegistry_();
  return items.map(function(item) {
    return { key: item.targetCode, label: registry[item.targetCode].label,
      reason: getApplicantApplicationCorrectionReason_(item) };
  });
}

function getApplicantApplicationCorrections(applicationId, secureToken) {
  const record = requireApplicantCorrectionRecord_(applicationId, secureToken).record;
  const items = getPendingApplicationCorrectionItems_(record);
  const registry = getApplicationCorrectionRegistry_();
  return { ok: true, revision: getApplicationCorrectionRevision_(record), fields: items.map(function(item) {
    const field = registry[item.targetCode];
    return { key: field.key, label: field.label, type: field.type, required: field.required,
      choices: getApplicationCorrectionChoices_(field), value: getApplicationCorrectionValue_(record, field.key),
      reason: getApplicantApplicationCorrectionReason_(item) };
  }) };
}

function getApplicationCorrectionRevision_(record) {
  // Include legacy notes so an edited request invalidates an already-open editor.
  const source = JSON.stringify([record.rowNumber,
    getAdminRecordValue_(record, 'Application Correction Fields JSON'),
    getAdminRecordValue_(record, 'Application Information Verified At'),
    getAdminRecordValue_(record, 'Application Information Review Notes')]);
  return Utilities.base64EncodeWebSafe(Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, source));
}

function validateApplicationCorrectionValue_(field, value) {
  if (field.type === 'multiselect') {
    if (!Array.isArray(value) || !value.length || value.some(function(choice) { return typeof choice !== 'string' || getApplicationCorrectionChoices_(field).indexOf(choice) === -1; })) {
      throw new Error('Select valid values for ' + field.label + '.');
    }
    return Array.from(new Set(value)).join(', ');
  }
  if (typeof value !== 'string') { throw new Error('Enter a valid value for ' + field.label + '.'); }
  const clean = value.trim();
  if (field.required && !clean) { throw new Error(field.label + ' is required.'); }
  if (clean.length > 10000) { throw new Error(field.label + ' is too long.'); }
  if (clean && field.type === 'email' && !/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(clean)) { throw new Error('Enter a valid email address.'); }
  if (clean && field.type === 'select' && getApplicationCorrectionChoices_(field).indexOf(clean) === -1) { throw new Error('Select a valid ' + field.label + '.'); }
  if (clean && field.type === 'date') {
    const date = new Date(clean + 'T00:00:00Z');
    if (!/^\d{4}-\d{2}-\d{2}$/.test(clean) || isNaN(date.getTime()) || date.toISOString().slice(0, 10) !== clean) { throw new Error('Enter a valid date for ' + field.label + '.'); }
  }
  return clean;
}

function submitApplicantApplicationCorrections(payload) {
  const input = payload || {};
  // Coordinate with both Admin reviews (script lock) and existing applicant writes (document lock).
  const scriptLock = LockService.getScriptLock();
  scriptLock.waitLock(30000);
  let documentLock;
  try {
    documentLock = LockService.getDocumentLock();
    if (documentLock) { documentLock.waitLock(30000); }
    const context = requireApplicantCorrectionRecord_(input.applicationId, input.secureToken);
    const record = context.record;
    const items = getPendingApplicationCorrectionItems_(record);
    if (!input.revision || input.revision !== getApplicationCorrectionRevision_(record)) { throw new Error('This correction request has changed. Reload the correction form.'); }
    const values = input.values;
    if (!values || typeof values !== 'object' || Array.isArray(values)) { throw new Error('Correction values are required.'); }
    const keys = items.map(function(item) { return item.targetCode; });
    if (Object.keys(values).length !== keys.length || Object.keys(values).some(function(key) { return keys.indexOf(key) === -1; })) { throw new Error('Submit all and only the requested correction fields.'); }
    const registry = getApplicationCorrectionRegistry_();
    const validated = {};
    keys.forEach(function(key) { validated[key] = validateApplicationCorrectionValue_(registry[key], values[key]); });
    // Preserve the Form's optional State/Other rules. Validate dependencies only
    // when a requested dependent value is supplied; do not force unrelated edits.
    const nationality = validated.nationality === undefined ? getApplicationCorrectionValue_(record, 'nationality') : validated.nationality;
    if (validated.state_of_origin && nationality !== 'Nigeria') { throw new Error('State of Origin is only applicable to Nigerian applicants.'); }
    const areas = validated.area_of_practice === undefined ? getApplicationCorrectionValue_(record, 'area_of_practice') : validated.area_of_practice;
    if (validated.other_area_of_practice && areas.split(',').map(function(area) { return area.trim(); }).indexOf('Other') === -1) { throw new Error('Other Area of Practice applies only when Other is selected.'); }
    const submittedAt = new Date().toISOString();
    keys.forEach(function(key) {
      if (key === 'other') { return; } // Generic explanation is retained in the existing JSON, not an invented Sheet field.
      const header = getApplicationCorrectionHeader_(record, key);
      const value = validated[key];
      // Prevent submitted text being interpreted as a Sheet formula.
      const stored = registry[key].type === 'date' && value
        ? Utilities.parseDate(value, Session.getScriptTimeZone(), 'yyyy-MM-dd')
        : value.charAt(0) === '=' ? "'" + value : value;
      setAdminRecordMemory_(record, header, stored);
    });
    const completed = items.map(function(item) {
      return Object.assign({}, item, { targetLabel: registry[item.targetCode].label, status: 'submitted',
        submittedAt: submittedAt, submittedValue: validated[item.targetCode] });
    });
    const completedJson = JSON.stringify(completed);
    if (completedJson.length > 45000) { throw new Error('The correction responses are too long. Please shorten them.'); }
    setAdminRecordMemory_(record, 'Application Correction Fields JSON', completedJson);
    if (Object.prototype.hasOwnProperty.call(validated, 'other')) {
      setAdminRecordMemory_(record, 'Application Information Review Notes',
        String(getAdminRecordValue_(record, 'Application Information Review Notes') || '') +
        '\nApplicant response (Other Application Information): ' + validated.other);
    }
    setAdminRecordMemory_(record, 'Application Information Status', 'Pending');
    setAdminRecordMemory_(record, 'Application Information Verified At', '');
    setAdminRecordMemory_(record, 'Application Information Verified By', '');
    setAdminRecordMemory_(record, 'Verification Status', 'Pending');
    setAdminRecordMemory_(record, 'Verification Notes', '');
    setAdminRecordMemory_(record, 'Verification Completed At', '');
    setAdminRecordMemory_(record, 'Verification Completed By', '');
    setAdminRecordMemory_(record, 'Record Status', 'Resubmission');
    // One authoritative write includes corrected fields and their completion state.
    commitAdminRecord_(context.sheet, record);
    SpreadsheetApp.flush();
    try {
      invalidateAdminCachesAfterWrite_(String(input.applicationId).trim());
    } catch (cacheError) {
      // The authoritative commit succeeded. A cache failure must not invite a
      // duplicate submission or report that the saved corrections were lost.
      console.error('Correction saved; cache invalidation failed:', cacheError.message);
    }
    return { ok: true, message: 'Your corrections have been saved and submitted for Admin review.' };
  } finally {
    if (documentLock && documentLock.hasLock()) { documentLock.releaseLock(); }
    scriptLock.releaseLock();
  }
}

// Old public-Form correction links must not supersede an authenticated request.
function isSecureApplicationCorrectionRequired_(row) {
  if (String(row['Application Information Status'] || '').trim().toLowerCase() === 'correction required') { return true; }
  return parseApplicationCorrectionFields_(row['Application Correction Fields JSON']).some(function(item) {
    return item && (item.status === 'pending' || item.status === 'submitted');
  });
}
