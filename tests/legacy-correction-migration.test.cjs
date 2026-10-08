// Isolated migration tests: all Sheets, locks, logs and services are in memory.
const { test } = require('node:test');
const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const vm = require('node:vm');
const root = path.resolve(__dirname, '..');
const source = fs.readdirSync(root).filter(file => file.endsWith('.js')).map(file => fs.readFileSync(path.join(root, file), 'utf8')).join('\n');
function fixture(overrides = [{}], options = {}) {
  let sequence = 0;
  const logs = [], writes = [], invalidated = [], held = { script: false, document: false };
  const ctx = vm.createContext({ Date, console: { log: text => logs.push(JSON.parse(text)), warn() {}, error() {} },
    Utilities: { getUuid: () => 'migration-request-' + ++sequence },
    SpreadsheetApp: { getActiveSpreadsheet: () => ({}) },
    FormApp: new Proxy({}, { get() { assert.fail('migration must not use Google Forms'); } }),
    DriveApp: new Proxy({}, { get() { assert.fail('migration must not use Drive'); } }),
    LockService: Object.fromEntries(['script', 'document'].map(kind => ['get' + kind[0].toUpperCase() + kind.slice(1) + 'Lock', () => options.noDocumentLock && kind === 'document' ? null : ({
      waitLock() { assert.equal(held[kind], false); held[kind] = true; }, hasLock: () => held[kind], releaseLock() { held[kind] = false; }
    })]))
  });
  vm.runInContext(source, ctx);
  const base = Object.fromEntries(Object.values(ctx.getApplicationCorrectionRegistry_()).filter(field => field.key !== 'other').map(field => [field.label, 'Existing value']));
  Object.assign(base, { 'Application ID': 'SYNTH-1001', 'Secure Token': 'private-token', 'Full Name': 'Synthetic Applicant',
    'Submission Count': 2, 'Record Status': 'Resubmission', 'Application Information Status': 'Correction Required',
    'Application Correction Fields JSON': '', 'Application Information Review Notes': 'Company RC Number requires correction.',
    'Payment Status': 'Confirmed', 'Payment Reference': 'PAY-KEEP', 'Document Status': 'Complete', 'CAC Document File ID': 'FILE-KEEP',
    'CAC Document URL': 'https://example.test/cac', 'Licence Status': 'Not Generated', 'Licence Number': 'LIC-KEEP',
    'Verification Status': 'Pending', 'Verification Notes': 'KEEP', 'Company Name': '=UPPER("keep")' });
  const headers = Object.keys(base).filter(header => !(options.omitHeaders || []).includes(header));
  const rows = overrides.map(override => headers.map(header => ({ ...base, ...override })[header]));
  const snapshot = () => JSON.parse(JSON.stringify(rows));
  const sheet = { getLastRow: () => rows.length + 1, getLastColumn: () => headers.length,
    getRange(r, c, n = 1, w = 1) {
      const slice = () => Array.from({ length: n }, (_, i) => (r + i === 1 ? headers : rows[r + i - 2]).slice(c - 1, c - 1 + w));
      return { getValues: slice,
        getDisplayValues: () => slice().map(row => row.map(value => value == null ? '' : String(value))),
        getFormulas: () => slice().map(row => row.map(value => typeof value === 'string' && value.startsWith('=') ? value : '')),
        setValues(values) {
          assert.ok(held.script && (options.noDocumentLock || held.document), 'writes require available locks');
          assert.ok(r >= 2 && r <= rows.length + 1, 'no appends');
          assert.equal(n, 1); assert.equal(w, 1);
          assert.equal(headers[c - 1], 'Application Correction Fields JSON', 'only JSON cell can be written');
          assert.equal(logs.at(-1).phase, 'BEFORE_WRITE', 'backup logged before cell write');
          writes.push({ rowNumber: r, header: headers[c - 1], value: values[0][0] });
          rows[r - 2][c - 1] = values[0][0];
        }
      };
    }
  };
  ctx.getResponseSheet_ = () => sheet;
  ctx.invalidateApplicantCacheByApplicationId_ = id => invalidated.push(id);
  ctx.onFormSubmit = () => assert.fail('Form submission is forbidden');
  ctx.commitAdminRecord_ = () => assert.fail('full-row writes are forbidden');
  const run = opts => ctx.migrateLegacyApplicationCorrections_(opts);
  const set = (row, header, value) => { rows[row - 2][headers.indexOf(header)] = value; };
  const append = override => rows.push(headers.map(header => ({ ...base, ...override })[header]));
  return { ctx, headers, rows, logs, writes, invalidated, held, snapshot, run, set, append };
}
const first = s => s.run().entries[0];
const preview = entry => JSON.parse(entry.generatedJsonPreview);

for (const [note, key] of [
  ['Company RC Number requires correction.', 'company_rc_number'],
  ['RC Number requires correction.', 'company_rc_number'],
  ['CAC Registration Number requires correction.', 'company_rc_number'],
  ['CRFFN Membership Number is missing.', 'crffn_membership_number'],
  ['CRFFN Corporate Membership Number is missing.', 'crffn_membership_number'],
  ['State of Origin requires correction.', 'state_of_origin'],
  ['State of Origin — Nigerian Applicants Only requires correction.', 'state_of_origin'],
  ['Company TIN appears invalid.', 'company_tin'],
  ['Company Tax Identification Number (TIN) appears invalid.', 'company_tin'],
  ['Tax Identification Number appears invalid.', 'company_tin'],
  ['TIN appears invalid.', 'company_tin']
]) test('migration reuses registry alias: ' + note, () => {
  const s = fixture([{ 'Application Information Review Notes': note }]);
  const entry = first(s);
  assert.equal(entry.migrationStatus, 'SAFE_TO_MIGRATE');
  assert.deepEqual(Array.from(entry.resolvedTargets), [key]);
  const items = preview(entry); assert.equal(items.length, 1);
  assert.equal(items[0].targetLabel, s.ctx.getApplicationCorrectionRegistry_()[key].label);
  assert.equal(items[0].status, 'pending'); assert.equal(s.writes.length, 0);
});
test('multiple legacy lines produce one array and one shared requestId', () => {
  const s = fixture([{ 'Application Information Review Notes': 'Company RC Number requires correction.\nCRFFN Membership Number is missing.\nCompany TIN appears invalid.' }]);
  const items = preview(first(s));
  assert.deepEqual(items.map(item => item.targetCode), ['company_rc_number', 'crffn_membership_number', 'company_tin']);
  assert.equal(new Set(items.map(item => item.requestId)).size, 1);
  assert.deepEqual(items.map(item => item.reason), ['requires correction', 'is missing', 'appears invalid']);
});
test('custom Other instruction is a reason, never a submitted correction value', () => {
  for (const phrase of ['has another issue', 'HAS ANOTHER ISSUE']) {
    const reason = 'Your valid corporate membership is RFFC-349547.';
    const s = fixture([{ 'Application Information Review Notes': 'CRFFN Membership Number ' + phrase + ': ' + reason }]);
    const before = s.snapshot(), item = preview(first(s))[0];
    assert.equal(item.reason, reason); assert.equal(item.customDetails, reason);
    assert.equal(s.ctx.getApplicantApplicationCorrectionReason_(item), reason);
    assert.equal(Object.hasOwn(item, 'submittedValue'), false);
    assert.equal(Object.hasOwn(item, 'submittedAt'), false);
    assert.deepEqual(s.snapshot(), before);
  }
});
test('duplicated legacy another-issue prefixes are removed only from generated reason and customDetails', () => {
  const note = 'CRFFN Membership Number has another issue: CRFFN Membership Number has another issue: Kindly update your profile and enter the corporate number.';
  const s = fixture([{ 'Application Information Review Notes': note }]);
  const before = s.snapshot();
  const item = preview(first(s))[0];
  assert.equal(item.reason, 'Kindly update your profile and enter the corporate number.');
  assert.equal(item.customDetails, item.reason);
  assert.equal(item.targetCode, 'crffn_membership_number');
  assert.equal(s.snapshot()[0][s.headers.indexOf('Application Information Review Notes')], note);
  assert.deepEqual(s.snapshot(), before);
});
for (const note of ['Please correct your company information.', 'Other Application Information has another issue: Explain.',
  'Company Name and Company Address requires correction.', 'Company RC Number requires correction.\nCompany RC Number appears invalid.',
  'Company RC Number requires correction.\nUnknown field is missing.']) test('unsafe legacy note skipped: ' + note, () => {
  const s = fixture([{ 'Application Information Review Notes': note }]); const before = s.snapshot();
  assert.equal(first(s).migrationStatus, 'SKIPPED_UNSAFE_CORRECTION');
  assert.deepEqual(s.snapshot(), before); assert.equal(s.writes.length, 0);
});
test('blank note is separately reported', () => {
  assert.equal(first(fixture([{ 'Application Information Review Notes': ' \n ' }])).migrationStatus, 'SKIPPED_NO_CORRECTION_NOTE');
});
for (const header of ['Application ID', 'Secure Token']) test('missing credential skipped: ' + header, () => {
  const s = fixture([{ [header]: '' }]);
  assert.equal(first(s).migrationStatus, 'SKIPPED_MISSING_CREDENTIALS'); assert.equal(s.writes.length, 0);
});
test('same-credential historical duplicate is untouched; only latest physical row is eligible', () => {
  const s = fixture([{}, {}]); const before = s.snapshot();
  const report = s.run({ dryRun: false });
  assert.deepEqual(Array.from(report.entries, entry => entry.migrationStatus), ['SKIPPED_NOT_AUTHORITATIVE', 'MIGRATED']);
  assert.equal(s.writes.length, 1); assert.equal(s.writes[0].rowNumber, 3); assert.deepEqual(s.rows[0], before[0]);
});
test('same-ID multi-token histories are all skipped even when each pair has an authoritative row', () => {
  const s = fixture([{}, { 'Secure Token': 'different-token' }]); const before = s.snapshot();
  assert.ok(s.run({ dryRun: false }).entries.every(entry => entry.migrationStatus === 'SKIPPED_NOT_AUTHORITATIVE'));
  assert.deepEqual(s.snapshot(), before); assert.equal(s.writes.length, 0);
});
test('a non-correction or missing-token duplicate still makes multi-token authority unsafe', () => {
  for (const token of ['different-token', '']) {
    const s = fixture([{}, { 'Secure Token': token, 'Application Information Status': 'Confirmed' }]);
    assert.equal(first(s).migrationStatus, 'SKIPPED_NOT_AUTHORITATIVE');
  }
});
for (const status of ['pending', 'submitted', 'resolved']) test('APP-0005 existing structured ' + status + ' is never overwritten', () => {
  const json = JSON.stringify([{ targetCode: 'company_rc_number', status, reason: 'Manually tested', requestId: 'legacy-APP-0005' }]);
  const s = fixture([{ 'Application ID': 'APP-0005', 'Application Correction Fields JSON': json }]); const before = s.snapshot();
  assert.equal(s.run({ dryRun: false }).entries[0].migrationStatus, 'SKIPPED_ALREADY_STRUCTURED');
  assert.deepEqual(s.snapshot(), before); assert.equal(s.writes.length, 0);
});
for (const json of ['{bad', '[]', 'null', '{}', 'false', '0', '[{"targetCode":"unknown"}]', '=IF(TRUE,"","")']) test('nonblank malformed JSON/formula never overwritten: ' + json, () => {
  const s = fixture([{ 'Application Correction Fields JSON': json }]); const before = s.snapshot();
  const entry = s.run({ dryRun: false }).entries[0];
  assert.equal(entry.migrationStatus, 'SKIPPED_MALFORMED_JSON'); assert.equal(entry.detail, 'MALFORMED_EXISTING_JSON');
  assert.deepEqual(s.snapshot(), before); assert.equal(s.writes.length, 0);
});
test('missing Expiry Date column produces no-existing-field without adding headers', () => {
  const s = fixture([{ 'Application Information Review Notes': 'Expiry Date requires correction.' }], { omitHeaders: ['Expiry Date'] });
  const before = s.snapshot(); assert.equal(first(s).migrationStatus, 'SKIPPED_NO_EXISTING_FIELD');
  assert.deepEqual(s.snapshot(), before); assert.equal(s.headers.includes('Expiry Date'), false);
});
test('missing required storage header fails closed without creating a column', () => {
  const s = fixture([{}], { omitHeaders: ['Application Correction Fields JSON'] }); const before = s.snapshot();
  assert.throws(() => s.run(), /exactly one existing column/); assert.deepEqual(s.snapshot(), before);
});
test('default, explicit dry-run, and nonboolean false modes perform zero writes or cache mutations', () => {
  for (const options of [undefined, {}, { dryRun: true }, { dryRun: 'false' }, { dryRun: 0 }]) {
    const s = fixture([{ 'Application Correction Fields JSON': ' \n ' }]); const before = s.snapshot();
    const report = s.run(options); assert.equal(report.dryRun, true);
    assert.equal(report.entries[0].migrationStatus, 'SAFE_TO_MIGRATE');
    assert.deepEqual(s.snapshot(), before); assert.deepEqual(s.writes, []); assert.deepEqual(s.invalidated, []);
    assert.equal(JSON.stringify(report).includes('private-token'), false);
    assert.equal(JSON.stringify(s.logs).includes('private-token'), false);
  }
});
test('mocked explicit write changes only JSON, retains backup, and is idempotent', () => {
  const s = fixture([{ 'Application Correction Fields JSON': '  ' }]); const before = s.snapshot();
  const report = s.run({ dryRun: false }), entry = report.entries[0];
  assert.equal(entry.migrationStatus, 'MIGRATED'); assert.equal(entry.backup.originalJson, '  ');
  assert.equal(entry.backup.newJson, s.writes[0].value); assert.ok(entry.backup.timestamp);
  assert.equal(s.rows.length, before.length); assert.equal(s.writes.length, 1);
  s.headers.forEach((header, column) => { if (header !== 'Application Correction Fields JSON') assert.deepEqual(s.rows[0][column], before[0][column], header); });
  const migrated = s.snapshot(); assert.equal(s.run({ dryRun: false }).entries[0].migrationStatus, 'SKIPPED_ALREADY_STRUCTURED');
  assert.deepEqual(s.snapshot(), migrated); assert.equal(s.writes.length, 1);
  assert.deepEqual(s.invalidated, ['SYNTH-1001']); assert.deepEqual(s.held, { script: false, document: false });
});
test('separate applications receive distinct request IDs', () => {
  const s = fixture([{}, { 'Application ID': 'SYNTH-1002' }]);
  const report = s.run(); assert.equal(report.entries.length, 2);
  assert.notEqual(preview(report.entries[0])[0].requestId, preview(report.entries[1])[0].requestId);
});
test('non-correction rows never become candidates and newer non-correction row blocks older candidate', () => {
  const s = fixture([{}, { 'Application Information Status': 'Pending' }]);
  const report = s.run(); assert.equal(report.entries.length, 1); assert.equal(report.entries[0].migrationStatus, 'SKIPPED_NOT_AUTHORITATIVE');
});
for (const change of ['json', 'status', 'notes', 'row', 'token']) test('write mode rechecks current state: ' + change, () => {
  const s = fixture();
  assert.equal(first(s).migrationStatus, 'SAFE_TO_MIGRATE');
  const selector = s.ctx.findApplicantApplicationRecordByCredentials_;
  s.ctx.findApplicantApplicationRecordByCredentials_ = (...args) => {
    if (change === 'json') s.set(2, 'Application Correction Fields JSON', '[{"targetCode":"company_tin","status":"submitted"}]');
    if (change === 'status') s.set(2, 'Application Information Status', 'Pending');
    if (change === 'notes') s.set(2, 'Application Information Review Notes', 'Unidentifiable new instruction.');
    if (change === 'row') s.append({});
    if (change === 'token') s.set(2, 'Secure Token', 'replaced-token');
    return selector(...args);
  };
  const report = s.run({ dryRun: false });
  assert.equal(s.writes.length, 0); assert.notEqual(report.entries[0].migrationStatus, 'MIGRATED');
  assert.deepEqual(s.held, { script: false, document: false });
});
test('write mode uses freshly parsed notes rather than the previous dry-run preview', () => {
  const s = fixture(); first(s);
  s.set(2, 'Application Information Review Notes', 'Company TIN appears invalid.');
  const report = s.run({ dryRun: false });
  assert.equal(JSON.parse(s.writes[0].value)[0].targetCode, 'company_tin');
  assert.equal(report.entries[0].correctionNote, 'Company TIN appears invalid.');
});
test('web-app context without a document lock remains protected by script lock', () => {
  const s = fixture([{}], { noDocumentLock: true });
  assert.equal(s.run({ dryRun: false }).entries[0].migrationStatus, 'MIGRATED');
  assert.equal(s.writes.length, 1); assert.equal(s.held.script, false);
});
test('example report includes APP-0005 skip and two synthetic legacy previews', () => {
  const s = fixture([
    { 'Application ID': 'APP-0005', 'Application Correction Fields JSON': '[{"targetCode":"company_rc_number","status":"pending"}]' },
    { 'Application ID': 'SYNTH-1001' },
    { 'Application ID': 'SYNTH-1002', 'Application Information Review Notes': 'CRFFN Membership Number has another issue: Please provide your corporate membership number.' }
  ]);
  const report = s.run();
  assert.deepEqual(Array.from(report.entries, entry => entry.migrationStatus), ['SKIPPED_ALREADY_STRUCTURED', 'SAFE_TO_MIGRATE', 'SAFE_TO_MIGRATE']);
  for (const entry of report.entries) for (const key of ['applicationId', 'rowNumber', 'applicantName', 'submissionCount', 'recordStatus', 'informationStatus', 'jsonBlank', 'correctionNote', 'resolvedTargets', 'migrationStatus']) assert.ok(Object.hasOwn(entry, key), key);
  assert.equal(s.writes.length, 0);
});
