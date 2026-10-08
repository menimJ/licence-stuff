// Isolated services only; never connects to production. Run: node --test tests/*.test.cjs
const { test } = require('node:test');
const assert = require('node:assert/strict');
const fs = require('node:fs');
const vm = require('node:vm');
const crypto = require('node:crypto');
const path = require('node:path');
const root = path.resolve(__dirname, '..');
const read = file => fs.readFileSync(path.join(root, file), 'utf8');
const source = fs.readdirSync(root).filter(f => f.endsWith('.js')).map(read).join('\n');
function fixture(options = {}) {
  let writes = 0, failWrite = false, sequence = 0;
  const locks = { script: false, document: false };
  const ctx = vm.createContext({ Date, console: { log() {}, error() {}, warn() {} },
    Utilities: {
      getUuid: () => 'request-' + ++sequence,
      DigestAlgorithm: { SHA_256: 'sha256' },
      computeDigest: (_, text) => crypto.createHash('sha256').update(text).digest(),
      base64EncodeWebSafe: bytes => Buffer.from(bytes).toString('base64url'),
      formatDate: date => date.toISOString().slice(0, 10),
      parseDate: text => new Date(text + 'T00:00:00Z')
    }, Session: { getScriptTimeZone: () => 'Africa/Lagos' },
    SpreadsheetApp: { getActiveSpreadsheet: () => ({}), flush() {} },
    LockService: Object.fromEntries(['script', 'document'].map(kind => ['get' + kind[0].toUpperCase() + kind.slice(1) + 'Lock', () => ({
      waitLock() { assert.equal(locks[kind], false); locks[kind] = true; },
      tryLock() { assert.equal(locks[kind], false); locks[kind] = true; return true; },
      hasLock() { return locks[kind]; }, releaseLock() { locks[kind] = false; }
    })]))
  });
  vm.runInContext(source, ctx);
  const registry = ctx.getApplicationCorrectionRegistry_();
  const initial = {};
  Object.values(registry).filter(f => f.key !== 'other').forEach(f => { initial[f.label] = 'Old value'; });
  Object.assign(initial, {
    'Application ID': 'APP-1', 'Secure Token': 'valid-token', 'Full Name': 'Applicant', 'Email Address': 'old@example.test',
    'Nationality': 'Nigeria', 'Area of Practice': 'Other',
    'Application Correction Fields JSON': '', 'Application Information Status': 'Confirmed',
    'Application Information Review Notes': '', 'Application Information Verified At': '', 'Application Information Verified By': '',
    'Verification Status': 'Pending', 'Verification Notes': '', 'Verification Completed At': '', 'Verification Completed By': '',
    'Payment Status': 'Confirmed', 'Payment Reference': 'payment-original', 'Receipt PDF URL': 'receipt',
    'Document Status': 'Complete', 'CAC Document File ID': 'cac-original', 'CAC Document Review Status': 'Approved',
    'Stamped Licence File ID': '', 'Licence Status': 'Not Generated', 'Licence Number': '',
    'Portal URL': 'https://example.test/portal', 'Record Status': 'Approved',
    ...options.values
  });
  if (options.legacy) { initial['CRFFN Membership Number'] = initial['CRFFN Corporate Membership Number']; delete initial['CRFFN Corporate Membership Number']; }
  if (options.noExpiry) delete initial['Expiry Date'];
  if (options.aliasSwap) { const [key, alias] = options.aliasSwap; initial[alias] = initial[registry[key].label]; delete initial[registry[key].label]; }
  const headers = Object.keys(initial);
  let row = Object.values(initial);
  const sheet = {
    getLastRow: () => 2, getLastColumn: () => headers.length,
    getRange(r, col, count, width) {
      const slice = () => r === 1 ? headers.slice(col - 1, col - 1 + width) : row.slice(col - 1, col - 1 + width);
      return {
        getValues: () => [slice()], getDisplayValues: () => [slice().map(v => v instanceof Date ? v.toISOString() : String(v))],
        getFormulas: () => [slice().map(() => '')],
        setValues(values) { if (failWrite) throw Error('Mock Sheet write failed'); row.splice(col - 1, width, ...values[0]); writes++; }
      };
    }
  };
  ctx.getResponseSheet_ = () => sheet;
  ctx.requireAdminAccess_ = () => ({ email: 'admin@example.test', role: 'Admin' });
  ctx.invalidateAdminCachesAfterWrite_ = () => {};
  ctx.makeAdminBackgroundJobKey_ = () => 'job'; ctx.queueSystemJob_ = () => ({ ok: true });
  const snapshot = () => Object.fromEntries(headers.map((h, i) => [h, row[i]]));
  const request = (keys, details = '') => ctx.rejectAdminReviewStage('APP-1', 'information', keys.map(targetCode => ({ targetCode, issueCode: details ? 'other' : 'incorrect' })), details, 'admin-session');
  const load = () => ctx.getApplicantApplicationCorrections('APP-1', 'valid-token');
  const submit = (values, overrides = {}) => ctx.submitApplicantApplicationCorrections({ applicationId: 'APP-1', secureToken: 'valid-token', revision: load().revision, values, ...overrides });
  return { ctx, registry, request, load, submit, snapshot, locks, writes: () => writes, failWrites: () => { failWrite = true; } };
}
// Reported production case FIRST, using the actual Admin rejection and row-update path.
test('reported CRFFN case: structured request, corporate display, RFFC-349547 persisted and identity/workflows retained', () => {
  const s = fixture();
  s.request(['crffn_membership_number'], 'Your valid corporate membership is RFFC-349547.');
  const before = s.snapshot();
  const request = JSON.parse(before['Application Correction Fields JSON']);
  assert.equal(request[0].targetCode, 'crffn_membership_number');
  assert.equal(request[0].targetLabel, 'CRFFN Corporate Membership Number');
  assert.equal(request[0].reason, 'Your valid corporate membership is RFFC-349547.');
  assert.equal(s.load().fields[0].value, 'Old value');
  assert.equal(s.submit({ crffn_membership_number: 'RFFC-349547' }).ok, true);
  const after = s.snapshot();
  assert.equal(after['CRFFN Corporate Membership Number'], 'RFFC-349547');
  assert.equal(s.ctx.getCrffnMembershipNumber_(after), 'RFFC-349547');
  assert.equal(after['Application Information Status'], 'Pending');
  assert.equal(after['Record Status'], 'Resubmission');
  assert.equal(JSON.parse(after['Application Correction Fields JSON'])[0].status, 'submitted');
  for (const key of ['Application ID', 'Secure Token', 'Payment Status', 'Payment Reference', 'Receipt PDF URL', 'Document Status', 'CAC Document File ID', 'CAC Document Review Status', 'Licence Status']) assert.equal(after[key], before[key], key);
});
const valid = {
  full_name: 'New Name', gender: 'Female', date_of_birth: '1992-02-29', nationality: 'Ghana', state_of_origin: 'Lagos',
  residential_address: 'New Residence', phone_number: '+2348012345678', email_address: 'new@example.test',
  means_of_identification: 'National ID', id_number: 'NIN-123', expiry_date: '2030-01-01',
  area_of_practice: ['Warehousing', 'Other'], other_area_of_practice: 'Marine Logistics', company_name: 'New Company',
  company_rc_number: 'RC123', company_tin: 'TIN123', crffn_membership_number: 'RFFC-349547', company_address: 'Registered Address',
  position_held: 'Director', other: 'The requested explanation'
};
for (const [key, value] of Object.entries(valid)) test(`complete single-field correction: ${key}`, () => {
  const s = fixture(); s.request([key], key === 'other' ? 'Explain the issue' : '');
  assert.equal(s.load().fields[0].key, key);
  const before = s.snapshot(); const n = s.writes();
  assert.equal(s.submit({ [key]: value }).ok, true); assert.equal(s.writes(), n + 1);
  const after = s.snapshot(); const header = s.registry[key].label;
  const persisted = key === 'other' ? JSON.parse(after['Application Correction Fields JSON'])[0].submittedValue : after[header];
  assert.equal(persisted instanceof Date ? persisted.toISOString().slice(0, 10) : persisted, Array.isArray(value) ? value.join(', ') : value);
  for (const other of Object.keys(valid).filter(k => k !== key && k !== 'other')) assert.deepEqual(after[s.registry[other].label], before[s.registry[other].label], other);
});
test('multi-field correction is complete and committed once', () => {
  const s = fixture(); s.request(['crffn_membership_number', 'company_address']);
  const before = s.snapshot();
  assert.throws(() => s.submit({ crffn_membership_number: 'RFFC-349547' }), /all and only/);
  assert.deepEqual(s.snapshot(), before);
  s.submit({ crffn_membership_number: 'RFFC-349547', company_address: 'New Address' });
  assert.equal(JSON.parse(s.snapshot()['Application Correction Fields JSON']).filter(i => i.status === 'submitted').length, 2);
});
test('every required field rejects empty submitted value without completing request', () => {
  for (const key of Object.keys(valid)) {
    const s = fixture(); if (!s.registry[key].required) continue;
    s.request([key], key === 'other' ? 'Explain' : ''); const before = s.snapshot();
    assert.throws(() => s.submit({ [key]: key === 'area_of_practice' ? [] : '' }), /required|Select valid/);
    assert.deepEqual(s.snapshot(), before, key);
  }
});
test('invalid date/email/choice/checkbox values are rejected', () => {
  for (const [key, value] of [['date_of_birth', '2026-02-30'], ['email_address', 'not email'], ['gender', 'Invalid'], ['area_of_practice', ['Unknown']]]) {
    const s = fixture(); s.request([key]); assert.throws(() => s.submit({ [key]: value }), /valid/);
  }
});
test('conditional dependent values use submitted and existing context; optional fields stay optional', () => {
  const s = fixture(); s.request(['nationality', 'state_of_origin']);
  assert.throws(() => s.submit({ nationality: 'Ghana', state_of_origin: 'Lagos' }), /Nigerian/);
  s.submit({ nationality: 'Ghana', state_of_origin: '' });
  const t = fixture(); t.request(['area_of_practice', 'other_area_of_practice']);
  assert.throws(() => t.submit({ area_of_practice: ['Warehousing'], other_area_of_practice: 'Marine' }), /Other Area/);
  t.submit({ area_of_practice: ['Other'], other_area_of_practice: '' });
});
test('both membership labels resolve to same key; legacy records persist in existing legacy header', () => {
  for (const legacy of [false, true]) {
    const s = fixture({ legacy });
    for (const label of ['CRFFN Membership Number', 'CRFFN Corporate Membership Number']) assert.equal(s.ctx.resolveApplicationCorrectionKey_(label), 'crffn_membership_number');
    s.request(['crffn_membership_number']);
    assert.equal(s.load().fields[0].label, 'CRFFN Corporate Membership Number');
    s.submit({ crffn_membership_number: 'RFFC-349547' });
    assert.equal(s.snapshot()[legacy ? 'CRFFN Membership Number' : 'CRFFN Corporate Membership Number'], 'RFFC-349547');
  }
});
test('state aliases resolve; absent legacy Expiry Date fails before request write', () => {
  const s = fixture({ noExpiry: true });
  assert.equal(s.ctx.resolveApplicationCorrectionKey_('State of Origin'), 'state_of_origin');
  assert.equal(s.ctx.resolveApplicationCorrectionKey_('State of Origin — Nigerian Applicants Only'), 'state_of_origin');
  assert.throws(() => s.request(['expiry_date']), /no existing column/); assert.equal(s.writes(), 0);
});
test('email is data only: correction never reads or writes email-to-application properties', () => {
  const s = fixture(); s.ctx.PropertiesService = new Proxy({}, { get() { throw Error('Email identity lookup is forbidden'); } });
  s.request(['email_address']); s.submit({ email_address: 'belongs-to-another-application@example.test' });
  assert.equal(s.snapshot()['Application ID'], 'APP-1'); assert.equal(s.snapshot()['Secure Token'], 'valid-token');
});
test('invalid/revoked token, missing token and mismatched Application ID rejected', () => {
  const s = fixture(); s.request(['full_name']); const before = s.snapshot();
  for (const credentials of [{ secureToken: 'revoked-token' }, { secureToken: '' }, { applicationId: 'APP-OTHER' }]) {
    assert.throws(() => s.submit({ full_name: 'Changed' }, credentials), /invalid|incomplete|not found/);
    assert.deepEqual(s.snapshot(), before);
  }
});
test('unknown targets, extra fields, duplicates, and prototype keys rejected', () => {
  const s = fixture();
  for (const key of ['unknown', '__proto__', 'constructor']) assert.throws(() => s.request([key]), /invalid|Unknown/);
  assert.throws(() => s.request(['full_name', 'full_name']), /duplicate/);
  s.request(['full_name']);
  assert.throws(() => s.submit({ full_name: 'Name', company_name: 'Unrequested' }), /all and only/);
});
test('stale request and duplicate submission rejected; reissued correction succeeds', () => {
  const s = fixture(); s.request(['full_name']); const old = s.load().revision;
  s.request(['company_name']);
  assert.throws(() => s.submit({ company_name: 'New' }, { revision: old }), /changed/);
  const revision = s.load().revision;
  s.submit({ company_name: 'New' });
  assert.throws(() => s.ctx.submitApplicantApplicationCorrections({ applicationId: 'APP-1', secureToken: 'valid-token', revision, values: { company_name: 'Again' } }), /no active/);
  s.request(['company_name']); s.submit({ company_name: 'Second correction' });
  assert.equal(s.snapshot()['Company Name'], 'Second correction');
});
test('Admin approval retains correction history and marks targets resolved', () => {
  const s = fixture(); s.request(['full_name']); s.submit({ full_name: 'Corrected Name' });
  s.ctx.updateAdminReviewStage_({ applicationId: 'APP-1', stage: 'information', approved: true, notes: '' }, 'session');
  const item = JSON.parse(s.snapshot()['Application Correction Fields JSON'])[0];
  assert.equal(item.status, 'resolved'); assert.equal(item.submittedValue, 'Corrected Name');
});
test('missing/corrupt structured request never falls back to free text', () => {
  for (const json of ['', '{bad', '[]', '[{"targetCode":"unknown"}]']) {
    const s = fixture({ values: { 'Application Information Status': 'Correction Required', 'Application Correction Fields JSON': json, 'Application Information Review Notes': 'CRFFN Membership Number has another issue: RFFC-349547.' } });
    assert.throws(() => s.load(), /reissued|invalid/);
  }
});
test('failed Sheet commit leaves corrected values and completion state unchanged', () => {
  const s = fixture(); s.request(['full_name']); const before = s.snapshot(); s.failWrites();
  assert.throws(() => s.submit({ full_name: 'New' }), /write failed/); assert.deepEqual(s.snapshot(), before);
  assert.equal(s.locks.script, false); assert.equal(s.locks.document, false);
});
test('licence processing blocks correction without changing stamped/payment/document fields', () => {
  const s = fixture({ values: { 'Application Information Status': 'Correction Required', 'Licence Status': 'Released', 'Stamped Licence File ID': 'final' } });
  assert.throws(() => s.load(), /licence processing/); assert.equal(s.writes(), 0);
});
test('formula-like text cannot become an executable Sheet formula', () => {
  const s = fixture(); s.request(['company_name']); s.submit({ company_name: '=IMPORTXML("url")' });
  assert.equal(s.snapshot()['Company Name'], '\'=IMPORTXML("url")');
});
test('old public Form submissions cannot replace active or submitted secure requests', () => {
  const s = fixture(); s.request(['full_name']); assert.equal(s.ctx.isSecureApplicationCorrectionRequired_(s.snapshot()), true);
  s.submit({ full_name: 'Name' }); assert.equal(s.ctx.isSecureApplicationCorrectionRequired_(s.snapshot()), true);
  assert.equal(s.ctx.isSecureApplicationCorrectionRequired_({ 'Application Information Status': 'Confirmed' }), false);
});
test('registry inventory covers all 20 requested keys without changing document keys', () => {
  const s = fixture(); assert.deepEqual(Object.keys(s.registry), Object.keys(valid));
  for (const [alias, type] of [['cac_document','cac'],['passport_photograph','passport'],['educational_certificates','education'],['proof_of_experience','experience'],['means_of_identification','identification']]) assert.equal(s.ctx.resolveSupportingDocumentConfig_(alias).type, type);
});
test('cache invalidation failure after commit still reports a saved submission', () => {
  const s = fixture(); s.request(['full_name']); s.ctx.invalidateAdminCachesAfterWrite_ = () => { throw Error('cache unavailable'); };
  assert.equal(s.submit({ full_name: 'Saved' }).ok, true); assert.equal(s.snapshot()['Full Name'], 'Saved');
});
test('web-app context without a document lock uses its script lock without crashing', () => {
  const s = fixture(); s.request(['full_name']); s.ctx.LockService.getDocumentLock = () => null;
  assert.equal(s.submit({ full_name: 'Saved' }).ok, true); assert.equal(s.locks.script, false);
});
test('actual old Form trigger quarantines unbound response instead of clearing active request', () => {
  const s = fixture(); s.request(['full_name']); const before = s.snapshot(); let submittedRow;
  s.ctx.ensureCustomColumns_ = () => {};
  s.ctx.nextSerialNumber_ = () => 3;
  s.ctx.PropertiesService = { getDocumentProperties: () => ({ getProperty: key => key.startsWith('EMAIL_TO_APP') ? 'APP-1' : key.startsWith('EMAIL_TO_TOKEN') ? 'valid-token' : '1' }) };
  s.ctx.findPreviousApplicationRow_ = () => before;
  s.ctx.writeRowValues_ = (sheet, row, headers, values) => { submittedRow = values; };
  s.ctx.onFormSubmit({ range: { getSheet: () => ({ getLastColumn: () => 2, getRange: () => ({ getValues: () => [['Full Name','Email Address']] }) }), getRow: () => 3 }, namedValues: { 'Full Name': ['New name'], 'Email Address': ['old@example.test'] } });
  assert.equal(submittedRow['Record Status'], 'Needs Review');
  assert.equal(submittedRow['Application ID'], undefined); assert.deepEqual(s.snapshot(), before);
});
test('all HTML script/template blocks parse with the shared registry injection', () => {
  for (const name of ['ApplicantPortal.html','AdminDashboardPage.html','AdminPortalPage.html','AdminDashboardScripts.html']) {
    const html = read(name);
    // Compile template control flow, leaving literal output out of the program.
    let templateCode = '';
    for (const match of html.matchAll(/<\?([\s\S]*?)\?>/g)) {
      templateCode += match[1].startsWith('!=') ? '\nvoid (' + match[1].slice(2).trim().replace(/;$/, '') + ');' : match[1].startsWith('=') ? '\nvoid (' + match[1].slice(1).trim().replace(/;$/, '') + ');' : '\n' + match[1];
    }
    new vm.Script(templateCode, { filename: name + ':template' });
    for (const match of html.matchAll(/<script[^>]*>([\s\S]*?)<\/script>/g)) {
      new vm.Script(match[1].replace(/<\?[\s\S]*?\?>/g, 'null'), { filename: name + ':script' });
    }
  }
});
test('applicant UI renders current values and submits stable keys with original credentials', () => {
  const elements = new Map();
  class Element {
    constructor(tag = 'div') { this.tag = tag; this.style = {}; this.children = []; this.options = []; this.value = ''; }
    set id(id) { this._id = id; elements.set(id, this); } get id() { return this._id; }
    append(...children) { this.children.push(...children); }
    replaceChildren() { this.children = []; }
    add(option) { this.options.push(option); if (option.selected) this.value = option.value; }
    get selectedOptions() { return this.options.filter(option => option.selected); }
  }
  for (const id of ['applicationCorrectionMessage','applicationCorrectionFields','applicationCorrectionSubmit','applicationCorrectionForm','applicationCorrectionOpen']) { const e = new Element(); e.id = id; }
  const requests = [];
  const ui = vm.createContext({ document: { getElementById: id => elements.get(id), createElement: tag => new Element(tag) },
    Option: function(label, value, defaultSelected, selected) { return { label, value, selected }; },
    getPortalCredentials: () => ({ applicationId: 'APP-1', secureToken: 'valid-token' }), reloadPractitionerPortal_() {}
  });
  ui.google = { script: { get run() { const rpc = { withSuccessHandler(fn) { this.success = fn; return this; }, withFailureHandler(fn) { this.failure = fn; return this; },
    getApplicantApplicationCorrections(...args) { this.args = args; this.type = 'load'; requests.push(this); },
    submitApplicantApplicationCorrections(payload) { this.payload = payload; this.type = 'submit'; requests.push(this); }
  }; return rpc; } } };
  const html = read('ApplicantPortal.html'); vm.runInContext(html.slice(html.indexOf('  let applicationCorrectionRequest_'), html.lastIndexOf('  </script>')), ui);
  const s = fixture(); s.request(Object.keys(valid), 'Please correct the requested fields');
  ui.loadApplicationCorrections_(); requests[0].success(s.load());
  assert.equal(elements.get('applicationCorrectionFields').children.length, 20);
  assert.equal(elements.get('application-correction-crffn_membership_number').value, 'Old value');
  elements.get('application-correction-crffn_membership_number').value = 'RFFC-349547';
  ui.submitApplicationCorrections_({ preventDefault() {} });
  assert.equal(requests[1].payload.values.crffn_membership_number, 'RFFC-349547');
  assert.equal(requests[1].payload.applicationId, 'APP-1'); assert.equal(requests[1].payload.secureToken, 'valid-token');
  assert.equal(Object.keys(requests[1].payload.values).length, 20);
  ui.submitApplicationCorrections_({ preventDefault() {} }); assert.equal(requests.length, 2);
  requests[1].failure(Error('retry')); assert.equal(elements.get('applicationCorrectionSubmit').disabled, false);
});

test('all registry aliases read and persist to existing legacy headers', () => {
  const registry = fixture().registry;
  for (const [key, field] of Object.entries(registry)) {
    if (key === 'other') continue;
    for (const alias of field.aliases.slice(1)) {
      const s = fixture({ aliasSwap: [key, alias] });
      assert.equal(s.ctx.resolveApplicationCorrectionKey_(alias), key, alias);
      s.request([key]);
      s.submit({ [key]: valid[key] });
      const persisted = s.snapshot()[alias];
      assert.equal(persisted instanceof Date ? persisted.toISOString().slice(0, 10) : persisted,
        Array.isArray(valid[key]) ? valid[key].join(', ') : valid[key], alias);
      assert.equal(Object.hasOwn(s.snapshot(), field.label), false);
    }
  }
});
