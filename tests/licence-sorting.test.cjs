// Run locally with node --test tests/licence-sorting.test.cjs. No live services.
const { test } = require('node:test');
const assert = require('node:assert/strict');
const vm = require('node:vm');
const fs = require('node:fs');
const path = require('node:path');
const root = path.resolve(__dirname, '..');
const read = name => fs.readFileSync(path.join(root, name), 'utf8');
function server(role = 'Admin') {
  const ctx = vm.createContext({ Date, console });
  for (const name of ['ApplicationCorrections.js', 'AdminAuth.js', 'AdminPortal.js', 'Portal.js']) vm.runInContext(read(name), ctx);
  const headers = ['Application ID', 'Full Name', 'Company Name', 'Licence Number', 'Licence Status', 'Timestamp', 'Licence Generated At', 'Licence Released At', 'Payment Status', 'Record Status'];
  const rows = Array.from({ length: 47 }, (_, i) => [
    `APP-${String(i + 1).padStart(3, '0')}`, i % 2 ? 'alice' : 'Zoe', i % 3 ? 'Beta' : 'alpha', `LIC-${String(47 - i).padStart(3, '0')}`,
    i % 2 ? 'Generated' : 'Released', new Date(Date.UTC(2026, 0, i + 1)), new Date(Date.UTC(2026, 0, 47 - i)), i % 2 ? '' : new Date(Date.UTC(2026, 1, i + 1)), 'Confirmed', 'Approved'
  ]);
  // Ineligible rows must never enter the list, regardless of sort.
  rows.push(['APP-999', 'Aaron', '', '', 'Generated', new Date('2027-01-01'), '', '', 'Pending', 'Approved']);
  rows.push(['APP-998', 'Aaron', '', '', 'Generated', new Date('2027-01-01'), '', '', 'Confirmed', 'Rejected']);
  let rawReads = 0;
  const sheet = { getLastRow: () => rows.length + 1, getLastColumn: () => headers.length,
    getRange(row, col, count) {
      return { getDisplayValues() {
        return row === 1 ? [headers] : rows.slice(row - 2, row - 2 + count).map(r => r.map(value => value instanceof Date ? 'DISPLAY DATE IS NOT SORTABLE' : value));
      }, getValues() { rawReads++; return rows.slice(row - 2, row - 2 + count); } };
    }
  };
  ctx.requireAdminSession_ = () => ({ role });
  ctx.SpreadsheetApp = { getActiveSpreadsheet: () => ({}) };
  ctx.getResponseSheet_ = () => sheet;
  return { ctx, rows, rawReads: () => rawReads, list: options => ctx.getAdminLicencesPage('session', options) };
}
const ids = result => Array.from(result.items, row => row.applicationId);
test('default submission date DESC precedes pagination and uses raw timestamps', () => {
  const s = server();
  assert.deepEqual(ids(s.list({ pageSize: 3 })), ['APP-047', 'APP-046', 'APP-045']);
  assert.deepEqual(ids(s.list({ pageSize: 3, page: 2 })), ['APP-044', 'APP-043', 'APP-042']);
  assert.equal(s.list({}).totalItems, 47);
  assert.equal(s.list({}).sortBy, 'submittedAt');
  assert.ok(s.rawReads() > 0);
});
for (const role of ['Admin', 'Super Admin']) test(`${role} has identical sorting permissions`, () => {
  const s = server(role);
  assert.deepEqual(ids(s.list({ sortBy: 'applicationId', sortDirection: 'asc', pageSize: 3 })), ['APP-001', 'APP-002', 'APP-003']);
  assert.deepEqual(ids(s.list({ sortBy: 'applicationId', sortDirection: 'desc', pageSize: 3 })), ['APP-047', 'APP-046', 'APP-045']);
});
test('case-insensitive names and stable Application ID ties across pages', () => {
  const s = server();
  assert.deepEqual(ids(s.list({ sortBy: 'applicantName', sortDirection: 'asc', pageSize: 3 })), ['APP-002', 'APP-004', 'APP-006']);
  assert.deepEqual(ids(s.list({ sortBy: 'applicantName', sortDirection: 'asc', pageSize: 3, page: 2 })), ['APP-008', 'APP-010', 'APP-012']);
  assert.deepEqual(ids(s.list({ sortBy: 'applicantName', sortDirection: 'desc', pageSize: 3 })), ['APP-001', 'APP-003', 'APP-005']);
});
test('all existing meaningful columns sort in both directions', () => {
  const s = server();
  for (const sortBy of ['applicationId', 'applicantName', 'companyName', 'licenceNumber', 'licenceStatus', 'licenceGeneratedAt', 'licenceReleasedAt']) {
    const asc = s.list({ sortBy, sortDirection: 'asc', pageSize: 100 });
    const desc = s.list({ sortBy, sortDirection: 'desc', pageSize: 100 });
    assert.equal(asc.totalItems, 47); assert.equal(desc.totalItems, 47);
    assert.notDeepEqual(ids(asc), ids(desc), sortBy);
  }
  assert.deepEqual(ids(s.list({ sortBy: 'licenceStatus', sortDirection: 'asc', pageSize: 2 })), ['APP-002', 'APP-004']);
  assert.deepEqual(ids(s.list({ sortBy: 'licenceNumber', sortDirection: 'asc', pageSize: 2 })), ['APP-047', 'APP-046']);
});
test('generated/released dates sort by underlying Date, with missing dates last', () => {
  const s = server();
  assert.deepEqual(ids(s.list({ sortBy: 'licenceGeneratedAt', sortDirection: 'asc', pageSize: 2 })), ['APP-047', 'APP-046']);
  for (const sortDirection of ['asc', 'desc']) {
    const result = s.list({ sortBy: 'licenceReleasedAt', sortDirection, pageSize: 100 });
    assert.equal(result.items[46].licenceReleasedAt, '');
  }
});
test('date aliases and ISO raw text are supported; ambiguous display dates are not parsed', () => {
  const s = server();
  assert.equal(s.ctx.getLicenceSortTimestamp_({ 'Submitted At': new Date('2026-10-02') }, ['Timestamp', 'Submitted At']), Date.parse('2026-10-02'));
  assert.equal(s.ctx.getLicenceSortTimestamp_({ Timestamp: '2026-10-02T10:00:00Z' }, ['Timestamp']), Date.parse('2026-10-02T10:00:00Z'));
  assert.equal(s.ctx.getLicenceSortTimestamp_({ Timestamp: '02/10/2026' }, ['Timestamp']), null);
});
test('search/status filters and eligibility are unchanged', () => {
  const s = server();
  const result = s.list({ sortBy: 'applicationId', sortDirection: 'desc', search: 'alice', status: 'generated', pageSize: 100 });
  assert.equal(result.totalItems, 23);
  assert.ok(result.items.every(r => r.applicantName === 'alice' && r.licenceStatus === 'Generated'));
  assert.equal(s.list({ search: 'APP-999' }).totalItems, 0);
  assert.equal(s.list({ search: 'APP-998' }).totalItems, 0);
});
test('unrecognized sort fields/expressions fall back to submission date DESC', () => {
  const s = server();
  for (const sortBy of ['Secure Token', '__proto__', 'Timestamp DESC', 'constructor', null]) {
    assert.deepEqual(ids(s.list({ sortBy, sortDirection: 'asc', pageSize: 2 })), ['APP-047', 'APP-046']);
  }
});
test('other dashboard consumers do not request raw values or get added date metadata', () => {
  const s = server(); const rows = s.ctx.getAdminDashboardLatestRows_('session');
  assert.equal(s.rawReads(), 0); assert.equal(rows[0].licenceSortDates, undefined);
});
function frontend() {
  const source = read('AdminDashboardScripts.html');
  const headers = ['applicationId', 'applicantName', 'companyName', 'licenceNumber', 'licenceStatus', 'licenceGeneratedAt', 'licenceReleasedAt'].map(field => ({
    field, attributes: {}, indicator: {}, getAttribute() { return field; }, setAttribute(key, value) { this.attributes[key] = value; },
    querySelector(selector) { return selector === 'button' ? { textContent: field } : this.indicator; }
  }));
  const elements = { licencesLoading: {}, licencesSortSummary: {} };
  const requests = []; const rendered = [];
  const ctx = vm.createContext({ document: { getElementById: id => elements[id], querySelectorAll: () => headers },
    ADMIN_SESSION_TOKEN: 'session', WORKFLOW_PAGE_SIZE: 20, setWorkflowLoading_() {}, getErrorMessage_: e => e.message,
    renderLicencesPage_: result => rendered.push(result)
  });
  ctx.google = { script: { get run() { const r = { withSuccessHandler(fn) { this.success = fn; return this; }, withFailureHandler(fn) { this.failure = fn; return this; }, getAdminLicencesPage(token, options) { this.options = options; requests.push(this); } }; return r; } } };
  vm.runInContext(source.slice(source.indexOf('      const licencesState ='), source.indexOf('      function applyPaymentFilters()')), ctx);
  vm.runInContext(source.slice(source.indexOf('      function sortDashboardLicences_('), source.indexOf('      function renderLicencesPage_(')), ctx);
  return { ctx, headers, requests, rendered, elements };
}
test('header clicks toggle, switching field starts ascending, and paging preserves sort', () => {
  const s = frontend();
  s.ctx.sortDashboardLicences_('applicantName');
  assert.equal(s.requests[0].options.sortDirection, 'asc');
  s.ctx.sortDashboardLicences_('applicantName');
  assert.equal(s.requests[1].options.sortDirection, 'desc');
  s.ctx.loadLicencesPage(2);
  assert.equal(s.requests[2].options.page, 2); assert.equal(s.requests[2].options.sortDirection, 'desc');
  s.ctx.sortDashboardLicences_('licenceStatus');
  assert.equal(s.requests[3].options.sortDirection, 'asc'); assert.equal(s.requests[3].options.page, 1);
  assert.equal(s.headers.filter(h => h.attributes['aria-sort'] !== 'none').length, 1);
  assert.equal(s.headers.find(h => h.field === 'licenceStatus').indicator.textContent, '↑');
});
test('older responses cannot overwrite newer sort selection', () => {
  const s = frontend(); s.ctx.loadLicencesPage(); s.ctx.sortDashboardLicences_('applicationId');
  s.requests[1].success({ marker: 'new' }); s.requests[0].success({ marker: 'old' });
  assert.deepEqual(s.rendered, [{ marker: 'new' }]);
});
test('source syntax and existing dependency smoke test', () => {
  const ctx = vm.createContext({ console: { log() {} } });
  vm.runInContext(fs.readdirSync(root).filter(f => f.endsWith('.js')).map(read).join('\n'), ctx);
  assert.equal(vm.runInContext('testCRFFNPerformanceDependencies().ok', ctx), true);
  new vm.Script(read('AdminDashboardScripts.html').replace(/^\s*<script>/, '').replace(/<\/script>\s*$/, ''));
});
