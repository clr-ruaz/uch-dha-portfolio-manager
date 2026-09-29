const assert = require('node:assert/strict');
const { test } = require('node:test');
const fs = require('node:fs');
const ts = require('typescript');

// Run focused source tests without adding a test framework or generated files.
for (const extension of ['.ts', '.tsx']) {
  require.extensions[extension] = (module, filename) => module._compile(ts.transpileModule(
    fs.readFileSync(filename, 'utf8'),
    { compilerOptions: { module: ts.ModuleKind.CommonJS, jsx: ts.JsxEmit.React, target: ts.ScriptTarget.ES2019 } }
  ).outputText, filename);
}
require.extensions['.scss'] = (module) => { module.exports = { default: {} }; };
const dates = require('../src/webparts/dhaPortfolioManager/components/hapExpirationDates.ts');
const { readAllSharePointItems } = require('../src/webparts/dhaPortfolioManager/components/sharePointPaging.ts');
const React = require('react');
const { renderToStaticMarkup } = require('react-dom/server');
const HapExpirationsView = require('../src/webparts/dhaPortfolioManager/components/HapExpirations.tsx').default;
function HapExpirations(props) {
  const [filter, onFilterChange] = React.useState('upcoming');
  return React.createElement(HapExpirationsView, { ...props, filter, onFilterChange });
}

test('expiration boundaries are exclusive and upcoming excludes today and expired', () => {
  const cases = [[-1, 'expired'], [0, 'today'], [1, '1-30'], [30, '1-30'], [31, '31-60'], [60, '31-60'], [61, '61-90'], [90, '61-90'], [91, undefined]];
  for (const [days, expected] of cases) {
    assert.equal(dates.expirationBucket(days), expected);
    assert.equal(dates.matchesExpiration(expected, 'upcoming'), days >= 1 && days <= 90);
  }
  assert.equal(dates.expirationBucket(undefined), undefined);
  assert.equal(dates.expirationBucket(NaN), undefined);
});

test('stored dates preserve their date portion and reject invalid dates', () => {
  const day = dates.storedCalendarDay('2026-09-30');
  assert.equal(dates.storedCalendarDay('2026-09-30T00:00:00Z'), day);
  assert.equal(dates.storedCalendarDay('2026-09-30T23:00:00-05:00'), day);
  for (const value of [null, undefined, '', 'invalid', '2026-02-30', '2026-02-30T00:00:00Z', '2026-13-01', '2026-00-01', '2026-09-00', '2026-09-30Trash']) {
    assert.equal(dates.storedCalendarDay(value), undefined, String(value));
  }
  assert.equal(dates.expirationDateLabel('2026-09-30T00:00:00Z'), '9/30/2026');
});

test('Dallas midnight, daylight saving, month/year rollover and leap day', () => {
  for (const [instant, expected] of [
    ['2026-09-30T04:59:59Z', '2026-09-29'], ['2026-09-30T05:00:00Z', '2026-09-30'],
    ['2026-03-08T07:59:59Z', '2026-03-08'], ['2026-03-08T08:00:00Z', '2026-03-08'],
    ['2026-11-01T06:59:59Z', '2026-11-01'], ['2026-11-01T07:00:00Z', '2026-11-01'],
    ['2027-01-01T05:59:59Z', '2026-12-31'], ['2027-01-01T06:00:00Z', '2027-01-01'],
  ]) assert.equal(dates.dallasToday(new Date(instant)), dates.storedCalendarDay(expected));
  for (const [start, end, expected] of [
    ['2026-03-08', '2026-03-09', 1], ['2026-11-01', '2026-11-02', 1],
    ['2026-12-31', '2027-01-01', 1], ['2028-02-28', '2028-03-01', 2], ['2026-09-30', '2026-10-01', 1],
  ]) assert.equal(dates.storedCalendarDay(end) - dates.storedCalendarDay(start), expected);
});

test('results are independent of the viewer time zone', () => {
  const previous = process.env.TZ;
  try {
    for (const zone of ['Asia/Manila', 'America/Los_Angeles', 'UTC']) {
      process.env.TZ = zone;
      assert.equal(dates.storedCalendarDay('2026-09-30T00:00:00Z') - dates.dallasToday(new Date('2026-09-30T04:00:00Z')), 1);
      assert.equal(dates.expirationDateLabel('2026-09-30T00:00:00Z'), '9/30/2026');
    }
  } finally { if (previous === undefined) delete process.env.TZ; else process.env.TZ = previous; }
});

test('SharePoint paging includes records beyond 5000 and every supported next-link format', async () => {
  const first = Array.from({ length: 5000 }, (_, index) => ({ Id: index + 1 }));
  const responses = {
    first: { value: first, '@odata.nextLink': 'second' },
    second: { value: [{ Id: 5001 }], 'odata.nextLink': 'third' },
    third: { d: { results: [{ Id: 5002 }], __next: 'fourth' } },
    fourth: { value: [{ Id: 5003 }] },
  };
  const calls = [];
  const result = await readAllSharePointItems('first', async (url) => { calls.push(url); return responses[url]; });
  assert.equal(result.length, 5003);
  assert.deepEqual(result.slice(-3).map((item) => item.Id), [5001, 5002, 5003]);
  assert.deepEqual(calls, ['first', 'second', 'third', 'fourth']);
});

test('later-page failure rejects instead of returning incomplete results', async () => {
  await assert.rejects(readAllSharePointItems('first', async (url) => {
    if (url === 'first') return { value: [{ Id: 1 }], '@odata.nextLink': 'second' };
    throw new Error('Access denied');
  }), /Access denied/);
});

const today = dates.storedCalendarDay('2026-09-29');
const record = (id, days, resident = `Resident ${id}`) => ({
  id, resident, property: 'Synthetic Property', unit: String(id), start: '2025-09-29',
  end: days === undefined ? '' : new Date((today + days) * 86400000).toISOString(),
  documentUrl: id === 1 ? 'https://example.test/documents/' : undefined,
});
const render = (records, overrides = {}) => renderToStaticMarkup(React.createElement(HapExpirations, {
  records, today, loading: false, error: '', filterKey: '', onOpen() {}, ...overrides,
}));

test('initial view shows upcoming rows sorted by end then resident, separate counts and document links', () => {
  const html = render([record(1, 30, 'Zulu'), record(2, 1, 'Beta'), record(3, 1, 'Alpha'), record(4, 31), record(5, 61), record(6, 0), record(7, -1), record(8, 91), record(9, undefined)]);
  assert.match(html, /HAP Expirations <span>\(5\)<\/span>/);
  assert.ok(!html.includes('<caption'));
  assert.ok(!html.includes('Resident table filters do not apply'));
  assert.ok(!html.includes('>Refresh</button>'));
  assert.match(html, /1 records have missing or invalid HAP end dates/);
  assert.ok(html.indexOf('Open resident Alpha') < html.indexOf('Open resident Beta'));
  assert.ok(html.indexOf('Open resident Beta') < html.indexOf('Open resident Zulu'));
  for (const id of [6, 7, 8, 9]) assert.ok(!html.includes(`Open resident Resident ${id}`));
  assert.match(html, /href="https:\/\/example.test\/documents\/"/);
  assert.match(html, /Not uploaded/);
  assert.match(html, /aria-pressed="true"/);
});

test('pagination does not limit bucket counts and edited dates leave the upcoming view', () => {
  const records = Array.from({ length: 12 }, (_, index) => record(index + 1, 1));
  const html = render(records);
  assert.equal((html.match(/Open resident /g) || []).length, 10);
  assert.match(html, /of 12 records/);
  assert.match(html, /Page 1 of 2/);
  assert.match(html, /1–30 days<\/span><strong>12<\/strong>/);
  assert.match(render([record(1, 91)]), /No HAP contracts match/);
});

test('loading and errors hide stale rows and counts', () => {
  for (const overrides of [{ loading: true }, { error: 'Synthetic failure' }]) {
    const html = render([record(1, 1)], overrides);
    assert.ok(!html.includes('Open resident'));
    assert.ok(!html.includes('<strong>1</strong>'));
    assert.match(html, overrides.loading ? /Loading HAP expirations/ : /Use Refresh Data to retry/);
  }
});

test('bucket selection, pagination, sorting, editing refresh and row actions work together', () => {
  const { JSDOM } = require('jsdom');
  const dom = new JSDOM('<!doctype html><div id="root"></div>', { url: 'https://example.test/', pretendToBeVisual: true });
  const previousWindow = global.window;
  const previousDocument = global.document;
  const previousMessageChannel = global.MessageChannel;
  // React 17's browser scheduler otherwise leaves Node's MessageChannel ports open.
  global.MessageChannel = undefined;
  global.window = dom.window;
  global.document = dom.window.document;
  const ReactDOM = require('react-dom');
  const { act } = require('react-dom/test-utils');
  const container = document.getElementById('root');
  let opened;
  let props = {
    records: [...Array.from({ length: 12 }, (_, i) => record(i + 1, i + 1)), record(20, 31), record(21, 61), record(22, 0), record(23, -1)],
    today, loading: false, error: '', filterKey: 'all', onOpen(id) { opened = id; },
  };
  const update = () => act(() => { ReactDOM.render(React.createElement(HapExpirations, props), container); });
  const click = (label) => {
    const button = Array.from(container.querySelectorAll('button')).find((item) => item.textContent.startsWith(label));
    assert.ok(button, label);
    act(() => { button.dispatchEvent(new window.MouseEvent('click', { bubbles: true })); });
  };
  const rowCount = () => container.querySelectorAll('tbody button').length;
  try {
    update();
    click('Next');
    assert.equal(rowCount(), 4);
    click('31–60 days');
    assert.equal(rowCount(), 1);
    assert.equal(container.querySelector('h2').textContent, 'HAP Expirations (1)');
    assert.match(container.querySelector('tbody').textContent, /Resident 20/);
    assert.match(container.textContent, /Page 1 of 1/);
    click('61–90 days');
    assert.match(container.querySelector('tbody').textContent, /Resident 21/);
    click('Expiring today');
    assert.match(container.querySelector('tbody').textContent, /Resident 22/);
    click('Expired');
    assert.match(container.querySelector('tbody').textContent, /1 days overdue/);
    click('1–30 days');
    click('HAP end');
    assert.match(container.querySelector('tbody button').textContent, /Resident 12/);
    click('HAP end');
    click('Resident 1');
    assert.equal(opened, 1);
    opened = undefined;
    act(() => { container.querySelector('tbody a').dispatchEvent(new window.MouseEvent('click', { bubbles: true, cancelable: true })); });
    assert.equal(opened, undefined, 'document link must not open the editor');
    props = { ...props, records: props.records.map((item) => item.id === 1 ? record(1, 91) : item) };
    update();
    assert.ok(!container.querySelector('[aria-label="Open resident Resident 1"]'));
    click('Next');
    props = { ...props, filterKey: 'different property' };
    update();
    assert.match(container.textContent, /Page 1 of 2/);
    props = { ...props, today: today + 1 };
    update();
    click('Expiring today');
    assert.equal(rowCount(), 0);
  } finally {
    act(() => { ReactDOM.unmountComponentAtNode(container); });
    dom.window.close();
    global.window = previousWindow;
    global.document = previousDocument;
    global.MessageChannel = previousMessageChannel;
  }
});
