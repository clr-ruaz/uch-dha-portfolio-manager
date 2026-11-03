const assert = require('node:assert/strict');
const { test } = require('node:test');
const fs = require('node:fs');
const ts = require('typescript');
const Module = require('node:module');

// Load the real dashboard with the same lightweight setup as the editor tests.
for (const extension of ['.ts', '.tsx']) {
  require.extensions[extension] = (module, filename) => module._compile(ts.transpileModule(
    fs.readFileSync(filename, 'utf8'),
    { compilerOptions: { module: ts.ModuleKind.CommonJS, jsx: ts.JsxEmit.React, target: ts.ScriptTarget.ES2019 } }
  ).outputText, filename);
}
require.extensions['.scss'] = (module) => {
  module.exports = { default: new Proxy({}, { get: (_, key) => key }) };
};
const originalLoad = Module._load;
Module._load = function (name, ...args) {
  if (name === '@microsoft/sp-http') return { SPHttpClient: { configurations: { v1: {} } } };
  return originalLoad.call(this, name, ...args);
};
const Dashboard = require('../src/webparts/dhaPortfolioManager/components/Dashboard.tsx').default;
Module._load = originalLoad;

test('document completeness counts each document separately, ignores HAP dates, and respects filters across pages', async () => {
  const { JSDOM } = require('jsdom');
  const dom = new JSDOM('<!doctype html><div id="root"></div>', { url: 'https://example.test/', pretendToBeVisual: true });
  const globals = ['window', 'document', 'MessageChannel', 'ResizeObserver'];
  const previous = Object.fromEntries(globals.map((key) => [key, global[key]]));
  global.window = dom.window;
  global.document = dom.window.document;
  global.MessageChannel = undefined;
  global.ResizeObserver = class { observe() {} disconnect() {} };
  const React = require('react');
  const ReactDOM = require('react-dom');
  const { act, Simulate } = require('react-dom/test-utils');
  const root = document.getElementById('root');
  const fields = ['HAPContract', 'WelcomeLetter', 'HAPContractStart', 'HAPContractEnd'];
  const values = ['https://example.test/contract.pdf', 'https://example.test/welcome.pdf', '2026-01-01', '2027-01-01'];
  // Every combination of present/missing fields, with a SharePoint ID on each.
  const records = Array.from({ length: 16 }, (_, mask) => ({
    Id: mask + 1, Title: `${(mask & 3) === 3 ? 'Ready' : 'Missing'} resident ${mask}`, TenantNameId: mask + 1,
    ...Object.fromEntries(fields.map((field, index) => [field, mask & (1 << index) ? values[index] : null])),
  }));
  const complete = Object.fromEntries(fields.map((field, index) => [field, values[index]]));
  for (const field of fields) records.push({
    ...complete, Id: records.length + 1, TenantNameId: records.length + 1,
    Title: `${field === 'HAPContract' || field === 'WelcomeLetter' ? 'Missing' : 'Ready'} blank ${field}`, [field]: '   ',
  });
  records.push({
    ...complete, Id: 21, TenantNameId: 21, Title: 'Missing link values',
    HAPContract: { __metadata: { type: 'SP.FieldUrlValue' } },
    WelcomeLetter: { Url: '', Description: '' },
  });
  records.push({
    ...complete, Id: 22, TenantNameId: 22, Title: 'Ready link objects',
    HAPContract: { Url: values[0], Description: 'Contract' },
    WelcomeLetter: { Url: values[1], Description: 'Welcome' },
  });
  const people = records.map((record) => ({
    Id: record.Id, personName: record.Title, propertyName: 'Test Property',
    unitName: String(record.Id), residencyStatus: 'Current',
  }));
  const props = {
    context: {
      instanceId: 'completeness-test',
      pageContext: { web: { absoluteUrl: 'https://example.test' }, user: { email: 'test@example.test' } },
      spHttpClient: {
        async get(url) {
          const value = url.includes("getbytitle('DHA Intake')/items") ? records
            : url.includes("getbytitle('ResMan People')/items") ? people : [];
          return { ok: true, json: async () => ({ value }) };
        },
      },
    },
  };
  const kpi = (title) => Array.from(root.querySelectorAll('.kpi'))
    .find((el) => el.querySelector('span').textContent === title).querySelector('strong').textContent;
  const search = (value) => act(() => Simulate.change(root.querySelector('input[placeholder^="Search residents"]'), { target: { value } }));
  try {
    await act(async () => {
      ReactDOM.render(React.createElement(Dashboard, props), root);
      await new Promise((resolve) => setImmediate(resolve));
    });
    assert.equal(kpi('Residents Shown'), '22');
    assert.equal(kpi('Document Completeness'), '55%', '24 recorded documents out of 44 expected, rounded');
    search('Missing resident 2');
    assert.equal(kpi('Residents Shown'), '1');
    assert.equal(kpi('Document Completeness'), '50%', 'a welcome letter alone earns half credit even without HAP dates');
    search('Missing blank WelcomeLetter');
    assert.equal(kpi('Residents Shown'), '1');
    assert.equal(kpi('Document Completeness'), '50%', 'a contract alone earns half credit; whitespace is missing');
    search('Missing resident 0');
    assert.equal(kpi('Residents Shown'), '1');
    assert.equal(kpi('Document Completeness'), '0%', 'a SharePoint ID alone earns no document credit');
    search('Ready resident 3');
    assert.equal(kpi('Residents Shown'), '1');
    assert.equal(kpi('Document Completeness'), '100%', 'both documents count even when both HAP dates are missing');
    search('Ready');
    assert.equal(kpi('Residents Shown'), '7');
    assert.equal(kpi('Document Completeness'), '100%', 'string and SharePoint URL object document values count');
    search('Missing');
    assert.equal(kpi('Residents Shown'), '15');
    assert.equal(kpi('Document Completeness'), '33%', 'partial records contribute 10 documents out of 30 expected');
    search('Missing link values');
    assert.equal(kpi('Document Completeness'), '0%', 'metadata and empty URL objects earn no credit');
    search('No matching resident');
    assert.equal(kpi('Residents Shown'), '0');
    assert.equal(kpi('Document Completeness'), '0%', 'empty results avoid division by zero');
    search('');
    assert.equal(kpi('Document Completeness'), '55%');
  } finally {
    act(() => { ReactDOM.unmountComponentAtNode(root); });
    dom.window.close();
    for (const key of globals) global[key] = previous[key];
  }
});
