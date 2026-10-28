const assert = require('node:assert/strict');
const { test } = require('node:test');
const fs = require('node:fs');
const ts = require('typescript');
const Module = require('node:module');

// Exercise the actual editor with synthetic SharePoint responses, using the
// project's existing Node test runner and DOM dependencies.
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

test('resident requirements share the edit/save/cancel lifecycle and preserve field values', async () => {
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
  const flagNames = [
    'MonthlyRentRequired', 'HAPPortionRequired', 'CaseworkerNameRequired',
    'CaseworkerContactRequired', 'HAPContractStartRequired', 'HAPContractEndRequired',
    'ApplicationFeeRequired', 'AdministrativeFeeRequired', 'RiskFeeRequired',
    'SecurityDepositRequired', 'LandlordIncentiveFeeRequired', 'HAPContractRequired',
    'WelcomeLetterRequired',
  ];
  let record = {
    Id: 1, Title: 'Synthetic resident', TenantNameId: 2, DHAStatus: 'Active',
    MonthlyRent: 1400, ApplicationFee: 0, LandlordIncentiveFee: 1000,
    HAPContractStart: '2026-01-01', HAPContractEnd: '2027-01-01',
    HAPAmendment1Date: '2026-02-01', HAPAmendment1Portion: 0,
    HAPAmendment2Date: '2026-03-01', HAPAmendment2Portion: 200,
    HAPAmendment3Date: '2026-04-01', HAPAmendment3Portion: 300,
    ApplicationFeeRequired: false, HAPPortionRequired: null,
  };
  const person = { Id: 2, personName: 'Synthetic resident', propertyName: 'Test Property', unitName: '101', residencyStatus: 'Current' };
  const requests = [];
  let rejectSave = false;
  const props = {
    context: {
      instanceId: 'requirements-test',
      pageContext: { web: { absoluteUrl: 'https://example.test' }, user: { email: 'test@example.test' } },
      spHttpClient: {
        async get(url) {
          const value = url.includes("getbytitle('DHA Intake')/items") ? [record]
            : url.includes("getbytitle('ResMan People')/items") ? [person] : [];
          return { ok: true, json: async () => ({ value }) };
        },
        async fetch(url, config, options) {
          const payload = JSON.parse(options.body);
          requests.push({ url, options, payload });
          if (rejectSave) return { ok: false, status: 500 };
          record = { ...record, ...payload };
          return { ok: true };
        },
      },
    },
  };
  const settle = async (action) => act(async () => {
    action();
    await new Promise((resolve) => setImmediate(resolve));
  });
  const dialog = () => root.querySelector('[role="dialog"]');
  const button = (label) => Array.from(dialog().querySelectorAll('button')).find((el) => el.textContent === label);
  const disclosure = () => dialog().querySelector('.residentRequirements');
  const checkbox = (label) => Array.from(disclosure().querySelectorAll('label')).find((el) => el.textContent === label).querySelector('input');
  const input = (label) => Array.from(dialog().querySelectorAll('.form section label')).find((el) => el.textContent.startsWith(label)).querySelector('input');
  const toggle = (label) => act(() => checkbox(label).click());
  const open = () => act(() => {
    const row = Array.from(root.querySelectorAll('tbody tr')).find((el) => el.textContent.includes('Synthetic resident'));
    assert.ok(row, 'resident must be loaded');
    row.click();
  });
  try {
    await settle(() => ReactDOM.render(React.createElement(Dashboard, props), root));
    open();
    assert.equal(disclosure().open, false);
    assert.equal(disclosure().querySelectorAll('input').length, 13);
    assert.equal(disclosure().querySelectorAll('input:checked').length, 12);
    assert.match(disclosure().querySelector('summary').textContent, /12 of 13 required/);
    assert.equal(checkbox('Application Fee').checked, false);
    assert.equal(checkbox('HAP Portion').checked, true, 'null retains required behavior');
    assert.equal(checkbox('DHA Contract').checked, true, 'absent flags retain required behavior');
    assert.ok(!disclosure().textContent.includes('DHA Notes'));
    assert.equal(disclosure().querySelector('.requirementsDescription').textContent,
      'Select the information and documents required for this resident.');
    assert.equal(input('Application Fee').value, '0', 'a legitimate zero remains visible');
    assert.equal(input('Application Fee').disabled, false);
    assert.match(input('Application Fee').parentElement.textContent, /Not applicable/);
    assert.equal(button('Save changes').disabled, true);

    act(() => disclosure().querySelector('summary').click());
    assert.equal(disclosure().open, true);
    const amendmentInputs = ['HAP Amend 1 Date', 'HAP Portion 1', 'HAP Amend 2 Date',
      'HAP Portion 2', 'HAP Amend 3 Date', 'HAP Portion 3'];
    const amendmentValues = amendmentInputs.map((label) => input(label).value);
    const assertAmendments = (notApplicable) => {
      for (const [index, label] of amendmentInputs.entries()) {
        assert.equal(input(label).parentElement.textContent.includes('Not applicable'), notApplicable, label);
        assert.equal(input(label).disabled, false, label);
        assert.equal(input(label).value, amendmentValues[index], label);
      }
      assert.ok(!input('HAP Portion').parentElement.textContent.includes('Not applicable'),
        'the main HAP portion retains its independent requirement');
    };
    assertAmendments(false);
    toggle('Contract Start');
    assertAmendments(true);
    assert.match(disclosure().querySelector('summary').textContent, /11 of 13 required/);
    toggle('Contract End');
    assertAmendments(true);
    toggle('Contract Start');
    assertAmendments(true);
    toggle('Contract End');
    assertAmendments(false);
    assert.equal(button('Save changes').disabled, true, 'restoring both dates restores the original draft');
    toggle('Landlord Incentive');
    assert.equal(button('Save changes').disabled, false, 'flag-only edit enables Save');
    assert.equal(input('Landlord Incentive Fee').value, '1000');
    assert.equal(input('Landlord Incentive Fee').disabled, false);
    assert.match(input('Landlord Incentive Fee').parentElement.textContent, /Not applicable/);
    toggle('Landlord Incentive');
    assert.equal(button('Save changes').disabled, true, 'reverting a flag clears dirty state');
    toggle('Landlord Incentive');
    act(() => button('Cancel').click());
    assert.equal(requests.length, 0);
    open();
    assert.equal(disclosure().open, false, 'reopening starts collapsed');
    assert.equal(checkbox('Landlord Incentive').checked, true, 'Cancel discards changes');

    toggle('Landlord Incentive');
    toggle('DHA Contract');
    toggle('Application Fee');
    act(() => Simulate.change(input('Monthly Rent'), { target: { value: '1500' } }));
    rejectSave = true;
    await settle(() => button('Save changes').click());
    assert.match(dialog().querySelector('[role="alert"]').textContent, /Save failed/);
    assert.equal(checkbox('Landlord Incentive').checked, false, 'failed save retains draft');
    rejectSave = false;
    await settle(() => button('Save changes').click());
    assert.equal(dialog(), null);
    const saved = requests.at(-1);
    assert.match(saved.url, /items\(1\)$/);
    assert.equal(saved.options.headers['X-HTTP-Method'], 'MERGE');
    assert.equal(saved.payload.MonthlyRent, 1500);
    assert.equal(saved.payload.LandlordIncentiveFee, 1000);
    assert.equal(saved.payload.ApplicationFee, 0);
    assert.equal(saved.payload.HAPContractStart, '2026-01-01');
    assert.equal(saved.payload.HAPContractEnd, '2027-01-01');
    assert.deepEqual(Object.keys(saved.payload).filter((key) => key.endsWith('Required')).sort(),
      ['ApplicationFeeRequired', 'HAPContractRequired', 'LandlordIncentiveFeeRequired']);
    assert.equal(saved.payload.ApplicationFeeRequired, true);
    assert.equal(saved.payload.LandlordIncentiveFeeRequired, false);
    assert.equal(saved.payload.HAPContractRequired, false);

    open();
    assert.equal(checkbox('DHA Contract').checked, false, 'saved flags survive reload');
    assert.equal(button('Save changes').disabled, true);
    toggle('Contract End');
    await settle(() => button('Save changes').click());
    assert.equal(requests.at(-1).payload.HAPContractEndRequired, false);
    assert.equal(requests.at(-1).payload.HAPAmendment1Date, '2026-02-01');
    assert.equal(requests.at(-1).payload.HAPAmendment1Portion, 0);
    assert.ok(!Object.keys(requests.at(-1).payload).some((key) => /^HAPAmendment.*Required$/.test(key)),
      'derived amendment applicability does not introduce SharePoint columns');
    open();
    assertAmendments(true);
    // All 13 flags must serialize as real booleans, including the all-No case.
    for (const label of Array.from(disclosure().querySelectorAll('label'))) {
      if (label.querySelector('input').checked) toggle(label.textContent);
    }
    assert.match(disclosure().querySelector('summary').textContent, /0 of 13 required/);
    await settle(() => button('Save changes').click());
    for (const key of flagNames) assert.equal(record[key], false, key);
    assert.equal(record.LandlordIncentiveFee, 1000);
    open();
    for (const label of Array.from(disclosure().querySelectorAll('label'))) toggle(label.textContent);
    await settle(() => button('Save changes').click());
    for (const key of flagNames) assert.equal(record[key], true, key);
  } finally {
    act(() => { ReactDOM.unmountComponentAtNode(root); });
    dom.window.close();
    for (const key of globals) global[key] = previous[key];
  }
});
