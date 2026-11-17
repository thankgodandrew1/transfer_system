// DOM interaction checks; no browser, server submissions or external resources.
const assert = require('node:assert/strict');
const fs = require('node:fs');
const { JSDOM } = require(process.env.REVIEW_TEST_JSDOM || 'jsdom');
const dom = new JSDOM(fs.readFileSync(process.argv[2], 'utf8'), {runScripts: 'outside-only', url: 'http://localhost/'});
const {window} = dom;
const doc = window.document;
window.matchMedia = () => ({matches: false});
window.eval(fs.readFileSync(process.argv[3], 'utf8'));
const get = selector => {
  const element = doc.querySelector(selector);
  assert.ok(element, `Missing ${selector}`);
  return element;
};
const change = (selector, value, type = 'change') => {
  const element = get(selector);
  element.value = value;
  element.dispatchEvent(new window.Event(type, {bubbles: true}));
};

assert.equal(get('#roster-tab-previous').getAttribute('aria-selected'), 'true');
assert.equal(get('#roster-panel-current').hidden, true);
change('#roster-search', 'GAMMA', 'input');
assert.match(get('#roster-results').textContent, /Showing 0 of 2/);
assert.equal(get('#roster-panel-previous .roster-empty').hidden, false);
get('#roster-tab-current').click();
assert.match(get('#roster-results').textContent, /Showing 1 of 3/);
change('#roster-zone-filter', 'UYO');
assert.match(get('#roster-results').textContent, /Showing 0 of 3/);
change('#roster-search', '', 'input');
assert.match(get('#roster-results').textContent, /Showing 2 of 3/);
change('#roster-zone-filter', '');
assert.match(get('#roster-results').textContent, /Showing 3 of 3/);

if (process.argv.includes('--completed')) {
  assert.equal(doc.querySelector('#roster-editor'), null);
  assert.equal(doc.querySelector('[data-edit-source]'), null);
  console.log('Completed review navigation passed.');
  process.exit(0);
}

assert.equal(get('#editor-fields').disabled, true);
assert.equal(get('#review-missionary').value, '');
get('[data-edit-source="current"][data-edit-index="0"]').click();
assert.equal(get('#review-source').value, 'current');
assert.equal(get('#review-missionary').value, '0');
assert.equal(get('#editor-fields').disabled, false);
assert.equal(get('#review-area').value, 'NEW AREA');
assert.match(get('#editor-context').textContent, /move TO/);
assert.equal(get('[data-source="current"][data-index="0"]').classList.contains('is-selected'), true);
let form = new window.FormData(get('#roster-editor'));
assert.equal(form.get('source'), 'current');
assert.equal(form.get('missionary'), '0');
assert.deepEqual(form.getAll('companions'), ['1']);
const gamma = get('#review-companions input[value="2"]');
gamma.checked = true;
gamma.dispatchEvent(new window.Event('change', {bubbles: true}));
change('#companion-search', 'gamma', 'input');
assert.match(get('#companion-count').textContent, /2 selected/);
assert.match(get('#assignment-scope').textContent, /update 3 people/);
assert.equal(get('#review-companions input[value="1"]').closest('label').hidden, true);
form = new window.FormData(get('#roster-editor'));
assert.deepEqual(form.getAll('companions'), ['1', '2']);
assert.equal(form.has('confirmed'), false);
get('#clear-companions').click();
assert.deepEqual(new window.FormData(get('#roster-editor')).getAll('companions'), []);
assert.match(get('#assignment-scope').textContent, /Only this missionary/);
change('#review-area', 'UNSAVED AREA', 'input');
get('#reset-assignment').click();
assert.equal(get('#review-area').value, 'NEW AREA');
assert.deepEqual(new window.FormData(get('#roster-editor')).getAll('companions'), ['1']);
get('#roster-tab-previous').click();
get('[data-edit-source="previous"][data-edit-index="1"]').click();
assert.equal(get('#review-source').value, 'previous');
assert.equal(get('#review-missionary').value, '1');
assert.equal(get('#review-area').value, 'OLD AREA');
assert.match(get('#editor-context').textContent, /move FROM/);
change('#review-source', 'current');
assert.equal(get('#editor-fields').disabled, true);
assert.equal(get('#review-missionary').value, '');
assert.equal(get('#roster-tab-current').getAttribute('aria-selected'), 'true');
get('#roster-tab-current').dispatchEvent(new window.KeyboardEvent('keydown', {key: 'ArrowLeft', bubbles: true}));
assert.equal(get('#roster-tab-previous').getAttribute('aria-selected'), 'true');
get('[data-open-fold]').click();
assert.equal(get('#review-attention').open, true);
console.log('Roster tabs, search, zone filter, editor, companion selection, reset and keyboard navigation passed.');
