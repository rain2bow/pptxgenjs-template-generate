#!/usr/bin/env node
'use strict';

const assert = require('node:assert/strict');
const {
  canonicalLayoutNames,
  layoutDefinition,
  validateCanonicalSpec,
  createRendererSlide,
} = require('./pptxgen/layout-schema');
const { exampleSlide } = require('./pptxgen/layout-examples');

const names = canonicalLayoutNames();
assert.equal(names.length, 31, 'canonical layout count changed; update schema tests and all-layout samples');
assert(names.every((name) => /^(deck|text|image|data)-/.test(name)), 'every layout must have a category prefix');
assert(names.every((name) => layoutDefinition(name)), 'every canonical name must resolve to a definition');
validateCanonicalSpec({ slides: names.map(exampleSlide) }, (message) => {
  throw new Error(message);
});

const timeline = createRendererSlide({ layout: 'text-timeline', title: 'T', items: [{ title: 'A', body: 'B' }] });
assert.equal(timeline.layout, 'timeline');
assert.deepEqual(timeline.steps, [{ title: 'A', body: 'B' }]);
assert.equal(timeline.items, undefined);

const dashboard = createRendererSlide({ layout: 'data-dashboard', title: 'D', items: [{ label: 'K', value: '1' }], charts: [{ values: [1], labels: ['A'] }, { values: [2], labels: ['B'] }] });
assert.equal(dashboard.layout, 'dashboard');
assert.equal(dashboard.metrics.length, 1);
assert.equal(dashboard.charts.length, 2);

expectFailure(
  { slides: [{ layout: 'fourCards', title: 'Old', items: [{ title: 'A', body: 'B' }] }] },
  'Rename it to "text-cards"'
);
expectFailure(
  { slides: [{ layout: 'text-grid', title: 'Old field', sections: [{ title: 'A', body: 'B' }] }] },
  'Rename it to "items"'
);
expectFailure(
  { slides: [{ layout: 'image-statement', title: 'Old media', image: 'a.png' }] },
  'Use "images"'
);
expectFailure(
  { slides: [{ layout: 'data-chart', title: 'Old chart', chart: { values: [1], labels: ['A'] }, items: [{ title: 'A', body: 'B' }] }] },
  'Use "charts"'
);

console.log(`Canonical layout schema check passed: ${names.length} layout(s).`);

function expectFailure(spec, expected) {
  let message = '';
  try {
    validateCanonicalSpec(spec, (value) => {
      message = String(value);
      throw new Error('__expected_failure__');
    });
  } catch (error) {
    if (error.message !== '__expected_failure__') throw error;
    assert(message.includes(expected), `expected error to include ${JSON.stringify(expected)}; got ${message}`);
    return;
  }
  throw new Error(`expected schema validation to fail with: ${expected}`);
}
