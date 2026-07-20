#!/usr/bin/env node
'use strict';

const assert = require('node:assert/strict');
const {
  canonicalLayoutNames,
  layoutDefinition,
  validateCanonicalSpec,
  createRendererSlide,
} = require('../../scripts/pptxgen/layout-schema');
const { exampleSlide } = require('../../scripts/pptxgen/layout-examples');

const names = canonicalLayoutNames();
assert.equal(names.length, 47, 'canonical layout count changed; update schema tests and all-layout samples');
assert(names.every((name) => /^(deck|text|image|data)-/.test(name)), 'every layout must have a category prefix');
assert(names.every((name) => layoutDefinition(name)), 'every canonical name must resolve to a definition');
validateCanonicalSpec({ slides: names.map(exampleSlide) }, (message) => {
  throw new Error(message);
});

names.filter((name) => name.startsWith('text-')).forEach((textName) => {
  const textDefinition = layoutDefinition(textName);
  const imageName = textDefinition.counterpart;
  const imageDefinition = layoutDefinition(imageName);
  assert(imageDefinition, `${textName} is missing image counterpart ${imageName}`);
  assert.equal(imageDefinition.counterpart, textName, `${imageName} must point back to ${textName}`);
  assert.deepEqual(imageDefinition.publicFields, textDefinition.publicFields, `${textName} and ${imageName} must expose the same public content field contract`);
  assert.deepEqual(exampleKeys(exampleSlide(textName)), exampleKeys(exampleSlide(imageName)), `${textName} and ${imageName} must expose identical fields except layout/images`);
});

expectSuccess({ slides: [{ layout: 'text-statement', title: 'T', body: 'B', callout: 'C' }] });
expectSuccess({ slides: [{ layout: 'image-statement', title: 'T', body: 'B', callout: 'C', images: ['a.png'] }] });
expectSuccess({ slides: [{ layout: 'text-quote', title: 'T', body: 'B', caption: 'C' }] });
expectSuccess({ slides: [{ layout: 'image-quote', title: 'T', body: 'B', caption: 'C', images: ['a.png'] }] });

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
expectFailure(
  { slides: [{ layout: 'text-grid', title: 'Wrong media', items: [{ title: 'A', body: 'B' }], images: ['a.png'] }] },
  'field-compatible image counterpart "image-grid"'
);
expectFailure(
  { slides: [{ layout: 'image-grid', title: 'Missing media', items: [{ title: 'A', body: 'B' }] }] },
  'field-compatible text counterpart "text-grid"'
);
expectFailure(
  { slides: [{ layout: 'text-article', title: 'Unsupported pair field', conclusion: 'C', items: [{ title: 'A', body: 'B' }] }] },
  'counterpart "image-article"'
);
expectFailure(
  { slides: [{ layout: 'image-article', title: 'Unsupported pair field', conclusion: 'C', items: [{ title: 'A', body: 'B' }], images: ['a.png'] }] },
  'counterpart "text-article"'
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

function expectSuccess(spec) {
  validateCanonicalSpec(spec, (message) => {
    throw new Error(`expected canonical validation to succeed; got ${message}`);
  });
}

function exampleKeys(slide) {
  return Object.keys(slide).filter((key) => !['layout', 'images'].includes(key)).sort();
}
