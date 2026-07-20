'use strict';

const LAYOUTS = Object.freeze({
  'deck-cover': layout('deck', 'cover'),
  'deck-section': layout('deck', 'section'),
  'deck-closing': layout('deck', 'closing'),

  'text-statement': pairedLayout('text', 'statement', 'bigQuote'),
  'text-quote': pairedLayout('text', 'quote', 'bigQuote'),
  'text-article': pairedLayout('text', 'article', 'article', 'items'),
  'text-briefing': pairedLayout('text', 'briefing', 'briefing', 'items'),
  'text-feature': pairedLayout('text', 'feature', 'pairedText', 'items'),
  'text-list': pairedLayout('text', 'list', 'sectionList', 'items'),
  'text-grid': pairedLayout('text', 'grid', 'textGrid', 'items'),
  'text-cards': pairedLayout('text', 'cards', 'fourCards', 'items'),
  'text-weave': pairedLayout('text', 'weave', 'textWeave', 'items'),
  'text-agenda': pairedLayout('text', 'agenda', 'agenda', 'items'),
  'text-timeline': pairedLayout('text', 'timeline', 'timeline', 'items', 'steps'),
  'text-pipeline': pairedLayout('text', 'pipeline', 'pipeline', 'items', 'steps'),
  'text-roadmap': pairedLayout('text', 'roadmap', 'roadmap', 'items', 'steps'),
  'text-matrix': pairedLayout('text', 'matrix', 'matrix', 'items'),
  'text-radial': pairedLayout('text', 'radial', 'radial', 'items'),
  'text-pyramid': pairedLayout('text', 'pyramid', 'pyramid', 'items', 'layers'),
  'text-swimlane': pairedLayout('text', 'swimlane', 'swimlane', 'items', 'lanes'),
  'text-hero': pairedLayout('text', 'hero', 'pairedText', 'items'),
  'text-case-study': pairedLayout('text', 'case-study', 'pairedText', 'items'),

  'image-statement': pairedLayout('image', 'statement', 'statement', null, null, { minImages: 1, maxImages: 1 }),
  'image-quote': pairedLayout('image', 'quote', 'quoteImage', null, null, { minImages: 1, maxImages: 1 }),
  'image-article': pairedLayout('image', 'article', 'pairedMedia', 'items', null, { minImages: 1, maxImages: 6 }),
  'image-briefing': pairedLayout('image', 'briefing', 'pairedMedia', 'items', null, { minImages: 1, maxImages: 6 }),
  'image-feature': pairedLayout('image', 'feature', 'pairedMedia', 'items', null, { minImages: 1, maxImages: 6 }),
  'image-list': pairedLayout('image', 'list', 'pairedMedia', 'items', null, { minImages: 1, maxImages: 6 }),
  'image-grid': pairedLayout('image', 'grid', 'pairedMedia', 'items', null, { minImages: 1, maxImages: 6 }),
  'image-cards': pairedLayout('image', 'cards', 'pairedMedia', 'items', null, { minImages: 1, maxImages: 6 }),
  'image-weave': pairedLayout('image', 'weave', 'pairedMedia', 'items', null, { minImages: 1, maxImages: 6 }),
  'image-agenda': pairedLayout('image', 'agenda', 'pairedMedia', 'items', null, { minImages: 1, maxImages: 6 }),
  'image-timeline': pairedLayout('image', 'timeline', 'pairedMedia', 'items', null, { minImages: 1, maxImages: 6 }),
  'image-pipeline': pairedLayout('image', 'pipeline', 'pairedMedia', 'items', null, { minImages: 1, maxImages: 6 }),
  'image-roadmap': pairedLayout('image', 'roadmap', 'pairedMedia', 'items', null, { minImages: 1, maxImages: 6 }),
  'image-matrix': pairedLayout('image', 'matrix', 'pairedMedia', 'items', null, { minImages: 1, maxImages: 6 }),
  'image-radial': pairedLayout('image', 'radial', 'pairedMedia', 'items', null, { minImages: 1, maxImages: 6 }),
  'image-pyramid': pairedLayout('image', 'pyramid', 'pairedMedia', 'items', null, { minImages: 1, maxImages: 6 }),
  'image-swimlane': pairedLayout('image', 'swimlane', 'pairedMedia', 'items', null, { minImages: 1, maxImages: 6 }),
  'image-hero': pairedLayout('image', 'hero', 'pairedMedia', 'items', null, { minImages: 1, maxImages: 6 }),
  'image-case-study': pairedLayout('image', 'case-study', 'pairedMedia', 'items', null, { minImages: 1, maxImages: 6 }),

  'data-numbers': layout('data', 'bigNumbers', 'items'),
  'data-kpis': layout('data', 'kpiTower', 'items'),
  'data-compare': layout('data', 'compare'),
  'data-chart': layout('data', 'chart', 'items', 'insights', { minCharts: 1, maxCharts: 1 }),
  'data-dashboard': layout('data', 'dashboard', 'items', 'metrics', { minCharts: 2, maxCharts: 2 }),
  'data-table': layout('data', 'dataSheet', 'items', 'notes', { requiresTable: true }),
});

const LEGACY_LAYOUTS = Object.freeze({
  cover: 'deck-cover',
  section: 'deck-section',
  closing: 'deck-closing',
  bigQuote: 'text-quote',
  article: 'text-article',
  briefing: 'text-briefing',
  executiveBrief: 'text-briefing',
  contentBrief: 'text-briefing',
  sectionList: 'text-list',
  textGrid: 'text-grid',
  fourCards: 'text-cards',
  textWeave: 'text-weave',
  contentSynthesis: 'text-weave',
  denseText: 'text-weave',
  agenda: 'text-agenda',
  timeline: 'text-timeline',
  pipeline: 'text-pipeline',
  roadmap: 'text-roadmap',
  matrix: 'text-matrix',
  radial: 'text-radial',
  pyramid: 'text-pyramid',
  swimlane: 'text-swimlane',
  statement: 'image-statement',
  'image-text': 'image-article',
  quoteImage: 'image-quote',
  textImage: 'image-article',
  media: 'image-feature',
  mediaGrid: 'image-grid',
  gallery: 'image-grid',
  imageGrid: 'image-grid',
  imageHero: 'image-hero',
  caseStudy: 'image-case-study',
  bigNumbers: 'data-numbers',
  kpiTower: 'data-kpis',
  compare: 'data-compare',
  duoCompare: 'data-compare',
  splitCompare: 'data-compare',
  chart: 'data-chart',
  dashboard: 'data-dashboard',
  dataSheet: 'data-table',
});

const DEPRECATED_COLLECTION_FIELDS = Object.freeze([
  'sections', 'columns', 'steps', 'nodes', 'layers', 'lanes', 'metrics', 'agenda', 'captions', 'points', 'insights', 'notes',
]);
const DEPRECATED_MEDIA_FIELDS = Object.freeze(['image', 'gallery', 'media']);
const DEPRECATED_CHART_FIELDS = Object.freeze(['chart']);

function layout(category, renderer, collection = null, rendererCollection = null, extra = {}) {
  return Object.freeze({ category, renderer, collection, rendererCollection: rendererCollection || collection, ...extra });
}

function pairedLayout(category, pairKey, renderer, collection = null, rendererCollection = null, extra = {}) {
  const counterpart = `${category === 'text' ? 'image' : 'text'}-${pairKey}`;
  return layout(category, renderer, collection, rendererCollection, { pairKey, counterpart, ...extra });
}

function canonicalLayoutNames() {
  return Object.keys(LAYOUTS);
}

function layoutDefinition(name) {
  return LAYOUTS[name] || null;
}

function defaultLayoutForStyle() {
  return 'image-statement';
}

function validateCanonicalSpec(spec, fail) {
  const errors = [];
  spec.slides.forEach((slide, index) => validateCanonicalSlide(slide, index, errors));
  if (errors.length) {
    fail(`Canonical layout validation failed:\n- ${errors.join('\n- ')}`);
  }
}

function validateCanonicalSlide(slide, index, errors) {
  const page = index + 1;
  const name = slide.layout || defaultLayoutForStyle();
  const definition = layoutDefinition(name);
  if (!definition) {
    const replacement = LEGACY_LAYOUTS[name];
    if (replacement) {
      errors.push(`slide ${page} uses legacy layout "${name}". Rename it to "${replacement}".`);
    } else {
      errors.push(`slide ${page} uses unsupported layout "${name}". Use one of: ${canonicalLayoutNames().join(', ')}.`);
    }
    return;
  }

  DEPRECATED_COLLECTION_FIELDS.forEach((field) => {
    if (slide[field] !== undefined) errors.push(`slide ${page} layout "${name}" uses legacy collection field "${field}". Rename it to "items".`);
  });
  DEPRECATED_MEDIA_FIELDS.forEach((field) => {
    if (slide[field] !== undefined) errors.push(`slide ${page} layout "${name}" uses legacy media field "${field}". Use "images": ["path/to/image"] instead.`);
  });
  DEPRECATED_CHART_FIELDS.forEach((field) => {
    if (slide[field] !== undefined) errors.push(`slide ${page} layout "${name}" uses legacy chart field "${field}". Use "charts": [chartSpec] instead.`);
  });

  if (!definition.collection && slide.items !== undefined && !['data-compare'].includes(name)) {
    errors.push(`slide ${page} layout "${name}" does not render "items". Remove it or choose a layout with content items.`);
  }
  if (definition.collection && !hasItems(slide.items) && !slide.allowSparseContent) {
    errors.push(`slide ${page} layout "${name}" requires a non-empty "items" array.`);
  }

  const images = Array.isArray(slide.images) ? slide.images : [];
  if (slide.images !== undefined && !Array.isArray(slide.images)) errors.push(`slide ${page} field "images" must be an array.`);
  if (definition.category !== 'image' && images.length) errors.push(`slide ${page} layout "${name}" does not accept images. Change layout to its field-compatible image counterpart "${definition.counterpart}".`);
  if (definition.category === 'image' && images.length) validateCount(page, name, 'images', images.length, definition.minImages, definition.maxImages, errors);
  if (definition.category === 'image' && !images.length) errors.push(`slide ${page} layout "${name}" requires image content. Provide "images", or change layout to its field-compatible text counterpart "${definition.counterpart}".`);

  const charts = Array.isArray(slide.charts) ? slide.charts : [];
  if (slide.charts !== undefined && !Array.isArray(slide.charts)) errors.push(`slide ${page} field "charts" must be an array.`);
  if (!['data-chart', 'data-dashboard'].includes(name) && charts.length) errors.push(`slide ${page} layout "${name}" does not accept charts. Choose a data-chart or data-dashboard layout.`);
  if (definition.minCharts) validateCount(page, name, 'charts', charts.length, definition.minCharts, definition.maxCharts, errors);

  if (definition.requiresTable && !slide.table && !slide.allowMissingTable) errors.push(`slide ${page} layout "${name}" requires "table" data.`);
  if (!definition.requiresTable && slide.table !== undefined) errors.push(`slide ${page} layout "${name}" does not render "table". Choose "data-table".`);
}

function validateCount(page, layoutName, field, count, min, max, errors) {
  if (min && count < min) errors.push(`slide ${page} layout "${layoutName}" requires at least ${min} ${field} entry; got ${count}.`);
  if (max && count > max) errors.push(`slide ${page} layout "${layoutName}" supports at most ${max} ${field} entries; got ${count}.`);
}

function hasItems(value) {
  if (Array.isArray(value)) return value.length > 0;
  return value !== undefined && value !== null && String(value).trim().length > 0;
}

function createRendererSpec(spec) {
  return {
    ...spec,
    slides: spec.slides.map(createRendererSlide),
  };
}

function createRendererSlide(slide) {
  const canonicalName = slide.layout || defaultLayoutForStyle();
  const definition = layoutDefinition(canonicalName);
  if (!definition) return { ...slide };
  const next = { ...slide, layout: definition.renderer, __canonicalLayout: canonicalName };

  if (definition.collection && slide.items !== undefined) {
    const target = definition.rendererCollection || 'items';
    next[target] = slide.items;
    if (target !== 'items') delete next.items;
  }
  if (Array.isArray(slide.images)) next.images = slide.images.slice();
  if (Array.isArray(slide.charts)) {
    next.charts = slide.charts.slice();
    if (canonicalName === 'data-chart') next.chart = slide.charts[0];
  }
  return next;
}

module.exports = {
  LAYOUTS,
  LEGACY_LAYOUTS,
  canonicalLayoutNames,
  layoutDefinition,
  defaultLayoutForStyle,
  validateCanonicalSpec,
  createRendererSpec,
  createRendererSlide,
};
