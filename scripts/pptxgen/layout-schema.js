'use strict';

const LAYOUTS = Object.freeze({
  'deck-cover': layout('deck', 'cover'),
  'deck-section': layout('deck', 'section'),
  'deck-closing': layout('deck', 'closing'),

  'text-quote': layout('text', 'bigQuote'),
  'text-article': layout('text', 'article', 'items'),
  'text-briefing': layout('text', 'briefing', 'items'),
  'text-list': layout('text', 'sectionList', 'items'),
  'text-grid': layout('text', 'textGrid', 'items'),
  'text-cards': layout('text', 'fourCards', 'items'),
  'text-weave': layout('text', 'textWeave', 'items'),
  'text-agenda': layout('text', 'agenda', 'items'),
  'text-timeline': layout('text', 'timeline', 'items', 'steps'),
  'text-pipeline': layout('text', 'pipeline', 'items', 'steps'),
  'text-roadmap': layout('text', 'roadmap', 'items', 'steps'),
  'text-matrix': layout('text', 'matrix', 'items'),
  'text-radial': layout('text', 'radial', 'items'),
  'text-pyramid': layout('text', 'pyramid', 'items', 'layers'),
  'text-swimlane': layout('text', 'swimlane', 'items', 'lanes'),

  'image-statement': layout('image', 'statement', null, null, { minImages: 1, maxImages: 1 }),
  'image-quote': layout('image', 'quoteImage', null, null, { minImages: 1, maxImages: 1 }),
  'image-text': layout('image', 'textImage', null, null, { minImages: 1, maxImages: 1 }),
  'image-feature': layout('image', 'media', 'items', null, { minImages: 1, maxImages: 1 }),
  'image-grid': layout('image', 'mediaGrid', 'items', 'captions', { minImages: 1, maxImages: 6 }),
  'image-hero': layout('image', 'imageHero', 'items', null, { minImages: 1, maxImages: 1 }),
  'image-case-study': layout('image', 'caseStudy', 'items', 'metrics', { minImages: 1, maxImages: 1 }),

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
  quoteImage: 'image-quote',
  textImage: 'image-text',
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
  if (definition.category !== 'image' && images.length) errors.push(`slide ${page} layout "${name}" does not accept images. Choose an image-* layout.`);
  if (definition.category === 'image') validateCount(page, name, 'images', images.length, definition.minImages, definition.maxImages, errors);

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
