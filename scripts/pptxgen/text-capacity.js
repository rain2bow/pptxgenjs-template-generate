const fs = require('node:fs');
const path = require('node:path');

const CAPACITY_SCALE_FOR_UNIFIED_TYPOGRAPHY = 0.82;

const COMMON_GUIDE = {
  cover: slots([['title', 'cover title', 8, 28], ['subtitle', 'cover subtitle', 12, 60]]),
  section: slots([['title', 'section title', 8, 30], ['subtitle', 'section subtitle', 12, 55]]),
  closing: slots([['title', 'closing title', 8, 30], ['subtitle', 'closing subtitle', 8, 45]]),
  statement: slots([['title', 'main statement title', 10, 36], ['body', 'statement body', 20, 80], ['subtitle', 'statement subtitle', 12, 60], ['callout', 'statement callout', 6, 36]]),
  bigQuote: slots([['quote', 'large quote', 8, 42], ['body', 'quote explanation', 12, 60]]),
  textImage: slots([['title', 'page title', 8, 32], ['body', 'body paragraph', 35, 110]]),
  article: multiCollectionSlots(['sections', 'items', 'columns'], 'article section', 2, 6, 6, 16, 25, 85),
  sectionList: multiCollectionSlots(['sections', 'items', 'columns'], 'section list item', 3, 7, 5, 18, 20, 80),
  textGrid: multiCollectionSlots(['sections', 'items', 'columns'], 'text grid card', 4, 9, 4, 14, 15, 48),
  fourCards: collectionSlots('items', 'card item', 1, 8, 4, 14, 15, 50),
  agenda: collectionSlots('items', 'agenda item', 3, 8, 4, 16, 10, 42),
  timeline: collectionSlots('steps', 'timeline step', 3, 6, 4, 14, 12, 45),
  pipeline: collectionSlots('steps', 'pipeline step', 3, 6, 4, 14, 12, 45),
  roadmap: collectionSlots('steps', 'roadmap step', 3, 6, 4, 14, 15, 50),
  radial: collectionSlots('items', 'radial node', 4, 8, 3, 12, 8, 32),
  pyramid: collectionSlots('layers', 'pyramid layer', 3, 5, 3, 12, 8, 35),
  swimlane: collectionSlots('lanes', 'swimlane lane', 2, 4, 4, 16, 8, 36),
  media: mediaSlots(),
  mediaGrid: mediaGridSlots('mediaGrid'),
  gallery: mediaGridSlots('gallery'),
  imageGrid: mediaGridSlots('imageGrid'),
  imageHero: imageHeroSlots(),
  chart: slots([['title', 'page title', 8, 32], ['insights[].title', 'chart insight title', 4, 14], ['insights[].body', 'chart insight body', 12, 48], ['notes[].body', 'chart note body', 12, 48]]),
  dataSheet: slots([['title', 'page title', 8, 32], ['notes[].title', 'side note title', 4, 14], ['notes[].body', 'side note body', 10, 42], ['insights[].body', 'side insight body', 10, 42]]),
  dashboard: metricSlots('metrics', 'dashboard metric', 1, 4, [['label', 'metric label', 2, 10], ['value', 'metric value', 1, 8]]),
  bigNumbers: metricSlots('items', 'number card', 1, 6, [['label', 'number label', 2, 12], ['value', 'number value', 1, 8], ['note', 'number note', 8, 32]]),
  kpiTower: metricSlots('items', 'KPI card', 1, 4, [['label', 'KPI label', 2, 12], ['value', 'KPI value', 1, 8], ['note', 'KPI note', 8, 32]]),
  compare: slots([['title', 'page title', 8, 32], ['before.title', 'left title', 4, 18], ['after.title', 'right title', 4, 18], ['before.items[].body', 'left item body', 8, 36], ['after.items[].body', 'right item body', 8, 36]]),
  duoCompare: null,
  splitCompare: null,
  caseStudy: caseStudySlots(),
};
COMMON_GUIDE.duoCompare = COMMON_GUIDE.compare;
COMMON_GUIDE.splitCompare = COMMON_GUIDE.compare;

const CMB_GUIDE = {
  ...COMMON_GUIDE,
  article: cmbBriefingSlots('article/briefing'),
  sectionList: cmbBriefingSlots('sectionList/briefing'),
  agenda: cmbAgendaSlots(),
  briefing: cmbBriefingSlots('briefing'),
  executiveBrief: cmbBriefingSlots('executiveBrief'),
  contentBrief: cmbBriefingSlots('contentBrief'),
  textGrid: cmbTextWeaveSlots('textGrid/textWeave'),
  fourCards: cmbTextWeaveSlots('fourCards/textWeave'),
  textWeave: cmbTextWeaveSlots('textWeave'),
  contentSynthesis: cmbTextWeaveSlots('contentSynthesis'),
  denseText: cmbTextWeaveSlots('denseText'),
};

function slots(items, description = '', extra = {}) {
  return { description, ...extra, slots: items.map(([field, label, min, max, note]) => {
    const adjustedMax = Math.max(min, Math.floor(max * CAPACITY_SCALE_FOR_UNIFIED_TYPOGRAPHY));
    return { field, label, min: recommendedCapacityMin(adjustedMax), max: adjustedMax, note };
  }) };
}

function collectionSlots(key, label, minItems, maxItems, titleMin, titleMax, bodyMin, bodyMax) {
  return multiCollectionSlots([key], label, minItems, maxItems, titleMin, titleMax, bodyMin, bodyMax);
}

function multiCollectionSlots(keys, label, minItems, maxItems, titleMin, titleMax, bodyMin, bodyMax) {
  const fields = [];
  keys.forEach((key) => {
    fields.push([key + '[].title', label + ' title', titleMin, titleMax]);
    fields.push([key + '[].body', label + ' body', bodyMin, bodyMax]);
  });
  return slots(fields, minItems + '-' + maxItems + ' items; keep every item title + body within range.', { collection: { keys, label, minItems, maxItems, titleMin, titleMax, bodyMin, bodyMax } });
}

function metricSlots(key, label, minItems, maxItems, fields) {
  return slots(
    [['title', 'page title', 8, 32], ...fields.map(([leaf, fieldLabel, min, max, note]) => [key + '[].' + leaf, fieldLabel, min, max, note])],
    minItems + '-' + maxItems + ' ' + label + '(s). Fill every listed metric field; do not use title/body for metric cards.',
    {
      collection: { keys: [key], label, minItems, maxItems, titleMin: 2, titleMax: 12, bodyMin: 6, bodyMax: 32 },
      plannedCollectionFields: fields.map(([leaf]) => leaf),
    }
  );
}

function mediaSlots() {
  return slots([
    ['title', 'page title', 8, 32],
    ['body', 'media body', 20, 80, 'Use as the short main paragraph only; side points are still required unless allowSparseContent:true.'],
    ['summary', 'media body alias', 20, 80, 'Alias of body; do not fill both body and summary. Side points are still required.'],
    ['items[].title', 'side point title', 4, 14],
    ['items[].body', 'side point body', 12, 45],
    ['insights[].title', 'side insight title', 4, 14],
    ['insights[].body', 'side insight body', 12, 45],
    ['points[].title', 'side point title alias', 4, 14],
    ['points[].body', 'side point body alias', 12, 45],
  ], 'media: one media/chart area + short main paragraph + 1-4 side points. Do not use body/summary alone, otherwise the page is too empty.', {
    collection: { keys: ['items', 'insights', 'points'], label: 'side points', minItems: 1, maxItems: 4, titleMin: 4, titleMax: 14, bodyMin: 12, bodyMax: 45 },
    plannedScalarFields: ['title', 'body'],
    plannedCollectionFields: ['title', 'body'],
  });
}

function caseStudySlots() {
  return slots([
    ['title', 'page title', 8, 32],
    ['caseTitle', 'case title', 6, 26],
    ['body', 'case body', 25, 90],
    ['summary', 'case body alias', 25, 90, 'Alias of body; do not fill both body and summary.'],
    ['metrics[].label', 'case metric label', 2, 10],
    ['metrics[].value', 'case metric value', 1, 8],
    ['metrics[].note', 'case metric note', 6, 28],
    ['items[].label', 'case metric label alias', 2, 10],
    ['items[].value', 'case metric value alias', 1, 8],
    ['items[].note', 'case metric note alias', 6, 28],
  ], 'caseStudy: case narrative + 1-3 metric chips. Do not provide only body/summary; metrics/items are required unless allowSparseContent:true.', {
    collection: { keys: ['metrics', 'items'], label: 'case metrics', minItems: 1, maxItems: 3, titleMin: 2, titleMax: 10, bodyMin: 6, bodyMax: 28 },
    plannedScalarFields: ['title', 'caseTitle', 'body'],
    plannedCollectionFields: ['label', 'value', 'note'],
  });
}

function mediaGridSlots(name) {
  return slots([
    ['title', 'page title', 8, 32],
    ['captions[].caption', 'media caption', 4, 28],
    ['captions[].title', 'media caption title alias', 4, 22],
    ['captions[].label', 'media caption label alias', 3, 18],
    ['items[].caption', 'media caption alias', 4, 28],
    ['items[].title', 'media caption title alias', 4, 22],
    ['items[].label', 'media caption label alias', 3, 18],
    ['sections[].caption', 'media caption alias', 4, 28],
    ['sections[].title', 'media caption title alias', 4, 22],
    ['sections[].label', 'media caption label alias', 3, 18],
  ], name + ': 1-6 media slots with short captions. Use captions/items/sections with caption/title/label only; long body text is not rendered in this layout.', {
    collection: { keys: ['captions', 'items', 'sections'], label: 'media captions', minItems: 1, maxItems: 6, titleMin: 4, titleMax: 28, bodyMin: 0, bodyMax: 0 },
    plannedScalarFields: ['title'],
    plannedCollectionFields: ['caption'],
  });
}

function imageHeroSlots() {
  return slots([
    ['title', 'page title', 8, 32],
    ['body', 'short hero explanation', 20, 72],
    ['subtitle', 'short hero explanation alias', 20, 72],
    ['items[].label', 'hero metric label', 2, 10],
    ['items[].value', 'hero metric value', 1, 8],
    ['items[].note', 'hero metric note', 6, 24],
  ], 'imageHero: large media area + short explanation + 1-3 metric chips. Metrics are required unless allowSparseContent:true.', {
    collection: { keys: ['items'], label: 'hero metrics', minItems: 1, maxItems: 3, titleMin: 2, titleMax: 10, bodyMin: 6, bodyMax: 24 },
    plannedScalarFields: ['title', 'body'],
    plannedCollectionFields: ['label', 'value', 'note'],
  });
}

function cmbAgendaSlots() {
  return slots([
    ['title', 'page title', 8, 32],
    ['agendaTitle', 'left rail title', 2, 8],
    ['agendaSubtitle', 'left rail subtitle', 6, 24],
    ['items[].title', 'agenda section title', 4, 16],
    ['items[].body', 'agenda section summary', 14, 44],
    ['agenda[].title', 'agenda section title alias', 4, 16],
    ['agenda[].body', 'agenda section summary alias', 14, 44],
    ['sections[].title', 'agenda section title alias', 4, 16],
    ['sections[].body', 'agenda section summary alias', 14, 44],
  ], 'agenda: dedicated CMB table-of-contents page with left rail and 1-8 chapter rows. Use concise section summaries; do not use it as a dense article page.');
}
function cmbBriefingSlots(name) {
  return slots([
    ['title', 'page title', 8, 32], ['summary', 'top summary body', 35, 90, 'Use body/lead if summary is omitted.'], ['body', 'top summary body alias', 35, 90], ['lead', 'top summary body alias', 35, 90],
    ['sections[].title', 'analysis card title', 4, 14], ['sections[].body', 'analysis card body', 18, 58], ['sections[].points[]', 'analysis card explicit bullet point', 8, 24], ['items[].title', 'analysis card title', 4, 14], ['items[].body', 'analysis card body', 18, 58], ['items[].points[]', 'analysis card explicit bullet point', 8, 24], ['agenda[].title', 'analysis card title', 4, 14], ['agenda[].body', 'analysis card body', 18, 58], ['agenda[].points[]', 'analysis card explicit bullet point', 8, 24],
    ['conclusion', 'bottom conclusion body', 12, 48], ['takeaway', 'bottom conclusion body alias', 12, 48],
  ], name + ': CMB top summary + middle analysis cards + bottom takeaway. For points[] arrays, line count is estimated per point; each point occupies at least one line.');
}

function cmbTextWeaveSlots(name) {
  return slots([
    ['title', 'page title', 8, 32], ['sections[].title', 'text card title', 4, 14], ['sections[].body', 'text card body', 15, 52], ['sections[].points[]', 'text card explicit bullet point', 8, 24], ['items[].title', 'text card title', 4, 14], ['items[].body', 'text card body', 15, 52], ['items[].points[]', 'text card explicit bullet point', 8, 24], ['columns[].title', 'text card title', 4, 14], ['columns[].body', 'text card body', 15, 52], ['columns[].points[]', 'text card explicit bullet point', 8, 24],
  ], name + ': CMB asymmetric 2-6 text cards: 1 lead card + at least 1 right-side card. 5-6 cards need shorter bodies. For points[] arrays, line count is estimated per point; each point occupies at least one line.');
}

function guideForStyle(style = 'swiss') {
  return style === 'cmb' ? CMB_GUIDE : COMMON_GUIDE;
}

function layoutCapacityGuide(style = 'swiss') {
  return {
    style,
    unit: 'visual characters; CJK counts as 1, Latin letters count as about 0.56',
    instruction: 'Use these ranges before writing JSON. Ranges are calibrated for unified typography: Microsoft YaHei / Times New Roman with 36, 28, 16, 14, 12 pt tiers. For CMB points[] arrays, capacity is also checked by estimated rendered line count: every point occupies at least one line and wraps by card width. If generation warns that a field exceeds max, shorten that field in JSON and regenerate the PPTX.',
    layouts: guideForStyle(style),
  };
}

function layoutCapacityMarkdown(style = 'swiss') {
  const guide = layoutCapacityGuide(style);
  const lines = ['# Layout text capacity guide: ' + style, '', 'Unit: ' + guide.unit, '', guide.instruction, ''];
  Object.entries(guide.layouts).forEach(([layout, info]) => {
    if (!info) return;
    lines.push('## ' + layout);
    if (info.description) lines.push('', info.description);
    lines.push('', '| Field | Suitable range | Label | Note |', '| --- | ---: | --- | --- |');
    info.slots.forEach((slot) => lines.push('| ' + slot.field + ' | ' + slot.min + '-' + slot.max + ' | ' + slot.label + ' | ' + (slot.note || '') + ' |'));
    lines.push('');
  });
  return lines.join('\n').trim() + '\n';
}

function writeLayoutCapacityGuide(style, outPath) {
  const target = path.resolve(outPath);
  fs.mkdirSync(path.dirname(target), { recursive: true });
  const content = target.toLowerCase().endsWith('.json') ? JSON.stringify(layoutCapacityGuide(style), null, 2) + '\n' : layoutCapacityMarkdown(style);
  fs.writeFileSync(target, content, 'utf8');
  console.log('Wrote layout capacity guide ' + target);
}

function layoutCapacityGuideForSpec(spec = {}) {
  validateCapacityPlan(spec);
  const style = spec.style || 'swiss';
  const slides = Array.isArray(spec.slides) ? spec.slides : [];
  return {
    style,
    unit: 'visual characters; CJK counts as 1, Latin letters count as about 0.56',
    instruction: 'This guide is calculated from the title-only deck plan. Fill only the listed fields. If a layout, media choice, or item count changes, regenerate this guide before writing full content.',
    slides: slides.map((slide, index) => {
      const planned = plannedSlideCapacity(style, slide || {}, index);
      planned.example = plannedSlideExample(slide || {}, planned);
      return planned;
    }),
  };
}

function layoutCapacityMarkdownForSpec(spec = {}) {
  const guide = layoutCapacityGuideForSpec(spec);
  const lines = ['# Planned deck text capacity guide: ' + guide.style, '', 'Unit: ' + guide.unit, '', guide.instruction, ''];
  guide.slides.forEach((slide) => {
    lines.push('## Slide ' + slide.slide + ': ' + slide.layout);
    if (slide.title) lines.push('', 'Planned title: ' + slide.title);
    if (slide.notes?.length) slide.notes.forEach((note) => lines.push('', '> ' + note));
    lines.push('', '| Field to fill | Suitable range | Planned label/title | Note |', '| --- | ---: | --- | --- |');
    if (slide.slots.length) {
      slide.slots.forEach((slot) => lines.push('| ' + slot.field + ' | ' + slot.min + '-' + slot.max + ' | ' + (slot.title || slot.label || '') + ' | ' + (slot.note || '') + ' |'));
    } else {
      lines.push('| - | - | - | This planned slide has no fillable body slots. |');
    }
    if (slide.example) {
      lines.push('', 'Example slide JSON:', '', '```json', JSON.stringify(slide.example, null, 2), '```');
    }
    lines.push('');
  });
  return lines.join('\n').trim() + '\n';
}

function writePlannedCapacityGuide(spec, outPath) {
  const target = path.resolve(outPath);
  fs.mkdirSync(path.dirname(target), { recursive: true });
  const content = target.toLowerCase().endsWith('.json')
    ? JSON.stringify(layoutCapacityGuideForSpec(spec), null, 2) + '\n'
    : layoutCapacityMarkdownForSpec(spec);
  fs.writeFileSync(target, content, 'utf8');
  console.log('Wrote planned layout capacity guide ' + target);
}

function validateCapacityPlan(spec = {}) {
  const errors = [];
  const style = spec.style || 'swiss';
  if (!Array.isArray(spec.slides) || !spec.slides.length) {
    errors.push('deck.plan.json must contain a non-empty slides array.');
  }
  (Array.isArray(spec.slides) ? spec.slides : []).forEach((slide, index) => {
    validatePlannedSlide(style, slide || {}, index, errors);
  });
  if (errors.length) {
    throw new Error('Capacity plan validation failed:\n- ' + errors.join('\n- '));
  }
}

function validatePlannedSlide(style, slide, index, errors) {
  const layout = slide.layout || 'statement';
  const layoutGuide = guideForStyle(style)[layout] || COMMON_GUIDE[layout];
  if (!layoutGuide || !Array.isArray(layoutGuide.slots)) {
    errors.push('slide ' + (index + 1) + ' uses unknown or unsupported layout "' + layout + '". Change slide.layout before generating capacity-guide.');
    return;
  }
  validatePlanCollections(style, layout, layoutGuide, slide, index, errors);
  validatePlanElementCounts(style, layout, layoutGuide, slide, index, errors);
  validatePlanScalars(layout, layoutGuide, slide, index, errors);
}

function validatePlanCollections(style, layout, layoutGuide, slide, index, errors) {
  const allowed = allowedPlanCollectionKeys(style, layout, layoutGuide);
  const collectionKeys = ['items', 'sections', 'columns', 'steps', 'nodes', 'layers', 'lanes', 'metrics', 'agenda', 'captions', 'notes', 'insights', 'points'];
  collectionKeys.forEach((key) => {
    if (slide[key] === undefined || slide[key] === null) return;
    if (!allowed.has(key)) {
      errors.push('slide ' + (index + 1) + ' layout "' + layout + '" does not use collection field "' + key + '". Use one of: ' + Array.from(allowed).join(', ') + '.');
      return;
    }
    validatePlanCollectionShape(slide[key], key, layout, index, errors);
  });
  if (['compare', 'duoCompare', 'splitCompare'].includes(layout)) {
    ['before', 'after'].forEach((key) => {
      if (slide[key] && typeof slide[key] === 'object' && slide[key].items !== undefined) {
        validatePlanCollectionShape(slide[key].items, key + '.items', layout, index, errors);
      }
    });
  }
}

function allowedPlanCollectionKeys(style, layout, layoutGuide) {
  const keys = new Set();
  (layoutGuide.slots || []).forEach((slot) => {
    const field = String(slot.field || '');
    const match = field.match(/^([^.[\]]+)\[\]/);
    if (match) keys.add(match[1]);
  });
  if (layoutGuide.collection?.keys) layoutGuide.collection.keys.forEach((key) => keys.add(key));
  if (style === 'cmb' && isCmbBriefingLayout(layout)) ['sections', 'items', 'columns', 'points', 'agenda'].forEach((key) => keys.add(key));
  if (style === 'cmb' && isCmbTextWeaveLayout(layout)) ['sections', 'items', 'columns', 'points', 'agenda'].forEach((key) => keys.add(key));
  return keys;
}

function validatePlanCollectionShape(value, key, layout, index, errors) {
  if (typeof value === 'string' || typeof value === 'number') return;
  const items = Array.isArray(value)
    ? value
    : (value && typeof value === 'object' ? Object.entries(value).map(([title, body]) => ({ title, body })) : null);
  if (!items) {
    errors.push('slide ' + (index + 1) + ' layout "' + layout + '" field "' + key + '" must be an array, object map, string, or number in deck.plan.json.');
    return;
  }
  items.forEach((item, itemIndex) => {
    if (item === null || item === undefined || typeof item === 'string' || typeof item === 'number') return;
    if (typeof item !== 'object' || Array.isArray(item)) {
      errors.push('slide ' + (index + 1) + ' field "' + key + '[' + itemIndex + ']" must be a title-only object or plain text.');
      return;
    }
    const allowed = new Set(['title', 'label', 'name', 'heading', 'value', 'unit', 'caption', 'year', 'number', 'icon', 'highlight']);
    const titleKeys = planTitleKeysForCollection(key, layout);
    if (!titleKeys.some((field) => String(item[field] ?? '').trim().length > 0)) {
      errors.push('slide ' + (index + 1) + ' field "' + key + '[' + itemIndex + ']" is missing a title field for deck.plan.json. Use one of: ' + titleKeys.join(', ') + '.');
    }
    Object.keys(item).forEach((field) => {
      if (!allowed.has(field)) {
        errors.push('slide ' + (index + 1) + ' field "' + key + '[' + itemIndex + '].' + field + '" should not be in deck.plan.json. Put body/detail text only after generating capacity-guide.');
      }
    });
  });
}

function planTitleKeysForCollection(key, layout) {
  if (['captions'].includes(key) || ['mediaGrid', 'gallery', 'imageGrid'].includes(layout)) return ['caption', 'title', 'label'];
  if (['metrics'].includes(key) || ['dashboard', 'imageHero', 'caseStudy', 'bigNumbers', 'kpiTower'].includes(layout)) return ['label', 'title', 'name', 'value'];
  if (['steps'].includes(key) || ['timeline', 'pipeline', 'roadmap'].includes(layout)) return ['title', 'label', 'year'];
  if (['lanes'].includes(key)) return ['title', 'label', 'name'];
  return ['title', 'label', 'name', 'heading'];
}

function validatePlanScalars(layout, layoutGuide, slide, index, errors) {
  const contentScalars = ['body', 'summary', 'lead', 'story', 'note', 'desc', 'detail', 'text', 'quote', 'callout', 'conclusion', 'takeaway', 'footerSummary', 'nextStep'];
  const invalidContentScalars = new Set();
  contentScalars.forEach((key) => {
    if (slide[key] !== undefined && slide[key] !== null && String(slide[key]).trim()) {
      invalidContentScalars.add(key);
      errors.push('slide ' + (index + 1) + ' field "' + key + '" should not be in deck.plan.json. Use title-only planning fields first, generate capacity-guide, then fill body text in the full JSON.');
    }
  });
  const mediaScalars = ['image', 'images', 'gallery', 'media', 'mediaCount', 'imageSlots', 'slotCount'];
  if (!planLayoutHasMediaSlot(layout)) {
    mediaScalars.forEach((key) => {
      if (slide[key] !== undefined && slide[key] !== null) {
        errors.push('slide ' + (index + 1) + ' layout "' + layout + '" does not render media field "' + key + '". Remove it or choose a media layout before generating capacity-guide.');
      }
    });
  } else if (!planHasMediaOrChart(slide)) {
    errors.push('slide ' + (index + 1) + ' layout "' + layout + '" requires image/images/media/gallery or chart/charts in deck.plan.json. mediaCount/imageSlots/allowEmptyMediaSlots are not enough because the page would start with an unfilled media slot.');
  }
  const allowed = new Set(['layout', 'title', 'kicker', 'subtitle', 'theme', 'style', 'itemCount', 'count', 'mediaCount', 'imageSlots', 'slotCount', 'allowEmptyMediaSlots', 'allowSparseContent', 'allowMissingChart', 'allowMissingTable']);
  ['image', 'images', 'gallery', 'media', 'chart', 'charts', 'table', 'before', 'after', 'left', 'right'].forEach((key) => allowed.add(key));
  ['summaryTitle', 'leadTitle', 'focusTitle', 'conclusionTitle', 'takeawayTitle', 'footerSummaryTitle', 'nextStepTitle', 'agendaTitle', 'agendaSubtitle'].forEach((key) => allowed.add(key));
  (layoutGuide.slots || []).forEach((slot) => {
    const field = String(slot.field || '');
    if (!field.includes('.') && !field.includes('[]')) allowed.add(field);
  });
  const collections = ['items', 'sections', 'columns', 'steps', 'nodes', 'layers', 'lanes', 'metrics', 'agenda', 'captions', 'notes', 'insights', 'points'];
  collections.forEach((key) => allowed.add(key));
  Object.keys(slide).forEach((key) => {
    if (invalidContentScalars.has(key)) return;
    if (!allowed.has(key)) {
      errors.push('slide ' + (index + 1) + ' layout "' + layout + '" has unsupported plan field "' + key + '". Remove it or use a supported layout field.');
    }
  });
}

function validatePlanElementCounts(style, layout, layoutGuide, slide, index, errors) {
  planCountRules(style, layout, layoutGuide).forEach((rule) => validatePlanCountRule(slide, index, layout, rule, errors));
  if (['compare', 'duoCompare', 'splitCompare'].includes(layout)) {
    validatePlanCountRule(slide.before || {}, index, layout, { keys: ['items'], min: 1, max: 6, label: 'before items', prefix: 'before.' }, errors);
    validatePlanCountRule(slide.after || {}, index, layout, { keys: ['items'], min: 1, max: 6, label: 'after items', prefix: 'after.' }, errors);
  }
  validatePlanMediaAssetCounts(layout, slide, index, errors);
  validatePlanDataCounts(style, layout, slide, index, errors);
  validatePlanCmbBriefingCounts(style, layout, slide, index, errors);
}

function planCountRules(style, layout, layoutGuide) {
  const sideMax = style === 'swiss' ? 5 : style === 'cmb' ? 4 : 3;
  const cmbTextGridMax = style === 'cmb' ? 6 : 9;
  const rules = {
    bigNumbers: [{ keys: ['items'], min: 1, max: 6, label: 'number cards' }],
    kpiTower: [{ keys: ['items'], min: 1, max: 4, label: 'KPI cards' }],
    pipeline: [{ keys: ['steps', 'items'], min: 3, max: 6, label: 'pipeline steps' }],
    timeline: [{ keys: ['items', 'steps'], min: 3, max: 6, label: 'timeline steps' }],
    matrix: [{ keys: ['items'], min: 1, max: 12, label: 'matrix cells' }],
    fourCards: [{ keys: ['items'], min: style === 'cmb' ? 2 : 1, max: style === 'cmb' ? 6 : 8, label: style === 'cmb' ? 'CMB text weave cards' : 'cards' }],
    article: [{ keys: style === 'cmb' ? ['sections', 'items', 'columns', 'points', 'agenda'] : ['sections', 'items', 'columns'], min: style === 'cmb' ? 2 : 2, max: 6, label: style === 'cmb' ? 'briefing text blocks' : 'article sections' }],
    sectionList: [{ keys: style === 'cmb' ? ['sections', 'items', 'columns', 'points', 'agenda'] : ['sections', 'items', 'columns'], min: style === 'cmb' ? 2 : 3, max: style === 'cmb' ? 6 : 7, label: style === 'cmb' ? 'briefing text blocks' : 'section list items' }],
    briefing: [{ keys: ['sections', 'items', 'columns', 'points', 'agenda'], min: 2, max: 6, label: 'briefing text blocks' }],
    executiveBrief: [{ keys: ['sections', 'items', 'columns', 'points', 'agenda'], min: 2, max: 6, label: 'briefing text blocks' }],
    contentBrief: [{ keys: ['sections', 'items', 'columns', 'points', 'agenda'], min: 2, max: 6, label: 'briefing text blocks' }],
    textGrid: [{ keys: ['sections', 'items', 'columns'], min: style === 'cmb' ? 2 : 4, max: cmbTextGridMax, label: style === 'cmb' ? 'CMB text weave cards' : 'text grid cells' }],
    textWeave: [{ keys: ['sections', 'items', 'columns', 'points', 'agenda'], min: 2, max: 6, label: 'CMB text weave cards' }],
    contentSynthesis: [{ keys: ['sections', 'items', 'columns', 'points', 'agenda'], min: 2, max: 6, label: 'CMB text weave cards' }],
    denseText: [{ keys: ['sections', 'items', 'columns', 'points', 'agenda'], min: 2, max: 6, label: 'dense text blocks' }],
    agenda: [{ keys: style === 'cmb' ? ['items', 'sections', 'agenda'] : ['items'], min: 3, max: 8, label: 'agenda items' }],
    pyramid: [{ keys: ['layers', 'items', 'sections'], min: 3, max: 5, label: 'pyramid layers' }],
    radial: [{ keys: ['items', 'nodes', 'sections'], min: 4, max: 8, label: 'radial nodes' }],
    roadmap: [{ keys: ['steps', 'items'], min: 3, max: 6, label: 'roadmap steps' }],
    swimlane: [{ keys: ['lanes', 'sections'], min: 2, max: 4, label: 'swimlanes' }],
    media: [{ keys: ['items', 'insights', 'points'], min: 1, max: sideMax, label: 'side points' }],
    mediaGrid: [{ keys: ['captions', 'items', 'sections'], min: 1, max: 6, label: 'media captions' }],
    gallery: [{ keys: ['captions', 'items', 'sections'], min: 1, max: 6, label: 'media captions' }],
    imageGrid: [{ keys: ['captions', 'items', 'sections'], min: 1, max: 6, label: 'media captions' }],
    imageHero: [{ keys: ['items'], min: 1, max: 3, label: 'image hero metrics' }],
    caseStudy: [{ keys: ['metrics', 'items'], min: 1, max: 3, label: 'case metrics' }],
    dataSheet: [{ keys: ['notes', 'insights'], min: 0, max: style === 'swiss' ? 4 : 3, label: 'side notes' }],
    chart: [{ keys: ['insights', 'notes'], min: 0, max: style === 'swiss' ? 3 : 4, label: 'chart insights' }],
    dashboard: [{ keys: ['metrics', 'items'], min: 1, max: style === 'swiss' ? 5 : 4, label: 'dashboard metrics' }],
  };
  if (rules[layout]) return rules[layout];
  if (layoutGuide.collection?.keys?.length) {
    const collection = layoutGuide.collection;
    return [{ keys: collection.keys, min: collection.minItems || 0, max: collection.maxItems || 0, label: collection.label || 'items' }];
  }
  return [];
}

function validatePlanCountRule(source, index, layout, rule, errors) {
  const present = rule.keys.filter((key) => source[key] !== undefined && source[key] !== null);
  const label = `${rule.prefix || ''}${rule.label}`;
  if (present.length > 1) {
    errors.push('slide ' + (index + 1) + ' deck.plan.json provides multiple fields for ' + label + ': ' + present.map((key) => (rule.prefix || '') + key).join(', ') + '. Keep only one field so the planned element count is unambiguous.');
  }
  const key = present[0];
  const count = key ? planCollectionCount(source[key]) : planExplicitItemCount(source);
  if (!key && count === null && rule.min > 0) {
    errors.push('slide ' + (index + 1) + ' layout "' + layout + '" must declare planned ' + label + ' count before capacity-guide. Add a title-only array such as ' + rule.keys[0] + ': [{ "title": "..." }], or set itemCount/count to the intended number.');
    return;
  }
  if (count === null) return;
  if (rule.max && count > rule.max) {
    errors.push('slide ' + (index + 1) + ' deck.plan.json declares ' + count + ' ' + label + ', but layout "' + layout + '" supports at most ' + rule.max + '. Reduce the count, split into another slide, or choose a different layout before generating capacity-guide.');
  }
  if (rule.min && count < rule.min) {
    errors.push('slide ' + (index + 1) + ' deck.plan.json declares ' + count + ' ' + label + ', but layout "' + layout + '" expects at least ' + rule.min + '. Add elements or choose a simpler layout before generating capacity-guide.');
  }
}

function planCollectionCount(value) {
  if (value === undefined || value === null) return null;
  if (typeof value === 'string' || typeof value === 'number') return 1;
  if (Array.isArray(value)) return value.length;
  if (typeof value === 'object') return Object.keys(value).length;
  return null;
}

function planExplicitItemCount(source) {
  const raw = source.itemCount ?? source.count;
  if (raw === undefined || raw === null || raw === '') return null;
  const count = Number(raw);
  return Number.isFinite(count) ? count : null;
}

function validatePlanMediaAssetCounts(layout, slide, index, errors) {
  if (!planLayoutHasMediaSlot(layout)) return;
  const imageCount = planAssetCount(slide.image) + planAssetCount(slide.images) + planAssetCount(slide.media) + planAssetCount(slide.gallery);
  const chartCount = planAssetCount(slide.chart) + planAssetCount(slide.charts);
  const assetCount = Math.max(imageCount, chartCount);
  const explicitCount = Number(slide.mediaCount || slide.imageSlots || slide.slotCount || 0);
  if (layout === 'statement' && imageCount > 1) {
    errors.push('slide ' + (index + 1) + ' deck.plan.json uses statement with ' + imageCount + ' image asset(s), but statement supports exactly one image slot. Use mediaGrid/imageGrid or split into another slide.');
  }
  if (layout === 'statement' && chartCount) {
    errors.push('slide ' + (index + 1) + ' deck.plan.json uses statement with chart data, but statement reserves the media area for one image. Use chart/media layout instead.');
  }
  if (layout !== 'statement' && imageCount > 6) {
    errors.push('slide ' + (index + 1) + ' deck.plan.json declares ' + imageCount + ' image asset(s), but media layouts support at most 6. Split images across slides.');
  }
  if (layout !== 'statement' && chartCount > 6) {
    errors.push('slide ' + (index + 1) + ' deck.plan.json declares ' + chartCount + ' chart asset(s), but media layouts support at most 6. Split charts across slides.');
  }
  if (explicitCount && explicitCount !== assetCount) {
    errors.push('slide ' + (index + 1) + ' deck.plan.json declares mediaCount/imageSlots/slotCount=' + explicitCount + ' but provides ' + assetCount + ' image/chart asset(s). Make the media count match before generating capacity-guide.');
  }
}

function validatePlanDataCounts(style, layout, slide, index, errors) {
  if (layout === 'dashboard') {
    const charts = planAssetCount(slide.charts);
    if (!charts && !slide.allowMissingChart) errors.push('slide ' + (index + 1) + ' deck.plan.json uses dashboard but has no charts[]. Provide 1-2 planned charts before generating capacity-guide.');
    if (charts > 2) errors.push('slide ' + (index + 1) + ' deck.plan.json declares ' + charts + ' dashboard charts, but dashboard supports at most 2.');
  }
  if (layout === 'chart' && !hasPlanAssetValue(slide.chart) && !slide.allowMissingChart) {
    errors.push('slide ' + (index + 1) + ' deck.plan.json uses chart layout but has no chart. Provide chart data before generating capacity-guide.');
  }
  if (layout === 'dataSheet' && !hasPlanAssetValue(slide.table) && !slide.allowMissingTable) {
    errors.push('slide ' + (index + 1) + ' deck.plan.json uses dataSheet but has no table. Provide table.headers/table.rows before generating capacity-guide.');
  }
}

function validatePlanCmbBriefingCounts(style, layout, slide, index, errors) {
  if (style !== 'cmb' || !isCmbBriefingLayout(layout)) return;
  const sourceKey = firstCollectionKey(slide, ['sections', 'items', 'columns', 'points', 'agenda']);
  const itemCount = sourceKey ? planCollectionCount(slide[sourceKey]) : (planExplicitItemCount(slide) || 0);
  const hasLead = hasAnyKey(slide, ['summaryTitle', 'leadTitle', 'focusTitle']);
  const hasConclusion = hasAnyKey(slide, ['conclusionTitle', 'takeawayTitle', 'footerSummaryTitle', 'nextStepTitle']);
  const restCount = hasLead ? itemCount : Math.max(0, itemCount - 1);
  const maxRest = hasConclusion ? 4 : 5;
  if (restCount > maxRest) {
    errors.push('slide ' + (index + 1) + ' deck.plan.json declares ' + restCount + ' middle briefing text block(s), but layout "' + layout + '" supports at most ' + maxRest + (hasConclusion ? ' when conclusion/takeaway is present.' : '.') + ' Reduce item count or use textWeave/denseText before generating capacity-guide.');
  }
}

function planAssetCount(value) {
  if (value === undefined || value === null) return 0;
  if (typeof value === 'string') return value.trim() ? 1 : 0;
  if (typeof value === 'number' || typeof value === 'boolean') return 1;
  if (Array.isArray(value)) return value.filter((item) => hasPlanAssetValue(item)).length;
  if (typeof value === 'object') return 1;
  return 0;
}

function planLayoutHasMediaSlot(layout) {
  return ['statement', 'media', 'mediaGrid', 'gallery', 'imageGrid', 'imageHero', 'quoteImage', 'textImage', 'caseStudy'].includes(layout || '');
}

function planHasMediaOrChart(slide) {
  return ['image', 'images', 'media', 'gallery', 'chart', 'charts'].some((key) => hasPlanAssetValue(slide[key]));
}

function hasPlanAssetValue(value) {
  if (value === undefined || value === null) return false;
  if (typeof value === 'string') return value.trim().length > 0;
  if (typeof value === 'number' || typeof value === 'boolean') return true;
  if (Array.isArray(value)) return value.some((item) => hasPlanAssetValue(item));
  if (typeof value === 'object') return true;
  return false;
}

function plannedSlideCapacity(style, slide, index) {
  const layout = slide.layout || 'statement';
  if (style === 'cmb' && isCmbTextWeaveLayout(layout)) return plannedCmbTextWeaveCapacity(slide, index, layout);
  if (style === 'cmb' && isCmbBriefingLayout(layout)) return plannedCmbBriefingCapacity(slide, index, layout);
  return plannedGenericCapacity(style, slide, index, layout);
}

function plannedGenericCapacity(style, slide, index, layout) {
  const guide = guideForStyle(style);
  const layoutGuide = guide[layout] || COMMON_GUIDE[layout];
  const result = { slide: index + 1, layout, title: slide.title || '', slots: [], notes: [] };
  if (!layoutGuide || !Array.isArray(layoutGuide.slots)) {
    result.notes.push('No text capacity metadata exists for this layout. Keep text short and validate after generation.');
    return result;
  }
  const collection = layoutGuide.collection;
  const collectionKey = collection ? firstCollectionKey(slide, collection.keys) : null;
  const collectionItems = collectionKey ? normalizeItems(slide[collectionKey]) : [];
  if (collection && Array.isArray(layoutGuide.plannedCollectionFields)) {
    pushPlannedScalarSlots(result, layoutGuide, slide);
    const plannedKey = collectionKey || primaryCollectionKey(collection);
    const plannedItems = collectionItems.length ? collectionItems : plannedPlaceholderItems(slide, collection);
    if (!collectionItems.length) {
      result.notes.push('This layout requires ' + (collection.label || 'collection') + '. Fill the listed ' + plannedKey + '[n] field(s), or add a title-only ' + plannedKey + ' array in the plan JSON and regenerate capacity-guide for exact per-item ranges. Supported field(s): ' + collection.keys.join(', ') + '.');
    }
    pushPlannedCollectionSlots(result, layoutGuide, plannedKey, plannedItems, collection, layoutGuide.plannedCollectionFields);
    return result;
  }
  layoutGuide.slots.forEach((slot) => {
    const nestedMatch = slot.field.match(/^([^.[\]]+)\.([^.[\]]+)\[\]\.(.+)$/);
    if (nestedMatch) {
      const [, parentKey, collectionName, leaf] = nestedMatch;
      const nestedCollection = slide[parentKey] && typeof slide[parentKey] === 'object' ? slide[parentKey][collectionName] : null;
      const nestedItems = nestedCollection ? normalizeItems(nestedCollection) : plannedPlaceholderItems(slide, null);
      if (!nestedCollection) {
        result.notes.push('This layout requires ' + parentKey + '.' + collectionName + ' content. Fill the listed ' + parentKey + '.' + collectionName + '[n] field(s), or add title-only nested items in the plan JSON and regenerate capacity-guide.');
      }
      nestedItems.forEach((item, itemIndex) => {
        result.slots.push(adjustPlannedSlot(slot, parentKey + '.' + collectionName + '[' + itemIndex + '].' + leaf, itemTitle(item) || parentKey + ' ' + (itemIndex + 1), nestedItems.length, null));
      });
      return;
    }
    const match = slot.field.match(/^([^.[\]]+)\[\]\.(.+)$/);
    if (match) {
      const [, key, leaf] = match;
      const actualItems = !collection && slide[key] !== undefined && slide[key] !== null ? normalizeItems(slide[key]) : [];
      if (!collection && !actualItems.length) return;
      const plannedKey = collection ? (collectionKey || primaryCollectionKey(collection)) : key;
      const plannedItems = collectionItems.length ? collectionItems : (actualItems.length ? actualItems : plannedPlaceholderItems(slide, collection));
      if (key !== plannedKey) return;
      if (!collectionItems.length && collection) {
        if (!result.notes.some((note) => note.includes('requires ' + (collection.label || 'collection')))) {
          result.notes.push('This layout requires ' + (collection.label || 'collection') + '. Fill the listed ' + plannedKey + '[n] field(s), or add a title-only collection in the plan JSON and regenerate capacity-guide for exact per-item ranges. Supported field(s): ' + collection.keys.join(', ') + '.');
        }
      }
      plannedItems.forEach((item, itemIndex) => {
        result.slots.push(adjustPlannedSlot(slot, plannedKey + '[' + itemIndex + '].' + leaf, itemTitle(item) || (collection?.label || plannedKey) + ' ' + (itemIndex + 1), plannedItems.length, collection));
      });
      return;
    }
    if (isFillableScalarSlot(slot.field)) result.slots.push({ ...slot, title: plannedScalarTitle(slide, slot.field) });
  });
  return result;
}

function pushPlannedCollectionSlots(result, layoutGuide, collectionKey, items, collection, leaves) {
  items.forEach((item, itemIndex) => {
    leaves.forEach((leaf) => {
      const slot = plannedCollectionSlot(layoutGuide, collectionKey, leaf, collection);
      result.slots.push(adjustPlannedSlot(slot, collectionKey + '[' + itemIndex + '].' + leaf, itemTitle(item) || (collection.label || collectionKey) + ' ' + (itemIndex + 1), items.length, collection));
    });
  });
}

function pushPlannedScalarSlots(result, layoutGuide, slide) {
  const fields = Array.isArray(layoutGuide.plannedScalarFields) ? layoutGuide.plannedScalarFields : null;
  layoutGuide.slots.forEach((slot) => {
    if (!isFillableScalarSlot(slot.field)) return;
    if (fields && !fields.includes(slot.field)) return;
    result.slots.push({ ...slot, title: plannedScalarTitle(slide, slot.field) });
  });
}

function plannedCollectionSlot(layoutGuide, collectionKey, leaf, collection) {
  const slot = layoutGuide.slots.find((candidate) => candidate.field === collectionKey + '[].' + leaf)
    || layoutGuide.slots.find((candidate) => candidate.field.endsWith('[].' + leaf));
  if (slot) return slot;
  const isTitle = ['title', 'label', 'caption', 'name', 'value'].includes(leaf);
  return {
    field: collectionKey + '[].' + leaf,
    label: (collection.label || collectionKey) + ' ' + leaf,
    min: recommendedCapacityMin(isTitle ? collection.titleMax : collection.bodyMax),
    max: isTitle ? collection.titleMax : collection.bodyMax,
  };
}

function primaryCollectionKey(collection) {
  return Array.isArray(collection?.keys) && collection.keys.length ? collection.keys[0] : 'items';
}

function plannedPlaceholderItems(slide, collection) {
  const count = plannedPlaceholderCount(slide, collection);
  return Array.from({ length: count }, (_, index) => ({ title: (collection?.label || 'item') + ' ' + (index + 1) }));
}

function plannedPlaceholderCount(slide, collection) {
  const minItems = Math.max(1, Number(collection?.minItems || 1));
  const maxItems = Math.max(minItems, Number(collection?.maxItems || minItems));
  const explicit = Number(slide.mediaCount || slide.itemCount || slide.count || 0);
  const mediaCount = Math.max(
    arrayLength(slide.images),
    arrayLength(slide.gallery),
    arrayLength(slide.media),
    arrayLength(slide.charts),
    slide.image || slide.chart ? 1 : 0
  );
  const planned = Math.max(minItems, explicit || 0, mediaCount || 0);
  return Math.min(maxItems, planned);
}

function arrayLength(value) {
  return Array.isArray(value) ? value.length : 0;
}

function adjustPlannedSlot(slot, field, title, count, collection) {
  const isTitle = ['.title', '.label', '.name', '.value', '.unit', '.caption'].some((suffix) => field.endsWith(suffix));
  let max = slot.max;
  if (collection && !isTitle) {
    const factor = Math.max(0.72, Math.min(2.05, Number(collection.maxItems || count || 1) / Math.max(1, count || 1)));
    max = Math.max(1, Math.floor(max * factor));
  }
  return { field, label: slot.label, title, min: recommendedCapacityMin(max), max, note: plannedSlotNote(slot, count, collection) };
}

function recommendedCapacityMin(max) {
  return Math.max(1, Math.floor(Number(max || 0) * 0.6));
}

function plannedPointSlot(field, title, cap, label) {
  const max = Math.max(8, Math.floor(cap.max / Math.max(2, Math.min(4, cap.maxLines))));
  return {
    field,
    min: recommendedCapacityMin(max),
    max,
    label,
    title,
    note: 'Use only when content should be bullets. Total estimated lines across points must stay <= ' + cap.maxLines + '.',
  };
}

function plannedSlotNote(slot, count, collection) {
  const notes = [];
  if (slot.note) notes.push(slot.note);
  if (collection && count) notes.push('Calculated for ' + count + ' planned ' + (collection.label || 'item') + '(s). Regenerate if count changes.');
  return notes.join(' ');
}

function isFillableScalarSlot(field) {
  return !field.includes('[]') && !['title', 'subtitle', 'kicker', 'agendaTitle', 'agendaSubtitle'].includes(field);
}

function plannedScalarTitle(slide, field) {
  const titleKey = field + 'Title';
  return slide[titleKey] || slide.title || '';
}

function plannedSlideExample(sourceSlide, planned) {
  const example = {};
  example.layout = planned.layout || sourceSlide.layout || 'statement';
  if (sourceSlide.title || planned.title) example.title = sourceSlide.title || planned.title;
  if (sourceSlide.kicker) example.kicker = sourceSlide.kicker;
  if (sourceSlide.subtitle) example.subtitle = sourceSlide.subtitle;
  copyExamplePlanFields(sourceSlide, example);
  (planned.slots || []).forEach((slot) => {
    if (!slot?.field || slot.field === '-') return;
    assignExampleField(example, slot.field, exampleValueForSlot(slot));
  });
  return example;
}

function copyExamplePlanFields(source, target) {
  ['style', 'theme', 'mediaCount', 'imageSlots', 'slotCount', 'allowEmptyMediaSlots', 'allowSparseContent', 'allowMissingChart', 'allowMissingTable'].forEach((key) => {
    if (source[key] !== undefined) target[key] = source[key];
  });
  ['image', 'images', 'gallery', 'media', 'chart', 'charts', 'table'].forEach((key) => {
    if (source[key] !== undefined) target[key] = cloneJson(source[key]);
  });
  ['sections', 'items', 'columns', 'steps', 'nodes', 'layers', 'lanes', 'metrics', 'agenda', 'captions', 'notes', 'insights'].forEach((key) => {
    if (source[key] !== undefined) target[key] = collectionSkeleton(source[key]);
  });
  ['before', 'after', 'left', 'right'].forEach((key) => {
    if (source[key] && typeof source[key] === 'object' && !Array.isArray(source[key])) target[key] = nestedPlanSkeleton(source[key]);
  });
}

function assignExampleField(target, field, value) {
  const parts = String(field || '').split('.');
  let current = target;
  parts.forEach((part, index) => {
    const last = index === parts.length - 1;
    const arrayMatch = part.match(/^(.+)\[(\d+)\]$/);
    const pushArrayMatch = part.match(/^(.+)\[\]$/);
    if (arrayMatch || pushArrayMatch) {
      const key = (arrayMatch || pushArrayMatch)[1];
      const itemIndex = arrayMatch ? Number(arrayMatch[2]) : 0;
      if (!Array.isArray(current[key])) current[key] = [];
      while (current[key].length <= itemIndex) current[key].push({});
      if (last) {
        current[key][itemIndex] = value;
      } else {
        if (!current[key][itemIndex] || typeof current[key][itemIndex] !== 'object') current[key][itemIndex] = {};
        current = current[key][itemIndex];
      }
      return;
    }
    if (last) {
      if (current[part] === undefined) current[part] = value;
      return;
    }
    if (!current[part] || typeof current[part] !== 'object' || Array.isArray(current[part])) current[part] = {};
    current = current[part];
  });
}

function exampleValueForSlot(slot) {
  const field = String(slot.field || '');
  const label = String(slot.label || '').toLowerCase();
  const title = String(slot.title || slot.label || '').trim();
  if (field.endsWith('points[]')) return samplePointText(title);
  if (field.endsWith('.value') || field === 'value') return 'XX';
  if (field.endsWith('.unit') || field === 'unit') return '%';
  if (field.endsWith('.label') || field === 'label') return title && !/^\w+\s+\d+$/.test(title) ? title : '指标';
  if (field.endsWith('.caption') || field === 'caption') return '图片说明';
  if (field.endsWith('.title') || field === 'caseTitle' || label.includes('title')) return title && !/^\w+\s+\d+$/.test(title) ? title : '小标题';
  if (field === 'quote') return '填写核心引述。';
  if (field === 'conclusion' || field === 'takeaway') return '填写结论或下一步行动。';
  if (field === 'summary' || field === 'body' || field === 'lead') return '填写本页核心正文，长度参考容量范围。';
  return '填写内容，长度参考容量范围。';
}

function samplePointText(title) {
  return (title && !/^\w+\s+\d+$/.test(title) ? title + '分点内容' : '分点内容');
}

function cloneJson(value) {
  if (value === undefined) return value;
  return JSON.parse(JSON.stringify(value));
}

function collectionSkeleton(value) {
  if (Array.isArray(value)) return value.map((item) => itemSkeleton(item));
  if (value && typeof value === 'object') {
    return Object.entries(value).map(([title, body]) => ({ title, ...(typeof body === 'string' && body.trim() ? { body } : {}) }));
  }
  if (typeof value === 'string' || typeof value === 'number') return [{ title: String(value) }];
  return cloneJson(value);
}

function itemSkeleton(item) {
  if (item === undefined || item === null) return {};
  if (typeof item === 'string' || typeof item === 'number') return { title: String(item) };
  if (typeof item !== 'object') return {};
  const next = {};
  ['title', 'label', 'name', 'heading', 'value', 'unit', 'caption', 'year', 'number', 'icon', 'highlight'].forEach((key) => {
    if (item[key] !== undefined) next[key] = cloneJson(item[key]);
  });
  return next;
}

function nestedPlanSkeleton(value) {
  const next = itemSkeleton(value);
  ['items', 'sections', 'points', 'bullets', 'list'].forEach((key) => {
    if (value[key] !== undefined) next[key] = collectionSkeleton(value[key]);
  });
  return next;
}

function plannedCmbTextWeaveCapacity(slide, index, layout) {
  const sourceKey = firstCollectionKey(slide, ['sections', 'items', 'columns', 'points', 'agenda']);
  const plannedKey = sourceKey || 'items';
  const items = sourceKey ? normalizeItems(slide[sourceKey]).slice(0, 6) : plannedPlaceholderItems(slide, { keys: ['items'], label: 'CMB text weave cards', minItems: 2, maxItems: 6 });
  const result = { slide: index + 1, layout, title: slide.title || '', slots: [], notes: [] };
  if (!sourceKey) {
    result.notes.push('This CMB layout requires at least 2 text cards. Fill the listed items[n] fields, or add a title-only sections/items/columns array in the plan JSON and regenerate capacity-guide for exact per-card ranges.');
  }
  if (items.length < 2) {
    result.notes.push('CMB ' + layout + ' needs at least 2 text cards so the right side is not empty. Add one more title-only item before writing body text.');
  }
  cmbTextWeaveCardBoxes(items.length, !!slide.subtitle).forEach(({ box, options }, itemIndex) => {
    const item = items[itemIndex];
    if (!item) return;
    const title = itemTitle(item) || String(itemIndex + 1).padStart(2, '0');
    const cap = cmbCardBodyCapacity(box, title, options);
    result.slots.push({ field: plannedKey + '[' + itemIndex + '].body', min: recommendedCapacityMin(cap.max), max: cap.max, label: 'card body', title, note: 'Calculated from actual card ' + (itemIndex + 1) + '/' + items.length + ': max ' + cap.maxLines + ' line(s), about ' + Math.floor(cap.charsPerLine) + ' CJK chars/line at 12pt.' });
    result.slots.push(plannedPointSlot(plannedKey + '[' + itemIndex + '].points[]', title, cap, 'explicit bullet point'));
  });
  return result;
}

function plannedCmbBriefingCapacity(slide, index, layout) {
  const sourceKey = firstCollectionKey(slide, ['sections', 'items', 'columns', 'points', 'agenda']);
  const plannedKey = sourceKey || 'items';
  const items = sourceKey ? normalizeItems(slide[sourceKey]).slice(0, 6) : plannedPlaceholderItems(slide, { keys: ['items'], label: 'briefing text blocks', minItems: 2, maxItems: 6 });
  const hasLead = hasAnyKey(slide, ['summary', 'body', 'lead', 'summaryTitle', 'leadTitle', 'focusTitle']);
  const hasConclusion = hasAnyKey(slide, ['conclusion', 'takeaway', 'footerSummary', 'nextStep', 'conclusionTitle', 'takeawayTitle', 'footerSummaryTitle', 'nextStepTitle']);
  const pseudo = { ...slide, summary: hasLead ? '__planned__' : '', conclusion: hasConclusion ? '__planned__' : '' };
  const entries = cmbBriefingCardEntries(pseudo, items.length);
  const result = { slide: index + 1, layout, title: slide.title || '', slots: [], notes: [] };
  if (!sourceKey && !hasLead) {
    result.notes.push('This CMB briefing layout requires briefing text blocks. Fill the listed items[n] fields, or add a title-only sections/items/columns array in the plan JSON and regenerate capacity-guide for exact card ranges.');
  }
  const plannedRestCount = hasLead ? items.length : Math.max(0, items.length - 1);
  const maxRestCount = hasConclusion ? 4 : 5;
  if (plannedRestCount > maxRestCount) {
    result.notes.push('This CMB briefing plan has ' + plannedRestCount + ' middle text block(s), but the layout supports at most ' + maxRestCount + (hasConclusion ? ' when conclusion/takeaway is present.' : '.') + ' Reduce item count or use textWeave/denseText.');
  }
  const leadEntry = entries.find((entry) => entry.kind === 'lead');
  if (leadEntry) {
    const title = slide.summaryTitle || slide.leadTitle || slide.focusTitle || (hasLead ? '摘要' : itemTitle(items[0]) || '摘要');
    const cap = cmbCardBodyCapacity(leadEntry.box, title, leadEntry.options);
    const field = hasLead ? 'summary' : plannedKey + '[0].body';
    result.slots.push({ field, min: recommendedCapacityMin(cap.max), max: cap.max, label: hasLead ? 'top summary body' : 'lead card body', title, note: 'Calculated from lead card: max ' + cap.maxLines + ' line(s), about ' + Math.floor(cap.charsPerLine) + ' CJK chars/line at 12pt.' });
    if (!hasLead) result.slots.push(plannedPointSlot(plannedKey + '[0].points[]', title, cap, 'lead card bullet point'));
  }
  const rest = hasLead ? items : items.slice(1);
  entries.filter((entry) => entry.kind === 'rest').forEach((entry, i) => {
    const sourceIndex = hasLead ? i : i + 1;
    const item = rest[i];
    if (!item) return;
    const title = itemTitle(item) || '分析' + (i + 1);
    const cap = cmbCardBodyCapacity(entry.box, title, entry.options);
    result.slots.push({ field: plannedKey + '[' + sourceIndex + '].body', min: recommendedCapacityMin(cap.max), max: cap.max, label: 'analysis card body', title, note: 'Calculated from actual card ' + (i + 1) + '/' + rest.length + ': max ' + cap.maxLines + ' line(s), about ' + Math.floor(cap.charsPerLine) + ' CJK chars/line at 12pt.' });
    result.slots.push(plannedPointSlot(plannedKey + '[' + sourceIndex + '].points[]', title, cap, 'analysis bullet point'));
  });
  const conclusionEntry = entries.find((entry) => entry.kind === 'conclusion');
  if (conclusionEntry) {
    const title = slide.conclusionTitle || slide.takeawayTitle || slide.footerSummaryTitle || slide.nextStepTitle || '结论';
    const cap = cmbCardBodyCapacity(conclusionEntry.box, title, conclusionEntry.options);
    result.slots.push({ field: 'conclusion', min: recommendedCapacityMin(cap.max), max: cap.max, label: 'bottom conclusion body', title, note: 'Calculated from bottom card: max ' + cap.maxLines + ' line(s), about ' + Math.floor(cap.charsPerLine) + ' CJK chars/line at 12pt.' });
  }
  return result;
}

function hasAnyKey(source, keys) {
  return keys.some((key) => Object.prototype.hasOwnProperty.call(source, key));
}
function warnSpecTextCapacity(spec) {
  const style = spec.style || 'swiss';
  const guide = guideForStyle(style);
  (spec.slides || []).forEach((slide, index) => {
    const layout = slide.layout || 'statement';
    if (style === 'cmb' && isCmbTextWeaveLayout(layout)) {
      warnCmbTextWeaveCapacity(slide, index, layout);
      return;
    }
    if (style === 'cmb' && isCmbBriefingLayout(layout)) {
      warnCmbBriefingTextCapacity(slide, index, layout);
      return;
    }
    const layoutGuide = guide[layout] || COMMON_GUIDE[layout];
    if (!layoutGuide || !Array.isArray(layoutGuide.slots)) return;
    layoutGuide.slots.forEach((slot) => warnSlot(slide, index, layout, slot));
  });
}

function warnSlot(slide, index, layout, slot) {
  readFieldValues(slide, slot.field).forEach(({ path: fieldPath, value }) => {
    const text = textValue(value);
    if (!text) return;
    const actual = Math.ceil(textVisualLength(text));
    if (actual > slot.max) {
      const preview = text.replace(/\s+/g, ' ').slice(0, 54);
      console.warn('Warning: slide ' + (index + 1) + ' layout "' + layout + '" field ' + fieldPath + ' has ' + actual + ' visual chars; recommended ' + slot.min + '-' + slot.max + '. Shorten this field in JSON and regenerate the PPTX: ' + preview);
    }
  });
}


function isCmbBriefingLayout(layout) {
  return ['article', 'sectionList', 'briefing', 'executiveBrief', 'contentBrief'].includes(layout);
}

function warnCmbBriefingTextCapacity(slide, index, layout) {
  const sourceKey = firstCollectionKey(slide, ['sections', 'items', 'columns', 'points', 'agenda']);
  const items = sourceKey ? normalizeItems(slide[sourceKey]).slice(0, 6) : [];
  const hasLead = !!(slide.summary || slide.body || slide.lead);
  const entries = cmbBriefingCardEntries(slide, items.length);
  const leadEntry = entries.find((entry) => entry.kind === 'lead');
  if (leadEntry) {
    const leadTitle = slide.summaryTitle || slide.leadTitle || slide.focusTitle || slide.kicker || itemTitle(items[0]) || '摘要';
    const leadPoints = hasLead ? [] : itemPoints(items[0]);
    const leadText = hasLead ? (slide.summary || slide.body || slide.lead) : itemBody(items[0]);
    warnCmbCardText(index, layout, hasLead ? 'summary/body/lead' : `${sourceKey || 'items'}[0].${leadPoints.length ? 'points' : 'body'}`, leadPoints.length ? leadPoints : leadText, leadEntry.box, leadTitle, leadEntry.options);
  }
  const rest = hasLead ? items : items.slice(1);
  entries.filter((entry) => entry.kind === 'rest').forEach((entry, i) => {
    const item = rest[i];
    if (!item) return;
    const sourceIndex = hasLead ? i : i + 1;
    const points = itemPoints(item);
    warnCmbCardText(index, layout, `${sourceKey || 'items'}[${sourceIndex}].${points.length ? 'points' : 'body'}`, points.length ? points : itemBody(item), entry.box, itemTitle(item) || `分析${i + 1}`, entry.options);
  });
  const conclusionEntry = entries.find((entry) => entry.kind === 'conclusion');
  if (conclusionEntry) {
    const title = slide.conclusionTitle || slide.takeawayTitle || slide.footerSummaryTitle || slide.nextStepTitle || '结论';
    warnCmbCardText(index, layout, 'conclusion/takeaway', slide.conclusion || slide.takeaway || slide.footerSummary || slide.nextStep, conclusionEntry.box, title, conclusionEntry.options);
  }
}

function cmbBriefingCardEntries(slide, itemCount) {
  const y0 = slide.subtitle ? 2.78 : 2.45;
  const hasLead = !!(slide.summary || slide.body || slide.lead);
  const conclusionText = slide.conclusion || slide.takeaway || slide.footerSummary || slide.nextStep;
  const leadH = 1.46;
  const leadGap = 0.22;
  const entries = [{ kind: 'lead', box: { x: 0.78, y: y0, w: 11.45, h: leadH }, options: { lead: true, accent: true, titleFontSize: 13.2 } }];
  const restCount = hasLead ? itemCount : Math.max(0, itemCount - 1);
  const midY = y0 + leadH + leadGap;
  const conclusionBox = { x: 0.78, y: 5.62, w: 11.45, h: 0.86 };
  const midBottom = conclusionText ? conclusionBox.y - 0.2 : 6.48;
  if (restCount <= 4) {
    const gap = 0.28;
    const w = (11.45 - gap * Math.max(0, restCount - 1)) / Math.max(1, restCount);
    const h = Math.max(1.15, midBottom - midY);
    for (let i = 0; i < restCount; i += 1) entries.push({ kind: 'rest', box: { x: 0.78 + i * (w + gap), y: midY, w, h }, options: {} });
  } else {
    const gapX = 0.28;
    const gapY = 0.18;
    const cols = 3;
    const rows = Math.ceil(restCount / cols);
    const w = (11.45 - gapX * (cols - 1)) / cols;
    const h = Math.max(0.95, (midBottom - midY - gapY * (rows - 1)) / rows);
    for (let i = 0; i < restCount; i += 1) entries.push({ kind: 'rest', box: { x: 0.78 + (i % cols) * (w + gapX), y: midY + Math.floor(i / cols) * (h + gapY), w, h }, options: {} });
  }
  if (conclusionText) entries.push({ kind: 'conclusion', box: conclusionBox, options: { compact: true, accent: true, titleFontSize: 13.2 } });
  return entries;
}

function warnCmbCardText(index, layout, field, textOrPoints, box, title, options = {}) {
  const points = Array.isArray(textOrPoints) ? textOrPoints.filter(Boolean) : [];
  const text = points.length ? points.join('\n') : String(textOrPoints || '');
  if (!text) return;
  const cap = cmbCardBodyCapacity(box, title, options);
  const lineInfo = estimateCmbBodyLines(points.length ? points : text, cap.contentW, cap.fontSize, { bullet: points.length > 1 });
  if (lineInfo.lines > cap.maxLines) {
    const preview = text.replace(/\s+/g, ' ').slice(0, 54);
    console.warn('Warning: slide ' + (index + 1) + ' layout "' + layout + '" field ' + field + ' may overflow card body; estimated ' + lineInfo.lines + ' line(s), available ' + cap.maxLines + ' at 12pt Microsoft YaHei. Shorten points/body in JSON, split the card, or use fewer cards: ' + preview);
  }
}

function isCmbTextWeaveLayout(layout) {
  return ['textGrid', 'fourCards', 'textWeave', 'contentSynthesis', 'denseText'].includes(layout);
}

function warnCmbTextWeaveCapacity(slide, index, layout) {
  const sourceKey = firstCollectionKey(slide, ['sections', 'items', 'columns', 'points', 'agenda']);
  if (!sourceKey) return;
  const items = normalizeItems(slide[sourceKey]).slice(0, 6);
  if (!items.length) return;
  cmbTextWeaveCardBoxes(items.length, !!slide.subtitle).forEach(({ box, options }, itemIndex) => {
    const item = items[itemIndex];
    if (!item) return;
    const title = itemTitle(item) || String(itemIndex + 1).padStart(2, '0');
    const body = itemBody(item);
    const points = itemPoints(item);
    const text = points.length ? points.join('\n') : body;
    if (!text) return;
    const cap = cmbCardBodyCapacity(box, title, options);
    const lineInfo = estimateCmbBodyLines(points.length ? points : text, cap.contentW, cap.fontSize, { bullet: points.length > 1 });
    if (lineInfo.lines > cap.maxLines) {
      const field = sourceKey + '[' + itemIndex + '].' + (points.length ? 'points' : 'body');
      const preview = text.replace(/\s+/g, ' ').slice(0, 54);
      console.warn('Warning: slide ' + (index + 1) + ' layout "' + layout + '" field ' + field + ' may overflow card body; estimated ' + lineInfo.lines + ' line(s), available ' + cap.maxLines + ' at 12pt Microsoft YaHei. Shorten points/body in JSON, split it, or use fewer cards: ' + preview);
    }
  });
}

function firstCollectionKey(slide, keys) {
  return keys.find((key) => slide[key] !== undefined && slide[key] !== null);
}

function cmbTextWeaveCardBoxes(count, hasSubtitle) {
  const y0 = hasSubtitle ? 2.78 : 2.42;
  const bottom = 6.48;
  if (count <= 5) {
    const leadW = 3.85;
    const gapX = 0.28;
    const boxes = [{ box: { x: 0.78, y: y0, w: leadW, h: bottom - y0 }, options: { lead: true } }];
    const right = count - 1;
    const rightX = 0.78 + leadW + gapX;
    const rightW = 11.45 - leadW - gapX;
    cmbTextWeaveRightBoxes(right, rightX, y0, rightW, bottom - y0, gapX, 0.24)
      .forEach((box) => boxes.push({ box, options: {} }));
    return boxes;
  }
  const boxes = [{ box: { x: 0.78, y: y0, w: 11.45, h: 0.9 }, options: { lead: true, compact: true } }];
  const rest = count - 1;
  const gridY = y0 + 1.12;
  const leftW = 3.72;
  const gapX = 0.28;
  boxes.push({ box: { x: 0.78, y: gridY, w: leftW, h: bottom - gridY }, options: {} });
  const cardW = (11.45 - leftW - gapX * 2) / 2;
  const rowGap = 0.2;
  const cardH = (bottom - gridY - rowGap) / 2;
  for (let i = 1; i < rest; i += 1) {
    boxes.push({ box: { x: 0.78 + leftW + gapX + ((i - 1) % 2) * (cardW + gapX), y: gridY + Math.floor((i - 1) / 2) * (cardH + rowGap), w: cardW, h: cardH }, options: {} });
  }
  return boxes;
}

function cmbTextWeaveRightBoxes(count, x, y, w, h, gapX, rowGap) {
  const halfW = (w - gapX) / 2;
  const halfH = (h - rowGap) / 2;
  if (count <= 0) return [];
  if (count === 1) return [{ x, y, w, h }];
  if (count === 2) return [
    { x, y, w: halfW, h },
    { x: x + halfW + gapX, y, w: halfW, h },
  ];
  if (count === 3) return [
    { x, y, w: halfW, h },
    { x: x + halfW + gapX, y, w: halfW, h: halfH },
    { x: x + halfW + gapX, y: y + halfH + rowGap, w: halfW, h: halfH },
  ];
  return Array.from({ length: Math.min(count, 4) }, (_, i) => ({
    x: x + (i % 2) * (halfW + gapX),
    y: y + Math.floor(i / 2) * (halfH + rowGap),
    w: halfW,
    h: halfH,
  }));
}

function cmbCardBodyCapacity(box, title, options = {}) {
  const padX = options.compact ? 0.2 : 0.26;
  const padTop = options.compact ? 0.12 : 0.18;
  const contentW = box.w - padX * 2;
  const titleW = Math.max(0.3, contentW - 0.38);
  const titleFont = options.titleFontSize || (options.lead ? 15.2 : options.compact ? 12.2 : 13.2);
  const titleH = estimateTextHeight(title, titleW, titleFont, { min: 0.26, max: options.compact ? 0.34 : 0.5, lineHeight: 1.14, padding: 0.02 });
  const bodyY = box.y + padTop + titleH + 0.13;
  const bodyH = roundCapacityDimension(Math.max(0.24, box.y + box.h - bodyY - (options.compact ? 0.12 : 0.18)));
  return estimatedChineseCapacity(contentW, bodyH, 12, 0.02);
}

function estimatedChineseCapacity(w, h, fontSize, margin = 0) {
  const boxW = roundCapacityDimension(Math.max(0.05, w - margin * 2));
  const boxH = roundCapacityDimension(Math.max(0.05, h - margin * 2));
  const charsPerLine = Math.max(1, (boxW * 72) / fontSize);
  const maxLines = Math.max(1, Math.floor((boxH * 72) / (fontSize * 1.12)));
  return { max: Math.max(1, Math.floor(charsPerLine * maxLines * 0.99)), maxLines, charsPerLine, contentW: boxW, bodyH: boxH, fontSize };
}

function estimateCmbBodyLines(textOrPoints, w, fontSize, options = {}) {
  const items = Array.isArray(textOrPoints) ? textOrPoints : String(textOrPoints || '').split(/\r?\n/);
  const bulletIndent = options.bullet ? 0.38 : 0;
  const effectiveW = Math.max(0.05, Number(w || 0) - bulletIndent);
  const charsPerLine = Math.max(1, (effectiveW * 72) / Math.max(1, fontSize));
  const lines = items.reduce((sum, item) => {
    const raw = String(item || '').trim();
    if (!raw) return sum;
    return sum + raw.split(/\r?\n/).reduce((partSum, line) => partSum + Math.max(1, Math.ceil(textVisualLength(line) / charsPerLine)), 0);
  }, 0);
  return { lines: Math.max(0, lines), charsPerLine };
}

function roundCapacityDimension(value) {
  return Math.round(Number(value) * 1000) / 1000;
}

function estimateTextHeight(text, boxW, fontSize, options = {}) {
  const raw = String(text || '').trim();
  if (!raw) return options.empty ?? 0;
  const charsPerLine = Math.max(4, (boxW * 72) / Math.max(1, fontSize));
  const lines = raw.split(/\r?\n/).reduce((sum, line) => sum + Math.max(1, Math.ceil(textVisualLength(line) / charsPerLine)), 0);
  const height = (lines * fontSize * (options.lineHeight || 1.18)) / 72 + (options.padding || 0.04);
  return Math.max(options.min ?? 0.2, Math.min(options.max ?? 10, height));
}

function itemTitle(item) {
  return String(item?.title || item?.label || item?.name || item?.heading || '').trim();
}

function itemBody(item) {
  if (!item || typeof item !== 'object') return typeof item === 'string' || typeof item === 'number' ? String(item) : '';
  return String(item.body || item.desc || item.note || item.summary || item.detail || item.text || item.story || '').trim();
}

function itemPoints(item) {
  const raw = item?.points || item?.bullets || item?.list;
  if (!Array.isArray(raw)) return [];
  return raw.map((point) => {
    if (point && typeof point === 'object') return String(point.body || point.text || point.title || point.label || '').trim();
    return String(point || '').trim();
  }).filter(Boolean);
}
function readFieldValues(source, field) {
  const parts = String(field || '').split('.');
  const results = [];
  walk(source, parts, [], results);
  return results;
}

function walk(current, parts, pathParts, results) {
  if (!parts.length) {
    results.push({ path: pathParts.join('.'), value: current });
    return;
  }
  const part = parts[0];
  const isArray = part.endsWith('[]');
  const key = isArray ? part.slice(0, -2) : part;
  if (!current || typeof current !== 'object' || current[key] === undefined || current[key] === null) return;
  const next = current[key];
  if (isArray) {
    const items = Array.isArray(next) ? next : normalizeItems(next);
    if (parts.length === 1) items.forEach((item, i) => results.push({ path: pathParts.concat(key + '[' + i + ']').join('.'), value: item }));
    else normalizeItems(next).forEach((item, i) => walk(item, parts.slice(1), pathParts.concat(key + '[' + i + ']'), results));
  }
  else walk(next, parts.slice(1), pathParts.concat(key), results);
}

function normalizeItems(value) {
  if (value === undefined || value === null) return [];
  if (typeof value === 'string' || typeof value === 'number') return [{ body: String(value) }];
  if (Array.isArray(value)) return value.map((item) => typeof item === 'string' || typeof item === 'number' ? { body: String(item) } : item).filter(Boolean);
  if (typeof value === 'object') return Object.entries(value).map(([title, body]) => ({ title, body: String(body ?? '') }));
  return [];
}

function textValue(value) {
  if (value === undefined || value === null) return '';
  if (Array.isArray(value)) return value.map(textValue).filter(Boolean).join(' ');
  if (typeof value === 'object') return String(value.body || value.text || value.title || value.label || '').trim();
  return String(value).trim();
}

function textVisualLength(text) {
  return String(text || '').split('').reduce((sum, ch) => {
    if (/\s/.test(ch)) return sum + 0.25;
    return /[\u2E80-\u9FFF\uF900-\uFAFF]/.test(ch) ? sum + 1 : sum + 0.56;
  }, 0);
}

module.exports = { layoutCapacityGuide, layoutCapacityGuideForSpec, layoutCapacityMarkdown, layoutCapacityMarkdownForSpec, writeLayoutCapacityGuide, writePlannedCapacityGuide, validateCapacityPlan, warnSpecTextCapacity, textVisualLength };
