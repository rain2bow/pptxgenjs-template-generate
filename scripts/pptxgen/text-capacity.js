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
  media: slots([['title', 'page title', 8, 32], ['body', 'media body', 20, 80], ['items[].title', 'side point title', 4, 14], ['items[].body', 'side point body', 12, 45], ['insights[].body', 'side insight body', 12, 45], ['points[].body', 'side point body', 12, 45]]),
  chart: slots([['title', 'page title', 8, 32], ['insights[].title', 'chart insight title', 4, 14], ['insights[].body', 'chart insight body', 12, 48], ['notes[].body', 'chart note body', 12, 48]]),
  dataSheet: slots([['title', 'page title', 8, 32], ['notes[].title', 'side note title', 4, 14], ['notes[].body', 'side note body', 10, 42], ['insights[].body', 'side insight body', 10, 42]]),
  dashboard: slots([['title', 'page title', 8, 32], ['metrics[].label', 'metric label', 2, 10], ['metrics[].value', 'metric value', 1, 8]]),
  bigNumbers: slots([['title', 'page title', 8, 32], ['items[].label', 'number label', 2, 12], ['items[].value', 'number value', 1, 8], ['items[].note', 'number note', 8, 32]]),
  kpiTower: slots([['title', 'page title', 8, 32], ['items[].label', 'KPI label', 2, 12], ['items[].value', 'KPI value', 1, 8], ['items[].note', 'KPI note', 8, 32]]),
  compare: slots([['title', 'page title', 8, 32], ['before.title', 'left title', 4, 18], ['after.title', 'right title', 4, 18], ['before.items[].body', 'left item body', 8, 36], ['after.items[].body', 'right item body', 8, 36]]),
  duoCompare: null,
  splitCompare: null,
  caseStudy: slots([['title', 'page title', 8, 32], ['caseTitle', 'case title', 6, 26], ['body', 'case body', 25, 90], ['metrics[].note', 'case metric note', 6, 28]]),
};
COMMON_GUIDE.duoCompare = COMMON_GUIDE.compare;
COMMON_GUIDE.splitCompare = COMMON_GUIDE.compare;

const CMB_GUIDE = {
  ...COMMON_GUIDE,
  article: cmbBriefingSlots('article/briefing'),
  sectionList: cmbBriefingSlots('sectionList/briefing'),
  agenda: cmbBriefingSlots('agenda/briefing'),
  briefing: cmbBriefingSlots('briefing'),
  executiveBrief: cmbBriefingSlots('executiveBrief'),
  contentBrief: cmbBriefingSlots('contentBrief'),
  textGrid: cmbTextWeaveSlots('textGrid/textWeave'),
  fourCards: cmbTextWeaveSlots('fourCards/textWeave'),
  textWeave: cmbTextWeaveSlots('textWeave'),
  contentSynthesis: cmbTextWeaveSlots('contentSynthesis'),
  denseText: cmbTextWeaveSlots('denseText'),
};

function slots(items, description = '') {
  return { description, slots: items.map(([field, label, min, max, note]) => {
    const adjustedMax = Math.max(min, Math.floor(max * CAPACITY_SCALE_FOR_UNIFIED_TYPOGRAPHY));
    return { field, label, min: Math.min(min, adjustedMax), max: adjustedMax, note };
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
  return slots(fields, minItems + '-' + maxItems + ' items; keep every item title + body within range.');
}

function cmbBriefingSlots(name) {
  return slots([
    ['title', 'page title', 8, 32], ['summary', 'top summary body', 35, 90, 'Use body/lead if summary is omitted.'], ['body', 'top summary body alias', 35, 90], ['lead', 'top summary body alias', 35, 90],
    ['sections[].title', 'analysis card title', 4, 14], ['sections[].body', 'analysis card body', 18, 58], ['items[].title', 'analysis card title', 4, 14], ['items[].body', 'analysis card body', 18, 58], ['agenda[].title', 'analysis card title', 4, 14], ['agenda[].body', 'analysis card body', 18, 58],
    ['conclusion', 'bottom conclusion body', 12, 48], ['takeaway', 'bottom conclusion body alias', 12, 48],
  ], name + ': CMB top summary + middle analysis cards + bottom takeaway.');
}

function cmbTextWeaveSlots(name) {
  return slots([
    ['title', 'page title', 8, 32], ['sections[].title', 'text card title', 4, 14], ['sections[].body', 'text card body', 15, 52], ['items[].title', 'text card title', 4, 14], ['items[].body', 'text card body', 15, 52], ['columns[].title', 'text card title', 4, 14], ['columns[].body', 'text card body', 15, 52],
  ], name + ': CMB asymmetric 1-6 text cards; 5-6 cards need shorter bodies.');
}

function guideForStyle(style = 'swiss') {
  return style === 'cmb' ? CMB_GUIDE : COMMON_GUIDE;
}

function layoutCapacityGuide(style = 'swiss') {
  return {
    style,
    unit: 'visual characters; CJK counts as 1, Latin letters count as about 0.56',
    instruction: 'Use these ranges before writing JSON. Ranges are calibrated for unified typography: Microsoft YaHei / Times New Roman with 36, 28, 16, 14, 12 pt tiers. If generation warns that a field exceeds max, shorten that field in JSON and regenerate the PPTX.',
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

function warnSpecTextCapacity(spec) {
  const style = spec.style || 'swiss';
  const guide = guideForStyle(style);
  (spec.slides || []).forEach((slide, index) => {
    const layout = slide.layout || 'statement';
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
  if (isArray) normalizeItems(next).forEach((item, i) => walk(item, parts.slice(1), pathParts.concat(key + '[' + i + ']'), results));
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
  if (typeof value === 'object') return '';
  return String(value).trim();
}

function textVisualLength(text) {
  return String(text || '').split('').reduce((sum, ch) => {
    if (/\s/.test(ch)) return sum + 0.25;
    return /[\u2E80-\u9FFF\uF900-\uFAFF]/.test(ch) ? sum + 1 : sum + 0.56;
  }, 0);
}

module.exports = { layoutCapacityGuide, layoutCapacityMarkdown, writeLayoutCapacityGuide, warnSpecTextCapacity, textVisualLength };
