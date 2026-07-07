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
    ['sections[].title', 'analysis card title', 4, 14], ['sections[].body', 'analysis card body', 18, 58], ['sections[].points[]', 'analysis card explicit numbered point', 8, 24], ['items[].title', 'analysis card title', 4, 14], ['items[].body', 'analysis card body', 18, 58], ['items[].points[]', 'analysis card explicit numbered point', 8, 24], ['agenda[].title', 'analysis card title', 4, 14], ['agenda[].body', 'analysis card body', 18, 58], ['agenda[].points[]', 'analysis card explicit numbered point', 8, 24],
    ['conclusion', 'bottom conclusion body', 12, 48], ['takeaway', 'bottom conclusion body alias', 12, 48],
  ], name + ': CMB top summary + middle analysis cards + bottom takeaway.');
}

function cmbTextWeaveSlots(name) {
  return slots([
    ['title', 'page title', 8, 32], ['sections[].title', 'text card title', 4, 14], ['sections[].body', 'text card body', 15, 52], ['sections[].points[]', 'text card explicit numbered point', 8, 24], ['items[].title', 'text card title', 4, 14], ['items[].body', 'text card body', 15, 52], ['items[].points[]', 'text card explicit numbered point', 8, 24], ['columns[].title', 'text card title', 4, 14], ['columns[].body', 'text card body', 15, 52], ['columns[].points[]', 'text card explicit numbered point', 8, 24],
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
    const leadText = hasLead ? (slide.summary || slide.body || slide.lead) : itemBody(items[0]);
    warnCmbCardText(index, layout, hasLead ? 'summary/body/lead' : `${sourceKey || 'items'}[0].body`, leadText, leadEntry.box, leadTitle, leadEntry.options);
  }
  const rest = hasLead ? items : items.slice(1);
  entries.filter((entry) => entry.kind === 'rest').forEach((entry, i) => {
    const item = rest[i];
    if (!item) return;
    const sourceIndex = hasLead ? i : i + 1;
    warnCmbCardText(index, layout, `${sourceKey || 'items'}[${sourceIndex}].body`, itemBody(item), entry.box, itemTitle(item) || `分析${i + 1}`, entry.options);
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
  const entries = [{ kind: 'lead', box: { x: 0.78, y: y0, w: 11.45, h: 1.12 }, options: { lead: true, accent: true, titleFontSize: 13.2 } }];
  const restCount = hasLead ? itemCount : Math.max(0, itemCount - 1);
  const midY = y0 + 1.34;
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

function warnCmbCardText(index, layout, field, text, box, title, options = {}) {
  if (!text) return;
  const cap = cmbCardBodyCapacity(box, title, options);
  const actual = Math.ceil(textVisualLength(text));
  if (actual > cap.max) {
    const preview = String(text).replace(/\s+/g, ' ').slice(0, 54);
    console.warn('Warning: slide ' + (index + 1) + ' layout "' + layout + '" field ' + field + ' has ' + actual + ' visual chars; estimated card capacity ' + cap.max + ' at 12pt Microsoft YaHei. Shorten this field in JSON, split it, or use fewer cards: ' + preview);
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
    const actual = Math.ceil(textVisualLength(text));
    if (actual > cap.max) {
      const field = sourceKey + '[' + itemIndex + '].' + (points.length ? 'points' : 'body');
      const preview = text.replace(/\s+/g, ' ').slice(0, 54);
      console.warn('Warning: slide ' + (index + 1) + ' layout "' + layout + '" field ' + field + ' has ' + actual + ' visual chars; estimated card capacity ' + cap.max + ' at 12pt Microsoft YaHei. Shorten this field in JSON, split it, or use fewer cards: ' + preview);
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
  return { max: estimatedChineseCapacity(contentW, bodyH, 12, 0.02) };
}

function estimatedChineseCapacity(w, h, fontSize, margin = 0) {
  const boxW = roundCapacityDimension(Math.max(0.05, w - margin * 2));
  const boxH = roundCapacityDimension(Math.max(0.05, h - margin * 2));
  const charsPerLine = Math.max(1, (boxW * 72) / fontSize);
  const lines = Math.max(1, Math.floor((boxH * 72) / (fontSize * 1.12)));
  return Math.max(1, Math.floor(charsPerLine * lines * 0.99));
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

module.exports = { layoutCapacityGuide, layoutCapacityMarkdown, writeLayoutCapacityGuide, warnSpecTextCapacity, textVisualLength };
