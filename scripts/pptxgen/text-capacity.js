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
  ], name + ': CMB asymmetric 1-6 text cards; 5-6 cards need shorter bodies. For points[] arrays, line count is estimated per point; each point occupies at least one line.');
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
  const style = spec.style || 'swiss';
  const slides = Array.isArray(spec.slides) ? spec.slides : [];
  return {
    style,
    unit: 'visual characters; CJK counts as 1, Latin letters count as about 0.56',
    instruction: 'This guide is calculated from the title-only deck plan. Fill only the listed fields. If a layout, media choice, or item count changes, regenerate this guide before writing full content.',
    slides: slides.map((slide, index) => plannedSlideCapacity(style, slide || {}, index)),
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
  layoutGuide.slots.forEach((slot) => {
    const match = slot.field.match(/^([^.[\]]+)\[\]\.(.+)$/);
    if (match) {
      const [, key, leaf] = match;
      if (collectionKey && key !== collectionKey) return;
      if (!collectionKey || !collectionItems.length) {
        if (!result.notes.some((note) => note.includes('title-only collection'))) result.notes.push('Add a title-only collection such as items/sections/steps in the plan JSON so per-card body capacity can be calculated.');
        return;
      }
      collectionItems.forEach((item, itemIndex) => {
        result.slots.push(adjustPlannedSlot(slot, collectionKey + '[' + itemIndex + '].' + leaf, itemTitle(item) || (collection.label || collectionKey) + ' ' + (itemIndex + 1), collectionItems.length, collection));
      });
      return;
    }
    if (isFillableScalarSlot(slot.field)) result.slots.push({ ...slot, title: plannedScalarTitle(slide, slot.field) });
  });
  return result;
}

function adjustPlannedSlot(slot, field, title, count, collection) {
  const isTitle = field.endsWith('.title') || field.endsWith('.label') || field.endsWith('.name');
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

function plannedCmbTextWeaveCapacity(slide, index, layout) {
  const sourceKey = firstCollectionKey(slide, ['sections', 'items', 'columns', 'points', 'agenda']);
  const items = sourceKey ? normalizeItems(slide[sourceKey]).slice(0, 6) : [];
  const result = { slide: index + 1, layout, title: slide.title || '', slots: [], notes: [] };
  if (!items.length) {
    result.notes.push('Add title-only sections/items/columns to the plan JSON so CMB text-card capacity can be calculated.');
    return result;
  }
  cmbTextWeaveCardBoxes(items.length, !!slide.subtitle).forEach(({ box, options }, itemIndex) => {
    const item = items[itemIndex];
    if (!item) return;
    const title = itemTitle(item) || String(itemIndex + 1).padStart(2, '0');
    const cap = cmbCardBodyCapacity(box, title, options);
    result.slots.push({ field: sourceKey + '[' + itemIndex + '].body', min: recommendedCapacityMin(cap.max), max: cap.max, label: 'card body', title, note: 'Calculated from actual card ' + (itemIndex + 1) + '/' + items.length + ': max ' + cap.maxLines + ' line(s), about ' + Math.floor(cap.charsPerLine) + ' CJK chars/line at 12pt.' });
    result.slots.push(plannedPointSlot(sourceKey + '[' + itemIndex + '].points[]', title, cap, 'explicit bullet point'));
  });
  return result;
}

function plannedCmbBriefingCapacity(slide, index, layout) {
  const sourceKey = firstCollectionKey(slide, ['sections', 'items', 'columns', 'points', 'agenda']);
  const items = sourceKey ? normalizeItems(slide[sourceKey]).slice(0, 6) : [];
  const hasLead = hasAnyKey(slide, ['summary', 'body', 'lead', 'summaryTitle', 'leadTitle', 'focusTitle']);
  const hasConclusion = hasAnyKey(slide, ['conclusion', 'takeaway', 'footerSummary', 'nextStep', 'conclusionTitle', 'takeawayTitle', 'footerSummaryTitle', 'nextStepTitle']);
  const pseudo = { ...slide, summary: hasLead ? '__planned__' : '', conclusion: hasConclusion ? '__planned__' : '' };
  const entries = cmbBriefingCardEntries(pseudo, items.length);
  const result = { slide: index + 1, layout, title: slide.title || '', slots: [], notes: [] };
  const plannedRestCount = hasLead ? items.length : Math.max(0, items.length - 1);
  const maxRestCount = hasConclusion ? 4 : 5;
  if (plannedRestCount > maxRestCount) {
    result.notes.push('This CMB briefing plan has ' + plannedRestCount + ' middle text block(s), but the layout supports at most ' + maxRestCount + (hasConclusion ? ' when conclusion/takeaway is present.' : '.') + ' Reduce item count or use textWeave/denseText.');
  }
  const leadEntry = entries.find((entry) => entry.kind === 'lead');
  if (leadEntry) {
    const title = slide.summaryTitle || slide.leadTitle || slide.focusTitle || (hasLead ? '摘要' : itemTitle(items[0]) || '摘要');
    const cap = cmbCardBodyCapacity(leadEntry.box, title, leadEntry.options);
    const field = hasLead ? 'summary' : (sourceKey || 'items') + '[0].body';
    result.slots.push({ field, min: recommendedCapacityMin(cap.max), max: cap.max, label: hasLead ? 'top summary body' : 'lead card body', title, note: 'Calculated from lead card: max ' + cap.maxLines + ' line(s), about ' + Math.floor(cap.charsPerLine) + ' CJK chars/line at 12pt.' });
  }
  const rest = hasLead ? items : items.slice(1);
  entries.filter((entry) => entry.kind === 'rest').forEach((entry, i) => {
    const sourceIndex = hasLead ? i : i + 1;
    const item = rest[i];
    if (!item) return;
    const title = itemTitle(item) || '分析' + (i + 1);
    const cap = cmbCardBodyCapacity(entry.box, title, entry.options);
    result.slots.push({ field: (sourceKey || 'items') + '[' + sourceIndex + '].body', min: recommendedCapacityMin(cap.max), max: cap.max, label: 'analysis card body', title, note: 'Calculated from actual card ' + (i + 1) + '/' + rest.length + ': max ' + cap.maxLines + ' line(s), about ' + Math.floor(cap.charsPerLine) + ' CJK chars/line at 12pt.' });
    result.slots.push(plannedPointSlot((sourceKey || 'items') + '[' + sourceIndex + '].points[]', title, cap, 'analysis bullet point'));
  });
  const conclusionEntry = entries.find((entry) => entry.kind === 'conclusion');
  if (conclusionEntry) {
    const title = slide.conclusionTitle || slide.takeawayTitle || slide.footerSummaryTitle || slide.nextStepTitle || '结论';
    const cap = cmbCardBodyCapacity(conclusionEntry.box, title, conclusionEntry.options);
    result.slots.push({ field: 'conclusion', min: recommendedCapacityMin(cap.max), max: cap.max, label: 'bottom conclusion body', title, note: 'Calculated from bottom card: max ' + cap.maxLines + ' line(s), about ' + Math.floor(cap.charsPerLine) + ' CJK chars/line at 12pt.' });
  }
  if (!items.length && !hasLead) result.notes.push('Add title-only sections/items or a summary field in the plan JSON so briefing capacity can be calculated.');
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

module.exports = { layoutCapacityGuide, layoutCapacityGuideForSpec, layoutCapacityMarkdown, layoutCapacityMarkdownForSpec, writeLayoutCapacityGuide, writePlannedCapacityGuide, warnSpecTextCapacity, textVisualLength };
