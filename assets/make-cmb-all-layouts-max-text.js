#!/usr/bin/env node
const fs = require('node:fs');
const path = require('node:path');
const vm = require('node:vm');
const { buildDeck } = require('../scripts/generate-pptx.js');
const { layoutCapacityGuide, textVisualLength } = require('../scripts/pptxgen/text-capacity');

const root = path.resolve(__dirname, '..');
const templatePath = path.join(root, 'assets', 'template-cmb-all-layouts.js');
const outJson = path.join(root, 'outputs', 'cmb-all-layouts-max-text.json');
const outPptx = path.join(root, 'outputs', 'cmb-all-layouts-max-text.pptx');
const READABILITY_MIN_FONT = 12;

const source = fs.readFileSync(templatePath, 'utf8');
const moduleStub = { exports: {} };
const sandbox = {
  require: (id) => {
    if (id === '../scripts/generate-pptx.js') return { buildDeck: async () => {} };
    return require(id);
  },
  module: moduleStub,
  exports: moduleStub.exports,
  __dirname: path.join(root, 'assets'),
  console,
  process: { exit() {}, argv: [] },
};
vm.runInNewContext(source.replace(/const outFile[\s\S]*$/m, 'module.exports = deckSpec;'), sandbox, { filename: 'template-cmb-all-layouts.js' });

const spec = JSON.parse(JSON.stringify(moduleStub.exports));
spec.title = 'CMB 全 Layout 最大文本容量检查';
spec.subtitle = '按当前估算上限填充文本区域，供人工检查容量是否偏大或偏小';
spec.slides = spec.slides.map((slide, i) => ({
  ...slide,
  title: `${String(i + 1).padStart(2, '0')} ${slide.layout} 最大文本检查`,
}));

function repeatToVisual(max, label = '') {
  const seed = `${label || '容量'}客户分层触达节奏风险预警服务闭环复盘机制责任动作指标安排`;
  let out = '';
  while (textVisualLength(out) < max) out += seed;
  const chars = [...out];
  while (textVisualLength(chars.join('')) > max) chars.pop();
  return chars.join('');
}

function normalizeItems(value) {
  if (!value) return [];
  if (Array.isArray(value)) return value.map((item) => typeof item === 'string' || typeof item === 'number' ? { body: String(item) } : item).filter(Boolean);
  if (typeof value === 'object') return Object.entries(value).map(([title, body]) => ({ title, body: String(body ?? '') }));
  return [{ body: String(value) }];
}

function collectionKey(slide, keys) {
  return keys.find((key) => slide[key] !== undefined && slide[key] !== null);
}

function clearExplicitPoints(slide) {
  walkObjects(slide, (obj) => {
    delete obj.points;
    delete obj.bullets;
    delete obj.list;
  });
}

function walkObjects(value, cb) {
  if (Array.isArray(value)) value.forEach((item) => walkObjects(item, cb));
  else if (value && typeof value === 'object') {
    cb(value);
    Object.values(value).forEach((child) => walkObjects(child, cb));
  }
}

function setPath(source, field, max) {
  const parts = field.split('.');
  function rec(current, idx) {
    if (!current || typeof current !== 'object') return;
    const raw = parts[idx];
    const isArray = raw.endsWith('[]');
    const key = isArray ? raw.slice(0, -2) : raw;
    if (idx === parts.length - 1) {
      if (isArray && Array.isArray(current[key])) current[key] = current[key].map((_, i) => repeatToVisual(max, `${key}${i + 1}`));
      else if (!isArray && current[key] !== undefined) current[key] = repeatToVisual(max, key);
      return;
    }
    const next = current[key];
    if (isArray && Array.isArray(next)) next.forEach((item) => rec(item, idx + 1));
    else if (!isArray) rec(next, idx + 1);
  }
  rec(source, 0);
}

function round(v) {
  return Math.round(Number(v) * 1000) / 1000;
}

function estimateTextHeight(text, boxW, fontSize, options = {}) {
  const raw = String(text || '').trim();
  if (!raw) return options.empty ?? 0;
  const charsPerLine = Math.max(4, (boxW * 72) / Math.max(1, fontSize));
  const lines = raw.split(/\r?\n/).reduce((sum, line) => sum + Math.max(1, Math.ceil(textVisualLength(line) / charsPerLine)), 0);
  const height = (lines * fontSize * (options.lineHeight || 1.18)) / 72 + (options.padding || 0.04);
  return Math.max(options.min ?? 0.2, Math.min(options.max ?? 10, height));
}

function estimatedChineseCapacity(w, h, fontSize = READABILITY_MIN_FONT, margin = 0) {
  const boxW = round(Math.max(0.05, w - margin * 2));
  const boxH = round(Math.max(0.05, h - margin * 2));
  const charsPerLine = Math.max(1, (boxW * 72) / fontSize);
  const lines = Math.max(1, Math.floor((boxH * 72) / (fontSize * 1.12)));
  return Math.max(1, Math.floor(charsPerLine * lines * 0.99));
}

function cmbTextWeaveRightBoxes(count, x, y, w, h, gapX, rowGap) {
  const halfW = (w - gapX) / 2;
  const halfH = (h - rowGap) / 2;
  if (count <= 0) return [];
  if (count === 1) return [{ x, y, w, h }];
  if (count === 2) return [{ x, y, w: halfW, h }, { x: x + halfW + gapX, y, w: halfW, h }];
  if (count === 3) return [{ x, y, w: halfW, h }, { x: x + halfW + gapX, y, w: halfW, h: halfH }, { x: x + halfW + gapX, y: y + halfH + rowGap, w: halfW, h: halfH }];
  return Array.from({ length: Math.min(count, 4) }, (_, i) => ({ x: x + (i % 2) * (halfW + gapX), y: y + Math.floor(i / 2) * (halfH + rowGap), w: halfW, h: halfH }));
}

function cmbTextWeaveBoxes(count, hasSubtitle) {
  const y0 = hasSubtitle ? 2.78 : 2.42;
  const bottom = 6.48;
  if (count <= 5) {
    const leadW = 3.85;
    const gapX = 0.28;
    const boxes = [{ box: { x: 0.78, y: y0, w: leadW, h: bottom - y0 }, options: { lead: true, accent: true } }];
    cmbTextWeaveRightBoxes(count - 1, 0.78 + leadW + gapX, y0, 11.45 - leadW - gapX, bottom - y0, gapX, 0.24)
      .forEach((box) => boxes.push({ box, options: {} }));
    return boxes;
  }
  const boxes = [{ box: { x: 0.78, y: y0, w: 11.45, h: 0.9 }, options: { lead: true, accent: true, compact: true } }];
  const gridY = y0 + 1.12;
  const leftW = 3.72;
  const gapX = 0.28;
  boxes.push({ box: { x: 0.78, y: gridY, w: leftW, h: bottom - gridY }, options: {} });
  const cardW = (11.45 - leftW - gapX * 2) / 2;
  const rowGap = 0.2;
  const cardH = (bottom - gridY - rowGap) / 2;
  for (let i = 0; i < count - 2; i += 1) boxes.push({ box: { x: 0.78 + leftW + gapX + (i % 2) * (cardW + gapX), y: gridY + Math.floor(i / 2) * (cardH + rowGap), w: cardW, h: cardH }, options: {} });
  return boxes;
}

function cmbBriefingBoxes(slide, itemCount) {
  const y0 = slide.subtitle ? 2.78 : 2.45;
  const hasLead = !!(slide.summary || slide.body || slide.lead);
  const conclusionText = slide.conclusion || slide.takeaway || slide.footerSummary || slide.nextStep;
  const boxes = [];
  boxes.push({ target: 'lead', box: { x: 0.78, y: y0, w: 11.45, h: 1.12 }, options: { lead: true, accent: true, titleFontSize: 13.2 } });
  const restCount = hasLead ? itemCount : Math.max(0, itemCount - 1);
  const midY = y0 + 1.34;
  const conclusionBox = { x: 0.78, y: 5.62, w: 11.45, h: 0.86 };
  const midBottom = conclusionText ? conclusionBox.y - 0.2 : 6.48;
  if (restCount <= 4) {
    const gap = 0.28;
    const w = (11.45 - gap * Math.max(0, restCount - 1)) / Math.max(1, restCount);
    const h = Math.max(1.15, midBottom - midY);
    for (let i = 0; i < restCount; i += 1) boxes.push({ target: 'rest', box: { x: 0.78 + i * (w + gap), y: midY, w, h }, options: {} });
  } else {
    const gapX = 0.28;
    const gapY = 0.18;
    const cols = 3;
    const rows = Math.ceil(restCount / cols);
    const w = (11.45 - gapX * (cols - 1)) / cols;
    const h = Math.max(0.95, (midBottom - midY - gapY * (rows - 1)) / rows);
    for (let i = 0; i < restCount; i += 1) boxes.push({ target: 'rest', box: { x: 0.78 + (i % cols) * (w + gapX), y: midY + Math.floor(i / cols) * (h + gapY), w, h }, options: {} });
  }
  if (conclusionText) boxes.push({ target: 'conclusion', box: conclusionBox, options: { compact: true, accent: true, titleFontSize: 13.2 } });
  return boxes;
}

function cardCapacity(box, title, options = {}) {
  const padX = options.compact ? 0.2 : 0.26;
  const padTop = options.compact ? 0.12 : 0.18;
  const contentW = box.w - padX * 2;
  const titleW = Math.max(0.3, contentW - 0.38);
  const titleFont = options.titleFontSize || (options.lead ? 15.2 : options.compact ? 12.2 : 13.2);
  const titleH = estimateTextHeight(title, titleW, titleFont, { min: 0.26, max: options.compact ? 0.34 : 0.5, lineHeight: 1.14, padding: 0.02 });
  const bodyY = box.y + padTop + titleH + 0.13;
  const bodyH = round(Math.max(0.24, box.y + box.h - bodyY - (options.compact ? 0.12 : 0.18)));
  return estimatedChineseCapacity(contentW, bodyH, READABILITY_MIN_FONT, 0.02);
}

function itemTitle(item, fallback) {
  return String(item?.title || item?.label || item?.name || item?.heading || fallback || '').trim();
}

function fillDynamicTextWeave(slide) {
  const key = collectionKey(slide, ['sections', 'items', 'columns', 'points', 'agenda']);
  if (!key) return;
  const items = normalizeItems(slide[key]).slice(0, 6);
  slide[key] = items;
  cmbTextWeaveBoxes(items.length, !!slide.subtitle).forEach(({ box, options }, i) => {
    const item = items[i];
    if (!item) return;
    item.title = itemTitle(item, `卡片${i + 1}`);
    item.body = repeatToVisual(cardCapacity(box, item.title, options), item.title);
  });
}

function fillDynamicBriefing(slide) {
  const key = collectionKey(slide, ['sections', 'items', 'columns', 'points', 'agenda']);
  const items = normalizeItems(key ? slide[key] : []).slice(0, 6);
  if (key) slide[key] = items;
  const hasLead = !!(slide.summary || slide.body || slide.lead);
  const boxes = cmbBriefingBoxes(slide, items.length);
  const leadBox = boxes.find((entry) => entry.target === 'lead');
  if (leadBox) {
    const leadTitle = slide.summaryTitle || slide.leadTitle || slide.focusTitle || slide.kicker || itemTitle(items[0], '摘要');
    if (hasLead) slide.summary = repeatToVisual(cardCapacity(leadBox.box, leadTitle, leadBox.options), leadTitle);
    else if (items[0]) items[0].body = repeatToVisual(cardCapacity(leadBox.box, itemTitle(items[0], '摘要'), leadBox.options), itemTitle(items[0], '摘要'));
  }
  const rest = hasLead ? items : items.slice(1);
  boxes.filter((entry) => entry.target === 'rest').forEach((entry, i) => {
    const item = rest[i];
    if (!item) return;
    item.title = itemTitle(item, `分析${i + 1}`);
    item.body = repeatToVisual(cardCapacity(entry.box, item.title, entry.options), item.title);
  });
  const conclusionBox = boxes.find((entry) => entry.target === 'conclusion');
  if (conclusionBox) {
    const title = slide.conclusionTitle || slide.takeawayTitle || slide.footerSummaryTitle || slide.nextStepTitle || '结论';
    slide.conclusion = repeatToVisual(cardCapacity(conclusionBox.box, title, conclusionBox.options), title);
  }
}

function fillDynamicAgenda(slide) {
  const key = collectionKey(slide, ['items', 'agenda', 'sections']);
  if (!key) return;
  let items = normalizeItems(slide[key]).slice(0, 8);
  if (items.length < 6) {
    const base = ['总览', '经营表现', '客户运营', '风险控制', '渠道协同', '推进计划'];
    items = base.map((title, i) => items[i] || { title });
  }
  slide[key] = items;
  slide.agendaTitle = slide.agendaTitle || '目录';
  slide.agendaSubtitle = slide.agendaSubtitle || '章节导航';
  const y0 = slide.subtitle ? 2.78 : 2.42;
  const bottom = 6.48;
  const leftW = 2.15;
  const gap = 0.28;
  const listW = 11.45 - leftW - gap;
  const cols = items.length > 4 ? 2 : 1;
  const rows = Math.ceil(Math.max(1, items.length) / cols);
  const colGap = cols > 1 ? 0.24 : 0;
  const rowGap = 0.16;
  const colW = (listW - colGap * (cols - 1)) / cols;
  const rowH = (bottom - y0 - rowGap * Math.max(0, rows - 1)) / rows;
  const bodyW = cols > 1 ? colW - 1.08 : colW - 4.84;
  const bodyH = cols > 1 ? Math.max(0.24, rowH - 0.5) : Math.max(0.26, rowH - 0.2);
  const cap = estimatedChineseCapacity(bodyW, bodyH, READABILITY_MIN_FONT, 0.02);
  items.forEach((item, i) => {
    item.title = itemTitle(item, `章节${i + 1}`);
    item.body = repeatToVisual(cap, item.title);
  });
}
const guide = layoutCapacityGuide('cmb');
const dynamicTextWeaveLayouts = new Set(['textGrid', 'fourCards', 'textWeave', 'contentSynthesis', 'denseText']);
const dynamicBriefingLayouts = new Set(['article', 'sectionList', 'briefing', 'executiveBrief', 'contentBrief']);

spec.slides.forEach((slide) => {
  clearExplicitPoints(slide);
  const layout = slide.layout || 'statement';
  if (dynamicTextWeaveLayouts.has(layout)) {
    fillDynamicTextWeave(slide);
    return;
  }
  if (dynamicBriefingLayouts.has(layout)) {
    fillDynamicBriefing(slide);
    return;
  }
  if (layout === 'agenda') {
    fillDynamicAgenda(slide);
    return;
  }
  const layoutGuide = guide.layouts[layout];
  if (layoutGuide?.slots) layoutGuide.slots.forEach((slot) => {
    if (slot.field.includes('points[]')) return;
    setPath(slide, slot.field, slot.max);
  });
});

fs.mkdirSync(path.dirname(outJson), { recursive: true });
fs.writeFileSync(outJson, JSON.stringify(spec, null, 2) + '\n', 'utf8');
buildDeck(spec, root, outPptx).then(() => {
  console.log('Wrote ' + outJson);
  console.log('Wrote ' + outPptx);
}).catch((error) => {
  console.error(error);
  process.exit(1);
});