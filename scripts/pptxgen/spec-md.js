const fs = require('node:fs');
const path = require('node:path');
const { fail } = require('./errors');
const { speakerNotesText } = require('./speaker-notes');

const LAYOUT_LABELS = {
  cover: '封面',
  section: '章节页',
  statement: '核心观点页',
  bigQuote: '大观点页',
  quoteImage: '观点图片页',
  textImage: '图文页',
  imageHero: '大图页',
  imageGrid: '图片宫格页',
  media: '媒体/图表页',
  mediaGrid: '多媒体页',
  gallery: '图库页',
  bigNumbers: '关键数字页',
  kpiTower: 'KPI 指标页',
  dashboard: '数据看板页',
  chart: '图表页',
  dataSheet: '表格数据页',
  article: '文章分栏页',
  sectionList: '分节列表页',
  textGrid: '文本卡片页',
  matrix: '矩阵页',
  fourCards: '卡片页',
  agenda: '议程/目录页',
  timeline: '时间线页',
  pipeline: '流程页',
  roadmap: '路线图页',
  swimlane: '泳道页',
  compare: '对比页',
  duoCompare: '双栏对比页',
  splitCompare: '双栏对比页',
  caseStudy: '案例页',
  pyramid: '金字塔页',
  radial: '放射关系页',
  closing: '结束页',
};

function specToMarkdown(spec) {
  if (!spec || typeof spec !== 'object') fail('Spec must be an object.');
  const slides = Array.isArray(spec.slides) ? spec.slides : [];
  const lines = [];
  const title = textValue(spec.title) || 'PPT 内容大纲';
  lines.push(`# ${title}`);
  lines.push('');
  addMeta(lines, '副标题', spec.subtitle);
  addMeta(lines, '风格', [spec.style, spec.theme].filter(Boolean).join(' / '));
  addMeta(lines, '作者', spec.author);
  lines.push(`- 总页数：${slides.length}`);
  lines.push('');
  lines.push('## 页面总览');
  lines.push('');
  if (!slides.length) {
    lines.push('暂无页面。');
    lines.push('');
    return lines.join('\n');
  }
  slides.forEach((slide, index) => {
    const pageTitle = pageDisplayTitle(slide, index);
    lines.push(`- 第 ${index + 1} 页：${pageTitle}（${layoutLabel(slide.layout)}）`);
  });
  lines.push('');
  slides.forEach((slide, index) => addSlideMarkdown(lines, slide, index, slides.length, spec));
  return `${lines.join('\n').replace(/\n{3,}/g, '\n\n').trim()}\n`;
}

function writeSpecMarkdown(spec, outPath) {
  fs.mkdirSync(path.dirname(outPath), { recursive: true });
  fs.writeFileSync(outPath, specToMarkdown(spec), 'utf8');
  console.log(`Wrote outline Markdown ${outPath}`);
}

function addSlideMarkdown(lines, slide, index, total, spec) {
  lines.push(`## 第 ${index + 1} / ${total} 页：${pageDisplayTitle(slide, index)}`);
  lines.push('');
  lines.push(`- 页面类型：${layoutLabel(slide.layout)}`);
  addMeta(lines, '页眉/标签', slide.kicker);
  addMeta(lines, '副标题', slide.subtitle);
  addMeta(lines, '主标题', slide.title || slide.quote || slide.caseTitle);
  addMeta(lines, '正文', slide.body || slide.story || slide.note || slide.cite);
  addMeta(lines, '重点提示', slide.callout || slide.center);
  addCompare(lines, slide);
  addCollections(lines, slide);
  addMetrics(lines, slide);
  addCharts(lines, slide);
  addTable(lines, slide);
  addMedia(lines, slide);
  addMeta(lines, '页脚/结束语', slide.footer || slide.closing);
  addMeta(lines, '演讲者备注', speakerNotesText(slide, { spec, index, total }));
  lines.push('');
}

function addCollections(lines, slide) {
  const metricItemLayouts = new Set(['dashboard', 'bigNumbers', 'kpiTower', 'imageHero', 'caseStudy']);
  const itemValue = metricItemLayouts.has(slide.layout) ? null : slide.items;
  const collections = [
    ['内容要点', slide.sections],
    ['内容要点', itemValue],
    ['内容要点', slide.columns],
    ['议程', slide.agenda],
    ['步骤', slide.steps],
    ['阶段', slide.milestones],
    ['节点', slide.nodes],
    ['层级', slide.layers],
    ['泳道', slide.lanes],
    ['图片说明', slide.captions],
    ['补充说明', slide.notes],
    ['洞察', slide.insights],
    ['侧边要点', slide.points],
  ];
  const used = new Set();
  collections.forEach(([label, value]) => {
    if (!value || used.has(value)) return;
    const items = normalizeItems(value);
    if (!items.length) return;
    used.add(value);
    lines.push(`- ${label}：`);
    items.forEach((item, i) => lines.push(`  ${i + 1}. ${formatItem(item)}`));
  });
}

function addMetrics(lines, slide) {
  const metricItemLayouts = new Set(['dashboard', 'bigNumbers', 'kpiTower', 'imageHero', 'caseStudy']);
  const metrics = normalizeItems(slide.metrics || (metricItemLayouts.has(slide.layout) ? slide.items : null));
  if (!metrics.length) return;
  lines.push('- 指标：');
  metrics.forEach((item, i) => {
    const value = [item.value, item.unit].filter(Boolean).join(' ');
    const label = item.label || item.title || `指标 ${i + 1}`;
    const note = item.note || item.body || item.desc;
    const detail = value && note ? `${value}（${note}）` : value || note;
    lines.push(`  ${i + 1}. ${[label, detail].filter(Boolean).join('：')}`);
  });
}

function addCharts(lines, slide) {
  const charts = [];
  if (slide.chart) charts.push(slide.chart);
  if (Array.isArray(slide.charts)) charts.push(...slide.charts);
  if (!charts.length) return;
  lines.push('- 图表：');
  charts.forEach((chart, i) => {
    const title = chart.title || `图表 ${i + 1}`;
    const type = chart.chartType || chart.type || 'chart';
    const labels = Array.isArray(chart.labels) ? chart.labels.join('、') : '';
    lines.push(`  ${i + 1}. ${title}（${type}${labels ? `；分类：${labels}` : ''}）`);
  });
}

function addTable(lines, slide) {
  const table = slide.table || (slide.rows || slide.headers ? slide : null);
  if (!table) return;
  const headers = Array.isArray(table.headers) ? table.headers.join(' / ') : '';
  const rows = Array.isArray(table.rows) ? table.rows : Array.isArray(table.data) ? table.data : [];
  lines.push(`- 表格：${headers || '未命名表格'}${rows.length ? `，${rows.length} 行数据` : ''}`);
}

function addMedia(lines, slide) {
  const images = [];
  if (slide.image) images.push(slide.image);
  if (Array.isArray(slide.images)) images.push(...slide.images);
  if (!images.length && !slide.mediaCount && !slide.imageSlots) return;
  lines.push('- 图片/媒体：');
  if (images.length) {
    images.forEach((image, i) => {
      const name = typeof image === 'string' ? image : image.path || image.src || image.caption || `图片 ${i + 1}`;
      lines.push(`  ${i + 1}. ${name}`);
    });
  } else {
    lines.push(`  - 预留 ${slide.mediaCount || slide.imageSlots} 个图片/媒体位置`);
  }
}

function addCompare(lines, slide) {
  const before = slide.before || slide.left;
  const after = slide.after || slide.right;
  if (!before && !after) return;
  lines.push('- 对比内容：');
  [
    ['左侧/之前', before],
    ['右侧/之后', after],
  ].forEach(([label, value]) => {
    if (!value) return;
    if (Array.isArray(value)) {
      lines.push(`  - ${label}：`);
      normalizeItems(value).forEach((item, i) => lines.push(`    ${i + 1}. ${formatItem(item)}`));
      return;
    }
    lines.push(`  - ${label}：${formatItem(value)}`);
    const items = normalizeItems(value.items || value.points || value.sections);
    items.forEach((item, i) => lines.push(`    ${i + 1}. ${formatItem(item)}`));
  });
}

function addMeta(lines, label, value) {
  const text = textValue(value);
  if (text) lines.push(`- ${label}：${text}`);
}

function normalizeItems(value) {
  if (!value) return [];
  if (typeof value === 'string' || typeof value === 'number') return [{ body: String(value) }];
  if (Array.isArray(value)) return value.filter((item) => item !== null && item !== undefined);
  if (typeof value === 'object') return Object.entries(value).map(([title, body]) => ({ title, body }));
  return [];
}

function formatItem(item) {
  if (typeof item === 'string' || typeof item === 'number') return String(item);
  if (!item || typeof item !== 'object') return '';
  const heading = textValue(item.title || item.label || item.name || item.text || item.value || item.metric);
  const body = textValue(item.body || item.desc || item.note || item.summary || item.detail || item.story);
  const unit = textValue(item.unit);
  if (heading && body) return `${heading}：${body}`;
  if (heading && unit) return `${heading} ${unit}`;
  return heading || body || JSON.stringify(item);
}

function pageDisplayTitle(slide, index) {
  return textValue(slide.title || slide.quote || slide.caseTitle || slide.kicker) || `未命名页面 ${index + 1}`;
}

function layoutLabel(layout) {
  if (!layout) return '默认页面';
  return `${LAYOUT_LABELS[layout] || layout}（${layout}）`;
}

function textValue(value) {
  if (value === null || value === undefined) return '';
  if (Array.isArray(value)) return value.map(textValue).filter(Boolean).join(' / ');
  if (typeof value === 'object') return '';
  return String(value).trim();
}

module.exports = {
  specToMarkdown,
  writeSpecMarkdown,
};
