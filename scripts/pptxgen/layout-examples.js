'use strict';

const fs = require('node:fs');
const path = require('node:path');
const { canonicalLayoutNames, layoutDefinition } = require('./layout-schema');

const STYLE_GUIDE = Object.freeze({
  cmb: {
    name: '招商银行商务风',
    description: '白色品牌页眉、招商红与灰白金融版式，适合银行经营、管理层汇报和正式商务场景。',
    theme: 'classic',
  },
  swiss: {
    name: '瑞士国际主义',
    description: '高亮色、直角网格、清晰信息层级，适合产品、战略、数据分析和方法论汇报。',
    theme: 'ikb',
  },
  magazine: {
    name: '电子杂志风',
    description: '衬线标题、编辑式图文节奏和纸感留白，适合叙事、观点、案例和图片材料较多的内容。',
    theme: 'ink',
  },
});

const LAYOUT_PURPOSES = Object.freeze({
  'deck-cover': '封面',
  'deck-section': '章节分隔',
  'deck-closing': '结束页',
  'text-statement': '纯文本核心陈述',
  'text-quote': '全宽核心观点或引用',
  'text-article': '结构化正文',
  'text-briefing': '总领加多项分析的简报',
  'text-feature': '重点说明与多项解释',
  'text-list': '分组列表',
  'text-grid': '紧凑文本网格',
  'text-cards': '自适应文本卡片',
  'text-weave': '非对称文本编织',
  'text-agenda': '目录或议程',
  'text-timeline': '时间节点',
  'text-pipeline': '流程管道',
  'text-roadmap': '阶段路线图',
  'text-matrix': '标题型矩阵',
  'text-radial': '中心辐射关系',
  'text-pyramid': '分层金字塔',
  'text-swimlane': '多角色泳道',
  'text-hero': '主观点加关键结果',
  'text-case-study': '纯文本案例复盘',
  'image-statement': '单图加核心陈述',
  'image-quote': '单图加引用',
  'image-article': '图片加结构化正文',
  'image-briefing': '图片加分析简报',
  'image-feature': '单图加多项解释',
  'image-list': '图片加分组列表',
  'image-grid': '图片加紧凑内容网格',
  'image-cards': '图片加自适应卡片',
  'image-weave': '图片加非对称文本编织',
  'image-agenda': '图片加目录或议程',
  'image-timeline': '图片加时间节点',
  'image-pipeline': '图片加流程管道',
  'image-roadmap': '图片加阶段路线图',
  'image-matrix': '图片加标题型矩阵',
  'image-radial': '图片加中心辐射关系',
  'image-pyramid': '图片加分层金字塔',
  'image-swimlane': '图片加多角色泳道',
  'image-hero': '主视觉图片加指标',
  'image-case-study': '案例图片加量化结果',
  'data-numbers': '大数字指标',
  'data-kpis': 'KPI 指标卡',
  'data-compare': '前后或左右对比',
  'data-chart': '单一主图表加洞察',
  'data-dashboard': '双图表加指标看板',
  'data-table': '原生表格加侧边说明',
});

function styleGuideMarkdown() {
  const lines = ['# 模板风格选择', ''];
  Object.entries(STYLE_GUIDE).forEach(([key, value]) => {
    lines.push(`- **${key} · ${value.name}**：${value.description}`);
  });
  lines.push('', '请选择一种风格后，再生成该风格的全布局 JSON 示例。');
  return `${lines.join('\n')}\n`;
}

function layoutExamplesMarkdown(style) {
  const guide = STYLE_GUIDE[style];
  if (!guide) throw new Error(`Unsupported style "${style}". Use one of: ${Object.keys(STYLE_GUIDE).join(', ')}.`);
  const lines = [
    `# ${guide.name} JSON 布局示例`,
    '',
    `- style: \`${style}\``,
    `- default theme: \`${guide.theme}\``,
    `- layouts: ${canonicalLayoutNames().length}`,
    '',
    '所有集合统一使用 `items`，图片统一使用 `images[]`，图表统一使用 `charts[]`。以下内容是字段结构示例，不包含字数限制；请按用户内容改写文本和素材路径。',
    '',
  ];
  canonicalLayoutNames().forEach((layoutName, index) => {
    const definition = layoutDefinition(layoutName);
    lines.push(`## ${index + 1}. ${layoutName}`, '');
    lines.push(`类型：${definition.category}；用途：${LAYOUT_PURPOSES[layoutName] || layoutName}。`, '');
    lines.push('```json', JSON.stringify(exampleSlide(layoutName), null, 2), '```', '');
  });
  return `${lines.join('\n')}\n`;
}

function writeLayoutExamples(style, outPath) {
  const target = path.resolve(outPath);
  fs.mkdirSync(path.dirname(target), { recursive: true });
  fs.writeFileSync(target, layoutExamplesMarkdown(style), 'utf8');
  console.log(`Wrote ${target}`);
  console.log(`Style: ${style}`);
  console.log(`Layouts: ${canonicalLayoutNames().length}`);
  return target;
}

function exampleSlide(layoutName) {
  const base = { layout: layoutName, kicker: '页面标签', title: LAYOUT_PURPOSES[layoutName] || '页面标题' };
  const definition = layoutDefinition(layoutName);
  if (layoutName === 'deck-cover') return { ...base, title: '演示文稿标题', subtitle: '副标题与汇报信息' };
  if (layoutName === 'deck-section') return { ...base, title: '章节标题', subtitle: '本章节关注的问题' };
  if (layoutName === 'deck-closing') return { ...base, title: '谢谢', subtitle: '联系人或结束语' };
  if (definition?.pairKey) {
    const paired = pairedContentExample(definition.pairKey, base);
    if (definition.category === 'image') paired.images = imageExamples(definition.pairKey);
    return paired;
  }
  if (layoutName === 'data-compare') return {
    ...base,
    before: { title: '改造前', items: textItems(3, '现状') },
    after: { title: '改造后', items: textItems(3, '目标') },
  };
  if (layoutName === 'data-chart') return { ...base, charts: [chart('趋势变化', 'line')], items: textItems(3, '洞察') };
  if (layoutName === 'data-dashboard') return { ...base, items: metricItems(4), charts: [chart('业务规模', 'column'), chart('业务结构', 'doughnut')] };
  if (layoutName === 'data-table') return {
    ...base,
    table: { headers: ['工作项', '状态', '负责人'], rows: [['客户运营', '进行中', '团队 A'], ['风险治理', '已完成', '团队 B']] },
    items: textItems(3, '说明'),
  };
  if (layoutName === 'data-numbers') return { ...base, items: metricItems(6) };
  if (layoutName === 'data-kpis') return { ...base, items: metricItems(4) };
  return base;
}

function pairedContentExample(pairKey, base) {
  if (pairKey === 'statement') return { ...base, title: '一句明确、可复述的核心判断', body: '补充这项判断成立的背景或依据。' };
  if (pairKey === 'quote') return { ...base, title: '一句明确、可复述的核心观点', body: '补充引用来源或解释。' };
  if (pairKey === 'matrix') return { ...base, items: Array.from({ length: 6 }, (_, index) => ({ label: String(index + 1).padStart(2, '0'), title: `矩阵项 ${index + 1}` })) };
  if (pairKey === 'swimlane') return {
    ...base,
    stages: ['现在', '下一步', '后续'],
    items: [
      { title: '业务', body: '业务推进路径', items: ['梳理', '试点', '推广'] },
      { title: '技术', body: '技术推进路径', items: ['设计', '开发', '运营'] },
      { title: '治理', body: '治理推进路径', items: ['规则', '监控', '复盘'] },
    ],
  };
  if (pairKey === 'radial') return { ...base, center: '中心主题', items: textItems(6, '节点') };
  if (pairKey === 'feature') return { ...base, body: '概括这一页的重点判断。', items: textItems(3, '要点') };
  if (pairKey === 'hero') return { ...base, body: '概括主视觉或主观点代表的关键结果。', items: textItems(3, '结果') };
  if (pairKey === 'case-study') return { ...base, body: '说明案例背景、做法与结果。', items: textItems(3, '结果') };
  const counts = { article: 4, briefing: 4, list: 5, grid: 6, cards: 5, weave: 3, agenda: 5, timeline: 4, pipeline: 5, roadmap: 4, pyramid: 5 };
  return { ...base, items: textItems(counts[pairKey] || 4, pairKey === 'agenda' ? '章节' : pairKey === 'pipeline' ? '步骤' : '要点') };
}

function imageExamples(pairKey) {
  if (pairKey === 'grid' || pairKey === 'cards') {
    return ['/path/to/image-1.jpg', '/path/to/image-2.jpg', '/path/to/image-3.jpg'];
  }
  return [`/path/to/${pairKey}.jpg`];
}

function textItems(count, prefix) {
  return Array.from({ length: count }, (_, index) => ({ title: `${prefix} ${index + 1}`, body: '说明该项的判断、依据和行动。' }));
}

function metricItems(count) {
  return Array.from({ length: count }, (_, index) => ({ label: `指标 ${index + 1}`, value: `${(index + 1) * 12}%`, note: '指标说明' }));
}

function chart(title, chartType) {
  return { chartType, title, labels: ['一季度', '二季度', '三季度', '四季度'], values: [32, 46, 58, 74] };
}

module.exports = {
  STYLE_GUIDE,
  styleGuideMarkdown,
  layoutExamplesMarkdown,
  writeLayoutExamples,
  exampleSlide,
};
