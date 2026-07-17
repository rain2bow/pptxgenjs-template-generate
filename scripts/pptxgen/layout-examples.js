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
  'text-quote': '全宽核心观点或引用',
  'text-article': '结构化正文',
  'text-briefing': '总领加多项分析的简报',
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
  'image-statement': '单图加核心陈述',
  'image-quote': '单图加引用',
  'image-text': '单图加正文',
  'image-feature': '单图加多项解释',
  'image-grid': '一至六图网格',
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
  if (layoutName === 'deck-cover') return { ...base, title: '演示文稿标题', subtitle: '副标题与汇报信息' };
  if (layoutName === 'deck-section') return { ...base, title: '章节标题', subtitle: '本章节关注的问题' };
  if (layoutName === 'deck-closing') return { ...base, title: '谢谢', subtitle: '联系人或结束语' };
  if (layoutName === 'text-quote') return { ...base, title: '一句明确、可复述的核心判断', body: '补充这项判断成立的背景或依据。' };
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
  if (layoutName === 'image-grid') return {
    ...base,
    images: ['/path/to/image-1.jpg', '/path/to/image-2.jpg', '/path/to/image-3.jpg'],
    items: [{ caption: '图片说明一' }, { caption: '图片说明二' }, { caption: '图片说明三' }],
  };
  if (layoutName === 'image-hero') return { ...base, body: '解释主视觉所代表的关键证据。', images: ['/path/to/hero.jpg'], items: metricItems(3) };
  if (layoutName === 'image-case-study') return { ...base, caseTitle: '案例名称', body: '案例背景、做法与结果。', images: ['/path/to/case.jpg'], items: metricItems(3) };
  if (layoutName === 'image-feature') return { ...base, body: '图片对应的总体说明。', images: ['/path/to/feature.jpg'], items: textItems(3, '要点') };
  if (layoutName === 'image-statement') return { ...base, title: '图片支持的一句核心判断', body: '补充判断依据。', images: ['/path/to/statement.jpg'] };
  if (layoutName === 'image-quote') return { ...base, title: '图片与观点共同构成叙事重点', body: '引用来源或解释。', images: ['/path/to/quote.jpg'] };
  if (layoutName === 'image-text') return { ...base, body: '与图片并列呈现的完整正文。', images: ['/path/to/text-image.jpg'] };
  if (layoutName === 'text-matrix') return { ...base, items: Array.from({ length: 6 }, (_, index) => ({ label: String(index + 1).padStart(2, '0'), title: `矩阵项 ${index + 1}` })) };
  if (layoutName === 'text-swimlane') return {
    ...base,
    stages: ['现在', '下一步', '后续'],
    items: [
      { title: '业务', body: '业务推进路径', items: ['梳理', '试点', '推广'] },
      { title: '技术', body: '技术推进路径', items: ['设计', '开发', '运营'] },
      { title: '治理', body: '治理推进路径', items: ['规则', '监控', '复盘'] },
    ],
  };
  if (layoutName === 'text-briefing') return { ...base, items: textItems(4, '简报') };
  if (layoutName === 'text-grid') return { ...base, items: textItems(6, '网格') };
  if (layoutName === 'text-cards') return { ...base, items: textItems(5, '卡片') };
  if (layoutName === 'text-weave') return { ...base, items: textItems(3, '主题') };
  if (layoutName === 'text-agenda') return { ...base, items: textItems(5, '章节') };
  if (layoutName === 'text-pyramid') return { ...base, items: textItems(5, '层级') };
  if (layoutName === 'text-radial') return { ...base, center: '中心主题', items: textItems(6, '节点') };
  if (layoutName === 'text-list') return { ...base, items: textItems(5, '列表') };
  if (layoutName === 'text-article') return { ...base, items: textItems(4, '正文') };
  return { ...base, items: textItems(layoutName === 'text-pipeline' ? 5 : 4, '步骤') };
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
