#!/usr/bin/env node
const path = require('node:path');
const fs = require('node:fs');
const root = path.resolve(__dirname, '../..');
const { buildDeck } = require('../../scripts/generate-pptx.js');

const styleArgIndex = process.argv.indexOf('--style');
const style = styleArgIndex >= 0 ? process.argv[styleArgIndex + 1] : 'cmb';
const themes = { cmb: 'classic', swiss: 'ikb', magazine: 'ink' };
if (!themes[style]) throw new Error(`Unsupported --style ${style}; use cmb, swiss, or magazine.`);

const logoHeader = 'logos/cmb-logo-lockup.png';
const logoMark = 'logos/cmb-logo-mark.svg';
const pointTitles = ['客群分层', '信号捕捉', '权益匹配', '风险协同', '渠道联动', '服务闭环', '数据治理', '复盘优化'];

function item(n, title = pointTitles[(n - 1) % pointTitles.length]) {
  return {
    icon: ['users', 'chart-line', 'shield-check', 'workflow', 'database', 'target'][n % 6],
    label: `0${n}`,
    title,
    body: '检查间距和字号。',
  };
}

function captionItem(n, title = pointTitles[(n - 1) % pointTitles.length]) {
  return {
    icon: ['image', 'chart-line', 'scan-search', 'file-text', 'database', 'target'][n % 6],
    label: `0${n}`,
    caption: title,
  };
}
function pointItem(n, title = pointTitles[(n - 1) % pointTitles.length]) {
  return {
    ...item(n, title),
    body: '明确关键判断、执行动作与责任边界。',
  };
}

function metric(label, value, note = '用于检查指标说明。') {
  return { label, value, note, valueNum: Number(String(value).replace(/[^0-9.-]/g, '')) || 1 };
}

function chart(title, chartType = 'column') {
  return { chartType, title, labels: ['一季度', '二季度', '三季度', '四季度'], values: [32, 46, 58, 74], showValue: chartType !== 'line' };
}

function table() {
  return {
    headers: ['工作项', '状态', '负责人', '下一步'],
    rows: [
      ['零售增长', '试点中', '团队A', '规模推广'],
      ['风险信号', '运行中', '团队B', '持续复盘'],
      ['服务流程', '设计中', '团队C', '上线验证'],
      ['数据集市', '建设中', '团队D', '口径校验'],
    ],
    caption: '原生可编辑表格',
  };
}

const deckSpec = {
  title: `${style} 全部支持布局`,
  subtitle: '统一 canonical 布局检查示例',
  author: 'PPTXGenJS 模板生成器',
  style,
  theme: themes[style],
  logoHeader,
  logoMark,
  logoHeaderW: 1.58,
  logoHeaderH: 0.5,
  logoHeaderBandH: 0.82,
  headY: 1.06,
  slides: [
    { layout: 'deck-cover', kicker: '招商银行布局检查 / 01', title: '招商银行封面页', subtitle: '检查透明标识位置、页眉和封面标题层级。' },
    { layout: 'deck-section', kicker: '招商银行布局检查 / 02', title: '章节分隔页', subtitle: '用于章节过渡和大段内容切换。' },
    { layout: 'image-statement', images: ['assets/logos/cmb-logo-lockup.png'], kicker: '核心观点', title: '用一句明确判断锚定整页叙事。', body: '本页用于承载核心结论、管理判断或阶段性决策。', callout: '稳健\n清晰\n协同' },
    { layout: 'data-kpis', kicker: '关键指标', title: '四项经营指标保持稳健改善', items: [metric('客户增长', '+18%'), metric('活跃提升', '+32%'), metric('覆盖率', '96%'), metric('周期缩短', '-24%')] },
    { layout: 'data-numbers', kicker: '核心数字', title: '核心数字页面', items: [metric('资产规模', '128'), metric('满意度', '91'), metric('覆盖网点', '42'), metric('风险项', '7')] },
    { layout: 'data-dashboard', kicker: '数据看板', title: '指标条与双图表组合', items: [{ label: '命中率', value: '87%' }, { label: '时效', value: '2.4h' }, { label: '自动化', value: '68%' }, { label: '准确率', value: '95%' }], charts: [chart('季度处置量', 'column'), chart('渠道结构', 'doughnut')] },
    { layout: 'data-chart', kicker: '图表页', title: '单一主图表与右侧洞察', charts: [chart('客户活跃趋势', 'line')], items: [item(1, '活跃'), item(2, '转化'), item(3, '留存')] },
    { layout: 'data-table', kicker: '数据表', title: '可编辑表格与侧边说明', table: table(), items: [item(1, '负责人'), item(2, '节奏'), item(3, '风险')] },
    { layout: 'image-feature', images: ['assets/logos/cmb-logo-lockup.png'], kicker: '媒体页', title: '媒体区域与侧边解释', body: '媒体槽位承载用户图片，右侧保留结构化解释。', items: [item(1, '信号'), item(2, '行动'), item(3, '复盘')] },
    { layout: 'image-grid', images: Array(4).fill('assets/logos/cmb-logo-lockup.png'), kicker: '图片网格', title: '自适应图片网格', items: [pointItem(1, '证据一'), pointItem(2, '证据二'), pointItem(3, '证据三'), pointItem(4, '证据四')] },
    { layout: 'image-hero', images: ['assets/logos/cmb-logo-lockup.png'], kicker: '主视觉媒体', title: '大图媒体与指标组合', body: '上方区域保留给关键证据图片。', items: [pointItem(1, '触达'), pointItem(2, '信任度'), pointItem(3, '时效')] },
    { layout: 'image-quote', images: ['assets/logos/cmb-logo-lockup.png'], kicker: '引用图片', title: '引用图片页面', quote: '当信号、服务和风险形成协同，客户价值会持续累积。', body: '主视觉媒体区域仍保持可编辑。' },
    { layout: 'image-article', images: ['assets/logos/cmb-logo-lockup.png'], kicker: '图文正文', title: '图文正文页面', items: [pointItem(1, '背景'), pointItem(2, '发现'), pointItem(3, '影响'), pointItem(4, '行动')] },
    { layout: 'image-briefing', images: ['assets/logos/cmb-logo-lockup.png'], kicker: '图文简报', title: '图片与简报要点', items: [pointItem(1), pointItem(2), pointItem(3), pointItem(4)] },
    { layout: 'image-list', images: ['assets/logos/cmb-logo-lockup.png'], kicker: '图文列表', title: '图片与分组列表', items: [pointItem(1), pointItem(2), pointItem(3), pointItem(4)] },
    { layout: 'image-cards', images: Array(3).fill('assets/logos/cmb-logo-lockup.png'), kicker: '图文卡片', title: '图片与自适应卡片', items: [pointItem(1), pointItem(2), pointItem(3), pointItem(4), pointItem(5)] },
    { layout: 'image-weave', images: ['assets/logos/cmb-logo-lockup.png'], kicker: '图文编织', title: '图片与文本编织', items: [pointItem(1), pointItem(2), pointItem(3)] },
    { layout: 'image-agenda', images: ['assets/logos/cmb-logo-lockup.png'], kicker: '图文目录', title: '图片与章节导航', items: [pointItem(1), pointItem(2), pointItem(3), pointItem(4), pointItem(5)] },
    { layout: 'image-timeline', images: ['assets/logos/cmb-logo-lockup.png'], kicker: '图文时间线', title: '图片与时间节点', items: [pointItem(1), pointItem(2), pointItem(3), pointItem(4)] },
    { layout: 'image-pipeline', images: ['assets/logos/cmb-logo-lockup.png'], kicker: '图文流程', title: '图片与流程节点', items: [pointItem(1), pointItem(2), pointItem(3), pointItem(4), pointItem(5)] },
    { layout: 'image-roadmap', images: ['assets/logos/cmb-logo-lockup.png'], kicker: '图文路线', title: '图片与阶段路线', items: [pointItem(1), pointItem(2), pointItem(3), pointItem(4)] },
    { layout: 'image-matrix', images: ['assets/logos/cmb-logo-lockup.png'], kicker: '图文矩阵', title: '图片与矩阵项', items: [1, 2, 3, 4, 5, 6].map((n) => ({ label: `0${n}`, title: pointTitles[n - 1] })) },
    { layout: 'image-radial', images: ['assets/logos/cmb-logo-lockup.png'], kicker: '图文辐射', title: '图片与辐射节点', center: '客户价值', items: [pointItem(1), pointItem(2), pointItem(3), pointItem(4), pointItem(5), pointItem(6)] },
    { layout: 'image-pyramid', images: ['assets/logos/cmb-logo-lockup.png'], kicker: '图文层级', title: '图片与分层结构', items: [pointItem(1), pointItem(2), pointItem(3), pointItem(4), pointItem(5)] },
    { layout: 'image-swimlane', images: ['assets/logos/cmb-logo-lockup.png'], kicker: '图文泳道', title: '图片与多角色泳道', stages: ['现在', '下一步', '后续'], items: [{ title: '业务', body: '业务推进路径', items: ['梳理', '试点', '推广'] }, { title: '技术', body: '技术推进路径', items: ['设计', '开发', '运营'] }] },
    { layout: 'image-case-study', images: ['assets/logos/cmb-logo-lockup.png'], kicker: '案例页', title: '案例研究页面', body: '紧凑故事区将证据媒体与量化经营结果组合呈现。', items: [pointItem(1, '提升'), pointItem(2, '触达'), pointItem(3, '时效')] },
    { layout: 'data-compare', kicker: '前后对比', title: '改造前后对比', before: { title: '改造前', items: [item(1, '人工处理'), item(2, '系统割裂'), item(3, '响应滞后')] }, after: { title: '改造后', items: [item(4, '自动触发'), item(5, '统一视图'), item(6, '过程可见')] } },
    { layout: 'text-timeline', kicker: '时间线', title: '时间线布局', items: [item(1, '规划'), item(2, '试点'), item(3, '推广'), item(4, '复盘')] },
    { layout: 'text-statement', kicker: '纯文本陈述', title: '一句明确判断锚定整页叙事', body: '该布局不要求图片，并与 image-statement 字段兼容。' },
    { layout: 'text-feature', kicker: '纯文本重点', title: '重点说明与多项解释', body: '概括本页的总体判断。', items: [pointItem(1), pointItem(2), pointItem(3)] },
    { layout: 'text-hero', kicker: '纯文本主观点', title: '主观点与关键结果', body: '在没有主视觉素材时保留相同内容字段。', items: [pointItem(1), pointItem(2), pointItem(3)] },
    { layout: 'text-case-study', kicker: '纯文本案例', title: '案例复盘页面', body: '不依赖图片的案例背景、做法与结果。', items: [pointItem(1), pointItem(2), pointItem(3)] },
    { layout: 'text-pipeline', kicker: '流程管道', title: '流程管道布局', items: [item(1, '输入'), item(2, '建模'), item(3, '复核'), item(4, '输出'), item(5, '反馈')] },
    { layout: 'text-roadmap', kicker: '路线图', title: '路线图布局', items: [item(1, '一季度'), item(2, '二季度'), item(3, '三季度'), item(4, '四季度')], highlightLast: true },
    { layout: 'text-grid', kicker: '文本网格', title: '多点紧凑文本页', items: [pointItem(1), pointItem(2), pointItem(3), pointItem(4), pointItem(5), pointItem(6)], highlightIndex: 1 },
    { layout: 'text-article', kicker: '正文页', title: '结构化正文页面', items: [pointItem(1, '背景'), pointItem(2, '发现'), pointItem(3, '影响'), pointItem(4, '行动')] },
    { layout: 'text-briefing', kicker: '经营简报', title: '高密度经营逻辑简报', items: [pointItem(1, '经营重点'), pointItem(2, '关键发现'), pointItem(3, '风险判断'), pointItem(4, '下一步行动')] },
    { layout: 'text-list', kicker: '列表页', title: '分组列表页面', items: [pointItem(1), pointItem(2), pointItem(3), pointItem(4)] },
    { layout: 'text-weave', kicker: '文本编织', title: '非对称文本编织页', items: [1, 2, 3].map((n) => pointItem(n, pointTitles[n - 1])) },
    { layout: 'text-cards', kicker: '自适应卡片', title: '自适应卡片布局', items: [pointItem(1), pointItem(2), pointItem(3), pointItem(4), pointItem(5)] },
    { layout: 'text-matrix', kicker: '矩阵', title: '矩阵布局', items: [1, 2, 3, 4, 5, 6, 7, 8].map((n) => ({ label: `0${n}`, title: pointTitles[n - 1] })), highlightIndex: 2 },
    { layout: 'text-agenda', kicker: '目录', title: '目录与章节导航', items: [item(1, '总览'), item(2, '经营表现'), item(3, '风险控制'), item(4, '推进路径')] },
    { layout: 'text-pyramid', kicker: '金字塔', title: '分层结构页面', items: [item(1, '数据'), item(2, '模型'), item(3, '流程'), item(4, '治理'), item(5, '价值')], note: '层级顺序和宽度应保持平衡。' },
    { layout: 'text-radial', kicker: '放射关系', title: '放射关系页面', center: '客户价值', items: [item(1, '数据'), item(2, '风险'), item(3, '服务'), item(4, '产品'), item(5, '渠道'), item(6, '增长')] },
    { layout: 'text-swimlane', kicker: '泳道矩阵', title: '泳道执行矩阵', stages: ['现在', '下一步', '后续'], items: [ { title: '零售', body: '零售经营推进事项', items: ['分层', '权益', '复盘'] }, { title: '风险', body: '风险治理推进事项', items: ['信号', '规则', '监测'] }, { title: '数据', body: '数据能力推进事项', items: ['清洗', '建模', '发布'] } ] },
    { layout: 'text-quote', kicker: '强调引用', title: '大引用页面', quote: '先让经营系统可见，再扩大经营节奏。', body: '作为内容序列中的阶段性收束观点。' },
    { layout: 'deck-closing', kicker: '谢谢', title: '招商银行结束页', subtitle: '全部 CMB 布局已检查。' },
  ],
};

const outDir = path.join(root, 'temp', 'outputs');
const outFile = path.join(outDir, `deck-${style}-all-layouts.pptx`);

main().catch((error) => {
  console.error(error);
  process.exit(1);
});

async function main() {
  fs.mkdirSync(outDir, { recursive: true });
  await buildDeck(deckSpec, root, outFile);
  console.log('\nNext checks:');
  console.log(`  node scripts/validate-pptx-native.js "${outFile}"`);
  console.log(`  node scripts/validate-pptx-layout.js "${outFile}"`);
}
