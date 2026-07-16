#!/usr/bin/env node
const path = require('node:path');
const { buildDeck } = require('../scripts/generate-pptx.js');

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
    body: '明确本项工作的关键判断，并拆解可执行动作与责任边界。',
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
  title: 'CMB 全部支持布局',
  subtitle: '招商银行风格布局检查示例',
  author: 'PPTXGenJS 模板生成器',
  style: 'cmb',
  theme: 'classic',
  logoHeader,
  logoMark,
  logoHeaderW: 1.58,
  logoHeaderH: 0.5,
  logoHeaderBandH: 0.82,
  headY: 1.06,
  slides: [
    { layout: 'cover', kicker: '招商银行布局检查 / 01', title: '招商银行封面页', subtitle: '检查透明标识位置、页眉和封面标题层级。' },
    { layout: 'section', kicker: '招商银行布局检查 / 02', title: '章节分隔页', subtitle: '用于章节过渡和大段内容切换。' },
    { layout: 'statement', image: 'assets/logos/cmb-logo-lockup.png', kicker: '核心观点', title: '用一句明确判断锚定整页叙事。', body: '本页用于承载核心结论、管理判断或阶段性决策。', callout: '稳健\n清晰\n协同' },
    { layout: 'kpiTower', kicker: '关键指标', title: '四项经营指标保持稳健改善', items: [metric('客户增长', '+18%'), metric('活跃提升', '+32%'), metric('覆盖率', '96%'), metric('周期缩短', '-24%')] },
    { layout: 'bigNumbers', kicker: '核心数字', title: '大数字兼容页面', items: [metric('资产规模', '128'), metric('满意度', '91'), metric('覆盖网点', '42'), metric('风险项', '7')] },
    { layout: 'dashboard', kicker: '数据看板', title: '指标条与双图表组合', metrics: [{ label: '命中率', value: '87%' }, { label: '时效', value: '2.4h' }, { label: '自动化', value: '68%' }, { label: '准确率', value: '95%' }], charts: [chart('季度处置量', 'column'), chart('渠道结构', 'doughnut')] },
    { layout: 'chart', kicker: '图表页', title: '单一主图表与右侧洞察', chart: chart('客户活跃趋势', 'line'), insights: [item(1, '活跃'), item(2, '转化'), item(3, '留存')] },
    { layout: 'dataSheet', kicker: '数据表', title: '可编辑表格与侧边说明', table: table(), notes: [item(1, '负责人'), item(2, '节奏'), item(3, '风险')] },
    { layout: 'media', kicker: '媒体页', title: '媒体区域与侧边解释', body: '媒体槽位可承载图表或图片；没有素材时应改用纯文本布局。', chart: chart('媒体区图表', 'bar'), items: [item(1, '信号'), item(2, '行动'), item(3, '复盘')] },
    { layout: 'mediaGrid', kicker: '媒体网格', title: '自适应媒体网格', captions: [captionItem(1, '槽位一'), captionItem(2, '槽位二'), captionItem(3, '槽位三'), captionItem(4, '槽位四')], charts: [chart('图表一'), chart('图表二', 'line'), chart('图表三', 'doughnut'), chart('图表四', 'column')] },
    { layout: 'gallery', kicker: '图片集', title: '图库别名页面', captions: [captionItem(1, '图片一'), captionItem(2, '图片二'), captionItem(3, '图片三')], charts: [chart('图库一'), chart('图库二', 'line'), chart('图库三', 'doughnut')] },
    { layout: 'imageGrid', kicker: '图片网格', title: '图片网格兼容页面', captions: [captionItem(1, '证据一'), captionItem(2, '证据二'), captionItem(3, '证据三'), captionItem(4, '证据四')], charts: [chart('证据一'), chart('证据二', 'line'), chart('证据三', 'doughnut'), chart('证据四', 'column')] },
    { layout: 'imageHero', kicker: '主视觉媒体', title: '大图媒体与指标组合', body: '上方区域保留给关键证据图片或图表。', chart: chart('关键证据', 'area'), items: [metric('触达', '46%'), metric('信任度', '91'), metric('时效', '2.4h')] },
    { layout: 'quoteImage', image: 'assets/logos/cmb-logo-lockup.png', kicker: '引用图片', title: '引用图片兼容页面', quote: '当信号、服务和风险形成协同，客户价值会持续累积。', body: '主视觉媒体区域仍保持可编辑。' },
    { layout: 'textImage', image: 'assets/logos/cmb-logo-lockup.png', kicker: '图文页', title: '图文兼容页面', body: '有图片或图表素材时使用本页；没有素材时应改用纯文本布局。' },
    { layout: 'compare', kicker: '前后对比', title: '改造前后对比', before: { title: '改造前', items: [item(1, '人工处理'), item(2, '系统割裂'), item(3, '响应滞后')] }, after: { title: '改造后', items: [item(4, '自动触发'), item(5, '统一视图'), item(6, '过程可见')] } },
    { layout: 'duoCompare', kicker: '双栏对比', title: '双栏对比别名页面', before: { title: '当前状态', items: [item(1, '工具分散'), item(2, '闭环偏弱')] }, after: { title: '目标状态', items: [item(3, '统一入口'), item(4, '闭环运营')] } },
    { layout: 'timeline', kicker: '时间线', title: '时间线布局', steps: [item(1, '规划'), item(2, '试点'), item(3, '推广'), item(4, '复盘')] },
    { layout: 'pipeline', kicker: '流程管道', title: '流程管道别名布局', steps: [item(1, '输入'), item(2, '建模'), item(3, '复核'), item(4, '输出'), item(5, '反馈')] },
    { layout: 'roadmap', kicker: '路线图', title: '路线图布局', steps: [item(1, '一季度'), item(2, '二季度'), item(3, '三季度'), item(4, '四季度')], highlightLast: true },
    { layout: 'textGrid', kicker: '文本网格', title: '多点紧凑文本页', sections: [pointItem(1), pointItem(2), pointItem(3), pointItem(4), pointItem(5), pointItem(6)], highlightIndex: 1 },
    { layout: 'article', kicker: '正文页', title: '正文兼容网格', sections: [pointItem(1, '背景'), pointItem(2, '发现'), pointItem(3, '影响'), pointItem(4, '行动')] },
    { layout: 'briefing', kicker: '经营简报', title: '高密度经营逻辑简报', summaryTitle: '经营重点', summary: '通过顶部总览、中部分析和底部结论承载高密度内容，避免普通分栏导致信息割裂。', sections: [1, 2, 3, 4].map((n) => pointItem(n, pointTitles[n - 1])), conclusionTitle: '结论', conclusion: '底部应保留最终判断或下一步行动，便于管理层快速抓住重点。' },
    { layout: 'textWeave', kicker: '文本编织', title: '非对称文本编织页', sections: [1, 2, 3].map((n) => pointItem(n, pointTitles[n - 1])) },
    { layout: 'fourCards', kicker: '自适应卡片', title: '自适应卡片布局', items: [pointItem(1), pointItem(2), pointItem(3), pointItem(4), pointItem(5)] },
    { layout: 'matrix', kicker: '矩阵', title: '矩阵布局', items: [1, 2, 3, 4, 5, 6, 7, 8].map((n) => ({ label: `0${n}`, title: pointTitles[n - 1] })), highlightIndex: 2 },
    { layout: 'agenda', kicker: '目录', title: '目录与章节导航', items: [item(1, '总览'), item(2, '经营表现'), item(3, '风险控制'), item(4, '推进路径')] },
    { layout: 'caseStudy', kicker: '案例页', title: '案例研究页面', caseTitle: '零售客户经营试点', body: '紧凑故事区将证据媒体与量化经营结果组合呈现。', chart: chart('案例证据', 'line'), metrics: [metric('提升', '+18%'), metric('触达', '42k'), metric('时效', '2.4h')] },
    { layout: 'pyramid', kicker: '金字塔', title: '分层结构页面', layers: [item(1, '数据'), item(2, '模型'), item(3, '流程'), item(4, '治理'), item(5, '价值')], note: '层级顺序和宽度应保持平衡。' },
    { layout: 'radial', kicker: '放射关系', title: '放射关系页面', center: '客户价值', items: [item(1, '数据'), item(2, '风险'), item(3, '服务'), item(4, '产品'), item(5, '渠道'), item(6, '增长')] },
    { layout: 'swimlane', kicker: '泳道矩阵', title: '泳道执行矩阵', stages: ['现在', '下一步', '后续'], lanes: [ { title: '零售', items: ['分层', '权益', '复盘'] }, { title: '风险', items: ['信号', '规则', '监测'] }, { title: '数据', items: ['清洗', '建模', '发布'] } ] },
    { layout: 'bigQuote', kicker: '强调引用', title: '大引用兼容页面', quote: '先让经营系统可见，再扩大经营节奏。', body: '作为内容序列中的阶段性收束观点。' },
    { layout: 'closing', kicker: '谢谢', title: '招商银行结束页', subtitle: '全部 CMB 布局已检查。' },
  ],
};

const outFile = path.join(__dirname, 'outputs', 'deck-cmb-all-layouts.pptx');

main().catch((error) => {
  console.error(error);
  process.exit(1);
});

async function main() {
  await buildDeck(deckSpec, __dirname, outFile);
  console.log('\nNext checks:');
  console.log(`  node scripts/validate-pptx-native.js "${outFile}"`);
  console.log(`  node scripts/validate-pptx-layout.js "${outFile}"`);
}
