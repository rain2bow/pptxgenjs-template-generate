#!/usr/bin/env node
const path = require('node:path');
const { buildDeck } = require('../scripts/generate-pptx.js');

/* ============================================================================
   PPTXGenJS Template / Style A
   电子杂志 x 电子墨水 / Native PPTX Seed

   这个文件是 magazine 风格的可运行 JSON spec 示例，输出为
   pptxgenjs 原生可编辑 PPTX：
   1. 复制本文件到项目目录，例如 项目/XXX/pptx/deck.js
   2. 只修改 deckSpec：标题、主题、slides、图片路径、文案
   3. 运行 node deck.js

   硬规则：
   - 只用 pptxgenjs 原生生成，保持文本、形状、图片、表格和图表可编辑。
   - 保留杂志页眉页脚、衬线大标题、纸感底色、细线、图片槽位和数据大字报。
   - 背景纹理、色块、图片先画；正文、标题、页眉页脚最后画，避免文字被覆盖。
   - 内容过多时改写或拆页，不要缩到不可读。
   - readable projection fonts are preferred; split dense content instead of shrinking small text

   可选主题：
   - ink      / 墨水经典，默认
   - indigo   / 靛蓝瓷
   - forest   / 森林墨
   - kraft    / 牛皮纸
   - dune     / 沙丘
   - cmb      / 招商银行红

   支持版式：
   - cover       杂志封面
   - section     章节幕封
   - bigNumbers  数据大字报
   - quoteImage  左文右图 + callout
   - imageGrid   3-6 图网格
   - pipeline    3-6 步流程
   - bigQuote    金句页
   - compare     Before / After
   - textImage   正文 + 辅助图
   - also accepts Swiss layouts: statement, kpiTower, duoCompare, timeline, matrix, fourCards, imageHero, textGrid
   - article     Dense multi-column text page
   - dataSheet   Table + notes page
   - chart       Native PPT chart page
   - dashboard   KPI + chart dashboard
   - blocks/charts/tables can be inserted only when x/y/w/h are explicit; unpositioned data blocks are skipped
   - shared media layouts: media, mediaGrid, gallery; user images determine slot count, explicit charts render as charts, otherwise IMAGE SLOT placeholders are shown
   - shared rich layouts: agenda, caseStudy, pyramid, radial, roadmap, swimlane
   - multi-item pages auto-adapts column count; set columnsCount only when fixed columns are required
   - avoid long runs of the same layout; rotate page types for visual variety
   - repeated layouts only emit suggestions by default; use --diversify-layouts --write-normalized-spec to change layout and sync JSON
   - every card/point should include body/desc/note; title-only pages trigger warnings
   - closing     收束页
   ============================================================================ */

const deckSpec = {
  title: '[必填] PPT 标题',
  subtitle: '[必填] 副标题 / 演讲场景',
  author: '[必填] 作者 / 团队',
  style: 'magazine',
  theme: 'ink',

  slides: [
    {
      layout: 'cover',
      chromeLeft: 'Magazine Note / [必填] 主题',
      chromeRight: 'MG / 01 / NN',
      kicker: 'A Talk / 2026',
      title: '[必填] 主标题',
      subtitle: '[必填] 一句话副标题，保留杂志封面留白',
      author: '[必填] 作者 / 团队',
      foot: '[必填] 场景 / 日期 / 来源',
    },
    {
      layout: 'section',
      chromeLeft: 'Act I / Context',
      kicker: 'Act I',
      title: '[必填] 章节标题',
      subtitle: '[必填] 章节转场短句',
    },
    {
      layout: 'bigNumbers',
      chromeLeft: 'Proof / Numbers',
      kicker: 'Proof',
      title: '[必填] 先看数字',
      subtitle: '[选填] 用一句话解释这组数据',
      items: [
        { label: 'Duration', value: '64', unit: '天', note: '[必填] 指标解释' },
        { label: 'Lines', value: '110K+', note: '[必填] 指标解释' },
        { label: 'Commits', value: '608', note: '[必填] 指标解释' },
        { label: 'Platforms', value: '9', note: '[必填] 指标解释' },
        { label: 'Providers', value: '19', note: '[必填] 指标解释' },
        { label: 'Stars', value: '5K+', note: '[必填] 指标解释' },
      ],
    },
    {
      layout: 'quoteImage',
      chromeLeft: 'Story / Contrast',
      kicker: 'But',
      title: '[必填] 反差标题',
      body: '[必填] 1-2 句叙事正文，适合身份反差、案例背景或关键判断。',
      callout: '[选填] 一句强调判断，像杂志里的引文框。',
      image: { path: 'images/04-story-16x10.png', caption: 'Story visual / 16:10' },
    },
    {
      layout: 'imageGrid',
      chromeLeft: 'Evidence / Gallery',
      kicker: 'Evidence',
      title: '[必填] 证据网格',
      images: [
        { path: 'images/evidence-01.png', caption: '[必填] 图注 01' },
        { path: 'images/evidence-02.png', caption: '[必填] 图注 02' },
        { path: 'images/evidence-03.png', caption: '[必填] 图注 03' },
      ],
    },
    {
      layout: 'pipeline',
      chromeLeft: 'Workflow / Steps',
      kicker: 'Workflow',
      title: '[必填] 一条流水线',
      steps: [
        { title: 'Draft', desc: '[必填] 起草' },
        { title: 'Polish', desc: '[必填] 润色' },
        { title: 'Morph', desc: '[必填] 改写' },
        { title: 'Illustrate', desc: '[必填] 配图' },
        { title: 'Distribute', desc: '[必填] 分发' },
      ],
    },
    {
      layout: 'compare',
      chromeLeft: 'Before / After',
      kicker: 'Compare',
      title: '[必填] 两种状态的对比',
      before: {
        label: 'Before',
        title: '[必填] 旧状态',
        items: [
          { icon: 'circle-alert', text: '[必填] 痛点 1' },
          { icon: 'circle-x', text: '[必填] 痛点 2' },
          { icon: 'minus-circle', text: '[必填] 痛点 3' },
        ],
      },
      after: {
        label: 'After',
        title: '[必填] 新状态',
        items: [
          { icon: 'check-circle', text: '[必填] 改变 1' },
          { icon: 'arrow-right-circle', text: '[必填] 改变 2' },
          { icon: 'sparkles', text: '[必填] 改变 3' },
        ],
      },
    },
    {
      layout: 'article',
      chromeLeft: 'Dense / Article',
      kicker: 'Analysis',
      title: '[Required] Dense information page',
      subtitle: '[Optional] Use this when normal pages are too sparse.',
      sections: [
        { title: 'Context', body: '[Required] 2-3 sentences. Keep each section compact.' },
        { title: 'Finding', body: '[Required] Evidence, mechanism, or conclusion.' },
        { title: 'Risk', body: '[Required] State what changes if this is ignored.' },
        { title: 'Action', body: '[Required] Concrete next move and owner.' },
      ],
    },
    {
      layout: 'dataSheet',
      chromeLeft: 'Data / Table',
      kicker: 'Data Sheet',
      title: '[Required] Editable table page',
      body: '[Optional] One sentence explaining the table.',
      table: {
        headers: ['Metric', 'Current', 'Target', 'Note'],
        rows: [
          ['Activation', '42%', '60%', 'Main bottleneck'],
          ['Retention', '31%', '45%', 'Needs cohort split'],
          ['Cost', '$18', '$12', 'Automation opportunity'],
          ['Cycle', '9d', '5d', 'Process redesign'],
        ],
        caption: 'Native PPT table / editable cells',
      },
      notes: [
        { title: 'Read first', body: 'Highlight the row that drives the decision.' },
        { title: 'Then compare', body: 'Use target gap instead of raw values.' },
      ],
    },
    {
      layout: 'chart',
      chromeLeft: 'Data / Chart',
      kicker: 'Trend',
      title: '[Required] Native chart page',
      chart: {
        type: 'line',
        title: 'Monthly trend',
        labels: ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun'],
        series: [
          { name: 'Actual', values: [12, 18, 21, 25, 31, 38] },
          { name: 'Target', values: [15, 18, 22, 28, 34, 40] },
        ],
      },
      insights: [
        { title: 'Inflection', body: 'Growth accelerated after March.' },
        { title: 'Gap', body: 'Target is close but not yet stable.' },
      ],
    },
    {
      layout: 'dashboard',
      chromeLeft: 'Data / Dashboard',
      kicker: 'Dashboard',
      title: '[Required] KPI dashboard with charts',
      metrics: [
        { label: 'Revenue', value: '$2.4M', note: '+18% QoQ' },
        { label: 'Users', value: '128K', note: '+9% MoM' },
        { label: 'Cost', value: '$18', note: '-12%' },
        { label: 'NPS', value: '54', note: '+6 pts' },
      ],
      charts: [
        { type: 'column', title: 'Quarter result', labels: ['Q1', 'Q2', 'Q3', 'Q4'], values: [24, 32, 38, 51] },
        { type: 'pie', title: 'Channel mix', labels: ['Direct', 'Partner', 'Paid'], values: [46, 31, 23], showLegend: true },
      ],
    },    {
      layout: 'bigQuote',
      chromeLeft: 'Takeaway',
      kicker: 'Takeaway',
      quote: '[必填] 一句可被记住的结论。',
      body: '[选填] 英文原句 / 来源 / 行动建议',
    },
  ],
};

const outFile = path.join(__dirname, 'outputs', 'deck-magazine.pptx');

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


