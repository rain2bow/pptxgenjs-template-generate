#!/usr/bin/env node
const path = require('node:path');
const { buildDeck } = require('../scripts/generate-pptx.js');

/* ============================================================================
   Guizang PPTXGenJS Template / Style B
   瑞士国际主义 / Native PPTX Seed

   这个文件对应原 HTML 技能里的 assets/template-swiss.html，但输出为
   pptxgenjs 原生可编辑 PPTX：
   1. 复制本文件到项目目录，例如 项目/XXX/pptx/deck.js
   2. 只修改 deckSpec：标题、主题、slides、图片路径、文案
   3. 运行 node deck.js

   硬规则：
   - 只用 pptxgenjs 原生生成；不要截图、不要 HTML 转图片、不要整页背景图。
   - 全程无衬线、直角、纯色、无阴影、无渐变、单一 accent。
   - 色块/图片先画，文字最后画；内容过多时拆页，不要硬塞。
   - readable projection fonts are preferred; split dense content instead of shrinking small text
   - Swiss 页面的标题默认左上内容轴，不要做普通 PPT 的居中标题页。

   可选主题：
   - ikb     / 克莱因蓝，默认
   - lemon   / 柠檬黄
   - green   / 柠檬绿
   - orange  / 安全橙
   - cmb     / 招商银行红

   支持版式：
   - cover       满屏 accent 封面
   - statement   极简论点页
   - kpiTower    KPI 高度对比
   - duoCompare  Before / After
   - timeline    横向时间线
   - matrix      8-12 项矩阵
   - fourCards   1-8 项自适应卡片网格
   - imageHero   顶部 21:9 主图 + KPI
   - also accepts magazine layouts: section, bigNumbers, quoteImage, imageGrid, pipeline, bigQuote, compare, textImage, article
   - textGrid    Dense 3-column text grid
   - dataSheet   Table + side notes
   - chart       Native PPT chart page
   - dashboard   KPI strip + chart dashboard
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
  style: 'swiss',
  theme: 'ikb',

  slides: [
    {
      layout: 'cover',
      chromeLeft: 'Swiss Field Note / [必填] 主题',
      chromeRight: 'SS / 01 / NN',
      kicker: 'Swiss Field Note / 2026',
      title: '[必填] 主标题',
      subtitle: '[必填] 一句话定调，控制在 18-28 个中文字符',
      foot: '[必填] 作者 / 日期 / 场景',
    },
    {
      layout: 'statement',
      chromeLeft: 'Thesis / 01',
      kicker: 'Thesis / 01',
      title: '[必填] 一个强论点。',
      body: '[必填] 1-2 行解释，不要居中，不要变成普通正文页。',
    },
    {
      layout: 'kpiTower',
      chromeLeft: 'Proof / Numbers',
      kicker: 'Proof / Numbers',
      title: '[必填] 用真实数字支撑',
      items: [
        { label: 'Metric 01', value: '64d', valueNum: 64 },
        { label: 'Metric 02', value: '110K+', valueNum: 110 },
        { label: 'Metric 03', value: '9', valueNum: 9 },
        { label: 'Metric 04', value: '608', valueNum: 608 },
      ],
    },
    {
      layout: 'duoCompare',
      chromeLeft: 'Before / After',
      kicker: 'Before / After',
      title: '[必填] 从旧模式到新模式',
      before: {
        label: 'Before',
        title: '[必填] 旧模式',
        items: [
          { icon: 'chart-line', text: '[必填] 要点 1' },
          { icon: 'circle-alert', text: '[必填] 要点 2' },
          { icon: 'minus-circle', text: '[必填] 要点 3' },
        ],
      },
      after: {
        label: 'After',
        title: '[必填] 新模式',
        items: [
          { icon: 'chart-line', text: '[必填] 要点 1' },
          { icon: 'circle-alert', text: '[必填] 要点 2' },
          { icon: 'minus-circle', text: '[必填] 要点 3' },
        ],
      },
    },
    {
      layout: 'timeline',
      chromeLeft: 'Timeline / Process',
      kicker: 'Process / Timeline',
      title: '[必填] 一条清晰时间线',
      items: [
        { label: '01', title: '阶段一', desc: '[必填] 短句' },
        { label: '02', title: '阶段二', desc: '[必填] 短句' },
        { label: '03', title: '阶段三', desc: '[必填] 短句' },
        { label: '04', title: '阶段四', desc: '[必填] 短句' },
      ],
    },
    {
      layout: 'matrix',
      chromeLeft: 'Matrix / System',
      kicker: 'Matrix / 12 Cells',
      title: '[必填] 一个系统被拆成若干模块',
      highlightLast: true,
      items: [
        { title: '模块 01' }, { title: '模块 02' }, { title: '模块 03' }, { title: '模块 04' },
        { title: '模块 05' }, { title: '模块 06' }, { title: '模块 07' }, { title: '模块 08' },
        { title: '模块 09' }, { title: '模块 10' }, { title: '模块 11' }, { title: '关键模块' },
      ],
      heroStat: { value: '12', label: 'registered modules / native PPTX cells' },
    },
    {
      layout: 'fourCards',
      chromeLeft: 'Mechanism / 04',
      kicker: 'Mechanism / 04',
      title: '[必填] 四个等权模块',
      items: [
        { title: '模块一', desc: '[必填] 简短说明' },
        { title: '模块二', desc: '[必填] 简短说明' },
        { title: '模块三', desc: '[必填] 简短说明' },
        { title: '模块四', desc: '[必填] 简短说明' },
      ],
    },
    {
      layout: 'imageHero',
      chromeLeft: 'Case / Image Hero',
      title: '[必填] 图像证据',
      body: '[必填] 1-2 行解释这张图为什么重要，不要重复标题。',
      image: { path: 'images/08-hero-21x9.png', caption: 'Hero visual / 21:9' },
      items: [
        { label: 'Metric 01', value: '12x', note: '[必填] 解释' },
        { label: 'Metric 02', value: '3.4h', note: '[必填] 解释' },
        { label: 'Metric 03', value: '100%', note: '[必填] 解释' },
      ],
    },
    {
      layout: 'textGrid',
      chromeLeft: 'System / Dense Grid',
      kicker: 'System / Dense Grid',
      title: '[Required] Dense Swiss text grid',
      subtitle: '[Optional] For 6-9 compact points without turning into tiny body text.',
      highlightIndex: 0,
      sections: [
        { label: '01', title: 'Signal', body: 'Most important fact or decision.' },
        { label: '02', title: 'Constraint', body: 'Boundary condition or hard limit.' },
        { label: '03', title: 'Dependency', body: 'What must be ready first.' },
        { label: '04', title: 'Risk', body: 'Failure mode to watch.' },
        { label: '05', title: 'Metric', body: 'How success will be measured.' },
        { label: '06', title: 'Owner', body: 'Team or role accountable.' },
      ],
    },
    {
      layout: 'dataSheet',
      chromeLeft: 'Data / Table',
      kicker: 'Data / Table',
      title: '[Required] Swiss editable table',
      table: {
        headers: ['Area', 'Now', 'Next', 'Status'],
        rows: [
          ['Research', 'Manual', 'Agent assisted', 'Ready'],
          ['Design', 'Static', 'Tokenized', 'Building'],
          ['Delivery', 'One-off', 'Reusable', 'Pilot'],
          ['Review', 'Visual only', 'Native validation', 'Ready'],
        ],
        caption: 'Native PPT table / editable cells',
      },
      notes: [
        { title: 'Decision', body: 'Use the status column to drive discussion.' },
        { title: 'Scope', body: 'Keep columns short; split page if needed.' },
      ],
    },
    {
      layout: 'chart',
      chromeLeft: 'Data / Chart',
      kicker: 'Data / Chart',
      title: '[Required] Swiss native chart',
      chart: {
        type: 'column',
        title: 'Adoption by quarter',
        labels: ['Q1', 'Q2', 'Q3', 'Q4'],
        values: [18, 31, 47, 66],
        showValue: true,
      },
      insights: [
        { value: '+35', title: 'Growth', body: 'Q2 to Q4 adoption lift.' },
        { value: '66%', title: 'Latest', body: 'Latest quarter level.' },
      ],
    },
    {
      layout: 'dashboard',
      chromeLeft: 'Data / Dashboard',
      kicker: 'Data / Dashboard',
      title: '[Required] Swiss KPI dashboard',
      metrics: [
        { label: 'QUALITY', value: '92%' },
        { label: 'COST', value: '-18%' },
        { label: 'CYCLE', value: '5d' },
        { label: 'USAGE', value: '128K' },
        { label: 'NPS', value: '54' },
      ],
      charts: [
        { type: 'line', title: 'Usage trend', labels: ['W1', 'W2', 'W3', 'W4'], values: [24, 38, 44, 61] },
        { type: 'doughnut', title: 'Work mix', labels: ['Build', 'Review', 'Ops'], values: [52, 28, 20], showLegend: true },
      ],
    },    {
      layout: 'closing',
      chromeLeft: 'Closing',
      kicker: 'Takeaway',
      title: '[必填] 收束金句。',
      subtitle: '[选填] 英文原句 / 行动建议',
    },
  ],
};

const outFile = path.join(__dirname, 'outputs', 'deck-swiss.pptx');

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
