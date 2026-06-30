#!/usr/bin/env node
const path = require('node:path');
const { buildDeck } = require('../scripts/generate-pptx.js');

/* ============================================================================
   China Merchants Bank / 招商银行风格 PPTX Seed

   视觉方向：白底、招商红、黑灰正文、直角网格、克制分割线、可编辑图表。
   适合银行经营分析、数字化转型、风险管理、客户经营、项目汇报。

   Logo：
   - logoHeader: 白底 PNG 图片+中英文文字 logo，用于每页页眉。
   - logoFull: 白底 PNG 图片+中英文文字 logo，用于封面或收束页。
   - 页眉区域整条保持白色，避免红色页面中只出现局部白底 logo。
   - 不在招商银行模板页眉使用 SVG logo。

   规则：
   - 使用 style: 'swiss' + theme: 'cmb'。
   - 不把 logo 做成整页背景图；logo 只作为品牌识别元素。
   - 业务页优先使用 statement / kpiTower / chart / dashboard / dataSheet / media / roadmap。
   - 所有图表和表格保持 PowerPoint 原生可编辑。
   ============================================================================ */

const deckSpec = {
  title: '招商银行经营汇报模板',
  subtitle: 'China Merchants Bank / Business Review',
  author: '招商银行 / 项目团队',
  style: 'swiss',
  theme: 'cmb',
  logoHeader: 'logos/cmb-logo-lockup.png',
  logoHeaderW: 1.72,
  logoHeaderH: 0.54,
  logoHeaderBand: true,
  logoHeaderBandH: 0.78,
  logoHeaderBandColor: 'FFFFFF',
  logoHeaderTextColor: '111111',
  headY: 1.12,
  logoFull: 'logos/cmb-logo-lockup.png',

  slides: [
    {
      layout: 'cover',
      chromeLeft: 'CMB / Business Review',
      chromeRight: 'CMB / 01',
      kicker: 'CHINA MERCHANTS BANK / 2026',
      title: '招商银行业务增长与数字化经营汇报',
      subtitle: '围绕客户经营、风险控制与效率提升的阶段性复盘',
      foot: '招商银行 / 内部汇报 / 2026',
    },
    {
      layout: 'statement',
      chromeLeft: 'Executive Summary / 01',
      kicker: 'Executive Summary',
      title: '以客户价值为核心，形成增长、风控与效率的闭环。',
      body: '本页用于放置全篇核心判断。建议用一句明确结论加一行依据，不要写成普通目录。',
    },
    {
      layout: 'kpiTower',
      chromeLeft: 'Performance / KPI',
      kicker: 'Performance Snapshot',
      title: '关键经营指标保持稳健改善',
      items: [
        { label: '零售客户增长', value: '+18%', valueNum: 18 },
        { label: '活跃客户提升', value: '+32%', valueNum: 32 },
        { label: '风险预警覆盖', value: '96%', valueNum: 96 },
        { label: '流程时效缩短', value: '-24%', valueNum: 24 },
      ],
    },
    {
      layout: 'media',
      chromeLeft: 'Customer / Insight',
      kicker: 'Customer Operation',
      title: '客户经营从单点触达转向分层运营',
      body: '通过客群分层、权益匹配与渠道协同，提升客户转化与长期价值。没有用户图片且显式提供 chart 时，右侧媒体区会插入可编辑图表；否则显示 IMAGE SLOT 占位符。',
      chart: {
        chartType: 'line',
        title: '客户活跃趋势',
        labels: ['Q1', 'Q2', 'Q3', 'Q4'],
        values: [42, 51, 63, 78],
        showValue: true,
      },
      items: [
        { icon: 'users', title: '客群分层', body: '按资产、行为与生命周期拆分运营策略。' },
        { icon: 'scan-search', title: '信号识别', body: '捕捉高价值触点与潜在流失风险。' },
        { icon: 'workflow', title: '渠道协同', body: '联动 App、网点、远程服务与客户经理。' },
      ],
    },
    {
      layout: 'dashboard',
      chromeLeft: 'Data / Dashboard',
      kicker: 'Risk & Efficiency Dashboard',
      title: '风险与效率指标联动监控',
      metrics: [
        { label: '风险命中', value: '87%' },
        { label: '处置时效', value: '2.4h' },
        { label: '自动化率', value: '68%' },
        { label: '复核准确', value: '95%' },
        { label: '客户满意', value: '91' },
      ],
      charts: [
        { chartType: 'column', title: '风险事件处置量', labels: ['1月', '2月', '3月', '4月'], values: [120, 148, 176, 214] },
        { chartType: 'doughnut', title: '渠道结构', labels: ['App', '网点', '远程', '客户经理'], values: [46, 18, 21, 15], showLegend: true },
      ],
    },
    {
      layout: 'roadmap',
      chromeLeft: 'Roadmap / Next',
      kicker: 'Implementation Roadmap',
      title: '下一阶段推进路径',
      steps: [
        { label: '01', title: '统一指标', body: '明确经营、风险和体验的核心口径。' },
        { label: '02', title: '试点验证', body: '选择重点客群和重点分行进行闭环验证。' },
        { label: '03', title: '规模推广', body: '沉淀可复制流程并接入自动化运营。' },
        { label: '04', title: '持续评估', body: '按月复盘指标并优化模型和策略。' },
      ],
    },
    {
      layout: 'dataSheet',
      chromeLeft: 'Table / Governance',
      kicker: 'Governance Table',
      title: '重点事项跟踪表',
      table: {
        headers: ['事项', '当前状态', '下一步', '负责人'],
        rows: [
          ['客户分层模型', '试点完成', '扩大样本验证', '零售团队'],
          ['风险预警规则', '规则上线', '接入处置闭环', '风控团队'],
          ['运营看板', '原型评审', '联调数据接口', '数据团队'],
          ['网点协同流程', '方案确认', '分行试运行', '运营团队'],
        ],
        caption: 'Native PPT table / editable cells',
      },
      notes: [
        { title: '口径一致', body: '所有事项应绑定统一指标和责任人。' },
        { title: '节奏明确', body: '周跟踪、月复盘、季度校准。' },
      ],
    },
    {
      layout: 'closing',
      chromeLeft: 'Closing / CMB',
      kicker: 'Takeaway',
      title: '稳健经营，持续创造客户价值。',
      subtitle: 'CHINA MERCHANTS BANK',
    },
  ],
};

const outFile = path.join(__dirname, 'outputs', 'deck-cmb.pptx');

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