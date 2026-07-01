#!/usr/bin/env node
const path = require('node:path');
const { buildDeck } = require('../scripts/generate-pptx.js');

const logoHeader = 'logos/cmb-logo-lockup.png';
const logoMark = 'logos/cmb-logo-mark.svg';

function item(n, title = `Point ${n}`) {
  return { icon: ['users', 'chart-line', 'shield-check', 'workflow', 'database', 'target'][n % 6], label: `0${n}`, title, body: 'Short supporting copy for density and spacing checks.' };
}

function metric(label, value, note = 'projection check') {
  return { label, value, note, valueNum: Number(String(value).replace(/[^0-9.-]/g, '')) || 1 };
}

function chart(title, chartType = 'column') {
  return { chartType, title, labels: ['Q1', 'Q2', 'Q3', 'Q4'], values: [32, 46, 58, 74], showValue: chartType !== 'line' };
}

function table() {
  return {
    headers: ['Workstream', 'Status', 'Owner', 'Next'],
    rows: [
      ['Retail growth', 'Pilot', 'Team A', 'Scale'],
      ['Risk signal', 'Live', 'Team B', 'Review'],
      ['Service flow', 'Design', 'Team C', 'Launch'],
      ['Data mart', 'Build', 'Team D', 'Verify'],
    ],
    caption: 'Native editable table',
  };
}

const deckSpec = {
  title: 'CMB All Supported Layouts',
  subtitle: 'Layout QA deck for China Merchants Bank style',
  author: 'Guizang PPTXGenJS Skill',
  style: 'cmb',
  theme: 'classic',
  logoHeader,
  logoMark,
  logoHeaderW: 1.58,
  logoHeaderH: 0.5,
  logoHeaderBandH: 0.82,
  headY: 1.06,
  slides: [
    { layout: 'cover', kicker: 'CMB LAYOUT QA / 01', title: 'CMB Cover', subtitle: 'Cover page with fixed SVG mark position.' },
    { layout: 'section', kicker: 'CMB LAYOUT QA / 02', title: 'Section Divider', subtitle: 'Accent page for chapter transitions.' },
    { layout: 'statement', kicker: 'statement', title: 'A single executive statement anchors the narrative.', body: 'Use this page for a strong conclusion, thesis, or management decision.', callout: 'Stable\nReadable\nAligned' },
    { layout: 'kpiTower', kicker: 'kpiTower', title: 'Four key performance cards', items: [metric('Growth', '+18%'), metric('Active', '+32%'), metric('Coverage', '96%'), metric('Cycle', '-24%')] },
    { layout: 'bigNumbers', kicker: 'bigNumbers', title: 'Big number compatibility page', items: [metric('AUM', '128'), metric('NPS', '91'), metric('Branches', '42'), metric('Risk', '7')] },
    { layout: 'dashboard', kicker: 'dashboard', title: 'KPI strip with dual editable charts', metrics: [metric('Hit', '87%'), metric('SLA', '2.4h'), metric('Auto', '68%'), metric('Accuracy', '95%')], charts: [chart('Volume by quarter', 'column'), chart('Channel share', 'doughnut')] },
    { layout: 'chart', kicker: 'chart', title: 'Single dominant chart with insights', chart: chart('Customer activity trend', 'line'), insights: [item(1, 'Activity'), item(2, 'Conversion'), item(3, 'Retention')] },
    { layout: 'dataSheet', kicker: 'dataSheet', title: 'Editable table with side notes', table: table(), notes: [item(1, 'Owner'), item(2, 'Cadence'), item(3, 'Risk')] },
    { layout: 'media', kicker: 'media', title: 'Media area with side explanation', body: 'The media slot can hold a chart, image, or placeholder.', chart: chart('Media chart', 'bar'), items: [item(1, 'Signal'), item(2, 'Action'), item(3, 'Review')] },
    { layout: 'mediaGrid', kicker: 'mediaGrid', title: 'Adaptive media grid', captions: [item(1, 'Slot A'), item(2, 'Slot B'), item(3, 'Slot C'), item(4, 'Slot D')], charts: [chart('A'), chart('B', 'line'), chart('C', 'doughnut'), chart('D', 'column')] },
    { layout: 'gallery', kicker: 'gallery', title: 'Gallery alias page', captions: [item(1, 'Image A'), item(2, 'Image B'), item(3, 'Image C')], charts: [chart('Gallery A'), chart('Gallery B', 'line'), chart('Gallery C', 'doughnut')] },
    { layout: 'imageGrid', kicker: 'imageGrid', title: 'Image grid compatibility page', captions: [item(1, 'Evidence A'), item(2, 'Evidence B'), item(3, 'Evidence C'), item(4, 'Evidence D')], charts: [chart('Evidence A'), chart('Evidence B', 'line'), chart('Evidence C', 'doughnut'), chart('Evidence D', 'column')] },
    { layout: 'imageHero', kicker: 'imageHero', title: 'Hero media with metrics', body: 'The top area is reserved for evidence media.', chart: chart('Hero evidence', 'area'), items: [metric('Reach', '46%'), metric('Trust', '91'), metric('Time', '2.4h')] },
    { layout: 'quoteImage', kicker: 'quoteImage', title: 'Quote image compatibility', quote: 'Customer value compounds when signals, service and risk are coordinated.', body: 'Hero media area remains editable.' },
    { layout: 'textImage', kicker: 'textImage', title: 'Text and image compatibility', body: 'When no image is provided, this behaves like a clean statement page for text-heavy messages.' },
    { layout: 'compare', kicker: 'compare', title: 'Before and after comparison', before: { title: 'Before', items: [item(1, 'Manual'), item(2, 'Fragmented'), item(3, 'Delayed')] }, after: { title: 'After', items: [item(4, 'Automated'), item(5, 'Unified'), item(6, 'Visible')] } },
    { layout: 'duoCompare', kicker: 'duoCompare', title: 'Two-column comparison alias', before: { title: 'Current', items: [item(1, 'Many tools'), item(2, 'Weak loop')] }, after: { title: 'Target', items: [item(3, 'One view'), item(4, 'Closed loop')] } },
    { layout: 'timeline', kicker: 'timeline', title: 'Timeline layout', steps: [item(1, 'Plan'), item(2, 'Pilot'), item(3, 'Scale'), item(4, 'Review')] },
    { layout: 'pipeline', kicker: 'pipeline', title: 'Pipeline alias layout', steps: [item(1, 'Input'), item(2, 'Model'), item(3, 'Review'), item(4, 'Output'), item(5, 'Feedback')] },
    { layout: 'roadmap', kicker: 'roadmap', title: 'Roadmap layout', steps: [item(1, 'Q1'), item(2, 'Q2'), item(3, 'Q3'), item(4, 'Q4')], highlightLast: true },
    { layout: 'textGrid', kicker: 'textGrid', title: 'Text grid with multiple compact points', sections: [item(1), item(2), item(3), item(4), item(5), item(6)], highlightIndex: 1 },
    { layout: 'article', kicker: 'article', title: 'Article compatibility grid', sections: [item(1, 'Context'), item(2, 'Finding'), item(3, 'Impact'), item(4, 'Action')] },
    { layout: 'fourCards', kicker: 'fourCards', title: 'Adaptive card layout', items: [item(1), item(2), item(3), item(4), item(5)] },
    { layout: 'matrix', kicker: 'matrix', title: 'Matrix layout', items: [item(1), item(2), item(3), item(4), item(5), item(6), item(7), item(8)], highlightIndex: 2 },
    { layout: 'agenda', kicker: 'agenda', title: 'Agenda and section navigation', items: [item(1, 'Overview'), item(2, 'Performance'), item(3, 'Risk'), item(4, 'Roadmap')] },
    { layout: 'caseStudy', kicker: 'caseStudy', title: 'Case study page', caseTitle: 'Retail customer operation pilot', body: 'A compact story area pairs evidence media with quantified business results.', chart: chart('Case evidence', 'line'), metrics: [metric('Lift', '+18%'), metric('Reach', '42k'), metric('SLA', '2.4h')] },
    { layout: 'pyramid', kicker: 'pyramid', title: 'Layered structure page', layers: [item(1, 'Data'), item(2, 'Model'), item(3, 'Process'), item(4, 'Governance'), item(5, 'Value')], note: 'Layer order and width should remain balanced.' },
    { layout: 'radial', kicker: 'radial', title: 'Radial relationship page', center: 'Customer Value', items: [item(1, 'Data'), item(2, 'Risk'), item(3, 'Service'), item(4, 'Product'), item(5, 'Channel'), item(6, 'Growth')] },
    { layout: 'swimlane', kicker: 'swimlane', title: 'Swimlane execution matrix', stages: ['Now', 'Next', 'Later'], lanes: [ { title: 'Retail', items: ['Segment', 'Offer', 'Review'] }, { title: 'Risk', items: ['Signal', 'Rule', 'Monitor'] }, { title: 'Data', items: ['Clean', 'Model', 'Publish'] } ] },
    { layout: 'bigQuote', kicker: 'bigQuote', title: 'Big quote compatibility', quote: 'Make the operating system visible before scaling the operating rhythm.', body: 'A closing thought inside the content sequence.' },
    { layout: 'closing', kicker: 'THANK YOU', title: 'CMB Closing', subtitle: 'All supported CMB layouts checked.' },
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