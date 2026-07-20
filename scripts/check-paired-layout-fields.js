#!/usr/bin/env node
'use strict';

const fs = require('node:fs');
const os = require('node:os');
const path = require('node:path');
const { spawnSync } = require('node:child_process');
const JSZip = require('jszip');

const root = path.resolve(__dirname, '..');
const generator = path.join(root, 'scripts', 'generate-pptx.js');
const imagePath = path.join(root, 'assets', 'logos', 'cmb-logo-lockup.png');
const tempDir = fs.mkdtempSync(path.join(os.tmpdir(), 'pptx-paired-fields-'));

main().catch((error) => {
  console.error(error.stack || error.message);
  process.exitCode = 1;
}).finally(() => fs.rmSync(tempDir, { recursive: true, force: true }));

async function main() {
  verifyDynamicStyleInstruction();

  const accepted = {
    title: 'Paired field rendering check',
    slides: [
      slide('text-statement', { body: '正文', callout: 'TEXT_STATEMENT_CALLOUT' }),
      slide('image-statement', { body: '正文', callout: 'IMAGE_STATEMENT_CALLOUT', images: [imagePath] }),
      slide('text-quote', { quote: '引用', body: '正文', caption: 'TEXT_QUOTE_CAPTION' }),
      slide('image-quote', { quote: '引用', body: '正文', caption: 'IMAGE_QUOTE_CAPTION', images: [imagePath] }),
      slide('text-swimlane', { stages: ['TEXT_STAGE_MARKER', '阶段乙'], items: swimlaneItems() }),
      slide('image-swimlane', { stages: ['IMAGE_STAGE_MARKER', '阶段乙'], items: swimlaneItems(), images: [imagePath] }),
      slide('image-grid', { highlightIndex: 1, items: items(), images: [imagePath] }),
      slide('image-roadmap', { highlightLast: true, items: items(), images: [imagePath] }),
    ],
  };
  for (const style of ['cmb', 'swiss', 'magazine']) {
    const acceptedOut = runDeck(`accepted-fields-${style}`, { ...accepted, style }, true);
    const xml = await slideXml(acceptedOut);
    ['TEXT_STATEMENT_CALLOUT', 'IMAGE_STATEMENT_CALLOUT', 'TEXT_QUOTE_CAPTION', 'IMAGE_QUOTE_CAPTION', 'TEXT_STAGE_MARKER', 'IMAGE_STAGE_MARKER']
      .forEach((marker) => assert(xml.includes(marker), `${style} generated PPTX is missing rendered marker ${marker}`));
  }

  const rejected = {
    title: 'Paired unsupported field check',
    style: 'cmb',
    theme: 'classic',
    slides: [
      slide('text-article', { conclusion: 'NOT_ALLOWED', items: items() }),
      slide('image-article', { conclusion: 'NOT_ALLOWED', items: items(), images: [imagePath] }),
    ],
  };
  const rejectedResult = runDeck('rejected-fields', rejected, false);
  assert(rejectedResult.status === 2, `unsupported paired fields should fail with exit 2; got ${rejectedResult.status}`);
  assert(rejectedResult.output.includes('counterpart "image-article"'), 'text-side error does not name image counterpart');
  assert(rejectedResult.output.includes('counterpart "text-article"'), 'image-side error does not name text counterpart');

  const pairedSource = fs.readFileSync(path.join(root, 'scripts', 'pptxgen', 'templates', 'paired-layouts.js'), 'utf8');
  ['stages', 'highlightIndex', 'highlightLast'].forEach((field) => assert(new RegExp(`\\b${field}\\b`).test(pairedSource), `paired renderer does not consume ${field}`));

  console.log('Paired layout field check passed: dynamic style instruction, symmetric validation, rendered markers in 3 styles, stages and highlight controls.');
}

function verifyDynamicStyleInstruction() {
  const skill = fs.readFileSync(path.join(root, 'SKILL.md'), 'utf8');
  assert(skill.includes('读取 `--style-guide` 的当前输出'), 'SKILL.md does not require the dynamic style guide');
  assert(skill.includes('所有已注册 style'), 'SKILL.md does not include registered plugin styles in the interaction');
  assert(!/展示 `cmb`、`swiss`、`magazine`/.test(skill), 'SKILL.md still hard-codes the three built-in styles');
}

function runDeck(name, spec, expectSuccess) {
  const specPath = path.join(tempDir, `${name}.json`);
  const outPath = path.join(tempDir, `${name}.pptx`);
  fs.writeFileSync(specPath, JSON.stringify(spec, null, 2), 'utf8');
  const result = spawnSync(process.execPath, [generator, '--spec', specPath, '--out', outPath], { cwd: root, encoding: 'utf8' });
  const output = `${result.stdout || ''}\n${result.stderr || ''}`;
  if (expectSuccess && result.status !== 0) throw new Error(`generation failed:\n${output}`);
  return expectSuccess ? outPath : { status: result.status, output };
}

async function slideXml(pptxPath) {
  const zip = await JSZip.loadAsync(fs.readFileSync(pptxPath));
  const names = Object.keys(zip.files).filter((name) => /^ppt\/slides\/slide\d+\.xml$/.test(name));
  return (await Promise.all(names.map((name) => zip.file(name).async('string')))).join('\n');
}

function slide(layout, extra) {
  return { layout, kicker: '验证', title: `${layout} 字段验证`, ...extra };
}

function items() {
  return [{ title: '要点一', body: '说明一' }, { title: '要点二', body: '说明二' }];
}

function swimlaneItems() {
  return [
    { title: '业务', body: '推进说明', items: ['准备', '执行'] },
    { title: '技术', body: '技术说明', items: ['设计', '上线'] },
  ];
}

function assert(condition, message) {
  if (!condition) throw new Error(message);
}
