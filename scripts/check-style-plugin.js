#!/usr/bin/env node
'use strict';

const fs = require('node:fs');
const os = require('node:os');
const path = require('node:path');
const { spawnSync } = require('node:child_process');

const root = path.resolve(__dirname, '..');
const pluginRoot = path.join(__dirname, 'fixtures', 'style-plugins');
const tempDir = fs.mkdtempSync(path.join(os.tmpdir(), 'pptx-style-plugin-'));
const outPath = path.join(tempDir, 'registry-test.pptx');
const examplePath = path.join(tempDir, 'registry-test-layouts.md');
const env = { ...process.env, PPTXGEN_STYLE_PATHS: pluginRoot };

try {
  const guide = run(['scripts/generate-pptx.js', '--style-guide'], env);
  assertIncludes(guide, 'registry-test · 注册机制测试风格', 'style guide did not discover the additive plugin');
  assertIncludes(guide, 'skipped invalid style plugin', 'invalid plugin was not isolated with a warning');

  run(['scripts/generate-pptx.js', '--layout-examples', 'registry-test', '--out', examplePath], env);
  if (!fs.existsSync(examplePath)) throw new Error('plugin layout examples were not generated');

  run(['scripts/generate-pptx.js', '--sample', '--sample-style', 'registry-test', '--out', outPath], env);
  run(['scripts/validate-pptx-native.js', outPath], env);
  run(['scripts/validate-pptx-layout.js', outPath], env);

  const isolatedGuide = run(['scripts/generate-pptx.js', '--style-guide'], { ...process.env, PPTXGEN_STYLE_PATHS: '' });
  if (isolatedGuide.includes('registry-test')) throw new Error('fixture plugin leaked into the default registry');

  console.log('Additive style plugin check passed: discovery, examples, generation, validation, isolation.');
} finally {
  fs.rmSync(tempDir, { recursive: true, force: true });
}

function run(args, childEnv) {
  const result = spawnSync(process.execPath, args, { cwd: root, env: childEnv, encoding: 'utf8' });
  if (result.status !== 0) throw new Error(`command failed (${args.join(' ')}):\n${result.stdout || ''}\n${result.stderr || ''}`);
  return `${result.stdout || ''}\n${result.stderr || ''}`;
}

function assertIncludes(text, expected, message) {
  if (!text.includes(expected)) throw new Error(`${message}; expected ${JSON.stringify(expected)} in:\n${text}`);
}
