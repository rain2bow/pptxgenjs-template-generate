const fs = require('node:fs');
const path = require('node:path');
const process = require('node:process');
const { spawnSync } = require('node:child_process');

if (require.main === module) {
  main();
}

function main() {
  const args = parseArgs(process.argv.slice(2));
  if (!args.pptx) {
    fail('Usage: node scripts/test-inherited-template.js --pptx template.pptx [--out-dir outputs/inherited-template-test] [--no-repair-z-order]');
  }

  const pptxPath = path.resolve(args.pptx);
  const outDir = path.resolve(args.outDir || 'outputs/inherited-template-test');
  const slotsPath = path.join(outDir, 'template.slots.json');
  const filledPath = path.join(outDir, 'filled-inherited.pptx');
  const reportJsonPath = path.join(outDir, 'test-report.json');
  const reportMdPath = path.join(outDir, 'test-report.md');

  fs.mkdirSync(outDir, { recursive: true });

  const steps = [];
  steps.push(runNode('extract-inherited-template.js', ['--pptx', pptxPath, '--out', slotsPath]));
  const fillArgs = ['--pptx', pptxPath, '--plan', slotsPath, '--out', filledPath];
  if (args.repairZOrder === false) fillArgs.push('--no-repair-z-order');
  steps.push(runNode('fill-inherited-template.js', fillArgs));
  const originalNative = runNode('validate-pptx-native.js', [pptxPath]);
  const originalLayout = runNode('validate-pptx-layout.js', [pptxPath], { allowFailure: true });
  const generatedNative = runNode('validate-pptx-native.js', [filledPath]);
  const generatedLayout = runNode('validate-pptx-layout.js', [filledPath]);
  steps.push(originalNative, originalLayout, generatedNative, generatedLayout);

  const slots = JSON.parse(fs.readFileSync(slotsPath, 'utf8'));
  const report = {
    source: pptxPath,
    outDir,
    repairZOrder: args.repairZOrder !== false,
    slides: slots.slides.length,
    textSlots: slots.slides.reduce((sum, slide) => sum + slide.textSlots.length, 0),
    files: {
      slots: slotsPath,
      filled: filledPath,
      reportJson: reportJsonPath,
      reportMd: reportMdPath,
    },
    checks: {
      originalNative: summarize(originalNative),
      originalLayout: summarize(originalLayout),
      generatedNative: summarize(generatedNative),
      generatedLayout: summarize(generatedLayout),
    },
  };

  fs.writeFileSync(reportJsonPath, `${JSON.stringify(report, null, 2)}\n`, 'utf8');
  fs.writeFileSync(reportMdPath, markdownReport(report), 'utf8');

  const failed = steps.filter((step) => !step.allowFailure && step.status !== 0);
  console.log(`Wrote ${reportJsonPath}`);
  console.log(`Wrote ${reportMdPath}`);
  if (failed.length) {
    failed.forEach((step) => console.error(`Failed: ${step.name}`));
    process.exit(1);
  }
}

function parseArgs(argv) {
  const args = {};
  for (let i = 0; i < argv.length; i += 1) {
    const arg = argv[i];
    if (arg === '--pptx') args.pptx = argv[++i];
    else if (arg === '--out-dir') args.outDir = argv[++i];
    else if (arg === '--no-repair-z-order') args.repairZOrder = false;
    else if (arg === '--help' || arg === '-h') args.help = true;
  }
  return args;
}

function runNode(scriptName, args, options = {}) {
  const scriptPath = path.join(__dirname, scriptName);
  const result = spawnSync(process.execPath, [scriptPath, ...args], {
    cwd: path.resolve(__dirname, '..'),
    encoding: 'utf8',
    windowsHide: true,
  });
  const step = {
    name: scriptName,
    command: `node scripts/${scriptName} ${args.map(shellToken).join(' ')}`,
    status: result.status ?? 1,
    allowFailure: options.allowFailure === true,
    stdout: result.stdout || '',
    stderr: result.stderr || '',
  };
  if (step.stdout.trim()) process.stdout.write(step.stdout);
  if (step.stderr.trim()) process.stderr.write(step.stderr);
  return step;
}

function summarize(step) {
  return {
    status: step.status,
    passed: step.status === 0,
    command: step.command,
    stdout: step.stdout.trim(),
    stderr: step.stderr.trim(),
  };
}

function markdownReport(report) {
  const checks = report.checks;
  return `# Inherited PPTX Template Test\n\n` +
    `Source: ${report.source}\n\n` +
    `Slides: ${report.slides}\n\n` +
    `Text slots: ${report.textSlots}\n\n` +
    `Z-order repair: ${report.repairZOrder ? 'enabled' : 'disabled'}\n\n` +
    `## Files\n\n` +
    `- Slots JSON: ${report.files.slots}\n` +
    `- Filled PPTX: ${report.files.filled}\n` +
    `- JSON report: ${report.files.reportJson}\n\n` +
    `## Checks\n\n` +
    `| Check | Status |\n| --- | --- |\n` +
    `| Original native | ${statusLabel(checks.originalNative)} |\n` +
    `| Original layout | ${statusLabel(checks.originalLayout)} |\n` +
    `| Generated native | ${statusLabel(checks.generatedNative)} |\n` +
    `| Generated layout | ${statusLabel(checks.generatedLayout)} |\n\n` +
    `## Generated Layout Output\n\n` +
    codeBlock(checks.generatedLayout.stdout || checks.generatedLayout.stderr || 'No output.') + '\n';
}

function statusLabel(check) {
  return check.passed ? 'PASS' : `FAIL (${check.status})`;
}

function codeBlock(text) {
  return `\`\`\`text\n${String(text || '').trim()}\n\`\`\``;
}

function shellToken(value) {
  const text = String(value);
  return /\s/.test(text) ? JSON.stringify(text) : text;
}

function fail(message) {
  console.error(message);
  process.exit(2);
}
