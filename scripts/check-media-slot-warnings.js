#!/usr/bin/env node
const fs = require('fs');
const os = require('os');
const path = require('path');
const { spawnSync } = require('child_process');

const MEDIA_SLOT_LAYOUTS = [
  'image-statement',
  'image-quote',
  'image-text',
  'image-feature',
  'image-grid',
  'image-hero',
  'image-case-study',
];

const spec = {
  title: 'Media slot error check',
  style: 'magazine',
  theme: 'ink',
  slides: MEDIA_SLOT_LAYOUTS.map((layout) => ({
    layout,
    title: `${layout} without media`,
    body: 'This slide intentionally omits images and charts.',
    caseTitle: layout === 'image-case-study' ? 'Case without media' : undefined,
    items: ['image-feature', 'image-grid', 'image-hero', 'image-case-study'].includes(layout)
      ? [{ label: 'Metric', value: '1', note: 'No media asset' }]
      : undefined,
  })),
};

const tempDir = fs.mkdtempSync(path.join(os.tmpdir(), 'pptx-media-slot-check-'));
const specPath = path.join(tempDir, 'missing-media.spec.json');
const outPath = path.join(tempDir, 'missing-media.pptx');
fs.writeFileSync(specPath, JSON.stringify(spec, null, 2), 'utf8');

const result = spawnSync(process.execPath, [
  path.join(__dirname, 'generate-pptx.js'),
  '--spec',
  specPath,
  '--out',
  outPath,
], {
  cwd: path.resolve(__dirname, '..'),
  encoding: 'utf8',
});

try {
  fs.rmSync(tempDir, { recursive: true, force: true });
} catch (_) {
  // Best-effort cleanup only; a failed cleanup should not hide validation results.
}

if (result.status === 0) {
  console.error('Expected media slot validation to fail, but generation succeeded.');
  process.exit(1);
}

const output = `${result.stdout || ''}\n${result.stderr || ''}`;
const missing = MEDIA_SLOT_LAYOUTS.filter((layout, index) => {
  const slideNumber = index + 1;
  return !output.includes(`slide ${slideNumber} layout "${layout}" requires at least 1 images entry; got 0`);
});

if (missing.length) {
  console.error(`Missing media slot error for layout(s): ${missing.join(', ')}`);
  process.exit(1);
}

console.log(`Media slot error check passed: ${MEDIA_SLOT_LAYOUTS.length} layout(s).`);
