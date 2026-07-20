#!/usr/bin/env node
const fs = require('fs');
const os = require('os');
const path = require('path');
const { spawnSync } = require('child_process');
const { canonicalLayoutNames, layoutDefinition } = require('./pptxgen/layout-schema');
const { exampleSlide } = require('./pptxgen/layout-examples');

const MEDIA_SLOT_LAYOUTS = canonicalLayoutNames().filter((name) => name.startsWith('image-'));
const TEXT_LAYOUTS = canonicalLayoutNames().filter((name) => name.startsWith('text-'));

const spec = {
  title: 'Media slot error check',
  style: 'magazine',
  theme: 'ink',
  slides: [
    ...MEDIA_SLOT_LAYOUTS.map((layout) => {
      const slide = exampleSlide(layout);
      delete slide.images;
      return slide;
    }),
    ...TEXT_LAYOUTS.map((layout) => ({ ...exampleSlide(layout), images: ['missing-image.png'] })),
  ],
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
  return !output.includes(`slide ${slideNumber} layout "${layout}" requires image content. Provide "images", or change layout to its field-compatible text counterpart "${layoutDefinition(layout).counterpart}"`);
});

const wrongTextLayouts = TEXT_LAYOUTS.filter((layout, index) => {
  const slideNumber = MEDIA_SLOT_LAYOUTS.length + index + 1;
  return !output.includes(`slide ${slideNumber} layout "${layout}" does not accept images. Change layout to its field-compatible image counterpart "${layoutDefinition(layout).counterpart}"`);
});

if (missing.length || wrongTextLayouts.length) {
  if (missing.length) console.error(`Missing media slot error for layout(s): ${missing.join(', ')}`);
  if (wrongTextLayouts.length) console.error(`Missing text-to-image counterpart error for layout(s): ${wrongTextLayouts.join(', ')}`);
  process.exit(1);
}

console.log(`Media counterpart error check passed: ${MEDIA_SLOT_LAYOUTS.length} image layout(s), ${TEXT_LAYOUTS.length} text layout(s).`);
