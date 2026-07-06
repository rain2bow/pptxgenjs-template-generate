#!/usr/bin/env node
const { normalizeSpec } = require('./pptxgen/engine');

const MEDIA_SLOT_LAYOUTS = [
  'statement',
  'media',
  'mediaGrid',
  'gallery',
  'imageGrid',
  'imageHero',
  'quoteImage',
  'textImage',
  'caseStudy',
];

const spec = {
  title: 'Media slot warning check',
  style: 'magazine',
  theme: 'ink',
  slides: MEDIA_SLOT_LAYOUTS.map((layout) => ({
    layout,
    title: `${layout} without media`,
    body: 'This slide intentionally omits images and charts.',
    caseTitle: layout === 'caseStudy' ? 'Case without media' : undefined,
    items: layout === 'imageHero' || layout === 'caseStudy'
      ? [{ label: 'Metric', value: '1', note: 'No media asset' }]
      : undefined,
  })),
};

const originalWarn = console.warn;
const warnings = [];
console.warn = (message) => warnings.push(String(message));

try {
  normalizeSpec(spec, { specDir: process.cwd() });
} finally {
  console.warn = originalWarn;
}

const missing = MEDIA_SLOT_LAYOUTS.filter((layout, index) => {
  const slideNumber = index + 1;
  return !warnings.some((message) => (
    message.includes(`slide ${slideNumber} uses layout "${layout}"`)
    && message.includes('with media/image slot(s) but provides no images or charts')
  ));
});

if (missing.length) {
  console.error(`Missing media slot warning for layout(s): ${missing.join(', ')}`);
  process.exit(1);
}

console.log(`Media slot warning check passed: ${MEDIA_SLOT_LAYOUTS.length} layout(s).`);
