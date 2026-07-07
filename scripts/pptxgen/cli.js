const path = require('node:path');
const process = require('node:process');
const { buildDeck, normalizeSpec } = require('./engine');
const { sampleSpec } = require('./samples');
const { parseArgs, loadSpecFile, writeNormalizedSpec } = require('./spec-io');
const { fail } = require('./errors');
const { layoutCapacityMarkdown, writeLayoutCapacityGuide } = require('./text-capacity');

async function main(argv = process.argv.slice(2)) {
  const args = parseArgs(argv);
  if (args.capacityGuide) {
    const style = typeof args.capacityGuide === 'string' ? args.capacityGuide : (args.sampleStyle || 'swiss');
    if (args.out) writeLayoutCapacityGuide(style, args.out);
    else process.stdout.write(layoutCapacityMarkdown(style));
    return;
  }

  if (!args.sample && !args.spec) {
    fail('Usage: node scripts/generate-pptx.js --spec deck.json --out deck.pptx\n       node scripts/generate-pptx.js --spec deck.json --out deck.pptx --diversify-layouts --write-normalized-spec deck.normalized.json\n       node scripts/generate-pptx.js --capacity-guide cmb --out outputs/cmb-capacity.md\n       node scripts/generate-pptx.js --sample --out outputs/sample-deck.pptx');
  }

  const specPath = args.spec ? path.resolve(args.spec) : null;
  const specDir = specPath ? path.dirname(specPath) : process.cwd();
  const outPath = path.resolve(args.out || 'deck.pptx');
  const spec = args.sample ? sampleSpec(args.sampleStyle) : loadSpecFile(specPath);

  const normalizedSpecPath = args.writeNormalizedSpec ? path.resolve(args.writeNormalizedSpec) : null;
  normalizeSpec(spec, { diversifyLayouts: !!args.diversifyLayouts, writeNormalizedSpec: normalizedSpecPath, specDir });
  if (normalizedSpecPath) writeNormalizedSpec(spec, normalizedSpecPath);
  await buildDeck(spec, specDir, outPath);
}

module.exports = { main };

