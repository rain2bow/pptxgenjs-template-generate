const path = require('node:path');
const process = require('node:process');
const { buildDeck, normalizeSpec } = require('./engine');
const { sampleSpec } = require('./samples');
const { parseArgs, loadSpecFile, writeNormalizedSpec } = require('./spec-io');
const { fail } = require('./errors');
const { styleGuideMarkdown, layoutExamplesMarkdown, writeLayoutExamples } = require('./layout-examples');

async function main(argv = process.argv.slice(2)) {
  const args = parseArgs(argv);
  if (args.unsupportedFlags?.length) {
    fail(args.unsupportedFlags.join(', ') + ' is no longer supported. Edit slide.layout in the JSON manually; the generator never changes layout automatically.');
  }

  if (args.removedCapacityGuide) {
    fail('--capacity-guide has been removed. Choose a style first, then run --layout-examples <style> to generate complete JSON examples without word-count limits.');
  }

  if (args.styleGuide) {
    process.stdout.write(styleGuideMarkdown());
    return;
  }

  if (args.layoutExamples) {
    const style = typeof args.layoutExamples === 'string' ? args.layoutExamples : args.sampleStyle;
    if (!style) fail('Missing style. Use --layout-examples cmb|swiss|magazine.');
    try {
      if (args.out) writeLayoutExamples(style, args.out);
      else process.stdout.write(layoutExamplesMarkdown(style));
    } catch (error) {
      fail(error.message);
    }
    return;
  }

  if (!args.sample && !args.spec) {
    fail('Usage: node scripts/generate-pptx.js --style-guide\n       node scripts/generate-pptx.js --layout-examples cmb --out outputs/cmb-layout-examples.md\n       node scripts/generate-pptx.js --spec deck.json --out deck.pptx\n       node scripts/generate-pptx.js --sample --out outputs/sample-deck.pptx');
  }

  const specPath = args.spec ? path.resolve(args.spec) : null;
  const specDir = specPath ? path.dirname(specPath) : process.cwd();
  const outPath = path.resolve(args.out || 'deck.pptx');
  const spec = args.sample ? sampleSpec(args.sampleStyle) : loadSpecFile(specPath);

  const normalizedSpecPath = args.writeNormalizedSpec ? path.resolve(args.writeNormalizedSpec) : null;
  normalizeSpec(spec, { writeNormalizedSpec: normalizedSpecPath, specDir });
  if (normalizedSpecPath) writeNormalizedSpec(spec, normalizedSpecPath);
  await buildDeck(spec, specDir, outPath);
}

module.exports = { main };
