#!/usr/bin/env node
const path = require('node:path');
const { extractDocxContent } = require('./pptxgen/docx-import');

function parseArgs(argv) {
  const args = {};
  for (let i = 2; i < argv.length; i += 1) {
    const arg = argv[i];
    if (!arg.startsWith('--')) continue;
    const key = arg.slice(2);
    const next = argv[i + 1];
    if (!next || next.startsWith('--')) {
      args[key] = true;
    } else {
      args[key] = next;
      i += 1;
    }
  }
  return args;
}

function defaultOutputPath(docxPath, suffix) {
  const parsed = path.parse(docxPath);
  return path.join('outputs', parsed.name.replace(/[^\p{L}\p{N}._-]+/gu, '-') + suffix);
}

async function main() {
  const args = parseArgs(process.argv);
  if (args.help || !args.docx) {
    console.log([
      'Usage:',
      '  node scripts/docx-to-pptx.js --docx input.docx --write-extracted outputs/docx.extracted.json --write-md outputs/docx.extracted.md',
      '',
      'This command only extracts DOCX content and processed image assets.',
      'It does not generate PPT JSON or PPTX by rules. Write the PPT JSON semantically from the extracted content, then run scripts/generate-pptx.js.',
      '',
      'Options:',
      '  --docx path              Input DOCX file.',
      '  --asset-dir path         Directory for extracted/cropped image assets.',
      '  --write-extracted path   Write parsed DOCX text/image metadata JSON.',
      '  --write-md path          Write user-facing extraction Markdown.',
      '  --out-dir path           Base directory for default extracted files when write paths are omitted.',
    ].join('\n'));
    process.exit(args.help ? 0 : 2);
  }
  if (args.out || args['write-spec']) {
    throw new Error('DOCX rule-based PPT generation has been removed. Do not use --out or --write-spec here. First extract DOCX content, then create a semantic PPT JSON and run scripts/generate-pptx.js.');
  }

  const outDir = args['out-dir'] || path.dirname(args['write-extracted'] || args['write-md'] || defaultOutputPath(args.docx, '.extracted.json'));
  const writeExtracted = args['write-extracted'] || path.join(outDir, path.basename(defaultOutputPath(args.docx, '.extracted.json')));
  const writeMarkdown = args['write-md'] || path.join(outDir, path.basename(defaultOutputPath(args.docx, '.extracted.md')));
  const result = await extractDocxContent({
    docxPath: args.docx,
    outputDir: outDir,
    assetDir: args['asset-dir'],
    writeExtracted,
    writeMarkdown,
  });
  console.log(`DOCX source: ${path.resolve(args.docx)}`);
  console.log(`Extracted JSON: ${path.resolve(writeExtracted)}`);
  console.log(`Extracted Markdown: ${path.resolve(writeMarkdown)}`);
  console.log(`Extracted blocks: ${result.extracted.blocks.length}`);
  console.log(`Extracted images: ${result.extracted.images.length}`);
  if (result.extracted.warnings.length) {
    result.extracted.warnings.forEach((warning) => console.warn(`Warning: ${warning}`));
  }
  console.log('Next: write a semantic PPT JSON from the extracted content, then run node scripts/generate-pptx.js --spec <json> --out <pptx>.');
}

main().catch((error) => {
  console.error(error && error.stack ? error.stack : String(error));
  process.exitCode = 2;
});
