#!/usr/bin/env node
const path = require('node:path');
const process = require('node:process');
const { loadSpecFile } = require('./pptxgen/spec-io');
const { writeSpecMarkdown } = require('./pptxgen/spec-md');
const { fail } = require('./pptxgen/errors');

function parseArgs(argv) {
  const args = {};
  for (let i = 0; i < argv.length; i += 1) {
    const arg = argv[i];
    if (arg === '--spec') args.spec = argv[++i];
    else if (arg === '--out') args.out = argv[++i];
    else if (arg === '--help' || arg === '-h') args.help = true;
  }
  return args;
}

function usage() {
  return 'Usage: node scripts/spec-to-md.js --spec path/to/deck.json --out path/to/deck.md';
}

function main(argv = process.argv.slice(2)) {
  const args = parseArgs(argv);
  if (args.help || !args.spec) fail(usage());
  const specPath = path.resolve(args.spec);
  const outPath = path.resolve(args.out || specPath.replace(/\.(json|jsonc)$/i, '.md'));
  const spec = loadSpecFile(specPath);
  writeSpecMarkdown(spec, outPath);
}

if (require.main === module) {
  try {
    main();
  } catch (error) {
    console.error(error.message || error);
    process.exit(1);
  }
}

module.exports = { main };
