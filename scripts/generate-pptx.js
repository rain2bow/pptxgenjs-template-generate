#!/usr/bin/env node
const api = require('./pptxgen');

if (require.main === module) {
  require('./pptxgen/cli').main().catch((error) => {
    console.error(error);
    process.exit(1);
  });
}

module.exports = api;
