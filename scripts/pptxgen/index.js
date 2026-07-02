const { buildDeck, normalizeSpec } = require('./engine');
const { sampleSpec } = require('./samples');
const specIo = require('./spec-io');

module.exports = {
  buildDeck,
  normalizeSpec,
  sampleSpec,
  ...specIo,
};
