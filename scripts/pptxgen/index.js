const { buildDeck, normalizeSpec } = require('./engine');
const { sampleSpec } = require('./samples');
const specMd = require('./spec-md');
const speakerNotes = require('./speaker-notes');
const specIo = require('./spec-io');

module.exports = {
  buildDeck,
  normalizeSpec,
  sampleSpec,
  ...specMd,
  ...speakerNotes,
  ...specIo,
};
