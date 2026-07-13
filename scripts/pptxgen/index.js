const { buildDeck, normalizeSpec } = require('./engine');
const { sampleSpec } = require('./samples');
const specMd = require('./spec-md');
const speakerNotes = require('./speaker-notes');
const specIo = require('./spec-io');
const textCapacity = require('./text-capacity');
const docxImport = require('./docx-import');

module.exports = {
  buildDeck,
  normalizeSpec,
  sampleSpec,
  ...specMd,
  ...speakerNotes,
  ...specIo,
  ...textCapacity,
  ...docxImport,
};
