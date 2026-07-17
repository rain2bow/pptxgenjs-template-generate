const { buildDeck, normalizeSpec } = require('./engine');
const { sampleSpec } = require('./samples');
const specMd = require('./spec-md');
const speakerNotes = require('./speaker-notes');
const specIo = require('./spec-io');
const docxImport = require('./docx-import');
const layoutSchema = require('./layout-schema');
const layoutExamples = require('./layout-examples');

module.exports = {
  buildDeck,
  normalizeSpec,
  sampleSpec,
  ...specMd,
  ...speakerNotes,
  ...specIo,
  ...layoutSchema,
  ...layoutExamples,
  ...docxImport,
};
