const fs = require('node:fs');
const path = require('node:path');
const { fail } = require('./errors');

function parseArgs(argv) {
  const result = {};
  for (let i = 0; i < argv.length; i += 1) {
    const a = argv[i];
    if (a === '--sample') result.sample = true;
    else if (a === '--sample-style') result.sampleStyle = argv[++i];
    else if (a === '--spec') result.spec = argv[++i];
    else if (a === '--out') result.out = argv[++i];
    else if (a === '--capacity-guide') {
      const next = argv[i + 1];
      result.capacityGuide = next && !next.startsWith('--') ? argv[++i] : true;
    }
    else if (a === '--diversify-layouts') result.diversifyLayouts = true;
    else if (a === '--write-normalized-spec') result.writeNormalizedSpec = argv[++i];
    else if (a === '--help' || a === '-h') result.help = true;
  }
  return result;
}

function loadSpecFile(specPath) {
  const raw = fs.readFileSync(specPath, 'utf8');
  const result = parseSpecJson(raw, specPath);
  if (result.repaired) {
    console.warn(`Warning: ${path.basename(specPath)} was parsed after JSON quote/comment/trailing-comma normalization. Use --write-normalized-spec to emit strict UTF-8 JSON.`);
  }
  return result.value;
}

function parseSpecJson(raw, sourceName = 'spec') {
  const candidates = [];
  const addCandidate = (label, text, repaired = false) => {
    if (typeof text !== 'string') return;
    const cleaned = text.replace(/^\uFEFF/, '').trim();
    if (!cleaned) return;
    if (!candidates.some((item) => item.text === cleaned)) candidates.push({ label, text: cleaned, repaired });
  };
  addCandidate('strict JSON', raw, false);
  const extracted = extractJsonPayload(raw);
  addCandidate('extracted JSON payload', extracted, extracted !== raw);
  addCandidate('normalized loose JSON', normalizeLooseJson(extracted), true);

  const errors = [];
  for (const candidate of candidates) {
    try {
      return { value: JSON.parse(candidate.text), repaired: candidate.repaired || candidate.label !== 'strict JSON' };
    } catch (error) {
      errors.push(`${candidate.label}: ${error.message}`);
    }
  }
  fail(`Could not parse ${sourceName} as JSON. ${formatJsonParseHelp(raw)}\nAttempts:\n- ${errors.join('\n- ')}`);
}

function extractJsonPayload(raw) {
  const text = String(raw || '').replace(/^\uFEFF/, '').trim();
  const fence = text.match(/```(?:json|jsonc)?\s*([\s\S]*?)```/i);
  if (fence) return fence[1].trim();
  const objectStart = text.indexOf('{');
  const arrayStart = text.indexOf('[');
  const starts = [objectStart, arrayStart].filter((index) => index >= 0);
  if (!starts.length) return text;
  const start = Math.min(...starts);
  const endChar = text[start] === '[' ? ']' : '}';
  const end = text.lastIndexOf(endChar);
  return end > start ? text.slice(start, end + 1) : text;
}

function normalizeLooseJson(raw) {
  return removeTrailingCommas(normalizeSmartJsonQuotes(stripJsonComments(String(raw || '').replace(/^\\uFEFF/, '')))).trim();
}

function normalizeSmartJsonQuotes(text) {
  let out = '';
  let inString = false;
  let quote = '';
  let escaped = false;
  const openingQuotes = new Set(['"', '\'', '\u201C', '\u2018']);
  const matchingQuote = { '"': '"', '\'': '\'', '\u201C': '\u201D', '\u2018': '\u2019' };
  const isClosingContext = (index) => {
    for (let j = index + 1; j < text.length; j += 1) {
      const c = text[j];
      if (/\s/.test(c)) continue;
      return c === ':' || c === ',' || c === '}' || c === ']';
    }
    return true;
  };
  for (let i = 0; i < text.length; i += 1) {
    const ch = text[i];
    if (!inString) {
      if (openingQuotes.has(ch)) {
        inString = true;
        quote = ch;
        escaped = false;
        out += '"';
      } else {
        out += ch;
      }
      continue;
    }
    if (escaped) {
      out += ch;
      escaped = false;
      continue;
    }
    if (ch === '\\') {
      out += ch;
      escaped = true;
      continue;
    }
    if (ch === matchingQuote[quote] && isClosingContext(i)) {
      out += '"';
      inString = false;
      quote = '';
      continue;
    }
    if (ch === '"' && quote !== '"') out += '\\"';
    else out += ch;
  }
  return out;
}
function stripJsonComments(text) {
  let out = '';
  let inString = false;
  let quote = '';
  let escaped = false;
  const openingQuotes = new Set(['"', '\'', '\u201C', '\u2018']);
  const matchingQuote = { '"': '"', '\'': '\'', '\u201C': '\u201D', '\u2018': '\u2019' };
  for (let i = 0; i < text.length; i += 1) {
    const ch = text[i];
    const next = text[i + 1];
    if (inString) {
      out += ch;
      if (escaped) escaped = false;
      else if (ch === '\\') escaped = true;
      else if (ch === matchingQuote[quote]) inString = false;
      continue;
    }
    if (openingQuotes.has(ch)) {
      inString = true;
      quote = ch;
      escaped = false;
      out += ch;
      continue;
    }
    if (ch === '/' && next === '/') {
      while (i < text.length && text[i] !== '\n') i += 1;
      out += '\n';
      continue;
    }
    if (ch === '/' && next === '*') {
      i += 2;
      while (i < text.length && !(text[i] === '*' && text[i + 1] === '/')) i += 1;
      i += 1;
      continue;
    }
    out += ch;
  }
  return out;
}

function removeTrailingCommas(text) {
  let out = '';
  let inString = false;
  let escaped = false;
  for (let i = 0; i < text.length; i += 1) {
    const ch = text[i];
    if (inString) {
      out += ch;
      if (escaped) escaped = false;
      else if (ch === '\\') escaped = true;
      else if (ch === '"') inString = false;
      continue;
    }
    if (ch === '"') {
      inString = true;
      out += ch;
      continue;
    }
    if (ch === ',') {
      let j = i + 1;
      while (j < text.length && /\s/.test(text[j])) j += 1;
      if (text[j] === '}' || text[j] === ']') continue;
    }
    out += ch;
  }
  return out;
}

function formatJsonParseHelp(raw) {
  const text = String(raw || '');
  const lineCount = text.split(/\r?\n/).length;
  return `Check quote escaping around strings. Prefer JSON.stringify output; use Chinese corner quotes 「」 or escaped \\" inside JSON strings. File has ${lineCount} line(s).`;
}

function writeNormalizedSpec(spec, outPath) {
  const clean = JSON.parse(JSON.stringify(spec));
  delete clean.__layoutDiversified;
  delete clean.__normalized;
  fs.mkdirSync(path.dirname(outPath), { recursive: true });
  fs.writeFileSync(outPath, `${JSON.stringify(clean, null, 2)}\n`, 'utf8');
  console.log(`Wrote normalized spec ${outPath}`);
}

module.exports = {
  parseArgs,
  loadSpecFile,
  parseSpecJson,
  extractJsonPayload,
  normalizeLooseJson,
  normalizeSmartJsonQuotes,
  stripJsonComments,
  removeTrailingCommas,
  writeNormalizedSpec,
};
