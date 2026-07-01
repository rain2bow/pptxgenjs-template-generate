const fs = require('node:fs');
const path = require('node:path');
const process = require('node:process');

let JSZip;
try {
  JSZip = require('jszip');
} catch (error) {
  console.error('Missing dependency "jszip". Run "npm install" in the skill directory.');
  process.exit(2);
}

const EMU_PER_INCH = 914400;
const DEFAULT_SLIDE = { w: 13.333, h: 7.5 };

if (require.main === module) {
  main().catch((error) => {
    console.error(error.stack || error.message || error);
    process.exit(1);
  });
}

async function main() {
  const args = parseArgs(process.argv.slice(2));
  if (!args.pptx || !args.out) {
    fail('Usage: node scripts/extract-inherited-template.js --pptx template.pptx --out template.slots.json');
  }
  const pptxPath = path.resolve(args.pptx);
  const outPath = path.resolve(args.out);
  const zip = await JSZip.loadAsync(fs.readFileSync(pptxPath));
  const model = await extractTemplate(zip, pptxPath);
  fs.mkdirSync(path.dirname(outPath), { recursive: true });
  fs.writeFileSync(outPath, `${JSON.stringify(model, null, 2)}\n`, 'utf8');
  console.log(`Wrote ${outPath}`);
  console.log(`Slides: ${model.slides.length}`);
  console.log(`Text slots: ${model.slides.reduce((sum, slide) => sum + slide.textSlots.length, 0)}`);
}

async function extractTemplate(zip, pptxPath) {
  const presentationXml = await readText(zip, 'ppt/presentation.xml');
  const presentationRels = await readText(zip, 'ppt/_rels/presentation.xml.rels', '');
  const slideSize = parseSlideSize(presentationXml);
  const orderedSlidePaths = presentationSlidePaths(presentationXml, presentationRels, zip);
  const slides = [];

  for (let i = 0; i < orderedSlidePaths.length; i += 1) {
    const slidePath = orderedSlidePaths[i];
    const xml = await readText(zip, slidePath);
    const textSlots = parseTextSlots(xml, slideSize);
    slides.push({
      index: i + 1,
      sourceSlideNo: slideNumber(slidePath),
      sourceSlidePath: slidePath,
      pageType: inferPageType(i, orderedSlidePaths.length, textSlots, xml),
      textSlots,
    });
  }

  return {
    mode: 'strict-inherit',
    source: pptxPath,
    title: path.basename(pptxPath, '.pptx'),
    slideSize,
    totalSlides: slides.length,
    note: 'Edit slides[].textSlots[].text or set newText, then run fill-inherited-template.js. Only non-empty text slots are kept.',
    slides,
  };
}

function parseArgs(argv) {
  const args = {};
  for (let i = 0; i < argv.length; i += 1) {
    const arg = argv[i];
    if (arg === '--pptx') args.pptx = argv[++i];
    else if (arg === '--out') args.out = argv[++i];
    else if (arg === '--help' || arg === '-h') args.help = true;
  }
  return args;
}

function presentationSlidePaths(presentationXml, relsXml, zip) {
  const rels = parseRelationships(relsXml);
  const paths = [];
  for (const match of presentationXml.matchAll(/<p:sldId\b[^>]*\br:id="([^"]+)"[^>]*\/>/g)) {
    const target = rels[match[1]];
    if (!target) continue;
    const slidePath = normalizePptTarget('ppt/presentation.xml', target);
    if (zip.files[slidePath]) paths.push(slidePath);
  }
  if (paths.length) return paths;
  return Object.keys(zip.files)
    .filter((name) => /^ppt\/slides\/slide\d+\.xml$/i.test(name))
    .sort((a, b) => slideNumber(a) - slideNumber(b));
}

function parseTextSlots(xml, slideSize) {
  let slotIndex = 0;
  return blocks(xml, 'p:sp')
    .filter((block) => block.includes('<p:txBody'))
    .map((block) => {
      const paragraphs = blocks(block, 'a:p')
        .map((p) => joinMatches(p, /<a:t[^>]*>([\s\S]*?)<\/a:t>/g))
        .map(cleanText)
        .filter(Boolean);
      const text = paragraphs.join('\n').trim();
      if (!text) return null;
      slotIndex += 1;
      const id = attr(block, /<p:cNvPr\b[^>]*\bid="([^"]+)"/);
      const name = decodeXml(attr(block, /<p:cNvPr\b[^>]*\bname="([^"]*)"/) || `Text ${slotIndex}`);
      const box = parseXfrm(block, slideSize);
      return {
        slotId: `s${slotIndex}`,
        shapeId: id,
        name,
        role: inferTextRole(text, box, slideSize, slotIndex),
        text,
        paragraphs,
        box,
      };
    })
    .filter(Boolean)
    .sort((a, b) => (a.box.y - b.box.y) || (a.box.x - b.box.x));
}

function inferPageType(index, total, textSlots, xml) {
  if (index === 0) return 'cover';
  const text = textSlots.map((slot) => slot.text).join(' ');
  if (index === total - 1 && /thank|thanks|谢谢|感谢|致谢|结束/i.test(text)) return 'closing';
  if (/<a:tbl[\s>]/.test(xml)) return 'data';
  if (/<c:chart\b|chart\/chart\d+\.xml/i.test(xml)) return 'chart';
  if (/<p:pic\b/.test(xml)) return 'media';
  if (textSlots.length >= 5) return 'content-grid';
  return 'content';
}

function inferTextRole(text, box, slideSize, slotIndex) {
  const compact = String(text || '').replace(/\s/g, '');
  if (box.y < 1.45 && box.h >= 0.35 && compact.length <= 42) return slotIndex === 1 ? 'title' : 'kicker';
  if (box.y < 2.2 && compact.length <= 80) return 'subtitle';
  if (box.y + box.h > slideSize.h - 0.7 && box.h < 0.55) return 'footer';
  if (/^\d+\s*[\)\.、]/.test(compact) || /^[•\-]/.test(compact)) return 'bullet';
  return 'body';
}

function parseSlideSize(xml) {
  const m = xml.match(/<p:sldSz\b[^>]*\bcx="(\d+)"[^>]*\bcy="(\d+)"/);
  if (!m) return DEFAULT_SLIDE;
  return { w: Number(m[1]) / EMU_PER_INCH, h: Number(m[2]) / EMU_PER_INCH };
}

function parseXfrm(block, slideSize) {
  const off = block.match(/<a:off\b[^>]*\bx="(-?\d+)"[^>]*\by="(-?\d+)"/);
  const ext = block.match(/<a:ext\b[^>]*\bcx="(\d+)"[^>]*\bcy="(\d+)"/);
  return {
    x: off ? Number(off[1]) / EMU_PER_INCH : 0,
    y: off ? Number(off[2]) / EMU_PER_INCH : 0,
    w: ext ? Number(ext[1]) / EMU_PER_INCH : slideSize.w,
    h: ext ? Number(ext[2]) / EMU_PER_INCH : 0.5,
  };
}

function parseRelationships(xml) {
  const rels = {};
  for (const rel of xml.matchAll(/<Relationship\b[^>]*>/g)) {
    const node = rel[0];
    const id = attr(node, /\bId="([^"]+)"/);
    const target = attr(node, /\bTarget="([^"]+)"/);
    if (id && target) rels[id] = target;
  }
  return rels;
}

function normalizePptTarget(fromPath, target) {
  if (target.startsWith('/')) return target.replace(/^\//, '');
  return slash(path.posix.normalize(path.posix.join(path.posix.dirname(fromPath), target)));
}

async function readText(zip, name, fallback = null) {
  const file = zip.files[name];
  if (!file) {
    if (fallback !== null) return fallback;
    fail(`Missing ${name} in PPTX.`);
  }
  return file.async('text');
}

function blocks(xml, tag) {
  const escaped = tag.replace(':', '\\:');
  return xml.match(new RegExp(`<${escaped}\\b[\\s\\S]*?<\\/${escaped}>`, 'g')) || [];
}

function joinMatches(text, regex) {
  return Array.from(text.matchAll(regex), (m) => decodeXml(m[1])).join('');
}

function attr(text, regex) {
  const m = text.match(regex);
  return m ? m[1] : '';
}

function cleanText(text) {
  return decodeXml(String(text || '')).replace(/[ \t\r\f\v]+/g, ' ').trim();
}

function decodeXml(text) {
  return String(text || '')
    .replace(/&lt;/g, '<')
    .replace(/&gt;/g, '>')
    .replace(/&quot;/g, '"')
    .replace(/&apos;/g, "'")
    .replace(/&amp;/g, '&');
}

function slideNumber(name) {
  const m = name.match(/slide(\d+)\.xml$/i);
  return m ? Number(m[1]) : 0;
}

function slash(value) {
  return String(value).replace(/\\/g, '/');
}

function fail(message) {
  console.error(message);
  process.exit(2);
}

module.exports = { extractTemplate };
