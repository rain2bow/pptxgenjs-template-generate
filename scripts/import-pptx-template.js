#!/usr/bin/env node
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

const { buildDeck } = require('./generate-pptx.js');

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
  if (!args.pptx || !args.outSpec) {
    fail('Usage: node scripts/import-pptx-template.js --pptx template.pptx --out-spec imported.json [--out-pptx new.pptx] [--style swiss|magazine] [--theme ikb|ink]');
  }

  const pptxPath = path.resolve(args.pptx);
  const outSpec = path.resolve(args.outSpec);
  const specDir = path.dirname(outSpec);
  const assetDir = path.resolve(args.assetDir || path.join(specDir, `${path.basename(outSpec, path.extname(outSpec))}.assets`));
  const buffer = fs.readFileSync(pptxPath);
  const zip = await JSZip.loadAsync(buffer);
  const parsed = await parsePptx(zip, pptxPath, assetDir, specDir);
  const style = args.style || inferDeckStyle(parsed) || 'swiss';
  const theme = args.theme || (style === 'magazine' ? 'ink' : 'ikb');
  const spec = buildSpec(parsed, { style, theme, title: args.title || parsed.title || path.basename(pptxPath, '.pptx') });

  fs.mkdirSync(path.dirname(outSpec), { recursive: true });
  fs.writeFileSync(outSpec, `${JSON.stringify(spec, null, 2)}\n`, 'utf8');
  console.log(`Wrote ${outSpec}`);
  console.log(`Slides: ${spec.slides.length}`);
  console.log(`Style: ${spec.style} / ${spec.theme}`);

  if (args.outPptx) {
    const outPptx = path.resolve(args.outPptx);
    await buildDeck(spec, specDir, outPptx);
  }
}

function parseArgs(argv) {
  const result = {};
  for (let i = 0; i < argv.length; i += 1) {
    const arg = argv[i];
    if (arg === '--pptx') result.pptx = argv[++i];
    else if (arg === '--out-spec') result.outSpec = argv[++i];
    else if (arg === '--out-pptx') result.outPptx = argv[++i];
    else if (arg === '--asset-dir') result.assetDir = argv[++i];
    else if (arg === '--style') result.style = argv[++i];
    else if (arg === '--theme') result.theme = argv[++i];
    else if (arg === '--title') result.title = argv[++i];
    else if (arg === '--help' || arg === '-h') result.help = true;
  }
  return result;
}

async function parsePptx(zip, pptxPath, assetDir, specDir) {
  const presentationXml = await readText(zip, 'ppt/presentation.xml');
  const slideSize = parseSlideSize(presentationXml);
  const slidePaths = Object.keys(zip.files)
    .filter((name) => /^ppt\/slides\/slide\d+\.xml$/i.test(name))
    .sort((a, b) => slideNumber(a) - slideNumber(b));
  if (!slidePaths.length) fail('No slides found in PPTX.');

  const slides = [];
  for (const slidePath of slidePaths) {
    const slideNo = slideNumber(slidePath);
    const relsPath = `ppt/slides/_rels/slide${slideNo}.xml.rels`;
    const xml = await readText(zip, slidePath);
    const relsXml = await readText(zip, relsPath, '');
    const relationships = parseRelationships(relsXml);
    slides.push(await parseSlide(zip, xml, relationships, slideNo, slideSize, assetDir, specDir));
  }

  return {
    title: path.basename(pptxPath, '.pptx'),
    source: pptxPath,
    slideSize,
    slides,
  };
}

async function parseSlide(zip, xml, relationships, slideNo, slideSize, assetDir, specDir) {
  const texts = parseTextShapes(xml, slideSize);
  const pictures = await parsePictureShapes(zip, xml, relationships, slideNo, slideSize, assetDir, specDir);
  const hasTable = /<a:tbl[\s>]/.test(xml);
  const hasChart = /<c:chart\b|chart\/chart\d+\.xml/i.test(xml);
  return {
    slideNo,
    texts,
    pictures,
    hasTable,
    hasChart,
    slideSize,
    rawStats: {
      shapeCount: countMatches(xml, /<p:sp[\s>]/g),
      pictureCount: pictures.length,
      textBoxCount: texts.length,
    },
  };
}

function parseTextShapes(xml, slideSize) {
  return blocks(xml, 'p:sp')
    .filter((block) => block.includes('<p:txBody'))
    .map((block, index) => {
      const paragraphs = blocks(block, 'a:p')
        .map((p) => joinMatches(p, /<a:t[^>]*>([\s\S]*?)<\/a:t>/g))
        .map(cleanText)
        .filter(Boolean);
      const text = paragraphs.join('\n').trim();
      if (!text) return null;
      const box = parseXfrm(block, slideSize);
      return {
        id: attr(block, /<p:cNvPr\b[^>]*\bid="([^"]+)"/),
        name: decodeXml(attr(block, /<p:cNvPr\b[^>]*\bname="([^"]*)"/) || `Text ${index + 1}`),
        text,
        paragraphs,
        box,
      };
    })
    .filter(Boolean)
    .sort((a, b) => (a.box.y - b.box.y) || (a.box.x - b.box.x));
}

async function parsePictureShapes(zip, xml, relationships, slideNo, slideSize, assetDir, specDir) {
  const pics = [];
  const picBlocks = blocks(xml, 'p:pic');
  for (let i = 0; i < picBlocks.length; i += 1) {
    const block = picBlocks[i];
    const rid = attr(block, /<a:blip\b[^>]*(?:r:embed|embed)="([^"]+)"/);
    const target = rid ? relationships[rid] : null;
    const zipPath = target ? normalizePptTarget(`ppt/slides/slide${slideNo}.xml`, target) : null;
    const file = zipPath && zip.files[zipPath] ? zip.files[zipPath] : null;
    const relOut = file ? await writeMediaFile(file, zipPath, assetDir, specDir) : null;
    pics.push({
      id: attr(block, /<p:cNvPr\b[^>]*\bid="([^"]+)"/),
      name: decodeXml(attr(block, /<p:cNvPr\b[^>]*\bname="([^"]*)"/) || `Image ${i + 1}`),
      src: relOut,
      box: parseXfrm(block, slideSize),
    });
  }
  return pics;
}

async function writeMediaFile(file, zipPath, assetDir, specDir) {
  fs.mkdirSync(assetDir, { recursive: true });
  const base = path.basename(zipPath);
  const out = uniquePath(path.join(assetDir, base));
  const data = await file.async('nodebuffer');
  fs.writeFileSync(out, data);
  return slash(path.relative(specDir, out));
}

function buildSpec(parsed, options) {
  const slides = parsed.slides.map((slide, index) => inferSlideSpec(slide, index, parsed.slides.length, options.style));
  return {
    title: options.title,
    subtitle: `Imported from ${path.basename(parsed.source)}`,
    author: 'Guizang PPTXGenJS Skill',
    style: options.style,
    theme: options.theme,
    sourceTemplate: {
      file: parsed.source,
      slideSize: parsed.slideSize,
      importedAt: new Date().toISOString(),
      note: 'Inferred from editable PPTX XML. Review layout/content before final delivery.',
    },
    slides,
  };
}

function inferSlideSpec(slide, index, total, style) {
  const textBoxes = slide.texts;
  const titleBox = pickTitleBox(textBoxes);
  const title = titleBox?.text || `Slide ${index + 1}`;
  const rest = textBoxes.filter((box) => box !== titleBox);
  const paragraphs = rest.flatMap((box) => box.paragraphs.length ? box.paragraphs : [box.text]).map(cleanText).filter(Boolean);
  const pictures = slide.pictures.filter((pic) => pic.src && isContentPicture(pic, slide));
  const imageEntries = pictures.map((pic, i) => ({ path: pic.src, caption: pic.name || `Image ${i + 1}` }));

  const layout = inferLayout(slide, index, total, paragraphs, style);
  const base = {
    layout,
    kicker: inferKicker(slide, index),
    title,
  };

  if (layout === 'cover' || layout === 'closing' || layout === 'statement') {
    base.subtitle = paragraphs[0] || '';
    base.body = paragraphs.slice(1).join('\n');
    if (imageEntries[0]) base.image = imageEntries[0];
    return base;
  }

  if (layout === 'mediaGrid') {
    base.images = imageEntries;
    base.captions = imageEntries.map((img, i) => ({ title: img.caption || `Image ${i + 1}` }));
    if (!imageEntries.length) base.captions = paragraphs.slice(0, 6).map((text) => ({ title: text }));
    return base;
  }

  if (layout === 'media' || layout === 'imageHero' || layout === 'caseStudy' || layout === 'quoteImage') {
    if (imageEntries[0]) base.image = imageEntries[0];
    base.body = paragraphs[0] || '';
    const items = paragraphs.slice(1, layout === 'imageHero' ? 4 : 6).map((text) => textToItem(text));
    if (items.length) base.items = items;
    return base;
  }

  if (layout === 'dataSheet') {
    base.body = paragraphs[0] || '';
    base.table = { headers: ['Field', 'Value'], rows: paragraphs.slice(1, 7).map((text, i) => [`Item ${i + 1}`, text]) };
    return base;
  }

  if (layout === 'chart') {
    base.chart = placeholderChart(title, paragraphs);
    base.insights = paragraphs.slice(0, 3).map(textToItem).filter(slotItemHasText);
    return base;
  }

  if (layout === 'kpiTower' || layout === 'bigNumbers') {
    base.items = paragraphs.slice(0, style === 'swiss' ? 4 : 6).map((text, i) => numericItem(text, i));
    return base;
  }

  if (layout === 'timeline' || layout === 'pipeline' || layout === 'roadmap') {
    base.steps = paragraphs.slice(0, 6).map((text, i) => ({ label: String(i + 1).padStart(2, '0'), ...textToItem(text) }));
    return base;
  }

  const defaultItems = paragraphs.slice(0, layout === 'textGrid' ? 9 : 8).map(textToItem);
  if (['fourCards', 'matrix', 'agenda', 'radial'].includes(layout)) base.items = defaultItems;
  else base.sections = defaultItems;
  return base;
}

function inferLayout(slide, index, total, paragraphs, style) {
  const picCount = slide.pictures.length;
  const textCount = slide.texts.length;
  if (index === 0) return 'cover';
  const slideText = paragraphs.join(' ');
  if (index === total - 1 && (textCount <= 4 || /THANK\s*YOU|谢谢|感谢|致谢|观看/i.test(slideText))) return 'closing';
  if (slide.hasTable) return 'dataSheet';
  if (slide.hasChart) return 'chart';
  if (picCount >= 2) return 'mediaGrid';
  if (picCount === 1) return 'media';
  if (paragraphs.some((text) => /\d/.test(text)) && paragraphs.length <= 6) return style === 'swiss' ? 'kpiTower' : 'bigNumbers';
  if (looksSequential(paragraphs)) return style === 'swiss' ? 'timeline' : 'pipeline';
  if (paragraphs.length >= 6) return style === 'swiss' ? 'textGrid' : 'article';
  if (paragraphs.length >= 3) return 'fourCards';
  return 'statement';
}

function isContentPicture(pic, slide) {
  const box = pic.box || {};
  const slideH = slide.slideSize?.h || DEFAULT_SLIDE.h;
  const area = Number(box.w || 0) * Number(box.h || 0);
  if (area < 0.65) return false;
  if (Number(box.y || 0) < 0.95 && Number(box.h || 0) < 0.9) return false;
  if ((Number(box.y || 0) + Number(box.h || 0)) > slideH - 0.45 && Number(box.h || 0) < 0.6) return false;
  return true;
}
function inferDeckStyle(parsed) {
  const allText = parsed.slides.flatMap((slide) => slide.texts.map((shape) => shape.text)).join(' ');
  const upper = allText.toUpperCase();
  if (/SWISS|KPI|DASHBOARD|DATA|METRIC|ROADMAP/.test(upper)) return 'swiss';
  if (/MAGAZINE|STORY|QUOTE|ARTICLE/.test(upper)) return 'magazine';
  return parsed.slides.some((slide) => slide.hasChart || slide.hasTable) ? 'swiss' : 'magazine';
}

function pickTitleBox(textBoxes) {
  if (!textBoxes.length) return null;
  const candidates = textBoxes
    .map((box) => ({ box, score: scoreTitleBox(box) }))
    .sort((a, b) => b.score - a.score);
  return candidates[0].box;
}

function scoreTitleBox(box) {
  const text = box.text || '';
  const lengthPenalty = Math.min(text.length / 90, 1.5);
  return (box.box.h * box.box.w) + Math.max(0, 4 - box.box.y) + (text.length > 4 ? 0.6 : 0) - lengthPenalty;
}

function inferKicker(slide, index) {
  const top = slide.texts.find((box) => box.box.y < 1.2 && box.text.length < 80);
  return top?.text || `Imported / ${String(index + 1).padStart(2, '0')}`;
}

function textToItem(text) {
  const clean = cleanText(text).replace(/^[\u2022\-\d\.\)\s]+/, '').trim();
  if (!clean) return { title: 'Imported item', body: 'Review this placeholder.' };
  const parts = clean.split(/[：:。\.]/).map((part) => part.trim()).filter(Boolean);
  if (parts.length >= 2 && parts[0].length <= 24) return { title: parts[0], body: parts.slice(1).join('。') };
  return { title: clean.slice(0, 24), body: clean.length > 24 ? clean.slice(24) : clean };
}

function slotItemHasText(item) {
  return !!(item && (item.title || item.body || item.text || item.label || item.value));
}

function placeholderChart(title, paragraphs) {
  const labels = paragraphs.slice(0, 4).map((text, i) => cleanText(text).slice(0, 12) || `Item ${i + 1}`);
  const safeLabels = labels.length >= 2 ? labels : ['Item 1', 'Item 2', 'Item 3'];
  return {
    chartType: 'column',
    title: title || 'Imported chart placeholder',
    labels: safeLabels,
    values: safeLabels.map((_, i) => (i + 1) * 10),
    showValue: true,
    caption: 'Imported PPTX contained a chart. Review and replace placeholder values.',
  };
}

function numericItem(text, index) {
  const match = String(text).match(/([+-]?\d+(?:\.\d+)?%?|\d+[KkMm]?\+?)/);
  if (!match) return { label: `Metric ${index + 1}`, value: String(index + 1), note: cleanText(text) };
  return {
    label: cleanText(text).replace(match[0], '').trim() || `Metric ${index + 1}`,
    value: match[0],
    valueNum: Number(String(match[0]).replace(/[^0-9.-]/g, '')) || index + 1,
    note: cleanText(text),
  };
}

function looksSequential(paragraphs) {
  return paragraphs.length >= 3 && paragraphs.slice(0, 4).some((text) => /(^|\s)(0?1|1|step|phase|阶段|步骤|里程碑)/i.test(text));
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
  const re = new RegExp(`<${escaped}\\b[\\s\\S]*?<\\/${escaped}>`, 'g');
  return xml.match(re) || [];
}

function joinMatches(text, regex) {
  return Array.from(text.matchAll(regex), (m) => decodeXml(m[1])).join('');
}

function attr(text, regex) {
  const m = text.match(regex);
  return m ? m[1] : '';
}

function cleanText(text) {
  return decodeXml(String(text || '')).replace(/\s+/g, ' ').trim();
}

function decodeXml(text) {
  return String(text || '')
    .replace(/&lt;/g, '<')
    .replace(/&gt;/g, '>')
    .replace(/&quot;/g, '"')
    .replace(/&apos;/g, "'")
    .replace(/&amp;/g, '&');
}

function countMatches(text, regex) {
  return (text.match(regex) || []).length;
}

function slideNumber(name) {
  const m = name.match(/slide(\d+)\.xml$/i);
  return m ? Number(m[1]) : 0;
}

function uniquePath(file) {
  if (!fs.existsSync(file)) return file;
  const dir = path.dirname(file);
  const ext = path.extname(file);
  const base = path.basename(file, ext);
  for (let i = 2; i < 1000; i += 1) {
    const candidate = path.join(dir, `${base}-${i}${ext}`);
    if (!fs.existsSync(candidate)) return candidate;
  }
  return path.join(dir, `${base}-${Date.now()}${ext}`);
}

function slash(value) {
  return String(value).replace(/\\/g, '/');
}

function fail(message) {
  console.error(message);
  process.exit(2);
}

module.exports = {
  parsePptx,
  buildSpec,
  inferSlideSpec,
};