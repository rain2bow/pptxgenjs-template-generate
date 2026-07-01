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

if (require.main === module) {
  main().catch((error) => {
    console.error(error.stack || error.message || error);
    process.exit(1);
  });
}

async function main() {
  const args = parseArgs(process.argv.slice(2));
  if (!args.pptx || !args.plan || !args.out) {
    fail('Usage: node scripts/fill-inherited-template.js --pptx template.pptx --plan template.slots.json --out filled.pptx');
  }
  const pptxPath = path.resolve(args.pptx);
  const planPath = path.resolve(args.plan);
  const outPath = path.resolve(args.out);
  const zip = await JSZip.loadAsync(fs.readFileSync(pptxPath));
  const plan = JSON.parse(fs.readFileSync(planPath, 'utf8'));
  await fillInheritedTemplate(zip, plan);
  fs.mkdirSync(path.dirname(outPath), { recursive: true });
  const buffer = await zip.generateAsync({ type: 'nodebuffer', compression: 'DEFLATE' });
  fs.writeFileSync(outPath, buffer);
  console.log(`Wrote ${outPath}`);
  console.log(`Slides: ${Array.isArray(plan.slides) ? plan.slides.length : 0}`);
}

async function fillInheritedTemplate(zip, plan) {
  if (!Array.isArray(plan.slides) || !plan.slides.length) fail('Plan must include slides[].');
  const presentationXml = await readText(zip, 'ppt/presentation.xml');
  const presentationRelsXml = await readText(zip, 'ppt/_rels/presentation.xml.rels');
  const contentTypesXml = await readText(zip, '[Content_Types].xml');
  const sourcePaths = currentSlidePaths(presentationXml, presentationRelsXml, zip);
  if (!sourcePaths.length) fail('No source slides found in PPTX.');

  const sourceXml = new Map();
  const sourceRels = new Map();
  for (const sourcePath of sourcePaths) {
    sourceXml.set(sourcePath, await readText(zip, sourcePath));
    sourceRels.set(sourcePath, await readText(zip, relsPathForSlide(sourcePath), ''));
  }

  deleteCurrentSlideFiles(zip);

  const newSlidePaths = [];
  for (let i = 0; i < plan.slides.length; i += 1) {
    const slidePlan = plan.slides[i];
    const sourcePath = resolveSourceSlidePath(slidePlan, sourcePaths, i);
    let xml = sourceXml.get(sourcePath);
    if (!xml) fail(`Source slide not found for output slide ${i + 1}: ${sourcePath}`);
    xml = replaceSlideText(xml, slidePlan.textSlots || []);
    const newPath = `ppt/slides/slide${i + 1}.xml`;
    zip.file(newPath, xml);
    const rels = sourceRels.get(sourcePath);
    if (rels) zip.file(relsPathForSlide(newPath), rels);
    newSlidePaths.push(newPath);
  }

  zip.file('ppt/presentation.xml', rewritePresentationXml(presentationXml, newSlidePaths.length));
  zip.file('ppt/_rels/presentation.xml.rels', rewritePresentationRels(presentationRelsXml, newSlidePaths.length));
  zip.file('[Content_Types].xml', rewriteContentTypes(contentTypesXml, newSlidePaths.length));
}

function parseArgs(argv) {
  const args = {};
  for (let i = 0; i < argv.length; i += 1) {
    const arg = argv[i];
    if (arg === '--pptx') args.pptx = argv[++i];
    else if (arg === '--plan') args.plan = argv[++i];
    else if (arg === '--out') args.out = argv[++i];
    else if (arg === '--help' || arg === '-h') args.help = true;
  }
  return args;
}

function replaceSlideText(xml, textSlots) {
  if (!Array.isArray(textSlots) || !textSlots.length) return xml;
  const slots = textSlots.map((slot, index) => ({
    ...slot,
    index,
    value: resolveSlotText(slot),
  }));
  let seen = 0;
  return xml.replace(/<p:sp\b[\s\S]*?<\/p:sp>/g, (block) => {
    if (!block.includes('<p:txBody')) return block;
    const shapeId = attr(block, /<p:cNvPr\b[^>]*\bid="([^"]+)"/);
    const name = decodeXml(attr(block, /<p:cNvPr\b[^>]*\bname="([^"]*)"/));
    const slot = slots.find((item) => item.shapeId && item.shapeId === shapeId)
      || slots.find((item) => item.name && item.name === name)
      || slots.find((item) => item.slotId === `s${seen + 1}`)
      || slots[seen];
    seen += 1;
    if (!slot || slot.value === null || slot.value === undefined) return block;
    return replaceTextRuns(block, String(slot.value));
  });
}

function resolveSlotText(slot) {
  if (Object.prototype.hasOwnProperty.call(slot, 'newText')) return slot.newText;
  if (Object.prototype.hasOwnProperty.call(slot, 'targetText')) return slot.targetText;
  if (Array.isArray(slot.paragraphs) && slot.useParagraphs) return slot.paragraphs.join('\n');
  return slot.text;
}

function replaceTextRuns(block, value) {
  const escaped = escapeXml(value);
  let replaced = false;
  return block.replace(/<a:t([^>]*)>[\s\S]*?<\/a:t>/g, (node, attrs) => {
    if (!replaced) {
      replaced = true;
      const nextAttrs = /xml:space=/.test(attrs) || !/^\s|\s$|\n/.test(value) ? attrs : `${attrs} xml:space="preserve"`;
      return `<a:t${nextAttrs}>${escaped}</a:t>`;
    }
    return `<a:t${attrs}></a:t>`;
  });
}

function currentSlidePaths(presentationXml, relsXml, zip) {
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

function resolveSourceSlidePath(slidePlan, sourcePaths, index) {
  if (slidePlan.sourceSlidePath && sourcePaths.includes(slidePlan.sourceSlidePath)) return slidePlan.sourceSlidePath;
  if (slidePlan.sourceSlideNo) {
    const byNo = sourcePaths.find((slidePath) => slideNumber(slidePath) === Number(slidePlan.sourceSlideNo));
    if (byNo) return byNo;
  }
  if (slidePlan.copyFrom) {
    const copyIndex = Math.max(0, Math.min(sourcePaths.length - 1, Number(slidePlan.copyFrom) - 1));
    return sourcePaths[copyIndex];
  }
  return sourcePaths[Math.min(index, sourcePaths.length - 1)];
}

function deleteCurrentSlideFiles(zip) {
  Object.keys(zip.files)
    .filter((name) => /^ppt\/slides\/slide\d+\.xml$/i.test(name) || /^ppt\/slides\/_rels\/slide\d+\.xml\.rels$/i.test(name))
    .forEach((name) => zip.remove(name));
}

function rewritePresentationXml(xml, count) {
  const list = Array.from({ length: count }, (_, i) => `    <p:sldId id="${256 + i}" r:id="rId${1000 + i}"/>`).join('\n');
  const next = xml.replace(/<p:sldIdLst>[\s\S]*?<\/p:sldIdLst>/, `<p:sldIdLst>\n${list}\n  </p:sldIdLst>`);
  return next;
}

function rewritePresentationRels(xml, count) {
  const relationships = Array.from(xml.matchAll(/<Relationship\b[^>]*\/>/g), (m) => m[0])
    .filter((node) => !/\/relationships\/slide"/.test(node));
  for (let i = 0; i < count; i += 1) {
    relationships.push(`<Relationship Id="rId${1000 + i}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" Target="slides/slide${i + 1}.xml"/>`);
  }
  return xml.replace(/<Relationships\b([^>]*)>[\s\S]*?<\/Relationships>/, (_m, attrs) => `<Relationships${attrs}>\n  ${relationships.join('\n  ')}\n</Relationships>`);
}

function rewriteContentTypes(xml, count) {
  const withoutSlides = xml.replace(/\s*<Override\b[^>]*PartName="\/ppt\/slides\/slide\d+\.xml"[^>]*\/>/g, '');
  const overrides = Array.from({ length: count }, (_, i) => `<Override PartName="/ppt/slides/slide${i + 1}.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/>`).join('');
  return withoutSlides.replace('</Types>', `${overrides}</Types>`);
}

function relsPathForSlide(slidePath) {
  const name = path.posix.basename(slidePath);
  return `ppt/slides/_rels/${name}.rels`;
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

function attr(text, regex) {
  const m = text.match(regex);
  return m ? m[1] : '';
}

function decodeXml(text) {
  return String(text || '')
    .replace(/&lt;/g, '<')
    .replace(/&gt;/g, '>')
    .replace(/&quot;/g, '"')
    .replace(/&apos;/g, "'")
    .replace(/&amp;/g, '&');
}

function escapeXml(text) {
  return String(text || '')
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;');
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

module.exports = { fillInheritedTemplate };
