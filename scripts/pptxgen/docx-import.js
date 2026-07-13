const fs = require('node:fs');
const path = require('node:path');
const JSZip = require('jszip');
const { fail } = require('./errors');

const EMU_PER_INCH = 914400;
const IMAGE_REL_TYPE = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/image';

async function extractDocxContent(options = {}) {
  const docxPath = requiredPath(options.docxPath || options.docx, 'docxPath');
  const outputDir = path.resolve(options.outputDir || options.outDir || process.cwd());
  const specDir = path.resolve(options.specDir || (options.writeExtracted ? path.dirname(options.writeExtracted) : outputDir));
  const assetDir = path.resolve(options.assetDir || path.join(specDir, safeBaseName(docxPath) + '-docx-assets'));
  const extracted = await parseDocx(docxPath, { assetDir, specDir });
  if (options.writeExtracted) writeJson(options.writeExtracted, extracted);
  if (options.writeMarkdown) writeDocxExtractionMarkdown(options.writeMarkdown, extracted);
  return { extracted };
}

async function parseDocx(docxPath, options = {}) {
  const absoluteDocx = path.resolve(requiredPath(docxPath, 'docxPath'));
  if (!fs.existsSync(absoluteDocx)) fail(`DOCX file not found: ${absoluteDocx}`);
  const assetDir = path.resolve(options.assetDir || path.join(path.dirname(absoluteDocx), `${safeBaseName(absoluteDocx)}-assets`));
  const specDir = path.resolve(options.specDir || path.dirname(absoluteDocx));
  const zip = await JSZip.loadAsync(await fs.promises.readFile(absoluteDocx));
  const documentXml = await readZipText(zip, 'word/document.xml');
  if (!documentXml) fail('Invalid DOCX: word/document.xml is missing.');
  const rels = parseRelationships(await readZipText(zip, 'word/_rels/document.xml.rels'));

  fs.mkdirSync(assetDir, { recursive: true });
  const warnings = [];
  const paragraphs = [];
  const blocks = [];
  const paragraphMatches = [...documentXml.matchAll(/<w:p\b[\s\S]*?<\/w:p>/g)];
  let imageCounter = 0;

  for (let paragraphIndex = 0; paragraphIndex < paragraphMatches.length; paragraphIndex += 1) {
    const paragraphXml = paragraphMatches[paragraphIndex][0];
    const parsed = await parseParagraph(paragraphXml, {
      zip,
      rels,
      assetDir,
      specDir,
      paragraphIndex,
      imageCounter,
      warnings,
    });
    imageCounter = parsed.imageCounter;
    let paragraph = null;
    if (parsed.text.trim()) {
      paragraph = {
        index: paragraphIndex,
        paragraphIndex,
        blockIndex: null,
        text: normalizeWhitespace(parsed.text),
      };
      paragraphs.push(paragraph);
    }
    parsed.events.forEach((event) => {
      if (event.type === 'text') {
        const text = normalizeWhitespace(event.text);
        if (!text) return;
        const block = {
          type: 'text',
          index: paragraphIndex,
          paragraphIndex,
          runIndex: event.runIndex,
          blockIndex: blocks.length,
          text,
        };
        if (paragraph && paragraph.blockIndex == null) paragraph.blockIndex = block.blockIndex;
        blocks.push(block);
      } else if (event.type === 'image') {
        blocks.push({ type: 'image', blockIndex: blocks.length, ...event.image });
      }
    });
  }

  return {
    source: absoluteDocx,
    assetDir,
    specDir,
    paragraphs,
    images: blocks.filter((block) => block.type === 'image'),
    blocks,
    warnings,
  };
}

async function parseParagraph(paragraphXml, ctx) {
  const runMatches = [...paragraphXml.matchAll(/<w:r\b[\s\S]*?<\/w:r>/g)];
  const textParts = [];
  const events = [];
  let imageCounter = ctx.imageCounter;

  for (let runIndex = 0; runIndex < runMatches.length; runIndex += 1) {
    const runXml = runMatches[runIndex][0];
    const runText = extractRunText(runXml);
    if (runText) {
      textParts.push(runText);
      events.push({ type: 'text', runIndex, text: runText });
    }
    const drawingMatches = [...runXml.matchAll(/<w:drawing\b[\s\S]*?<\/w:drawing>/g)];
    for (const drawingMatch of drawingMatches) {
      imageCounter += 1;
      const image = await extractDrawingImage(drawingMatch[0], {
        ...ctx,
        runIndex,
        imageIndex: imageCounter,
      });
      if (image) events.push({ type: 'image', runIndex, image });
    }
  }

  return { text: textParts.join(''), events, imageCounter };
}

function extractRunText(runXml) {
  const parts = [];
  const textTag = /<w:t\b[^>]*>([\s\S]*?)<\/w:t>|<w:tab\b[^>]*\/>|<w:br\b[^>]*\/>/g;
  for (const match of runXml.matchAll(textTag)) {
    if (match[1] !== undefined) parts.push(xmlUnescape(match[1]));
    else if (match[0].startsWith('<w:tab')) parts.push('\t');
    else parts.push('\n');
  }
  return parts.join('');
}

async function extractDrawingImage(drawingXml, ctx) {
  const blipAttrs = drawingXml.match(/<a:blip\b([^>]*)>/)?.[1] || '';
  const relId = attrValue(blipAttrs, 'r:embed') || attrValue(blipAttrs, 'embed');
  const linkedId = attrValue(blipAttrs, 'r:link') || attrValue(blipAttrs, 'link');
  if (!relId && linkedId) {
    ctx.warnings.push(`paragraph ${ctx.paragraphIndex + 1} image ${ctx.imageIndex} is externally linked and was skipped.`);
    return null;
  }
  if (!relId || !ctx.rels[relId]) {
    ctx.warnings.push(`paragraph ${ctx.paragraphIndex + 1} image ${ctx.imageIndex} has no resolvable relationship id.`);
    return null;
  }
  const rel = ctx.rels[relId];
  const zipPath = resolveDocxTarget('word/document.xml', rel.target);
  const file = ctx.zip.file(zipPath);
  if (!file) {
    ctx.warnings.push(`paragraph ${ctx.paragraphIndex + 1} image ${ctx.imageIndex} target is missing in DOCX: ${zipPath}`);
    return null;
  }
  const sourceBuffer = await file.async('nodebuffer');
  const extent = parseExtent(drawingXml);
  const crop = parseSrcRect(drawingXml);
  const altText = parseAltText(drawingXml);
  const anchor = parseAnchor(drawingXml);
  const processed = await writeRenderedImage(sourceBuffer, {
    sourceName: path.basename(zipPath),
    assetDir: ctx.assetDir,
    specDir: ctx.specDir,
    imageIndex: ctx.imageIndex,
    extent,
    crop,
    warnings: ctx.warnings,
  });

  return {
    imageIndex: ctx.imageIndex,
    paragraphIndex: ctx.paragraphIndex,
    runIndex: ctx.runIndex,
    inline: /<wp:inline\b/.test(drawingXml),
    anchor: /<wp:anchor\b/.test(drawingXml) ? anchor : null,
    relId,
    originalTarget: zipPath,
    originalName: path.basename(zipPath),
    path: processed.relativePath,
    absolutePath: processed.absolutePath,
    altText,
    extent,
    crop: processed.crop,
    source: processed.source,
  };
}

async function writeRenderedImage(buffer, options) {
  const ext = path.extname(options.sourceName).toLowerCase();
  const base = `docx-image-${String(options.imageIndex).padStart(3, '0')}`;
  const targetPng = path.join(options.assetDir, `${base}.png`);
  const needsSharp = hasCrop(options.crop);
  if (!needsSharp) {
    const target = path.join(options.assetDir, `${base}${ext || '.bin'}`);
    await fs.promises.writeFile(target, buffer);
    return {
      absolutePath: target,
      relativePath: toPosixRelative(options.specDir, target),
      crop: normalizeCropForOutput(options.crop),
      source: {
        format: ext.replace('.', '') || 'unknown',
        processed: false,
        transform: 'copied-original',
      },
    };
  }

  const sharp = requireSharp();
  let pipeline;
  let metadata;
  try {
    pipeline = sharp(buffer, { animated: false, limitInputPixels: false }).rotate();
    metadata = await sharp(buffer, { animated: false, limitInputPixels: false }).metadata();
  } catch (error) {
    fail(`Cannot decode DOCX image ${options.imageIndex} (${options.sourceName}) with sharp. Install sharp-compatible image support or replace the image. ${error.message}`);
  }

  let cropBox = null;
  if (hasCrop(options.crop)) {
    if (!metadata.width || !metadata.height) fail(`Cannot apply DOCX crop to image ${options.imageIndex}: source dimensions are unavailable.`);
    cropBox = cropBoxFromSrcRect(metadata.width, metadata.height, options.crop);
    if (cropBox.width <= 0 || cropBox.height <= 0) fail(`Invalid DOCX crop rectangle for image ${options.imageIndex}.`);
    pipeline = pipeline.extract(cropBox);
  }
  const outputInfo = await pipeline.png().toFile(targetPng);
  return {
    absolutePath: targetPng,
    relativePath: toPosixRelative(options.specDir, targetPng),
    crop: { ...normalizeCropForOutput(options.crop), cropBox },
    source: {
      format: metadata.format || ext.replace('.', '') || 'unknown',
      width: metadata.width || null,
      height: metadata.height || null,
      outputWidth: outputInfo.width || null,
      outputHeight: outputInfo.height || null,
      processed: true,
      transform: 'cropped-source-pixels',
    },
  };
}

function parseRelationships(xml) {
  const rels = {};
  if (!xml) return rels;
  for (const match of xml.matchAll(/<Relationship\b([^>]*)\/>/g)) {
    const attrs = match[1];
    const id = attrValue(attrs, 'Id');
    const type = attrValue(attrs, 'Type');
    const target = attrValue(attrs, 'Target');
    const targetMode = attrValue(attrs, 'TargetMode');
    if (id && target && type === IMAGE_REL_TYPE && targetMode !== 'External') {
      rels[id] = { target, type };
    }
  }
  return rels;
}

function resolveDocxTarget(fromPart, target) {
  if (target.startsWith('/')) return target.replace(/^\/+/, '');
  const base = path.posix.dirname(fromPart.replace(/\\/g, '/'));
  return path.posix.normalize(path.posix.join(base, target.replace(/\\/g, '/')));
}

function parseExtent(xml) {
  const attrs = xml.match(/<wp:extent\b([^>]*)\/>/)?.[1] || xml.match(/<a:ext\b([^>]*)\/>/)?.[1];
  if (!attrs) return null;
  const cx = Number(attrValue(attrs, 'cx'));
  const cy = Number(attrValue(attrs, 'cy'));
  if (!Number.isFinite(cx) || !Number.isFinite(cy) || cx <= 0 || cy <= 0) return null;
  return {
    cx,
    cy,
    wIn: round(cx / EMU_PER_INCH, 4),
    hIn: round(cy / EMU_PER_INCH, 4),
  };
}

function parseSrcRect(xml) {
  const attrs = xml.match(/<a:srcRect\b([^>]*)\/>/)?.[1];
  if (!attrs) return null;
  return {
    l: percent100k(attrValue(attrs, 'l')),
    t: percent100k(attrValue(attrs, 't')),
    r: percent100k(attrValue(attrs, 'r')),
    b: percent100k(attrValue(attrs, 'b')),
  };
}

function parseAnchor(xml) {
  const hAttrs = xml.match(/<wp:positionH\b([^>]*)>/)?.[1] || '';
  const vAttrs = xml.match(/<wp:positionV\b([^>]*)>/)?.[1] || '';
  const hXml = xml.match(/<wp:positionH\b[\s\S]*?<\/wp:positionH>/)?.[0] || '';
  const vXml = xml.match(/<wp:positionV\b[\s\S]*?<\/wp:positionV>/)?.[0] || '';
  const xEmu = Number(xmlText(hXml, 'wp:posOffset')) || null;
  const yEmu = Number(xmlText(vXml, 'wp:posOffset')) || null;
  return {
    horizontalFrom: attrValue(hAttrs, 'relativeFrom') || null,
    verticalFrom: attrValue(vAttrs, 'relativeFrom') || null,
    xEmu,
    yEmu,
    xIn: xEmu ? round(xEmu / EMU_PER_INCH, 4) : null,
    yIn: yEmu ? round(yEmu / EMU_PER_INCH, 4) : null,
  };
}

function parseAltText(xml) {
  const attrs = xml.match(/<pic:cNvPr\b([^>]*)>/)?.[1] || xml.match(/<wp:docPr\b([^>]*)>/)?.[1] || '';
  return attrValue(attrs, 'descr') || attrValue(attrs, 'title') || attrValue(attrs, 'name') || '';
}

function cropBoxFromSrcRect(width, height, crop) {
  const left = Math.floor(width * crop.l / 100000);
  const top = Math.floor(height * crop.t / 100000);
  const right = Math.floor(width * crop.r / 100000);
  const bottom = Math.floor(height * crop.b / 100000);
  return {
    left,
    top,
    width: Math.max(1, width - left - right),
    height: Math.max(1, height - top - bottom),
  };
}

function hasCrop(crop) {
  return !!crop && ['l', 't', 'r', 'b'].some((key) => Number(crop[key]) > 0);
}

function normalizeCropForOutput(crop) {
  return crop || { l: 0, t: 0, r: 0, b: 0 };
}

function isSharpFriendly(ext) {
  return ['.png', '.jpg', '.jpeg', '.gif', '.webp', '.tif', '.tiff', '.svg', '.bmp'].includes(ext);
}

function requireSharp() {
  try {
    return require('sharp');
  } catch (error) {
    fail('DOCX image crop/scale processing requires sharp. Install dependencies with: npm install');
  }
}

async function readZipText(zip, name) {
  const file = zip.file(name);
  return file ? file.async('string') : '';
}

function attrValue(attrs, name) {
  const escaped = name.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
  const match = String(attrs || '').match(new RegExp(`${escaped}="([^"]*)"`, 'i'));
  return match ? xmlUnescape(match[1]) : '';
}

function xmlText(xml, tag) {
  const escaped = tag.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
  const match = String(xml || '').match(new RegExp(`<${escaped}\\b[^>]*>([\\s\\S]*?)<\\/${escaped}>`, 'i'));
  return match ? xmlUnescape(match[1]) : '';
}

function xmlUnescape(value) {
  return String(value || '')
    .replace(/&#x([0-9a-fA-F]+);/g, (_match, hex) => String.fromCodePoint(parseInt(hex, 16)))
    .replace(/&#([0-9]+);/g, (_match, dec) => String.fromCodePoint(parseInt(dec, 10)))
    .replace(/&lt;/g, '<')
    .replace(/&gt;/g, '>')
    .replace(/&amp;/g, '&')
    .replace(/&quot;/g, '"')
    .replace(/&apos;/g, "'");
}

function normalizeWhitespace(value) {
  return String(value || '')
    .replace(/\r/g, '')
    .replace(/[ \t]+\n/g, '\n')
    .replace(/\n{3,}/g, '\n\n')
    .replace(/[ \t]{2,}/g, ' ')
    .trim();
}

function percent100k(value) {
  const n = Number(value);
  return Number.isFinite(n) ? Math.max(0, Math.min(100000, n)) : 0;
}

function safeBaseName(filePath) {
  return path.basename(String(filePath || 'docx'), path.extname(String(filePath || 'docx')))
    .replace(/[^\p{L}\p{N}._-]+/gu, '-')
    .replace(/^-+|-+$/g, '') || 'docx';
}

function toPosixRelative(fromDir, targetPath) {
  const rel = path.relative(path.resolve(fromDir), path.resolve(targetPath));
  return rel.split(path.sep).join('/');
}

function requiredPath(value, name) {
  if (!value) fail(`${name} is required.`);
  return value;
}

function writeJson(filePath, value) {
  writeText(filePath, JSON.stringify(value, null, 2));
}

function writeDocxExtractionMarkdown(filePath, extracted) {
  const lines = [];
  lines.push('# DOCX 解析内容');
  lines.push('');
  lines.push('来源：' + extracted.source);
  lines.push('文本段落：' + extracted.paragraphs.length);
  lines.push('图片：' + extracted.images.length);
  lines.push('');
  if (extracted.warnings.length) {
    lines.push('## 解析警告');
    extracted.warnings.forEach((warning) => lines.push('- ' + warning));
    lines.push('');
  }
  lines.push('## 内容顺序');
  lines.push('');
  extracted.blocks.forEach((block) => {
    if (block.type === 'text') {
      lines.push('- 文本 block ' + block.blockIndex + ' / paragraph ' + block.paragraphIndex + ' / run ' + block.runIndex + ': ' + block.text);
    } else if (block.type === 'image') {
      lines.push('- 图片 block ' + block.blockIndex + ' / image ' + block.imageIndex + ' / paragraph ' + block.paragraphIndex + ' / run ' + block.runIndex + ': ' + block.path);
      if (block.altText) lines.push('  - alt: ' + block.altText.replace(/\s+/g, ' '));
      if (block.extent) lines.push('  - 显示尺寸: ' + block.extent.wIn + 'in x ' + block.extent.hIn + 'in');
      if (block.crop && ['l', 't', 'r', 'b'].some((key) => Number(block.crop[key]) > 0)) {
        lines.push('  - 已按 DOCX 裁剪处理: l=' + block.crop.l + ', t=' + block.crop.t + ', r=' + block.crop.r + ', b=' + block.crop.b);
      }
    }
  });
  lines.push('');
  lines.push('## 写 PPT JSON 时的要求');
  lines.push('');
  lines.push('- 不要按 DOCX 排版机械映射页面；应根据语义重新组织封面、目录、正文、图文页、总结页。');
  lines.push('- 插入图片时使用上方 `path`。无裁剪图片保留原始文件；带裁剪图片已按源像素裁剪为 PNG。PPT 生成阶段会按图片比例放入槽位。');
  lines.push('- 如果 DOCX 的 alt text 是工具自动生成的无意义描述，不要把它当作正文或标题。');
  lines.push('- PPT JSON 写完后再运行 `scripts/generate-pptx.js --spec ... --out ...` 生成演示文稿。');
  writeText(filePath, lines.join('\n') + '\n');
}

function writeText(filePath, value) {
  fs.mkdirSync(path.dirname(path.resolve(filePath)), { recursive: true });
  fs.writeFileSync(filePath, value, 'utf8');
}

function round(value, digits) {
  const scale = 10 ** digits;
  return Math.round(value * scale) / scale;
}

module.exports = {
  extractDocxContent,
  parseDocx,
};
