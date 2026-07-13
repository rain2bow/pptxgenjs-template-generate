'use strict';

module.exports = function createMediaTools(deps) {
  const {
    fs,
    path,
    pptx,
    FONTS,
    fail,
    svgDataUri,
    normalizeSections,
    normalizeChartData,
    addChartBlock,
  } = deps;

  const SVG_IMAGE_PNG_CACHE = new Map();

  const IMAGE_ASPECT_CACHE = new Map();

  let SHARP_MODULE = undefined;

  let SVG_IMAGE_RENDER_MODE = 'png';

  function loadSharpModule() {
    if (SHARP_MODULE) return SHARP_MODULE;
    try {
      SHARP_MODULE = require('sharp');
      return SHARP_MODULE;
    } catch (error) {
      return null;
    }
  }

  function requireSharpForPng(scope) {
    const sharp = loadSharpModule();
    if (sharp) return sharp;
    fail(scope + ' require sharp to be installed because PNG rasterization is the default. Run npm install in the skill directory, or explicitly set iconMode:"svg" / svgImageMode:"svg" if SVG output is intentional.');
  }

  function addImageAsset(slide, imagePath, box, options = {}) {
    const placedBox = fitImageBoxToAspect(imagePath, box);
    if (path.extname(imagePath).toLowerCase() === '.svg') {
      const pngData = svgImagePngData(imagePath, options.opacity);
      if (pngData) {
        slide.addImage({ data: pngData, ...placedBox });
      } else {
        const svg = readSvgWithOpacity(imagePath, options.opacity);
        slide.addImage({ data: svgDataUri(svg), ...placedBox });
      }
    } else {
      slide.addImage({ path: imagePath, ...placedBox });
    }
    return placedBox;
  }

  function svgImageCacheKey(imagePath, opacity) {
    const opacityKey = opacity == null ? 'default' : String(Math.max(0, Math.min(1, Number(opacity))));
    return path.resolve(imagePath) + ':' + opacityKey;
  }

  function svgImagePngData(imagePath, opacity) {
    if (SVG_IMAGE_RENDER_MODE === 'svg') return null;
    const data = SVG_IMAGE_PNG_CACHE.get(svgImageCacheKey(imagePath, opacity));
    if (data) return data;
    requireSharpForPng('SVG images/logos');
    fail('SVG image was not rasterized before insertion: ' + imagePath);
  }

  async function prepareSvgImageAssets(spec, specDir) {
    SVG_IMAGE_RENDER_MODE = String(spec.svgImageMode || spec.svgImageRenderMode || spec.imageMode || 'png').toLowerCase();
    if (SVG_IMAGE_RENDER_MODE === 'svg') return;
    requireSharpForPng('SVG images/logos');
    const images = collectSvgImageInputs(spec, specDir);
    if (spec.style === 'cmb') {
      const cmbMark = resolveImage(specDir, spec.logoMark || spec.logoSymbol || spec.brandLogoSymbol || 'logos/cmb-logo-mark.svg');
      if (cmbMark) images.add(cmbMark);
    }
    const jobs = [];
    for (const imagePath of images) {
      jobs.push(renderSvgImagePngToCache(imagePath, null));
      jobs.push(renderSvgImagePngToCache(imagePath, 0.2));
    }
    await Promise.all(jobs);
  }

  async function prepareImageAspectAssets(spec, specDir) {
    const images = collectLocalImageInputs(spec, specDir);
    if (!images.size) return;
    let sharp = null;
    try {
      sharp = require('sharp');
    } catch (_) {
      return;
    }
    await Promise.all(Array.from(images).map(async (imagePath) => {
      if (IMAGE_ASPECT_CACHE.has(imagePath)) return;
      try {
        const metadata = await sharp(imagePath).metadata();
        if (metadata.width && metadata.height) IMAGE_ASPECT_CACHE.set(imagePath, metadata.width / metadata.height);
      } catch (_) {
        // Keep the lightweight PNG/JPEG/SVG header readers as a fallback.
      }
    }));
  }

  function collectLocalImageInputs(value, specDir, images = new Set()) {
    if (!value) return images;
    if (typeof value === 'string') {
      if (/\.(png|jpe?g|webp|gif|bmp|tiff?|svg)(?:[?#].*)?$/i.test(value)) {
        const resolved = resolveImage(specDir, value);
        if (resolved) images.add(resolved);
      }
      return images;
    }
    if (Array.isArray(value)) {
      value.forEach((item) => collectLocalImageInputs(item, specDir, images));
      return images;
    }
    if (typeof value === 'object') Object.values(value).forEach((item) => collectLocalImageInputs(item, specDir, images));
    return images;
  }

  function collectSvgImageInputs(value, specDir, images = new Set()) {
    if (!value) return images;
    if (typeof value === 'string') {
      if (/\.svg(?:[?#].*)?$/i.test(value)) {
        const resolved = resolveImage(specDir, value);
        if (resolved) images.add(resolved);
      }
      return images;
    }
    if (Array.isArray(value)) {
      value.forEach((item) => collectSvgImageInputs(item, specDir, images));
      return images;
    }
    if (typeof value === 'object') {
      const raw = value.path || value.src || value.image || value.logo || value.logoMark || value.logoSymbol || value.brandLogoSymbol;
      if (typeof raw === 'string') collectSvgImageInputs(raw, specDir, images);
      Object.values(value).forEach((item) => collectSvgImageInputs(item, specDir, images));
    }
    return images;
  }

  async function renderSvgImagePngToCache(imagePath, opacity) {
    const key = svgImageCacheKey(imagePath, opacity);
    if (SVG_IMAGE_PNG_CACHE.has(key)) return;
    const sharp = requireSharpForPng('SVG images/logos');
    try {
      const svg = readSvgWithOpacity(imagePath, opacity);
      const buffer = await sharp(Buffer.from(svg), { density: 288 }).png().toBuffer();
      SVG_IMAGE_PNG_CACHE.set(key, 'data:image/png;base64,' + buffer.toString('base64'));
    } catch (error) {
      fail('Failed to rasterize SVG image/logo to PNG: ' + imagePath + '. ' + error.message);
    }
  }

  function fitImageBoxToAspect(imagePath, box) {
    const next = { ...box };
    const aspect = imageAspectRatio(imagePath);
    if (!aspect || aspect <= 0) return next;
    const hasW = Number(next.w) > 0;
    const hasH = Number(next.h) > 0;
    if (hasW && hasH) {
      const target = next.w / next.h;
      if (Math.abs(target - aspect) < 0.01) return next;
      if (target > aspect) {
        const fittedW = next.h * aspect;
        next.x += (next.w - fittedW) / 2;
        next.w = fittedW;
      } else {
        const fittedH = next.w / aspect;
        next.y += (next.h - fittedH) / 2;
        next.h = fittedH;
      }
    } else if (hasW) {
      next.h = next.w / aspect;
    } else if (hasH) {
      next.w = next.h * aspect;
    }
    return next;
  }

  function imageAspectRatio(imagePath) {
    if (!imagePath) return null;
    if (IMAGE_ASPECT_CACHE.has(imagePath)) return IMAGE_ASPECT_CACHE.get(imagePath);
    let ratio = null;
    const ext = path.extname(imagePath).toLowerCase();
    try {
      if (ext === '.svg') ratio = svgAspectRatio(fs.readFileSync(imagePath, 'utf8'));
      else if (ext === '.png') ratio = pngAspectRatio(fs.readFileSync(imagePath));
      else if (ext === '.jpg' || ext === '.jpeg') ratio = jpegAspectRatio(fs.readFileSync(imagePath));
    } catch (_) {
      ratio = null;
    }
    IMAGE_ASPECT_CACHE.set(imagePath, ratio);
    return ratio;
  }

  function svgAspectRatio(svg) {
    const viewBox = String(svg).match(/viewBox=["']\s*[-\d.]+\s+[-\d.]+\s+([\d.]+)\s+([\d.]+)\s*["']/i);
    if (viewBox) return Number(viewBox[1]) / Number(viewBox[2]);
    const width = String(svg).match(/\swidth=["']([\d.]+)(?:px)?["']/i);
    const height = String(svg).match(/\sheight=["']([\d.]+)(?:px)?["']/i);
    if (width && height) return Number(width[1]) / Number(height[1]);
    return null;
  }

  function pngAspectRatio(buffer) {
    if (!Buffer.isBuffer(buffer) || buffer.length < 24) return null;
    const isPng = buffer[0] === 0x89 && buffer[1] === 0x50 && buffer[2] === 0x4e && buffer[3] === 0x47;
    if (!isPng) return null;
    const width = buffer.readUInt32BE(16);
    const height = buffer.readUInt32BE(20);
    return width && height ? width / height : null;
  }

  function jpegAspectRatio(buffer) {
    if (!Buffer.isBuffer(buffer) || buffer.length < 4) return null;
    if (buffer[0] !== 0xff || buffer[1] !== 0xd8) return null;
    let offset = 2;
    while (offset + 9 < buffer.length) {
      if (buffer[offset] !== 0xff) {
        offset += 1;
        continue;
      }
      const marker = buffer[offset + 1];
      if (marker === 0xd9 || marker === 0xda) break;
      const length = buffer.readUInt16BE(offset + 2);
      if (!length || offset + 2 + length > buffer.length) break;
      const isSof = (
        (marker >= 0xc0 && marker <= 0xc3) ||
        (marker >= 0xc5 && marker <= 0xc7) ||
        (marker >= 0xc9 && marker <= 0xcb) ||
        (marker >= 0xcd && marker <= 0xcf)
      );
      if (isSof) {
        const height = buffer.readUInt16BE(offset + 5);
        const width = buffer.readUInt16BE(offset + 7);
        return width && height ? width / height : null;
      }
      offset += 2 + length;
    }
    return null;
  }

  function readSvgWithOpacity(imagePath, opacity) {
    let svg = fs.readFileSync(imagePath, 'utf8');
    if (opacity == null) return svg;
    const value = Math.max(0, Math.min(1, Number(opacity)));
    svg = svg.replace(/<\?xml[^>]*>\s*/i, '').replace(/<!DOCTYPE[^>]*>\s*/i, '').trim();
    return svg.replace(/<svg\b/i, `<svg opacity="${value}"`);
  }

  function addImagePlaceholder(slide, x, y, w, h, color, label = '图片占位') {
    slide.addShape(pptx.ShapeType.rect, { x, y, w, h, fill: { color: 'FFFFFF', transparency: 100 }, line: { color, transparency: 45, width: 0.8, dash: 'dash' } });
    slide.addText(label, { x: x + 0.25, y: y + h / 2 - 0.12, w: Math.max(0.5, w - 0.5), h: 0.24, fontFace: FONTS.mono, fontSize: 8, charSpace: 1.8, color, transparency: 45, align: 'center', margin: 0 });
  }

  function addImageOrPlaceholder(slide, ctx, image, x, y, w, h, color, label) {
    const imgPath = resolveImage(ctx.specDir, image);
    if (imgPath) {
      addImageAsset(slide, imgPath, { x, y, w, h });
      return;
    }
    addImagePlaceholder(slide, x, y, w, h, color, label || '图片占位');
  }

  function addStatementImageSlot(slide, ctx, box, color, label = '图片占位') {
    const images = normalizeMediaImages(ctx.slideSpec || {});
    if (images[0]) {
      addImageOrPlaceholder(slide, ctx, images[0], box.x, box.y, box.w, box.h, color, label);
      return 'image';
    }
    addImagePlaceholder(slide, box.x, box.y, box.w, box.h, color, label);
    return 'placeholder';
  }

  function normalizeMediaImages(data) {
    const raw = [];
    if (data.image) raw.push(data.image);
    if (Array.isArray(data.images)) raw.push(...data.images);
    if (Array.isArray(data.gallery)) raw.push(...data.gallery);
    if (Array.isArray(data.media)) raw.push(...data.media.filter((item) => item?.type === 'image' || item?.path || item?.src || typeof item === 'string'));
    return raw.filter((item) => typeof item === 'string' || item?.path || item?.src || item?.image).map((item) => (item?.image ? item.image : item));
  }

  function normalizeMediaCharts(data) {
    const raw = [];
    if (data.chart) raw.push(data.chart);
    if (Array.isArray(data.charts)) raw.push(...data.charts);
    if (Array.isArray(data.media)) raw.push(...data.media.filter((item) => item?.type === 'chart' || item?.chartType || item?.series || item?.values || item?.data));
    return raw.filter((chart) => normalizeChartData(chart).length);
  }

  function addMediaOrChart(slide, ctx, data, box, s, mode, label = 'MEDIA', index = 0) {
    const images = normalizeMediaImages(data);
    const image = images[index] || images[0];
    if (image) {
      addImageOrPlaceholder(slide, ctx, image, box.x, box.y, box.w, box.h, s.fg, label);
      return 'image';
    }
    const charts = normalizeMediaCharts(data);
    const chart = charts[index] || charts[0];
    if (chart) {
      addChartBlock(slide, ctx, { ...chart, x: box.x, y: box.y, w: box.w, h: box.h }, box, s, mode);
      return 'chart';
    }
    addImagePlaceholder(slide, box.x, box.y, box.w, box.h, s.fg, label || '图片占位');
    return 'placeholder';
  }

  function addMediaGrid(slide, ctx, data, boxes, s, mode) {
    const images = normalizeMediaImages(data);
    const charts = normalizeMediaCharts(data);
    const captions = normalizeSections(data.captions || data.items || data.sections || []);
    boxes.forEach((box, i) => {
      if (images[i]) {
        addImageOrPlaceholder(slide, ctx, images[i], box.x, box.y, box.w, box.h, s.fg, '图片占位');
      } else if (charts[i]) {
        addChartBlock(slide, ctx, { ...charts[i], x: box.x, y: box.y, w: box.w, h: box.h }, box, s, mode);
      } else {
        addImagePlaceholder(slide, box.x, box.y, box.w, box.h, s.fg, '图片占位');
      }
      const caption = imageCaption(images[i]) || charts[i]?.caption || captions[i]?.caption || captions[i]?.title || captions[i]?.label;
      addCaption(slide, caption, box.x, box.y + box.h + 0.08, box.w, s.fg, mode);
    });
  }

  function imageCaption(image) {
    return typeof image === 'string' ? '' : image?.caption || image?.title || image?.label || '';
  }

  function resolveImage(specDir, image) {
    const raw = typeof image === 'string' ? image : image?.path || image?.src;
    if (!raw || /^data:/i.test(raw) || /^https?:\/\//i.test(raw)) return null;
    const skillRoot = path.resolve(__dirname, '..', '..');
    const skillAssets = path.join(skillRoot, 'assets');
    const bases = [specDir, process.cwd(), skillAssets, skillRoot].filter(Boolean);
    const candidates = path.isAbsolute(raw)
      ? [raw]
      : bases.flatMap((base) => [path.resolve(base, raw)]);
    const seen = new Set();
    return candidates.find((p) => {
      const key = path.normalize(p).toLowerCase();
      if (seen.has(key)) return false;
      seen.add(key);
      return fs.existsSync(p);
    }) || null;
  }

  function addCaption(slide, text, x, y, w, color, mode) {
    if (!text) return;
    slide.addText(text, { x, y, w, h: 0.25, fontFace: mode === 'swiss' ? FONTS.mono : FONTS.mono, fontSize: 7.5, charSpace: 1.2, color, transparency: 35, margin: 0, fit: 'shrink' });
  }

  function addSwissBars(slide, x, y, w, h, color, transparency = 35) {
    const bars = [0.35, 0.7, 0.5, 0.95, 0.42, 0.82, 0.6];
    bars.forEach((height, i) => {
      slide.addShape(pptx.ShapeType.rect, { x: x + i * (w / bars.length), y: y + h * (1 - height), w: w / bars.length - 0.03, h: h * height, fill: { color, transparency }, line: { color, transparency: 100 } });
    });
  }


  return {
    addImageAsset,
    prepareSvgImageAssets,
    prepareImageAspectAssets,
    addImagePlaceholder,
    addImageOrPlaceholder,
    addStatementImageSlot,
    normalizeMediaImages,
    normalizeMediaCharts,
    addMediaOrChart,
    addMediaGrid,
    imageCaption,
    resolveImage,
    addCaption,
    addSwissBars,
  };
};
