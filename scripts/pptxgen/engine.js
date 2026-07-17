const fs = require('node:fs');
const path = require('node:path');
const process = require('node:process');
const pptxgen = require('pptxgenjs');
const JSZip = require('jszip');
const {
  SLIDE,
  THEMES,
  FONTS,
  TYPOGRAPHY,
  READABILITY,
  BASIC_ICON_NAMES,
  ICON_ALIASES,
  LUCIDE_STATIC_ICON_DIR,
  isSupportedStyle,
  defaultThemeForStyle,
} = require('./config');
const { fail } = require('./errors');
const { speakerNotesText } = require('./speaker-notes');
const createMagazineTemplate = require('./templates/magazine');
const createSwissTemplate = require('./templates/swiss');
const createCmbTemplate = require('./templates/cmb');
const createIconTools = require('./icons');
const createBlockTools = require('./blocks');
const createMediaTools = require('./media');
const createValidationTools = require('./validation');
const { validateCanonicalSpec, createRendererSpec, createRendererSlide } = require('./layout-schema');

const pptx = new pptxgen();
const CHART_TYPES = {
  bar: pptx.ChartType.bar,
  column: pptx.ChartType.bar,
  line: pptx.ChartType.line,
  pie: pptx.ChartType.pie,
  doughnut: pptx.ChartType.doughnut,
  area: pptx.ChartType.area,
  radar: pptx.ChartType.radar,
  scatter: pptx.ChartType.scatter,
};

const iconTools = createIconTools({ fs, path, process, pptx, FONTS, BASIC_ICON_NAMES, ICON_ALIASES, LUCIDE_STATIC_ICON_DIR, fail });
const {
  iconAlias,
  prepareIconAssets,
  normalizeHex,
  addSvgIcon,
  defaultContentIcon,
  itemIcon,
  addInlineIcon,
  addBulletIcon,
} = iconTools;

const blockTools = createBlockTools({ pptx, FONTS, TYPOGRAPHY, READABILITY, CHART_TYPES, fail, safeBox, chartPalette });
const {
  renderDataBlocks,
  hasExplicitBox,
  addTextBlock,
  addChartBlock,
  addTableBlock,
  normalizeChartData,
  normalizeTableRows,
} = blockTools;

const mediaTools = createMediaTools({ fs, path, pptx, FONTS, fail, svgDataUri, normalizeSections, normalizeChartData, addChartBlock });
const {
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
} = mediaTools;

const validationTools = createValidationTools({ fail, normalizeMediaImages, normalizeMediaCharts, normalizeChartData, normalizeTableRows, normalizeSections, resolveImage, clamp });
const {
  normalizeLayoutCompatibility,
  validateSpecSlots,
  warnLayoutVariety,
  resolveMediaSlotCount,
  normalizeSlotItemsForValidation,
} = validationTools;
function normalizeSpec(spec, options = {}) {
  spec.style = spec.style || 'magazine';
  if (!isSupportedStyle(spec.style)) fail(`Unsupported style: ${spec.style}`);
  spec.theme = spec.theme || defaultThemeForStyle(spec.style);
  if (!THEMES[spec.style][spec.theme]) fail(`Unsupported theme "${spec.theme}" for style "${spec.style}"`);
  if (!Array.isArray(spec.slides) || spec.slides.length === 0) fail('Spec must include a non-empty slides array.');
  if (spec.diversifyLayouts === true) {
    fail('diversifyLayouts is no longer supported. Edit each slide.layout in JSON manually; the generator never changes layout automatically.');
  }
  validateCanonicalSpec(spec, fail);
  const rendererSpec = createRendererSpec(spec);
  normalizeLayoutCompatibility(rendererSpec);
  validateSpecSlots(rendererSpec, { specDir: options.specDir || process.cwd() });
  warnLayoutVariety(rendererSpec);
  spec.__normalized = true;
}

async function buildDeck(spec, specDir, outPath) {
  if (!spec.__normalized) normalizeSpec(spec, { specDir });
  pptx.layout = 'LAYOUT_WIDE';
  pptx.author = spec.author || 'PPTXGenJS Template Generator';
  pptx.subject = spec.subtitle || spec.title || 'PPTX deck';
  pptx.title = spec.title || 'Untitled deck';
  pptx.company = spec.company || '';
  pptx.lang = 'zh-CN';
  pptx.theme = {
    headFontFace: FONTS.zh,
    bodyFontFace: FONTS.zh,
    lang: 'zh-CN',
  };
  pptx.defineLayout({ name: 'CUSTOM_WIDE', width: SLIDE.w, height: SLIDE.h });
  pptx.layout = 'CUSTOM_WIDE';

  const theme = THEMES[spec.style][spec.theme];
  await prepareIconAssets(spec, theme);
  await prepareImageAspectAssets(spec, specDir);
  await prepareSvgImageAssets(spec, specDir);
  spec.slides.forEach((slideSpec, index) => {
    const slide = pptx.addSlide();
    enforceReadableSlideText(slide);
    const rendererSlideSpec = createRendererSlide(slideSpec);
    const ctx = { spec, slideSpec: rendererSlideSpec, sourceSlideSpec: slideSpec, theme, specDir, index, total: spec.slides.length };
    renderByStyle(spec.style, slide, ctx);
    addSpeakerNotes(slide, slideSpec, ctx);
  });

  fs.mkdirSync(path.dirname(outPath), { recursive: true });
  await pptx.writeFile({ fileName: outPath });
  if (spec.disableTextAutofit !== false) await disableTextAutofit(outPath);
  console.log(`Wrote ${outPath}`);
  console.log(`Slides: ${spec.slides.length}`);
  console.log(`Style: ${spec.style} / ${theme.name}`);
}

function addSpeakerNotes(slide, slideSpec, ctx) {
  const notes = speakerNotesText(slideSpec, ctx);
  if (!notes) return;
  slide.addNotes(notes);
}

async function disableTextAutofit(pptxPath) {
  const buffer = await fs.promises.readFile(pptxPath);
  const zip = await JSZip.loadAsync(buffer);
  const slideFiles = Object.keys(zip.files).filter((name) => /^ppt\/slides\/slide\d+\.xml$/.test(name));
  const chartFiles = Object.keys(zip.files).filter((name) => /^ppt\/charts\/chart\d+\.xml$/.test(name));
  let changed = false;
  await Promise.all(slideFiles.map(async (name) => {
    const xml = await zip.file(name).async('string');
    const next = forceNoAutofitXml(xml);
    if (next !== xml) {
      zip.file(name, next);
      changed = true;
    }
  }));
  await Promise.all(chartFiles.map(async (name) => {
    const xml = await zip.file(name).async('string');
    const next = forceChartTypographyXml(xml);
    if (next !== xml) {
      zip.file(name, next);
      changed = true;
    }
  }));
  if (!changed) return;
  const output = await zip.generateAsync({ type: 'nodebuffer', compression: 'DEFLATE' });
  await fs.promises.writeFile(pptxPath, output);
}

function forceChartTypographyXml(xml) {
  return String(xml)
    .replace(/typeface="[^"]*"/g, 'typeface="' + FONTS.zh + '"')
    .replace(/(<a:defRPr\b[^>]*?)\ssz="\d+"/g, '$1 sz="' + (TYPOGRAPHY.dense * 100) + '"');
}

function forceNoAutofitXml(xml) {
  return String(xml)
    .replace(/<a:bodyPr\b([^>]*)\/>/g, '<a:bodyPr$1><a:noAutofit/></a:bodyPr>')
    .replace(/<a:bodyPr\b([^>]*)>([\s\S]*?)<\/a:bodyPr>/g, (_match, attrs, body) => {
      let nextBody = body.replace(/<a:(?:spAutoFit|normAutofit)\b[^>]*(?:\/>|>[\s\S]*?<\/a:(?:spAutoFit|normAutofit)>)/g, '');
      if (!/<a:noAutofit\b/.test(nextBody)) nextBody = insertNoAutofit(nextBody);
      return `<a:bodyPr${attrs}>${nextBody}</a:bodyPr>`;
    });
}

function insertNoAutofit(body) {
  const warp = body.match(/^\s*<a:prstTxWarp\b[^>]*(?:\/>|>[\s\S]*?<\/a:prstTxWarp>)/);
  if (!warp) return '<a:noAutofit/>' + body;
  return body.slice(0, warp[0].length) + '<a:noAutofit/>' + body.slice(warp[0].length);
}
function enforceReadableSlideText(slide) {
  if (slide.__readabilityPatched) return;
  const originalAddText = slide.addText.bind(slide);
  slide.addText = (text, options = {}) => {
    const boxOptions = readableTextOptions(options, text);
    const limitedText = warnTextExceedsBox(text, boxOptions);
    const readableText = Array.isArray(limitedText)
      ? limitedText.map((run) => ({ ...run, options: readableTextOptions(run.options || {}, run.text) }))
      : limitedText;
    return originalAddText(readableText, boxOptions);
  };
  slide.__readabilityPatched = true;
}

function readableTextOptions(options, text) {
  if (!options) return options;
  const next = { ...options };
  if (next.fit === 'shrink' && next.allowAutoFit !== true) delete next.fit;
  if (next.noReadabilityScale) return next;
  if (!isSymbolText(next, text)) {
    next.fontFace = typographyFontFace(text);
    next.fontSize = typographyFontSize(next.fontSize, next.typographyRole);
    delete next.typographyRole;
    const minBoxH = (READABILITY.minFontSize * 1.28) / 72;
    if (typeof next.h === 'number' && next.h < minBoxH) next.h = minBoxH;
  }
  return next;
}

function typographyFontFace(text) {
  return isPureEnglishText(text) ? FONTS.en : FONTS.zh;
}

function typographyFontSize(size, role) {
  if (role === 'coverTitle') return TYPOGRAPHY.coverTitle;
  if (role === 'pageTitle') return TYPOGRAPHY.pageTitle;
  if (role === 'itemTitle') return TYPOGRAPHY.itemTitle;
  if (role === 'body') return TYPOGRAPHY.body;
  if (role === 'dense') return TYPOGRAPHY.dense;
  const value = Number(size);
  if (!Number.isFinite(value) || value <= 0) return TYPOGRAPHY.body;
  if (value >= 40) return TYPOGRAPHY.coverTitle;
  if (value >= 23) return TYPOGRAPHY.pageTitle;
  if (value >= 15) return TYPOGRAPHY.itemTitle;
  if (value >= 13) return TYPOGRAPHY.body;
  return TYPOGRAPHY.dense;
}

function isPureEnglishText(text) {
  const raw = Array.isArray(text)
    ? richTextPlainText(text)
    : String(text || '');
  const cleaned = raw.trim();
  if (!cleaned) return false;
  if (/[\u2E80-\u9FFF\uF900-\uFAFF]/.test(cleaned)) return false;
  return /[A-Za-z0-9]/.test(cleaned);
}

function isSymbolText(options, text) {
  const raw = Array.isArray(text) ? '' : String(text || '').trim();
  return options.fontFace === 'Segoe UI Symbol' || (raw.length <= 1 && options.align === 'center' && options.valign === 'mid');
}

function warnTextExceedsBox(text, options = {}) {
  if (!options || options.noTextLimitWarning || options.noTextLimit || options.allowOverflowText) return text;
  const raw = Array.isArray(text)
    ? richTextPlainText(text)
    : String(text || '');
  if (!raw.trim() || isSymbolText(options, raw)) return text;
  const maxVisual = estimatedBoxTextCapacity(options, Array.isArray(text) ? richTextMaxFontSize(text, options) : undefined, raw);
  if (!Number.isFinite(maxVisual) || maxVisual <= 0) return text;
  const actualVisual = textVisualLength(raw);
  if (actualVisual > maxVisual) warnTextOverCapacity(raw, actualVisual, maxVisual, options);
  return text;
}

function richTextPlainText(runs) {
  return (runs || []).map((run) => String(run && run.text != null ? run.text : '') + (run && run.options && run.options.breakLine ? '\n' : '')).join('');
}

function richTextMaxFontSize(runs, options) {
  return runs.reduce((max, run) => {
    const size = Number(run && run.options && run.options.fontSize);
    return Number.isFinite(size) ? Math.max(max, size) : max;
  }, Number(options.fontSize) || READABILITY.minFontSize);
}

function estimatedBoxTextCapacity(options = {}, explicitFontSize, sampleText = '') {
  const w = Number(options.w);
  const h = Number(options.h);
  if (!Number.isFinite(w) || !Number.isFinite(h) || w <= 0 || h <= 0) return Infinity;
  const fontSize = Number(explicitFontSize || options.fontSize || READABILITY.minFontSize);
  if (!Number.isFinite(fontSize) || fontSize <= 0) return Infinity;
  const margins = textBoxMargins(options.margin);
  const boxW = roundCapacityDimension(Math.max(0.05, w - margins.left - margins.right));
  const boxH = roundCapacityDimension(Math.max(0.05, h - margins.top - margins.bottom));
  const lineHeight = Number(options.textLimitLineHeight || options.lineHeight) || defaultTextLimitLineHeight(sampleText);
  const charWidthRatio = Number(options.textLimitCharWidthRatio) || defaultTextLimitCharWidthRatio(sampleText);
  const charsPerLine = Math.max(1, (boxW * 72) / (fontSize * charWidthRatio));
  const lines = Math.max(1, Math.floor((boxH * 72) / (fontSize * lineHeight)));
  const reserve = Number(options.textLimitReserveRatio);
  const reserveRatio = Number.isFinite(reserve) ? clamp(reserve, 0.55, 1) : defaultTextLimitReserveRatio(sampleText);
  return Math.max(1, Math.floor(charsPerLine * lines * reserveRatio));
}

function roundCapacityDimension(value) {
  return Math.round(Number(value) * 1000) / 1000;
}

function defaultTextLimitCharWidthRatio(text = '') {
  return /[\u2E80-\u9FFF\uF900-\uFAFF]/.test(String(text || '')) ? 1 : 0.62;
}

function defaultTextLimitLineHeight(text = '') {
  return /[\u2E80-\u9FFF\uF900-\uFAFF]/.test(String(text || '')) ? 1.12 : 1.18;
}

function defaultTextLimitReserveRatio(text = '') {
  return /[\u2E80-\u9FFF\uF900-\uFAFF]/.test(String(text || '')) ? 0.99 : 0.92;
}

function textBoxMargins(margin) {
  if (typeof margin === 'number') return { left: margin, right: margin, top: margin, bottom: margin };
  if (Array.isArray(margin)) {
    const [top = 0, right = top, bottom = top, left = right] = margin.map((value) => Number(value) || 0);
    return { left, right, top, bottom };
  }
  if (margin && typeof margin === 'object') {
    return {
      left: Number(margin.left ?? margin.l ?? 0) || 0,
      right: Number(margin.right ?? margin.r ?? 0) || 0,
      top: Number(margin.top ?? margin.t ?? 0) || 0,
      bottom: Number(margin.bottom ?? margin.b ?? 0) || 0,
    };
  }
  return { left: 0, right: 0, top: 0, bottom: 0 };
}

let TEXT_CAPACITY_WARNING_COUNT = 0;
function warnTextOverCapacity(original, actualVisual, maxVisual, options = {}) {
  if (options.silentTextLimit || options.silentTextLimitWarning || TEXT_CAPACITY_WARNING_COUNT >= 30) return;
  TEXT_CAPACITY_WARNING_COUNT += 1;
  const preview = String(original || '').replace(/\s+/g, ' ').slice(0, 56);
  console.warn('Warning: text may overflow box (' + Number(options.w || 0).toFixed(2) + 'x' + Number(options.h || 0).toFixed(2) + '); estimated capacity ' + Math.floor(maxVisual) + ', got ' + Math.ceil(actualVisual) + '. Shorten this text in JSON and regenerate the PPTX, or split/enlarge the card: ' + preview);
}
function renderByStyle(style, slide, ctx) {
  const renderer = getTemplateRenderers()[style];
  if (!renderer) fail(`Unsupported style renderer: ${style}`);
  renderer(slide, ctx);
}

let TEMPLATE_RENDERERS = null;

function getTemplateRenderers() {
  if (TEMPLATE_RENDERERS) return TEMPLATE_RENDERERS;
  const api = createTemplateApi();
  const magazineTemplate = createMagazineTemplate(api);
  const swissTemplate = createSwissTemplate(api);
  const cmbTemplate = createCmbTemplate(api);
  TEMPLATE_RENDERERS = {
    magazine: magazineTemplate.render,
    swiss: swissTemplate.render,
    cmb: cmbTemplate.render,
  };
  return TEMPLATE_RENDERERS;
}

function createTemplateApi() {
  return {
    pptx,
    SLIDE,
    THEMES,
    FONTS,
    TYPOGRAPHY,
    READABILITY,
    BASIC_ICON_NAMES,
    ICON_ALIASES,
    CHART_TYPES,
    clamp,
    safeBox,
    addPageHead,
    addCallout,
    addBullets,
    normalizeBulletItem,
    addSvgIcon,
    addInlineIcon,
    addBulletIcon,
    defaultContentIcon,
    itemIcon,
    iconAlias,
    addImageAsset,
    addImagePlaceholder,
    addImageOrPlaceholder,
    addStatementImageSlot,
    normalizeMediaImages,
    normalizeMediaCharts,
    resolveMediaSlotCount,
    addMediaOrChart,
    addMediaGrid,
    imageCaption,
    addCaption,
    addSwissBars,
    renderDataBlocks,
    addTextBlock,
    addChartBlock,
    addTableBlock,
    normalizeChartData,
    normalizeTableRows,
    normalizeSections,
    autoColumns,
    autoCardColumns,
    clampColumns,
    addRadialConnector,
    chartPalette,
    addDecorativeBackground,
    addChrome,
    addBrandLogo,
    addFoot,
    resolveImage,
    estimateBulletedLineCount,
    estimateTextHeight,
    distributeRowHeights,
    fitTitle,
    pageHeadY,
    pageHeadSafeBottom,
    svgDataUri,
    svgEsc,
    normalizeHex,
    compareBulletItems,
    textVisualLength,
  };
}

function estimateBulletedLineCount(text, w, fontSize) {
  const effectiveW = Math.max(0.05, Number(w || 0) - 0.38);
  const charsPerLine = Math.max(1, (effectiveW * 72) / Math.max(1, fontSize));
  return String(text || '').split(/\r?\n/).reduce((sum, line) => sum + Math.max(1, Math.ceil(textVisualLength(line) / charsPerLine)), 0);
}
function hasBrandHeader(ctx) {
  return !!(ctx?.slideSpec?.logoHeader || ctx?.slideSpec?.brandLogoHeader || ctx?.spec?.logoHeader || ctx?.spec?.brandLogoHeader);
}

function pageHeadY(ctx, fallback) {
  return Number(ctx?.slideSpec?.headY || ctx?.spec?.headY) || fallback;
}

function pageHeadSafeBottom(ctx, fallback, gap = 0.24) {
  const y = pageHeadY(ctx, fallback);
  const hasSubtitle = !!ctx?.slideSpec?.subtitle;
  return y + (hasSubtitle ? 1.78 : 1.35) + gap;
}
function addDecorativeBackground(slide, ctx, mode, baseColor, emphasized = false) {
  const svg = decorativeBackgroundSvg(ctx, mode, baseColor, emphasized);
  slide.background = { path: `background-${mode}.svg`, data: svgDataUri(svg) };
}

function svgDataUri(svg) {
  return `data:image/svg+xml;base64,${Buffer.from(svg).toString('base64')}`;
}

function svgEsc(value) {
  return String(value || '')
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;');
}

function decorativeBackgroundSvg(ctx, mode, baseColor, emphasized) {
  const scale = 120;
  const w = Math.round(SLIDE.w * scale);
  const h = Math.round(SLIDE.h * scale);
  const headerH = hasBrandHeader(ctx) ? (Number(ctx.slideSpec.logoHeaderBandH || ctx.spec.logoHeaderBandH) || 0.78) * scale : 0;
  const bodyH = Math.max(1, h - headerH);
  const theme = ctx.theme;
  const base = normalizeHex(baseColor || theme.paper || 'FFFFFF');
  const header = normalizeHex(ctx.slideSpec.logoHeaderBandColor || ctx.spec.logoHeaderBandColor || 'FFFFFF');
  const shade = normalizeHex(theme.ink || theme.accent || '111111');
  const accent = normalizeHex(theme.accent || theme.inkTint || theme.ink || '111111');
  const grid = normalizeHex(emphasized ? theme.accentOn || theme.paper || 'FFFFFF' : theme.grey2 || theme.paperTint || 'D8D8D8');
  const gradientColor = emphasized ? shade : accent;
  const topOpacity = emphasized ? 0.08 : 0.02;
  const midOpacity = emphasized ? 0.2 : 0.075;
  const bottomOpacity = emphasized ? 0.38 : 0.19;
  const gridOpacity = mode === 'swiss' ? (emphasized ? 0.13 : 0.2) : (emphasized ? 0.07 : 0.11);
  const lineW = mode === 'swiss' ? 0.65 : 0.45;
  const gridLeft = 0.78 * scale;
  const gridRight = (SLIDE.w - 0.78) * scale;
  const colStep = 0.72 * scale;
  const rowStep = 0.5 * scale;
  const gridLines = [];
  for (let x = gridLeft; x <= gridRight + 0.1; x += colStep) {
    const xx = Math.round(x * 10) / 10;
    gridLines.push(`<line x1="${xx}" y1="${headerH}" x2="${xx}" y2="${h}" />`);
  }
  for (let y = headerH; y <= h + 0.1; y += rowStep) {
    const yy = Math.round(y * 10) / 10;
    gridLines.push(`<line x1="${gridLeft}" y1="${yy}" x2="${gridRight}" y2="${yy}" />`);
  }
  const headerRect = headerH > 0 ? `<rect x="0" y="0" width="${w}" height="${headerH}" fill="#${svgEsc(header)}"/>` : '';
  if (mode === 'swiss' && ctx?.spec?.style === 'swiss') {
    return `<svg xmlns="http://www.w3.org/2000/svg" width="${w}" height="${h}" viewBox="0 0 ${w} ${h}">
<rect x="0" y="0" width="${w}" height="${h}" fill="#${svgEsc(base)}"/>
${headerRect}
<g fill="none" stroke="#${svgEsc(grid)}" stroke-width="${lineW}" stroke-opacity="${gridOpacity}" vector-effect="non-scaling-stroke">
${gridLines.join('\n')}
</g>
</svg>`;
  }
  return `<svg xmlns="http://www.w3.org/2000/svg" width="${w}" height="${h}" viewBox="0 0 ${w} ${h}">
<defs>
  <linearGradient id="bgShade" x1="0" y1="${headerH}" x2="0" y2="${h}" gradientUnits="userSpaceOnUse">
    <stop offset="0%" stop-color="#${svgEsc(gradientColor)}" stop-opacity="${topOpacity}"/>
    <stop offset="56%" stop-color="#${svgEsc(gradientColor)}" stop-opacity="${midOpacity}"/>
    <stop offset="100%" stop-color="#${svgEsc(gradientColor)}" stop-opacity="${bottomOpacity}"/>
  </linearGradient>
</defs>
<rect x="0" y="0" width="${w}" height="${h}" fill="#${svgEsc(base)}"/>
${headerRect}
<rect x="0" y="${headerH}" width="${w}" height="${bodyH}" fill="url(#bgShade)"/>
<g fill="none" stroke="#${svgEsc(grid)}" stroke-width="${lineW}" stroke-opacity="${gridOpacity}" vector-effect="non-scaling-stroke">
${gridLines.join('\n')}
</g>
</svg>`;
}
function addChrome(slide, ctx, color, mode) {
  const left = ctx.slideSpec.chromeLeft || ctx.spec.title || 'PPTX Deck';
  const right = ctx.slideSpec.chromeRight || `${String(ctx.index + 1).padStart(2, '0')} / ${String(ctx.total).padStart(2, '0')}`;
  const font = mode === 'swiss' ? FONTS.mono : FONTS.mono;
  const hasHeaderLogo = !!(ctx.slideSpec.logoHeader || ctx.slideSpec.brandLogoHeader || ctx.spec.logoHeader || ctx.spec.brandLogoHeader);
  const logoBox = hasHeaderLogo
    ? {
        x: SLIDE.marginX,
        y: mode === 'swiss' ? 0.12 : 0.14,
        w: Number(ctx.slideSpec.logoHeaderW || ctx.spec.logoHeaderW) || 1.72,
        h: Number(ctx.slideSpec.logoHeaderH || ctx.spec.logoHeaderH) || 0.54,
      }
    : { x: SLIDE.marginX, y: mode === 'swiss' ? 0.25 : 0.28, w: 0.3, h: 0.3 };
  const headerBand = hasHeaderLogo && ctx.slideSpec.logoHeaderBand !== false && ctx.spec.logoHeaderBand !== false;
  const headerBandH = Number(ctx.slideSpec.logoHeaderBandH || ctx.spec.logoHeaderBandH) || 0.78;
  const headerTextColor = hasHeaderLogo ? (ctx.slideSpec.logoHeaderTextColor || ctx.spec.logoHeaderTextColor || ctx.theme.ink || '111111') : color;
  if (headerBand) {
    slide.addShape(pptx.ShapeType.rect, {
      x: 0,
      y: 0,
      w: SLIDE.w,
      h: headerBandH,
      fill: { color: ctx.slideSpec.logoHeaderBandColor || ctx.spec.logoHeaderBandColor || 'FFFFFF', transparency: 0 },
      line: { color: ctx.slideSpec.logoHeaderBandColor || ctx.spec.logoHeaderBandColor || 'FFFFFF', transparency: 100 },
    });
    if (ctx.theme.accent) {
      slide.addShape(pptx.ShapeType.line, { x: 0, y: headerBandH, w: SLIDE.w, h: 0, line: { color: ctx.theme.accent, transparency: 8, width: 1.1 } });
    }
  }
  const logoAdded = addBrandLogo(slide, ctx, logoBox, { variant: hasHeaderLogo ? 'header' : 'symbol', backplate: hasHeaderLogo && !headerBand });
  const leftX = logoAdded ? SLIDE.marginX + logoBox.w + 0.22 : SLIDE.marginX;
  slide.addText(left, {
    x: leftX,
    y: mode === 'swiss' ? 0.35 : 0.38,
    w: logoAdded ? Math.max(2.2, 6.2 - logoBox.w - 0.22) : 6.2,
    h: 0.2,
    fontFace: font,
    fontSize: mode === 'swiss' ? 7.8 : 7,
    charSpace: 1.3,
    color: headerTextColor,
    margin: 0,
    fit: 'shrink',
    breakLine: false,
  });
  slide.addText(right, {
    x: SLIDE.w - SLIDE.marginX - 2.5,
    y: mode === 'swiss' ? 0.35 : 0.38,
    w: 2.5,
    h: 0.2,
    fontFace: font,
    fontSize: mode === 'swiss' ? 7.8 : 7,
    charSpace: 1.3,
    color: headerTextColor,
    align: 'right',
    margin: 0,
    fit: 'shrink',
    breakLine: false,
  });
}

function addBrandLogo(slide, ctx, box, options = {}) {
  const logoPath = resolveBrandLogo(ctx, options.variant || 'symbol');
  if (!logoPath) return false;
  if (options.backplate) {
    slide.addShape(pptx.ShapeType.rect, {
      x: box.x - 0.08,
      y: box.y - 0.06,
      w: box.w + 0.16,
      h: box.h + 0.12,
      fill: { color: options.backplateColor || 'FFFFFF', transparency: options.backplateTransparency ?? 0 },
      line: { color: options.backplateColor || 'FFFFFF', transparency: 100 },
    });
  }
  addImageAsset(slide, logoPath, { x: box.x, y: box.y, w: box.w, h: box.h });
  return true;
}

function resolveBrandLogo(ctx, variant = 'symbol') {
  const data = ctx.slideSpec || {};
  const spec = ctx.spec || {};
  const raw = variant === 'header'
    ? data.logoHeader || data.brandLogoHeader || spec.logoHeader || spec.brandLogoHeader || data.logoFull || data.brandLogoFull || spec.logoFull || spec.brandLogoFull || data.logo || spec.logo || data.brandLogo || spec.brandLogo
    : variant === 'full'
      ? data.logoFull || data.brandLogoFull || spec.logoFull || spec.brandLogoFull || data.logoMark || spec.logoMark || data.logoSymbol || data.brandLogoSymbol || spec.logoSymbol || spec.brandLogoSymbol || data.logo || spec.logo || data.brandLogo || spec.brandLogo
      : data.logoMark || spec.logoMark || data.logoSymbol || data.brandLogoSymbol || spec.logoSymbol || spec.brandLogoSymbol || data.logo || spec.logo || data.brandLogo || spec.brandLogo;
  return resolveImage(ctx.specDir, raw);
}
function addFoot(slide, ctx, color, mode, text) {
  const y = 7.02;
  slide.addShape(pptx.ShapeType.line, {
    x: SLIDE.marginX,
    y: y - 0.1,
    w: SLIDE.w - SLIDE.marginX * 2,
    h: 0,
    line: { color, transparency: mode === 'swiss' ? 70 : 78, width: 0.6 },
  });
  slide.addText(text || ctx.slideSpec.foot || ctx.spec.subtitle || '', {
    x: SLIDE.marginX,
    y,
    w: 8,
    h: 0.22,
    fontFace: mode === 'swiss' ? FONTS.mono : FONTS.mono,
    fontSize: mode === 'swiss' ? 7.5 : 7,
    charSpace: 1.2,
    color,
    transparency: 35,
    margin: 0,
    fit: 'shrink',
  });
}

function cellText(cell) {
  if (cell == null) return '';
  if (typeof cell === 'string' || typeof cell === 'number') return String(cell);
  return String(cell.title || cell.text || cell.body || cell.label || '');
}
function addPageHead(slide, data, color, mode, y = 1.0) {
  y = Number(data.headY) || y;
  slide.addText(data.kicker || '', { x: 0.78, y, w: 5.8, h: 0.25, fontFace: mode === 'swiss' ? FONTS.mono : FONTS.mono, fontSize: mode === 'swiss' ? 8 : 8.3, charSpace: mode === 'swiss' ? 1.6 : 2, color, transparency: 35, margin: 0, fit: 'shrink' });
  slide.addText(data.title || '', { x: 0.75, y: y + 0.45, w: 10.8, h: 0.9, fontFace: mode === 'swiss' ? FONTS.sans : FONTS.serifZh, fontSize: fitTitle(data.title || '', mode === 'swiss' ? 34 : 36, mode === 'swiss' ? 25 : 27), bold: mode !== 'swiss', color, margin: 0, fit: 'shrink' , typographyRole: 'coverTitle' });
  if (data.subtitle) {
    slide.addText(data.subtitle, { x: 0.78, y: y + 1.38, w: 7.2, h: 0.4, fontFace: mode === 'swiss' ? FONTS.sansZh : FONTS.serifZh, fontSize: mode === 'swiss' ? 13 : 15, color, transparency: 30, margin: 0, fit: 'shrink' });
  }
}

function addCallout(slide, text, x, y, w, h, color, fill) {
  if (!text) return;
  slide.addShape(pptx.ShapeType.rect, { x, y, w, h, fill: { color: fill || 'FFFFFF', transparency: 45 }, line: { color, transparency: 100 } });
  slide.addShape(pptx.ShapeType.line, { x, y, w: 0, h, line: { color, transparency: 5, width: 2 } });
  slide.addText(text, { x: x + 0.22, y: y + 0.15, w: w - 0.35, h: h - 0.25, fontFace: FONTS.serifZh, fontSize: 13.5, color, margin: 0, fit: 'shrink', valign: 'mid' });
}

function addBullets(slide, items, x, y, w, h, color, transparency, mode) {
  if (!items.length) return;
  const normalized = items.map(normalizeBulletItem).filter((item) => item.text);
  if (!normalized.length) return;
  const rowH = clamp(h / normalized.length, 0.28, mode === 'swiss' ? 0.46 : 0.5);
  const fontSize = Math.min(mode === 'swiss' ? 10.5 : 11, Math.max(7.2, rowH * 23));
  const iconSize = clamp(rowH * 0.34, 0.09, 0.16);
  const iconGap = clamp(rowH * 0.28, 0.1, 0.16);
  normalized.forEach((item, i) => {
    const iy = y + i * rowH + Math.max(0.02, (rowH - iconSize) / 2);
    const textY = y + i * rowH + 0.02;
    addBulletIcon(slide, item.icon, x, iy, iconSize, item.iconColor || color, item.iconFill, transparency, mode, i + 1);
    slide.addText(item.text, {
      x: x + iconSize + iconGap,
      y: textY,
      w: Math.max(0.2, w - iconSize - iconGap),
      h: Math.max(0.18, rowH - 0.02),
      fontFace: FONTS.sansZh,
      fontSize,
      bold: !!item.bold,
      color: item.color || color,
      transparency: item.transparency ?? transparency,
      margin: 0,
      fit: 'shrink',
      breakLine: false,
      valign: 'mid',
    });
  });
}

function compareBulletItems(items) {
  return normalizeSections(items).map((item) => {
    const title = item.title || item.label || item.text || '';
    const body = item.body || item.desc || item.note || item.summary || item.detail || '';
    return {
      ...item,
      text: title && body ? title + ': ' + body : title || body,
    };
  });
}

function normalizeBulletItem(item) {
  if (typeof item === 'string') return { text: item, icon: 'dot' };
  return {
    text: item.text || item.title || item.label || '',
    icon: item.icon || item.bulletIcon || 'dot',
    iconColor: item.iconColor,
    iconFill: item.iconFill,
    color: item.color,
    transparency: item.transparency,
    bold: item.bold,
  };
}

function autoCardColumns(count) {
  const n = Number(count) || 0;
  if (n <= 1) return 1;
  if (n <= 4) return n;
  if (n <= 6) return 3;
  return 4;
}
function autoColumns(count, max = 3) {
  const n = Number(count) || 0;
  if (n <= 1) return 1;
  if (n === 2) return 2;
  if (n === 4) return 2;
  if (n <= max) return n;
  if (n <= 6) return Math.min(3, max);
  return max;
}

function clampColumns(value, min, max) {
  return Math.max(min, Math.min(max, Math.round(Number(value) || min)));
}
function normalizeSections(items) {
  if (!items) return [];
  if (typeof items === 'string') return [{ body: items }];
  if (!Array.isArray(items)) return Object.entries(items).map(([title, body]) => ({ title, body: String(body) }));
  return items.map((item) => typeof item === 'string' ? { body: item } : item);
}

function chartPalette(ctx, mode) {
  if (Array.isArray(ctx.theme.chartColors) && ctx.theme.chartColors.length) return ctx.theme.chartColors;
  if (mode === 'swiss') return [ctx.theme.accent, ctx.theme.ink, ctx.theme.grey3, ctx.theme.grey2, ctx.theme.grey1];
  return [ctx.theme.ink, ctx.theme.inkTint, '6B6258', 'A39A8F', ctx.theme.paperTint];
}

function safeBox(box, defaults) {
  const merged = { ...defaults, ...box };
  const maxBottom = merged.allowUnsafe ? SLIDE.h - 0.12 : 6.85;
  const x = clamp(Number(merged.x ?? defaults.x), 0.1, SLIDE.w - 0.2);
  const y = clamp(Number(merged.y ?? defaults.y), 0.55, maxBottom - 0.2);
  const w = clamp(Number(merged.w ?? defaults.w), 0.35, SLIDE.w - x - 0.15);
  const h = clamp(Number(merged.h ?? defaults.h), 0.2, maxBottom - y);
  return { x, y, w, h };
}
function fitTitle(text, max, min) {
  const len = String(text || '').replace(/\s/g, '').length;
  if (len <= 6) return max;
  if (len <= 12) return Math.max(min, max - 8);
  if (len <= 20) return Math.max(min, max - 14);
  return min;
}

function textVisualLength(text) {
  return String(text || '').split('').reduce((sum, ch) => {
    if (/\s/.test(ch)) return sum + 0.25;
    return /[\u2E80-\u9FFF\uF900-\uFAFF]/.test(ch) ? sum + 1 : sum + 0.56;
  }, 0);
}

function estimateTextHeight(text, boxW, fontSize, options = {}) {
  const raw = String(text || '').trim();
  if (!raw) return options.empty ?? 0;
  const charsPerLine = Math.max(4, (boxW * 72) / Math.max(1, fontSize * 0.58));
  const lines = raw.split(/\r?\n/).reduce((sum, line) => sum + Math.max(1, Math.ceil(textVisualLength(line) / charsPerLine)), 0);
  const height = (lines * fontSize * (options.lineHeight || 1.18)) / 72 + (options.padding || 0.04);
  return clamp(height, options.min ?? 0.2, options.max ?? 10);
}

function distributeRowHeights(demands, rows, cols, minH, maxTotal, gapY) {
  const rowHeights = Array.from({ length: rows }, (_, row) => {
    const rowItems = demands.slice(row * cols, row * cols + cols);
    return Math.max(minH, ...rowItems);
  });
  const available = Math.max(minH * rows, maxTotal - gapY * Math.max(0, rows - 1));
  const total = rowHeights.reduce((sum, h) => sum + h, 0);
  if (total <= available) return rowHeights;
  const baseTotal = minH * rows;
  if (available <= baseTotal) return rowHeights.map(() => available / rows);
  const extraTotal = rowHeights.reduce((sum, h) => sum + Math.max(0, h - minH), 0) || 1;
  const extraAvail = available - baseTotal;
  return rowHeights.map((h) => minH + Math.max(0, h - minH) / extraTotal * extraAvail);
}

function clamp(value, min, max) {
  return Math.max(min, Math.min(max, value));
}

function rectCenter(box) {
  return { x: box.x + box.w / 2, y: box.y + box.h / 2 };
}

function rectEdgePoint(box, dx, dy, inset = 0) {
  const c = rectCenter(box);
  if (!dx && !dy) return c;
  const scaleX = dx ? ((box.w / 2) - inset) / Math.abs(dx) : Infinity;
  const scaleY = dy ? ((box.h / 2) - inset) / Math.abs(dy) : Infinity;
  const scale = Math.max(0, Math.min(scaleX, scaleY));
  return { x: c.x + dx * scale, y: c.y + dy * scale };
}

function connectorBetweenRects(fromBox, toBox, gap = 0.05) {
  const from = rectCenter(fromBox);
  const to = rectCenter(toBox);
  const dx = to.x - from.x;
  const dy = to.y - from.y;
  if (!dx && !dy) return null;
  const length = Math.hypot(dx, dy) || 1;
  const ux = dx / length;
  const uy = dy / length;
  const start = rectEdgePoint(fromBox, dx, dy);
  const end = rectEdgePoint(toBox, -dx, -dy);
  return {
    x: start.x + ux * gap,
    y: start.y + uy * gap,
    w: end.x - start.x - ux * gap * 2,
    h: end.y - start.y - uy * gap * 2,
  };
}

function addRadialConnector(slide, fromBox, toBox, line) {
  const connector = connectorBetweenRects(fromBox, toBox);
  if (!connector) return;
  const endX = connector.x + connector.w;
  const endY = connector.y + connector.h;
  const shape = connector.w * connector.h < 0 ? pptx.ShapeType.lineInv : pptx.ShapeType.line;
  slide.addShape(shape, {
    x: Math.min(connector.x, endX),
    y: Math.min(connector.y, endY),
    w: Math.abs(connector.w),
    h: Math.abs(connector.h),
    line,
  });
}

module.exports = {
  buildDeck,
  normalizeSpec,
};
