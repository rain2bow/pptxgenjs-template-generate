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
const { warnSpecTextCapacity } = require('./text-capacity');

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
const LUCIDE_ICON_CACHE = new Map();
const ICON_PNG_CACHE = new Map();
const SVG_IMAGE_PNG_CACHE = new Map();
const IMAGE_ASPECT_CACHE = new Map();
let LUCIDE_MODULE = undefined;
let SHARP_MODULE = undefined;
let ICON_RENDER_MODE = 'png';
let SVG_IMAGE_RENDER_MODE = 'png';

function normalizeSpec(spec, options = {}) {
  spec.style = spec.style || 'magazine';
  if (!isSupportedStyle(spec.style)) fail(`Unsupported style: ${spec.style}`);
  spec.theme = spec.theme || defaultThemeForStyle(spec.style);
  if (!THEMES[spec.style][spec.theme]) fail(`Unsupported theme "${spec.theme}" for style "${spec.style}"`);
  if (!Array.isArray(spec.slides) || spec.slides.length === 0) fail('Spec must include a non-empty slides array.');
  if (spec.diversifyLayouts === true) {
    fail('diversifyLayouts is no longer supported. Edit each slide.layout in JSON manually; the generator never changes layout automatically.');
  }
  normalizeLayoutCompatibility(spec);
  validateSpecSlots(spec, { specDir: options.specDir || process.cwd() });
  warnSpecTextCapacity(spec);
  warnLayoutVariety(spec);
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
  await prepareSvgImageAssets(spec, specDir);
  spec.slides.forEach((slideSpec, index) => {
    const slide = pptx.addSlide();
    enforceReadableSlideText(slide);
    const ctx = { spec, slideSpec, theme, specDir, index, total: spec.slides.length };
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
  const renderers = {
    magazine: renderMagazine,
    swiss: renderSwiss,
    cmb: renderCmb,
  };
  const renderer = renderers[style];
  if (!renderer) fail(`Unsupported style renderer: ${style}`);
  renderer(slide, ctx);
}

function renderMagazine(slide, ctx) {
  const layout = ctx.slideSpec.layout || 'textImage';
  const dark = ['cover', 'section', 'statement', 'bigQuote', 'closing'].includes(layout) && ctx.slideSpec.tone !== 'light';
  const bg = dark ? ctx.theme.ink : ctx.theme.paper;
  const fg = dark ? ctx.theme.paper : ctx.theme.ink;
  if (ctx.slideSpec.headY == null) ctx.slideSpec.headY = Number(ctx.spec.headY) || 1.05;
  addDecorativeBackground(slide, ctx, 'magazine', bg, dark);
  addChrome(slide, ctx, fg, 'magazine');

  const renderers = {
    cover: magazineCover,
    section: magazineSection,
    bigNumbers: magazineBigNumbers,
    quoteImage: magazineQuoteImage,
    imageGrid: magazineImageGrid,
    media: magazineMedia,
    mediaGrid: magazineMediaGrid,
    gallery: magazineMediaGrid,
    pipeline: magazinePipeline,
    bigQuote: magazineBigQuote,
    compare: magazineCompare,
    splitCompare: magazineCompare,
    textImage: magazineTextImage,
    article: magazineArticle,
    briefing: magazineArticle,
    executiveBrief: magazineArticle,
    contentBrief: magazineArticle,
    sectionList: magazineSectionList,
    statement: magazineStatementCompat,
    kpiTower: magazineBigNumbers,
    duoCompare: magazineCompare,
    timeline: magazinePipeline,
    matrix: magazineMatrixCompat,
    fourCards: magazineMatrixCompat,
    imageHero: magazineImageHeroCompat,
    textGrid: magazineMatrixCompat,
    textWeave: magazineMatrixCompat,
    contentSynthesis: magazineMatrixCompat,
    denseText: magazineMatrixCompat,
    dataSheet: magazineDataSheet,
    chart: magazineChart,
    dashboard: magazineDashboard,
    agenda: magazineAgenda,
    caseStudy: magazineCaseStudy,
    pyramid: magazinePyramid,
    radial: magazineRadial,
    roadmap: magazineRoadmap,
    swimlane: magazineSwimlane,
    closing: magazineClosing,
  };
  const state = { dark, bg, fg };
  (renderers[layout] || magazineTextImage)(slide, ctx, state);
  renderDataBlocks(slide, ctx, state, 'magazine');
}

function renderSwiss(slide, ctx) {
  const layout = ctx.slideSpec.layout || 'statement';
  const accentLayouts = ['cover', 'section', 'closing'];
  const dark = ctx.slideSpec.tone === 'dark';
  const accent = accentLayouts.includes(layout) && ctx.slideSpec.tone !== 'light';
  const bg = accent ? ctx.theme.accent : dark ? ctx.theme.ink : ctx.theme.paper;
  const fg = accent ? ctx.theme.accentOn : dark ? ctx.theme.paper : ctx.theme.ink;
  if (ctx.slideSpec.headY == null) ctx.slideSpec.headY = Number(ctx.spec.headY) || 1.12;
  addDecorativeBackground(slide, ctx, 'swiss', bg, accent || dark);
  addChrome(slide, ctx, fg, 'swiss');

  const renderers = {
    cover: swissCover,
    section: swissSectionCompat,
    bigNumbers: swissKpiTower,
    quoteImage: swissQuoteImageCompat,
    imageGrid: swissImageGridCompat,
    media: swissMedia,
    mediaGrid: swissMediaGrid,
    gallery: swissMediaGrid,
    pipeline: swissTimeline,
    bigQuote: swissBigQuoteCompat,
    compare: swissDuoCompare,
    splitCompare: swissDuoCompare,
    textImage: swissTextImageCompat,
    article: swissTextGrid,
    briefing: swissSectionList,
    executiveBrief: swissSectionList,
    contentBrief: swissSectionList,
    sectionList: swissSectionList,
    statement: swissStatement,
    kpiTower: swissKpiTower,
    duoCompare: swissDuoCompare,
    timeline: swissTimeline,
    matrix: swissMatrix,
    fourCards: swissFourCards,
    imageHero: swissImageHero,
    textGrid: swissTextGrid,
    textWeave: swissTextGrid,
    contentSynthesis: swissTextGrid,
    denseText: swissTextGrid,
    dataSheet: swissDataSheet,
    chart: swissChart,
    dashboard: swissDashboard,
    agenda: swissAgenda,
    caseStudy: swissCaseStudy,
    pyramid: swissPyramid,
    radial: swissRadial,
    roadmap: swissRoadmap,
    swimlane: swissSwimlane,
    closing: swissClosing,
  };
  const state = { accent, dark, bg, fg };
  (renderers[layout] || swissStatement)(slide, ctx, state);
  renderDataBlocks(slide, ctx, state, 'swiss');
}

function renderCmb(slide, ctx) {
  const data = ctx.slideSpec;
  const layout = data.layout || 'statement';
  const accentLayouts = ['cover', 'section', 'closing'];
  const accent = accentLayouts.includes(layout) && data.tone !== 'light';
  const bg = accent ? ctx.theme.accent : ctx.theme.paper;
  const fg = accent ? ctx.theme.accentOn : ctx.theme.ink;
  if (data.headY == null) data.headY = Number(ctx.spec.headY) || 1.06;
  addCmbBackground(slide, ctx, accent);
  addCmbChrome(slide, ctx, fg);

  const state = { accent, dark: accent, bg, fg };
  const renderers = {
    cover: cmbCover,
    section: cmbSection,
    statement: cmbStatement,
    closing: cmbClosing,
    dashboard: swissDashboard,
    dataSheet: swissDataSheet,
    chart: swissChart,
    kpiTower: swissKpiTower,
    bigNumbers: swissKpiTower,
    media: swissMedia,
    mediaGrid: swissMediaGrid,
    gallery: swissMediaGrid,
    imageGrid: swissImageGridCompat,
    compare: swissDuoCompare,
    duoCompare: swissDuoCompare,
    splitCompare: swissDuoCompare,
    timeline: swissTimeline,
    pipeline: swissTimeline,
    roadmap: swissRoadmap,
    textGrid: cmbTextWeave,
    textWeave: cmbTextWeave,
    contentSynthesis: cmbTextWeave,
    denseText: cmbTextWeave,
    article: cmbBriefing,
    sectionList: cmbBriefing,
    briefing: cmbBriefing,
    executiveBrief: cmbBriefing,
    contentBrief: cmbBriefing,
    fourCards: cmbTextWeave,
    matrix: swissMatrix,
    agenda: cmbAgenda,
    caseStudy: swissCaseStudy,
    pyramid: swissPyramid,
    radial: swissRadial,
    swimlane: swissSwimlane,
    imageHero: swissImageHero,
    quoteImage: swissQuoteImageCompat,
    bigQuote: swissBigQuoteCompat,
    textImage: swissTextImageCompat,
  };
  (renderers[layout] || cmbStatement)(slide, ctx, state);
  renderDataBlocks(slide, ctx, state, 'swiss');
}

function addCmbBackground(slide, ctx, emphasized = false) {
  const theme = ctx.theme;
  const svg = cmbBackgroundSvg(ctx, emphasized);
  slide.background = { path: 'background-cmb.svg', data: svgDataUri(svg) };
  const headerH = Number(ctx.slideSpec.logoHeaderBandH || ctx.spec.logoHeaderBandH) || 0.82;
  slide.addShape(pptx.ShapeType.rect, { x: 0, y: 0, w: SLIDE.w, h: headerH, fill: { color: theme.headerBg || 'FFFFFF', transparency: 0 }, line: { color: theme.headerBg || 'FFFFFF', transparency: 100 } });
  slide.addShape(pptx.ShapeType.line, { x: 0, y: headerH, w: SLIDE.w, h: 0, line: { color: theme.headerLine || theme.accent, transparency: 7, width: 1.2 } });
  if (!emphasized) addCmbLogoWatermark(slide, ctx);
}

function cmbBackgroundSvg(ctx, emphasized = false) {
  const scale = 120;
  const w = Math.round(SLIDE.w * scale);
  const h = Math.round(SLIDE.h * scale);
  const headerH = Math.round((Number(ctx.slideSpec.logoHeaderBandH || ctx.spec.logoHeaderBandH) || 0.82) * scale);
  const theme = ctx.theme;
  const base = normalizeHex(emphasized ? theme.accent : theme.paper);
  const tint = normalizeHex(emphasized ? theme.accent2 : theme.paperTint || theme.grey1 || 'F8F1F1');
  const accent = normalizeHex(theme.accent || 'C8102E');
  const accent2 = normalizeHex(theme.accent2 || '8A1538');
  const grid = emphasized ? 'FFFFFF' : normalizeHex(theme.grey2 || 'DED8D6');
  const gridOpacity = emphasized ? 0.08 : 0.18;
  if (emphasized && theme.emphasisSolid) {
    return `<svg xmlns="http://www.w3.org/2000/svg" width="${w}" height="${h}" viewBox="0 0 ${w} ${h}">
<rect x="0" y="0" width="${w}" height="${h}" fill="#${svgEsc(base)}"/>
<rect x="0" y="0" width="${w}" height="${headerH}" fill="#${svgEsc(normalizeHex(theme.headerBg || 'FFFFFF'))}"/>
</svg>`;
  }
  const lines = [];
  const gridLeft = Math.round(0.78 * scale);
  const gridRight = Math.round((SLIDE.w - 0.78) * scale);
  for (let x = gridLeft; x <= gridRight + 1; x += 0.78 * scale) lines.push(`<line x1="${Math.round(x)}" y1="${headerH}" x2="${Math.round(x)}" y2="${h}"/>`);
  for (let y = headerH; y <= h + 1; y += 0.54 * scale) lines.push(`<line x1="${gridLeft}" y1="${Math.round(y)}" x2="${gridRight}" y2="${Math.round(y)}"/>`);
  return `<svg xmlns="http://www.w3.org/2000/svg" width="${w}" height="${h}" viewBox="0 0 ${w} ${h}">
<defs>
  <linearGradient id="body" x1="0" y1="${headerH}" x2="0" y2="${h}" gradientUnits="userSpaceOnUse">
    <stop offset="0%" stop-color="#${svgEsc(base)}"/>
    <stop offset="58%" stop-color="#${svgEsc(tint)}" stop-opacity="${emphasized ? 0.72 : 0.42}"/>
    <stop offset="100%" stop-color="#${svgEsc(accent2)}" stop-opacity="${emphasized ? 0.34 : 0.12}"/>
  </linearGradient>
</defs>
<rect x="0" y="0" width="${w}" height="${h}" fill="#${svgEsc(base)}"/>
<rect x="0" y="0" width="${w}" height="${headerH}" fill="#${svgEsc(normalizeHex(theme.headerBg || 'FFFFFF'))}"/>
<rect x="0" y="${headerH}" width="${w}" height="${h - headerH}" fill="url(#body)"/>
<rect x="0" y="${h - Math.round(0.18 * scale)}" width="${w}" height="${Math.round(0.18 * scale)}" fill="#${svgEsc(accent2)}" opacity="${emphasized ? 0.42 : 0.18}"/>
<rect x="0" y="${headerH}" width="${Math.round(0.16 * scale)}" height="${h - headerH}" fill="#${svgEsc(accent)}" opacity="${emphasized ? 0.9 : 0.78}"/>
<g fill="none" stroke="#${svgEsc(grid)}" stroke-width="0.7" stroke-opacity="${gridOpacity}">${lines.join('\n')}</g>
</svg>`;
}

function addCmbChrome(slide, ctx, color) {
  const headerH = Number(ctx.slideSpec.logoHeaderBandH || ctx.spec.logoHeaderBandH) || 0.82;
  const logoPath = resolveImage(ctx.specDir, ctx.slideSpec.logoHeader || ctx.slideSpec.brandLogoHeader || ctx.spec.logoHeader || ctx.spec.brandLogoHeader || 'logos/cmb-logo-lockup.png');
  const logoW = Number(ctx.slideSpec.logoHeaderW || ctx.spec.logoHeaderW) || 1.58;
  const logoH = Number(ctx.slideSpec.logoHeaderH || ctx.spec.logoHeaderH) || 0.5;
  if (logoPath) addImageAsset(slide, logoPath, { x: SLIDE.marginX, y: 0.16, w: logoW, h: logoH });
  const left = ctx.slideSpec.chromeLeft || ctx.spec.chromeLeft || ctx.spec.title || 'China Merchants Bank';
  const right = ctx.slideSpec.chromeRight || `${String(ctx.index + 1).padStart(2, '0')} / ${String(ctx.total).padStart(2, '0')}`;
  slide.addText(left, { x: SLIDE.marginX + logoW + 0.28, y: 0.34, w: 5.5, h: 0.2, fontFace: FONTS.sans, fontSize: 7.6, bold: true, charSpace: 0.8, color: ctx.theme.headerInk || ctx.theme.ink, transparency: 8, margin: 0, fit: 'shrink' });
  slide.addText(right, { x: SLIDE.w - SLIDE.marginX - 2.2, y: 0.34, w: 2.2, h: 0.2, fontFace: FONTS.sans, fontSize: 7.4, bold: true, charSpace: 0.8, color: ctx.theme.accent, align: 'right', margin: 0, fit: 'shrink' });
  slide.addShape(pptx.ShapeType.rect, { x: 0, y: headerH - 0.035, w: 1.85, h: 0.035, fill: { color: ctx.theme.accent, transparency: 0 }, line: { color: ctx.theme.accent, transparency: 100 } });
}

function cmbCover(slide, ctx, s) {
  const data = ctx.slideSpec;
  const headY = pageHeadY(ctx, 1.08);
  addCmbLogoMark(slide, ctx, { x: 10.62, y: 1.02, w: 1.72, h: 1.72 });
  slide.addText(data.kicker || '招商银行', { x: 0.78, y: headY, w: 6.8, h: 0.24, fontFace: FONTS.sans, fontSize: 8.6, bold: true, charSpace: 1.5, color: s.fg, transparency: 18, margin: 0, fit: 'shrink' });
  slide.addText(data.title || ctx.spec.title, { x: 0.78, y: headY + 0.6, w: 10.2, h: 2.15, fontFace: FONTS.sansZh, fontSize: fitTitle(data.title || ctx.spec.title, 47, 31), bold: true, color: s.fg, margin: 0, fit: 'shrink' , typographyRole: 'coverTitle' });
  slide.addShape(pptx.ShapeType.rect, { x: 0.78, y: headY + 3.15, w: 1.55, h: 0.09, fill: { color: ctx.theme.accent, transparency: 0 }, line: { color: ctx.theme.accent, transparency: 100 } });
  slide.addText(data.subtitle || ctx.spec.subtitle || '', { x: 0.78, y: headY + 3.5, w: 6.7, h: 0.72, fontFace: FONTS.sansZh, fontSize: 15.2, color: s.fg, transparency: 10, margin: 0, fit: 'shrink' });
  addFoot(slide, ctx, s.fg, 'swiss');
}

function cmbSection(slide, ctx, s) {
  cmbCover(slide, ctx, s);
}

function cmbStatement(slide, ctx, s) {
  const data = ctx.slideSpec;
  const headY = pageHeadY(ctx, 1.06);
  slide.addShape(pptx.ShapeType.rect, { x: 0.78, y: headY - 0.02, w: 0.12, h: 0.7, fill: { color: ctx.theme.accent, transparency: 0 }, line: { color: ctx.theme.accent, transparency: 100 } });
  slide.addText(data.kicker || '核心摘要', { x: 1.05, y: headY, w: 5.8, h: 0.24, fontFace: FONTS.sans, fontSize: 8, bold: true, charSpace: 1.1, color: ctx.theme.accent, margin: 0, fit: 'shrink' });
  slide.addText(data.title || '', { x: 1.02, y: headY + 0.48, w: 6.85, h: 1.58, fontFace: FONTS.sansZh, fontSize: fitTitle(data.title || '', 34, 26), bold: true, color: s.fg, margin: 0, fit: 'shrink' });
  if (data.body || data.subtitle) slide.addText(data.body || data.subtitle, { x: 1.05, y: headY + 2.42, w: 6.45, h: 1.1, fontFace: FONTS.sansZh, fontSize: 14.2, color: s.fg, transparency: 14, margin: 0, fit: 'shrink', valign: 'top' });
  addStatementImageSlot(slide, ctx, { x: 8.05, y: headY + 1.72, w: 3.85, h: 3.0 }, ctx.theme.accent, data.imageLabel || '图片占位');
  addFoot(slide, ctx, s.fg, 'swiss');
}

function cmbClosing(slide, ctx, s) {
  cmbCover(slide, { ...ctx, slideSpec: { ...ctx.slideSpec, kicker: ctx.slideSpec.kicker || '谢谢', title: ctx.slideSpec.title || '谢谢观看', subtitle: ctx.slideSpec.subtitle || '招商银行' } }, s);
}

function cmbBriefing(slide, ctx, s) {
  const data = ctx.slideSpec;
  addPageHead(slide, data, s.fg, 'swiss', 0.82);
  const items = cmbTextItems(data).slice(0, 6);
  const hasLead = !!(data.summary || data.body || data.lead);
  const lead = hasLead
    ? { title: data.summaryTitle || data.leadTitle || data.focusTitle || data.kicker || '经营重点', body: data.summary || data.body || data.lead }
    : (items[0] || { title: data.title || '', body: data.subtitle || '' });
  const conclusionText = data.conclusion || data.takeaway || data.footerSummary || data.nextStep || '';
  const restLimit = conclusionText ? 4 : 5;
  const rest = (hasLead ? items : items.slice(1)).slice(0, restLimit);
  const y0 = data.subtitle ? 2.78 : 2.45;
  addCmbTextCard(slide, ctx, lead, { x: 0.78, y: y0, w: 11.45, h: 1.12 }, 0, { lead: true, accent: true, label: '摘要', maxPoints: 3, titleFontSize: 13.2 });

  const midY = y0 + 1.34;
  const conclusionBox = { x: 0.78, y: 5.62, w: 11.45, h: 0.86 };
  const midBottom = conclusionText ? conclusionBox.y - 0.2 : 6.48;
  if (rest.length) {
    if (rest.length <= 4) {
      const gap = 0.28;
      const w = (11.45 - gap * (rest.length - 1)) / rest.length;
      const h = Math.max(1.15, midBottom - midY);
      rest.forEach((item, i) => addCmbTextCard(slide, ctx, item, { x: 0.78 + i * (w + gap), y: midY, w, h }, i + 1, { maxPoints: 3 }));
    } else {
      const gapX = 0.28;
      const gapY = 0.18;
      const cols = 3;
      const rows = Math.ceil(rest.length / cols);
      const w = (11.45 - gapX * (cols - 1)) / cols;
      const h = Math.max(0.95, (midBottom - midY - gapY * (rows - 1)) / rows);
      rest.forEach((item, i) => addCmbTextCard(slide, ctx, item, { x: 0.78 + (i % cols) * (w + gapX), y: midY + Math.floor(i / cols) * (h + gapY), w, h }, i + 1, { maxPoints: 3 }));
    }
  }

  if (conclusionText) {
    addCmbTextCard(slide, ctx, { title: data.conclusionTitle || data.takeawayTitle || data.footerSummaryTitle || data.nextStepTitle || '结论', body: conclusionText }, conclusionBox, 6, { compact: true, label: '结论', accent: true, maxPoints: 2, titleFontSize: 13.2 });
  }
  addFoot(slide, ctx, s.fg, 'swiss');
}

function cmbAgenda(slide, ctx, s) {
  const data = ctx.slideSpec;
  addPageHead(slide, data, s.fg, 'swiss', 0.82);
  const items = normalizeSections(data.items || data.agenda || data.sections || []).slice(0, 8)
    .map((item) => (item && typeof item === 'object') ? item : { title: String(item || '') })
    .filter((item) => cmbItemTitle(item) || cmbItemBody(item));
  const y0 = data.subtitle ? 2.78 : 2.42;
  const bottom = 6.48;
  const leftW = 2.15;
  const gap = 0.28;
  const listX = 0.78 + leftW + gap;
  const listW = 11.45 - leftW - gap;
  const count = Math.max(1, items.length);
  const cols = count > 4 ? 2 : 1;
  const rows = Math.ceil(count / cols);
  const colGap = cols > 1 ? 0.24 : 0;
  const rowGap = cols > 1 ? 0.16 : 0.16;
  const colW = (listW - colGap * (cols - 1)) / cols;
  const rowH = (bottom - y0 - rowGap * Math.max(0, rows - 1)) / rows;

  slide.addShape(pptx.ShapeType.rect, { x: 0.78, y: y0, w: leftW, h: bottom - y0, fill: { color: ctx.theme.accent, transparency: 0 }, line: { color: ctx.theme.accent, transparency: 100 } });
  slide.addText(data.agendaTitle || data.indexTitle || '目录', { x: 1.05, y: y0 + 0.28, w: leftW - 0.54, h: 0.5, fontFace: FONTS.sansZh, fontSize: 21, bold: true, color: ctx.theme.accentOn, margin: 0, fit: 'shrink' });
  slide.addText(data.agendaSubtitle || data.indexSubtitle || data.kicker || '', { x: 1.05, y: y0 + 0.98, w: leftW - 0.54, h: 1.28, fontFace: FONTS.sansZh, fontSize: 12, color: ctx.theme.accentOn, transparency: 12, margin: 0.02, fit: 'shrink', valign: 'top' });

  items.forEach((item, i) => {
    const col = i % cols;
    const row = Math.floor(i / cols);
    const x = listX + col * (colW + colGap);
    const y = y0 + row * (rowH + rowGap);
    const title = cmbItemTitle(item) || `章节${i + 1}`;
    const body = cmbItemBody(item);
    const hot = i === Number(data.highlightIndex || -1);
    slide.addShape(pptx.ShapeType.roundRect, { x, y, w: colW, h: rowH, rectRadius: 0.06, fill: { color: hot ? ctx.theme.accent : ctx.theme.paper, transparency: hot ? 0 : 2 }, line: { color: hot ? ctx.theme.accent : ctx.theme.grey2, transparency: hot ? 100 : 18, width: 0.7 } });
    slide.addText(String(i + 1).padStart(2, '0'), { x: x + 0.22, y: y + 0.13, w: 0.52, h: 0.28, fontFace: FONTS.sans, fontSize: 14.8, bold: true, color: hot ? ctx.theme.accentOn : ctx.theme.accent, margin: 0, fit: 'shrink' });
    if (cols > 1) {
      slide.addText(title, { x: x + 0.86, y: y + 0.12, w: colW - 1.08, h: 0.28, fontFace: FONTS.sansZh, fontSize: 15.0, bold: true, color: hot ? ctx.theme.accentOn : ctx.theme.ink, margin: 0, fit: 'shrink', valign: 'mid' });
      if (body) slide.addText(body, { x: x + 0.86, y: y + 0.42, w: colW - 1.08, h: Math.max(0.24, rowH - 0.5), fontFace: FONTS.sansZh, fontSize: READABILITY.minFontSize, color: hot ? ctx.theme.accentOn : ctx.theme.ink, transparency: hot ? 8 : 24, margin: 0.02, fit: 'shrink', valign: 'top' });
    } else {
      slide.addText(title, { x: x + 1.0, y: y + 0.13, w: body ? 3.1 : colW - 1.28, h: Math.min(0.38, rowH - 0.16), fontFace: FONTS.sansZh, fontSize: 15.5, bold: true, color: hot ? ctx.theme.accentOn : ctx.theme.ink, margin: 0, fit: 'shrink', valign: 'mid' });
      if (body) {
        slide.addShape(pptx.ShapeType.line, { x: x + 4.3, y: y + 0.18, w: 0, h: Math.max(0.18, rowH - 0.36), line: { color: hot ? ctx.theme.accentOn : ctx.theme.grey2, transparency: hot ? 55 : 20, width: 0.6 } });
        slide.addText(body, { x: x + 4.55, y: y + 0.12, w: colW - 4.84, h: Math.max(0.26, rowH - 0.2), fontFace: FONTS.sansZh, fontSize: READABILITY.minFontSize, color: hot ? ctx.theme.accentOn : ctx.theme.ink, transparency: hot ? 8 : 24, margin: 0.02, fit: 'shrink', valign: 'mid' });
      }
    }
  });
  addFoot(slide, ctx, s.fg, 'swiss');
}

function cmbTextWeave(slide, ctx, s) {
  const data = ctx.slideSpec;
  addPageHead(slide, data, s.fg, 'swiss', 0.82);
  const items = cmbTextItems(data).slice(0, 6);
  const y0 = data.subtitle ? 2.78 : 2.42;
  const bottom = 6.48;
  if (!items.length) {
    if (data.body) addCmbTextCard(slide, ctx, { title: data.title || '', body: data.body }, { x: 0.78, y: y0, w: 11.45, h: bottom - y0 }, 0, { lead: true, accent: true, maxPoints: 5 });
    addFoot(slide, ctx, s.fg, 'swiss');
    return;
  }

  if (items.length <= 5) {
    const leadW = 3.85;
    const gapX = 0.28;
    addCmbTextCard(slide, ctx, items[0], { x: 0.78, y: y0, w: leadW, h: bottom - y0 }, 0, { lead: true, accent: true, maxPoints: 5 });
    const right = items.slice(1);
    const rightX = 0.78 + leadW + gapX;
    const rightW = 11.45 - leadW - gapX;
    const rowGap = 0.24;
    const rightBoxes = cmbTextWeaveRightBoxes(right.length, rightX, y0, rightW, bottom - y0, gapX, rowGap);
    right.forEach((item, i) => addCmbTextCard(slide, ctx, item, rightBoxes[i], i + 1, { maxPoints: 3 }));
  } else {
    addCmbTextCard(slide, ctx, items[0], { x: 0.78, y: y0, w: 11.45, h: 0.9 }, 0, { lead: true, accent: true, compact: true, maxPoints: 2 });
    const rest = items.slice(1);
    const gridY = y0 + 1.12;
    const leftW = 3.72;
    const gapX = 0.28;
    addCmbTextCard(slide, ctx, rest[0], { x: 0.78, y: gridY, w: leftW, h: bottom - gridY }, 1, { maxPoints: 4 });
    const cardW = (11.45 - leftW - gapX * 2) / 2;
    const rowGap = 0.2;
    const cardH = (bottom - gridY - rowGap) / 2;
    rest.slice(1).forEach((item, i) => addCmbTextCard(slide, ctx, item, { x: 0.78 + leftW + gapX + (i % 2) * (cardW + gapX), y: gridY + Math.floor(i / 2) * (cardH + rowGap), w: cardW, h: cardH }, i + 2, { maxPoints: 3 }));
  }
  addFoot(slide, ctx, s.fg, 'swiss');
}

function cmbTextWeaveRightBoxes(count, x, y, w, h, gapX, rowGap) {
  const halfW = (w - gapX) / 2;
  const halfH = (h - rowGap) / 2;
  if (count <= 0) return [];
  if (count === 1) return [{ x, y, w, h }];
  if (count === 2) return [
    { x, y, w: halfW, h },
    { x: x + halfW + gapX, y, w: halfW, h },
  ];
  if (count === 3) return [
    { x, y, w: halfW, h },
    { x: x + halfW + gapX, y, w: halfW, h: halfH },
    { x: x + halfW + gapX, y: y + halfH + rowGap, w: halfW, h: halfH },
  ];
  return Array.from({ length: Math.min(count, 4) }, (_, i) => ({
    x: x + (i % 2) * (halfW + gapX),
    y: y + Math.floor(i / 2) * (halfH + rowGap),
    w: halfW,
    h: halfH,
  }));
}

function cmbTextItems(data) {
  return normalizeSections(data.sections || data.items || data.columns || data.points || data.agenda || [])
    .map((item) => (item && typeof item === 'object') ? item : { body: String(item || '') })
    .filter((item) => cmbItemTitle(item) || cmbItemBody(item) || cmbItemPoints(item).length);
}

function cmbItemTitle(item) {
  return String(item?.title || item?.label || item?.name || item?.heading || '').trim();
}

function cmbItemBody(item) {
  return String(item?.body || item?.desc || item?.note || item?.summary || item?.detail || item?.text || item?.story || '').trim();
}

function cmbItemPoints(item) {
  const raw = item?.points || item?.bullets || item?.list;
  if (!Array.isArray(raw)) return [];
  return raw.map((point) => {
    if (point && typeof point === 'object') return String(point.body || point.text || point.title || point.label || '').trim();
    return String(point || '').trim();
  }).filter(Boolean);
}

function addCmbTextCard(slide, ctx, item, box, index, options = {}) {
  const accent = ctx.theme.accent;
  const ink = ctx.theme.ink;
  const hot = !!options.accent;
  const fill = hot ? accent : ctx.theme.paper;
  const line = hot ? accent : ctx.theme.grey2;
  const color = hot ? ctx.theme.accentOn : ink;
  slide.addShape(pptx.ShapeType.rect, { x: box.x, y: box.y, w: box.w, h: box.h, fill: { color: fill, transparency: hot ? 0 : 4 }, line: { color: line, transparency: hot ? 100 : 18, width: 0.7 } });
  if (!hot) slide.addShape(pptx.ShapeType.rect, { x: box.x, y: box.y, w: 0.08, h: box.h, fill: { color: accent, transparency: 8 }, line: { color: accent, transparency: 100 } });

  const padX = options.compact ? 0.2 : 0.26;
  const padTop = options.compact ? 0.12 : 0.18;
  const title = cmbItemTitle(item);
  const body = cmbItemBody(item) || (!title ? String(item?.value || '') : '');
  const contentX = box.x + padX;
  const contentW = box.w - padX * 2;
  const headerY = box.y + padTop;
  const iconSize = options.compact ? 0.2 : 0.24;
  const iconColor = hot ? ctx.theme.accentOn : accent;
  const hasIcon = addInlineIcon(slide, item || {}, contentX, headerY + 0.02, iconSize, iconColor, 'swiss', { fallback: defaultContentIcon(index, 'swiss'), pad: 0.04 });
  const fallbackLabel = options.label || String(index + 1).padStart(2, '0');
  if (!hasIcon) slide.addText(fallbackLabel, { x: contentX, y: headerY + 0.02, w: 0.34, h: 0.2, fontFace: FONTS.mono, fontSize: 7.2, bold: true, color: iconColor, transparency: hot ? 4 : 18, margin: 0, fit: 'shrink' });
  const titleX = contentX + 0.38;
  const titleW = Math.max(0.3, contentW - 0.38);
  const titleText = title || fallbackLabel;
  const titleFont = Number(options.titleFontSize) || (options.lead ? 15.2 : options.compact ? 12.2 : 13.2);
  const titleH = estimateTextHeight(titleText, titleW, titleFont, { min: 0.26, max: options.compact ? 0.34 : 0.5, lineHeight: 1.14, padding: 0.02 });
  slide.addText(titleText, { x: titleX, y: headerY, w: titleW, h: titleH, fontFace: FONTS.sansZh, fontSize: titleFont, bold: true, color, margin: 0, valign: 'top' });

  const points = cmbItemPoints(item);
  if (!body && !points.length) return;
  const bodyY = headerY + titleH + 0.13;
  const bodyH = Math.max(0.24, box.y + box.h - bodyY - (options.compact ? 0.12 : 0.18));
  const fontSize = READABILITY.minFontSize;
  if (points.length > 1) {
    addBulletedCardBody(slide, points, { x: contentX, y: bodyY, w: contentW, h: bodyH }, { color, transparency: hot ? 4 : 22, fontSize });
  } else {
    slide.addText(points[0] || body, { x: contentX, y: bodyY, w: contentW, h: bodyH, fontFace: FONTS.sansZh, fontSize, color, transparency: hot ? 6 : 24, margin: 0.02, valign: 'top' });
  }
}

function addBulletedCardBody(slide, points, box, options = {}) {
  if (!Array.isArray(points) || points.length <= 1) return;
  const fontSize = options.fontSize || READABILITY.minFontSize;
  const runs = points.map((point, i) => ({
    text: String(point || '').trim(),
    options: {
      bullet: true,
      breakLine: i < points.length - 1,
      fontFace: FONTS.sansZh,
      fontSize,
      color: options.color || '111111',
      transparency: options.transparency ?? 24,
    },
  })).filter((run) => run.text);
  if (!runs.length) return;
  slide.addText(runs, {
    x: box.x,
    y: box.y,
    w: box.w,
    h: box.h,
    fontFace: FONTS.sansZh,
    fontSize,
    color: options.color || '111111',
    transparency: options.transparency ?? 24,
    margin: 0.01,
    valign: 'top',
  });
}
function estimateBulletedLineCount(text, w, fontSize) {
  const effectiveW = Math.max(0.05, Number(w || 0) - 0.38);
  const charsPerLine = Math.max(1, (effectiveW * 72) / Math.max(1, fontSize));
  return String(text || '').split(/\r?\n/).reduce((sum, line) => sum + Math.max(1, Math.ceil(textVisualLength(line) / charsPerLine)), 0);
}
function addCmbLogoMark(slide, ctx, box) {
  const logoPath = resolveCmbLogoMark(ctx);
  if (logoPath) addImageAsset(slide, logoPath, box);
}

function addCmbLogoWatermark(slide, ctx) {
  const logoPath = resolveCmbLogoMark(ctx);
  if (!logoPath) return;
  addImageAsset(slide, logoPath, { x: 10.62, y: 1.02, w: 1.72, h: 1.72 }, { opacity: 0.2 });
}

function resolveCmbLogoMark(ctx) {
  return resolveImage(ctx.specDir, ctx.slideSpec.logoMark || ctx.slideSpec.logoSymbol || ctx.slideSpec.brandLogoSymbol || ctx.spec.logoMark || ctx.spec.logoSymbol || ctx.spec.brandLogoSymbol || 'logos/cmb-logo-mark.svg');
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

function readSvgWithOpacity(imagePath, opacity) {
  let svg = fs.readFileSync(imagePath, 'utf8');
  if (opacity == null) return svg;
  const value = Math.max(0, Math.min(1, Number(opacity)));
  svg = svg.replace(/<\?xml[^>]*>\s*/i, '').replace(/<!DOCTYPE[^>]*>\s*/i, '').trim();
  return svg.replace(/<svg\b/i, `<svg opacity="${value}"`);
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

function magazineCover(slide, ctx, s) {
  const data = ctx.slideSpec;
  const headY = pageHeadY(ctx, 1.05);
  slide.addText(data.kicker || 'A Talk', {
    x: 0.75,
    y: headY,
    w: 8,
    h: 0.25,
    fontFace: FONTS.mono,
    fontSize: 9,
    charSpace: 2.2,
    color: s.fg,
    transparency: 22,
    margin: 0,
  });
  slide.addText(data.title || ctx.spec.title, {
    x: 0.72,
    y: headY + 0.47,
    w: 11.2,
    h: 2.2,
    fontFace: FONTS.serifZh,
    fontSize: fitTitle(data.title || ctx.spec.title, 54, 38),
    bold: true,
    typographyRole: 'coverTitle',
    color: s.fg,
    margin: 0,
    breakLine: false,
    fit: 'shrink',
  });
  slide.addText(data.subtitle || ctx.spec.subtitle || '', {
    x: 0.78,
    y: headY + 2.8,
    w: 7.2,
    h: 0.85,
    fontFace: FONTS.serifZh,
    fontSize: 21,
    color: s.fg,
    transparency: 18,
    margin: 0,
    fit: 'shrink',
  });
  slide.addShape(pptx.ShapeType.line, {
    x: 0.78,
    y: headY + 4.0,
    w: 3.2,
    h: 0,
    line: { color: s.fg, transparency: 45, width: 1 },
  });
  slide.addText(data.author || ctx.spec.author || '', {
    x: 0.78,
    y: headY + 4.23,
    w: 5.5,
    h: 0.3,
    fontFace: FONTS.mono,
    fontSize: 8.5,
    charSpace: 1.3,
    color: s.fg,
    transparency: 28,
    margin: 0,
  });
  addFoot(slide, ctx, s.fg, 'magazine');
}
function magazineSection(slide, ctx, s) {
  const data = ctx.slideSpec;
  slide.addText(data.kicker || `Act ${ctx.index + 1}`, {
    x: 0.78,
    y: 1.55,
    w: 3,
    h: 0.3,
    fontFace: FONTS.mono,
    fontSize: 9,
    charSpace: 2.2,
    color: s.fg,
    transparency: 28,
    margin: 0,
  });
  slide.addText(data.title, {
    x: 0.72,
    y: 2.15,
    w: 10.2,
    h: 1.55,
    fontFace: FONTS.serifZh,
    fontSize: fitTitle(data.title, 50, 35),
    bold: true,
    typographyRole: 'coverTitle',
    color: s.fg,
    margin: 0,
    fit: 'shrink',
  });
  slide.addText(data.subtitle || data.body || '', {
    x: 0.8,
    y: 4.08,
    w: 6.5,
    h: 0.85,
    fontFace: FONTS.serifZh,
    fontSize: 19,
    color: s.fg,
    transparency: 18,
    margin: 0,
    fit: 'shrink',
  });
  addFoot(slide, ctx, s.fg, 'magazine');
}

function magazineBigNumbers(slide, ctx, s) {
  const data = ctx.slideSpec;
  addPageHead(slide, data, s.fg, 'magazine');
  const items = (data.items || []).slice(0, 6);
  const cols = clampColumns(data.columnsCount || autoColumns(items.length, 3), 1, 3);
  const gap = 0.45;
  const cellW = (11.75 - gap * (cols - 1)) / cols;
  const cellH = 1.35;
  const startX = 0.78;
  const startY = items.length <= cols ? 3.0 : 2.7;
  items.forEach((item, i) => {
    const x = startX + (i % cols) * (cellW + gap);
    const y = startY + Math.floor(i / cols) * 1.75;
    slide.addShape(pptx.ShapeType.line, { x, y, w: cellW, h: 0, line: { color: s.fg, transparency: 55, width: 0.7 } });
    slide.addText(item.label || '', {
      x,
      y: y + 0.15,
      w: cellW,
      h: 0.18,
      fontFace: FONTS.mono,
      fontSize: 7.5,
      charSpace: 1.8,
      color: s.fg,
      transparency: 35,
      margin: 0,
      fit: 'shrink',
    });
    slide.addText([{ text: item.value || '', options: {} }, { text: item.unit ? ` ${item.unit}` : '', options: { fontFace: FONTS.serifZh, fontSize: 18, bold: false } }], {
      x,
      y: y + 0.38,
      w: cellW,
      h: 0.55,
      fontFace: FONTS.serifEn,
      fontSize: 35,
      bold: true,
      color: s.fg,
      margin: 0,
      fit: 'shrink',
    });
    slide.addText(item.note || '', {
      x,
      y: y + 1.02,
      w: cellW,
      h: 0.42,
      fontFace: FONTS.sansZh,
      fontSize: 10.5,
      color: s.fg,
      transparency: 22,
      margin: 0,
      fit: 'shrink',
      valign: 'top',
    });
  });
  addFoot(slide, ctx, s.fg, 'magazine');
}

function magazineQuoteImage(slide, ctx, s) {
  const data = ctx.slideSpec;
  const x1 = 0.78;
  addPageHead(slide, data, s.fg, 'magazine', 0.92);
  slide.addText(data.body || data.quote || '', {
    x: x1,
    y: 2.55,
    w: 5.9,
    h: 1.05,
    fontFace: FONTS.serifZh,
    fontSize: 19,
    color: s.fg,
    transparency: 12,
    margin: 0.02,
    fit: 'shrink',
    valign: 'mid',
  });
  addCallout(slide, data.callout || data.quote, x1, 4.22, 5.9, 1.25, s.fg, ctx.theme.paperTint);
  addMediaOrChart(slide, ctx, data, { x: 7.3, y: 1.65, w: 5.1, h: 3.8 }, s, 'magazine', '16:10');
  addCaption(slide, data.caption || data.image?.caption, 7.3, 5.58, 5.1, s.fg, 'magazine');
  addFoot(slide, ctx, s.fg, 'magazine');
}

function magazineImageGrid(slide, ctx, s) {
  magazineMediaGrid(slide, ctx, s);
}

function magazineMedia(slide, ctx, s) {
  const data = ctx.slideSpec;
  addPageHead(slide, data, s.fg, 'magazine', 0.88);
  const mediaBox = { x: 6.95, y: 2.2, w: 5.05, h: 3.6 };
  slide.addText(data.body || data.story || data.note || '', { x: 0.78, y: 2.55, w: 5.35, h: 1.18, fontFace: FONTS.sansZh, fontSize: 13.2, color: s.fg, transparency: 16, margin: 0.03, fit: 'shrink', valign: 'top' });
  const items = normalizeSections(data.items || data.insights || data.points || []).slice(0, 3);
  items.forEach((item, i) => {
    const y = 4.0 + i * 0.58;
    const hasIcon = addInlineIcon(slide, item, 0.82, y + 0.01, 0.22, s.fg, 'magazine', { fallback: defaultContentIcon(i, 'magazine'), pad: 0.04 });
    slide.addText(item.title || item.label || item.body || '', { x: hasIcon ? 1.14 : 0.82, y, w: hasIcon ? 4.98 : 5.3, h: 0.24, fontFace: FONTS.serifZh, fontSize: 11.8, bold: true, color: s.fg, margin: 0, fit: 'shrink' });
    slide.addText(item.body || item.desc || item.note || '', { x: hasIcon ? 1.14 : 0.82, y: y + 0.28, w: hasIcon ? 4.98 : 5.3, h: 0.22, fontFace: FONTS.sansZh, fontSize: 8.8, color: s.fg, transparency: 25, margin: 0, fit: 'shrink' });
  });
  addMediaOrChart(slide, ctx, data, mediaBox, s, 'magazine', 'MEDIA');
  addCaption(slide, data.caption || data.image?.caption || data.chart?.caption, mediaBox.x, mediaBox.y + mediaBox.h + 0.1, mediaBox.w, s.fg, 'magazine');
  addFoot(slide, ctx, s.fg, 'magazine');
}

function magazineMediaGrid(slide, ctx, s) {
  const data = ctx.slideSpec;
  addPageHead(slide, data, s.fg, 'magazine', 0.9);
  const count = resolveMediaSlotCount(data);
  const cols = count <= 2 ? count : 3;
  const rows = Math.ceil(count / cols);
  const gapX = 0.38;
  const gapY = rows > 1 ? 0.55 : 0.3;
  const gridW = 11.15;
  const w = (gridW - gapX * (cols - 1)) / cols;
  const h = rows > 1 ? 1.55 : 2.45;
  const startX = 0.78;
  const startY = rows > 1 ? 2.55 : 3.0;
  const boxes = Array.from({ length: count }, (_, i) => ({ x: startX + (i % cols) * (w + gapX), y: startY + Math.floor(i / cols) * (h + gapY), w, h }));
  addMediaGrid(slide, ctx, data, boxes, s, 'magazine');
  addFoot(slide, ctx, s.fg, 'magazine');
}

function magazinePipeline(slide, ctx, s) {
  const data = ctx.slideSpec;
  addPageHead(slide, data, s.fg, 'magazine');
  const items = (data.steps || data.items || []).slice(0, 6);
  const cols = items.length <= 4 ? items.length : 3;
  const cellW = cols === 3 ? 3.75 : cols === 4 ? 2.55 : 2.75;
  const startX = 0.78;
  const startY = 3.0;
  items.forEach((item, i) => {
    const x = startX + (i % cols) * (cellW + 0.45);
    const y = startY + Math.floor(i / cols) * 1.55;
    slide.addShape(pptx.ShapeType.line, { x, y, w: cellW, h: 0, line: { color: s.fg, transparency: 40, width: 1 } });
    slide.addText(item.number || String(i + 1).padStart(2, '0'), { x, y: y + 0.14, w: 0.6, h: 0.25, fontFace: FONTS.serifEn, italic: true, fontSize: 13, color: s.fg, transparency: 45, margin: 0 });
    slide.addText(item.title || item.label || '', { x, y: y + 0.48, w: cellW, h: 0.3, fontFace: FONTS.sansZh, bold: true, fontSize: 14, color: s.fg, margin: 0, fit: 'shrink' });
    slide.addText(item.desc || item.note || '', { x, y: y + 0.85, w: cellW, h: 0.48, fontFace: FONTS.sansZh, fontSize: 10.5, color: s.fg, transparency: 25, margin: 0, fit: 'shrink' });
  });
  addFoot(slide, ctx, s.fg, 'magazine');
}

function magazineBigQuote(slide, ctx, s) {
  const data = ctx.slideSpec;
  slide.addText(data.kicker || 'Quote', { x: 0.78, y: 1.15, w: 4, h: 0.3, fontFace: FONTS.mono, fontSize: 8.5, charSpace: 2, color: s.fg, transparency: 35, margin: 0 });
  slide.addText(data.quote || data.title, { x: 0.78, y: 1.85, w: 10.2, h: 2.25, fontFace: FONTS.serifZh, fontSize: fitTitle(data.quote || data.title, 38, 27), bold: true, color: s.fg, margin: 0, fit: 'shrink' });
  slide.addText(data.body || data.cite || '', { x: 0.82, y: 4.55, w: 7.2, h: 0.6, fontFace: FONTS.serifZh, fontSize: 16, color: s.fg, transparency: 30, margin: 0, fit: 'shrink' });
  addFoot(slide, ctx, s.fg, 'magazine');
}

function magazineCompare(slide, ctx, s) {
  const data = ctx.slideSpec;
  addPageHead(slide, data, s.fg, 'magazine', 0.9);
  const cols = [data.before || data.left || {}, data.after || data.right || {}];
  cols.forEach((col, i) => {
    const x = i === 0 ? 0.9 : 7.0;
    slide.addShape(pptx.ShapeType.line, { x, y: 2.5, w: 0, h: 3.4, line: { color: s.fg, transparency: i === 0 ? 55 : 15, width: 1.6 } });
    slide.addText(col.label || (i === 0 ? 'Before' : 'After'), { x: x + 0.25, y: 2.55, w: 4.8, h: 0.25, fontFace: FONTS.mono, fontSize: 8, charSpace: 1.5, color: s.fg, transparency: i === 0 ? 55 : 25, margin: 0 });
    slide.addText(col.title || '', { x: x + 0.25, y: 3.0, w: 4.9, h: 0.55, fontFace: FONTS.serifZh, fontSize: 22, bold: true, color: s.fg, transparency: i === 0 ? 38 : 0, margin: 0, fit: 'shrink' });
    addBullets(slide, compareBulletItems(col.items || []), x + 0.25, 3.82, 4.9, 1.65, s.fg, i === 0 ? 35 : 10, 'magazine');
  });
  addFoot(slide, ctx, s.fg, 'magazine');
}

function magazineTextImage(slide, ctx, s) {
  const data = ctx.slideSpec;
  addPageHead(slide, data, s.fg, 'magazine', 0.92);
  slide.addText(data.body || '', { x: 0.78, y: 2.6, w: 6.4, h: 1.75, fontFace: FONTS.sansZh, fontSize: 13.5, color: s.fg, transparency: 18, margin: 0.03, fit: 'shrink', valign: 'top', breakLine: false });
  if (data.callout) addCallout(slide, data.callout, 0.78, 4.7, 6.4, 0.85, s.fg, ctx.theme.paperTint);
  addMediaOrChart(slide, ctx, data, { x: 8.0, y: 2.15, w: 3.6, h: 3.9 }, s, 'magazine', '3:4');
  addCaption(slide, data.caption || data.image?.caption, 8.0, 6.18, 3.6, s.fg, 'magazine');
  addFoot(slide, ctx, s.fg, 'magazine');
}

function magazineArticle(slide, ctx, s) {
  const data = ctx.slideSpec;
  addPageHead(slide, data, s.fg, 'magazine', 0.82);
  const sections = normalizeSections(data.sections || data.items || data.columns || []).slice(0, data.maxItems || 6);
  if (!sections.length) {
    addFoot(slide, ctx, s.fg, 'magazine');
    return;
  }
  const cols = clampColumns(data.columnsCount || autoColumns(sections.length, 3), 1, 3);
  const rows = Math.ceil(sections.length / cols);
  const x0 = 0.78;
  const y0 = data.subtitle ? 2.68 : 2.32;
  const gapX = 0.34;
  const gapY = rows > 1 ? 0.26 : 0;
  const gridW = 11.75;
  const maxBottom = data.callout ? 5.58 : 6.35;
  const colW = (gridW - gapX * (cols - 1)) / cols;
  const titleFont = Math.max(15, READABILITY.minFontSize);
  const bodyFont = READABILITY.minFontSize;
  const articleBodyMax = Math.max(rows > 1 ? 1.62 : 2.8, maxBottom - y0 - 1.0);
  const demands = sections.map((section) => {
    const hasIconWidth = 0.36;
    const textW = colW - hasIconWidth;
    const title = section.title || section.label || '';
    const body = section.body || section.desc || (section.items || []).map((item) => typeof item === 'string' ? item : item.text || item.title || '').join('\n');
    const titleH = estimateTextHeight(title, textW, titleFont, { min: 0.38, max: 0.72, lineHeight: 1.28, padding: 0.03 });
    const bodyH = estimateTextHeight(body, textW, bodyFont, { min: body ? 0.68 : 0, empty: 0, max: articleBodyMax, lineHeight: rows > 1 ? 1.46 : 1.52, padding: 0.16 });
    return Math.max(rows > 1 ? 1.62 : 2.35, 0.2 + titleH + 0.2 + bodyH + 0.26);
  });
  const rowHeights = distributeRowHeights(demands, rows, cols, rows > 1 ? 1.54 : 2.35, maxBottom - y0, gapY);
  sections.forEach((section, i) => {
    const row = Math.floor(i / cols);
    const x = x0 + (i % cols) * (colW + gapX);
    const y = y0 + rowHeights.slice(0, row).reduce((sum, h) => sum + h, 0) + row * gapY;
    const h = rowHeights[row];
    slide.addShape(pptx.ShapeType.line, { x, y, w: colW, h: 0, line: { color: s.fg, transparency: 65, width: 0.6 } });
    const hasIcon = addInlineIcon(slide, section, x, y + 0.15, 0.26, s.fg, 'magazine', { fallback: defaultContentIcon(i, 'magazine'), pad: 0.045 });
    const titleX = hasIcon ? x + 0.36 : x;
    const textW = colW - (hasIcon ? 0.36 : 0);
    const titleText = section.title || section.label || '';
    const body = section.body || section.desc || (section.items || []).map((item) => typeof item === 'string' ? item : item.text || item.title || '').join('\n');
    const titleH = estimateTextHeight(titleText, textW, titleFont, { min: 0.38, max: 0.72, lineHeight: 1.28, padding: 0.03 });
    const bodyY = y + 0.18 + titleH + 0.2;
    const bodyH = Math.max(0.62, h - (bodyY - y) - 0.26);
    slide.addText(titleText, { x: titleX, y: y + 0.16, w: textW, h: titleH, fontFace: FONTS.serifZh, fontSize: titleFont, bold: true, color: s.fg, margin: 0, fit: 'shrink', valign: 'top' });
    slide.addText(body, { x: titleX, y: bodyY, w: textW, h: bodyH, fontFace: FONTS.sansZh, fontSize: bodyFont, color: s.fg, transparency: 18, margin: 0.03, fit: 'shrink', valign: 'top', breakLine: false });
  });
  if (data.callout) addCallout(slide, data.callout, 8.3, 5.88, 3.65, 0.68, s.fg, ctx.theme.paperTint);
  addFoot(slide, ctx, s.fg, 'magazine');
}

function magazineSectionList(slide, ctx, s) {
  const data = ctx.slideSpec;
  addPageHead(slide, data, s.fg, 'magazine', 0.82);
  const sections = normalizeSections(data.sections || data.items || data.columns || []).slice(0, data.maxItems || 7);
  if (!sections.length) {
    addFoot(slide, ctx, s.fg, 'magazine');
    return;
  }
  const y0 = data.subtitle ? 2.62 : 2.32;
  const bottom = 6.34;
  const gapY = 0.14;
  const rowMin = sections.length <= 4 ? 0.78 : 0.58;
  const textW = 8.95;
  const demands = sections.map((item) => {
    const title = item.title || item.label || item.name || '';
    const body = item.body || item.desc || item.note || item.summary || item.detail || item.text || '';
    const titleH = estimateTextHeight(title, textW, 14.2, { min: 0.26, max: 0.46, lineHeight: 1.2, padding: 0.02 });
    const bodyH = estimateTextHeight(body, textW, READABILITY.minFontSize, { min: body ? 0.32 : 0, empty: 0, max: 0.82, lineHeight: 1.35, padding: 0.06 });
    return Math.max(rowMin, titleH + (body ? 0.08 + bodyH : 0) + 0.2);
  });
  const rowHeights = distributeRowHeights(demands, sections.length, 1, rowMin, bottom - y0, gapY);
  let y = y0;
  sections.forEach((item, i) => {
    const h = rowHeights[i];
    const hot = i === data.highlightIndex;
    const accent = hot ? ctx.theme.inkTint : s.fg;
    slide.addShape(pptx.ShapeType.line, { x: 0.78, y, w: 10.95, h: 0, line: { color: accent, transparency: hot ? 8 : 62, width: hot ? 1.6 : 0.65 } });
    slide.addText(item.label || String(i + 1).padStart(2, '0'), { x: 0.82, y: y + 0.16, w: 0.58, h: 0.26, fontFace: FONTS.serifEn, fontSize: 13.5, italic: true, bold: hot, color: accent, transparency: hot ? 0 : 25, margin: 0 });
    const title = item.title || item.name || '';
    const body = item.body || item.desc || item.note || item.summary || item.detail || item.text || '';
    const titleH = estimateTextHeight(title, textW, 14.2, { min: 0.28, max: 0.48, lineHeight: 1.2, padding: 0.02 });
    slide.addText(title, { x: 1.55, y: y + 0.12, w: textW, h: titleH, fontFace: FONTS.serifZh, fontSize: 14.2, bold: true, color: s.fg, margin: 0, valign: 'top' });
    if (body) slide.addText(body, { x: 1.55, y: y + 0.16 + titleH, w: textW, h: Math.max(0.28, h - titleH - 0.22), fontFace: FONTS.sansZh, fontSize: READABILITY.minFontSize, color: s.fg, transparency: 18, margin: 0.02, valign: 'top' });
    y += h + gapY;
  });
  addFoot(slide, ctx, s.fg, 'magazine');
}

function magazineDataSheet(slide, ctx, s) {
  const data = ctx.slideSpec;
  addPageHead(slide, data, s.fg, 'magazine', 0.82);
  if (data.body) slide.addText(data.body, { x: 0.78, y: 2.18, w: 4.2, h: 0.78, fontFace: FONTS.sansZh, fontSize: 10.8, color: s.fg, transparency: 18, margin: 0.02, fit: 'shrink' });
  addTableBlock(slide, ctx, data.table || data, { x: 0.78, y: data.body ? 3.05 : 2.35, w: 7.45, h: data.body ? 3.25 : 3.95 }, s, 'magazine');
  const notes = normalizeSections(data.notes || data.insights || []).slice(0, 3);
  notes.forEach((note, i) => {
    const y = 2.45 + i * 1.18;
    const hasIcon = addInlineIcon(slide, note, 8.65, y - 0.02, 0.28, s.fg, 'magazine', { fallback: note.icon ? null : 'info', pad: 0.045 });
    if (!hasIcon) slide.addText(note.label || `0${i + 1}`, { x: 8.65, y, w: 0.55, h: 0.22, fontFace: FONTS.mono, fontSize: 7.5, color: s.fg, transparency: 35, margin: 0 });
    slide.addText(note.title || note.body || '', { x: 9.05, y: y - 0.04, w: 2.95, h: 0.32, fontFace: FONTS.serifZh, fontSize: 13.5, bold: true, color: s.fg, margin: 0, fit: 'shrink' });
    if (note.body && note.title) slide.addText(note.body, { x: 9.05, y: y + 0.36, w: 2.95, h: 0.52, fontFace: FONTS.sansZh, fontSize: 8.8, color: s.fg, transparency: 25, margin: 0, fit: 'shrink' });
  });
  addFoot(slide, ctx, s.fg, 'magazine');
}

function magazineChart(slide, ctx, s) {
  const data = ctx.slideSpec;
  addPageHead(slide, data, s.fg, 'magazine', 0.82);
  addChartBlock(slide, ctx, data.chart || data, { x: 0.8, y: 2.35, w: 7.6, h: 3.75 }, s, 'magazine');
  const insights = normalizeSections(data.insights || data.notes || []).slice(0, 4);
  insights.forEach((item, i) => {
    const y = 2.45 + i * 0.9;
    slide.addShape(pptx.ShapeType.line, { x: 8.85, y, w: 2.75, h: 0, line: { color: s.fg, transparency: 65, width: 0.5 } });
    const hasIcon = addInlineIcon(slide, item, 8.85, y + 0.1, 0.25, s.fg, 'magazine', { fallback: i === 0 ? 'trending-up' : null, pad: 0.04 });
    const tx = hasIcon ? 9.18 : 8.85;
    slide.addText(item.title || item.label || '', { x: tx, y: y + 0.12, w: 11.6 - tx, h: 0.26, fontFace: FONTS.serifZh, fontSize: 12.5, bold: true, color: s.fg, margin: 0, fit: 'shrink' });
    slide.addText(item.body || item.desc || '', { x: tx, y: y + 0.44, w: 11.6 - tx, h: 0.34, fontFace: FONTS.sansZh, fontSize: 8.5, color: s.fg, transparency: 25, margin: 0, fit: 'shrink' });
  });
  addFoot(slide, ctx, s.fg, 'magazine');
}

function magazineDashboard(slide, ctx, s) {
  const data = ctx.slideSpec;
  addPageHead(slide, data, s.fg, 'magazine', 0.76);
  const metricsY = Math.max(2.32, pageHeadSafeBottom(ctx, 0.76, 0.22));
  const valueY = metricsY + 0.5;
  const noteY = metricsY + 1.08;
  const chartY = Math.min(4.35, metricsY + 1.9);
  const chartH = Math.max(1.8, 6.3 - chartY);
  const metrics = (data.metrics || data.items || []).slice(0, 4);
  metrics.forEach((item, i) => {
    const x = 0.8 + i * 3.0;
    slide.addShape(pptx.ShapeType.line, { x, y: metricsY, w: 2.35, h: 0, line: { color: s.fg, transparency: 55, width: 0.7 } });
    const hasIcon = addInlineIcon(slide, item, x, metricsY + 0.16, 0.24, s.fg, 'magazine', { fallback: i === 0 ? 'chart-column' : null, pad: 0.04 });
    const tx = hasIcon ? x + 0.34 : x;
    slide.addText(item.label || '', { x: tx, y: metricsY + 0.2, w: 2.35 - (hasIcon ? 0.34 : 0), h: 0.2, fontFace: FONTS.mono, fontSize: 7, charSpace: 1.1, color: s.fg, transparency: 35, margin: 0, fit: 'shrink' });
    slide.addText(item.value || '', { x, y: valueY, w: 2.35, h: 0.52, fontFace: FONTS.serifEn, fontSize: 30, bold: true, color: s.fg, margin: 0, fit: 'shrink' });
    slide.addText(item.note || '', { x, y: noteY, w: 2.35, h: 0.32, fontFace: FONTS.sansZh, fontSize: 8.4, color: s.fg, transparency: 30, margin: 0, fit: 'shrink' });
  });
  const charts = data.charts || [];
  if (charts[0]) addChartBlock(slide, ctx, charts[0], { x: 0.8, y: chartY, w: 5.55, h: chartH }, s, 'magazine');
  if (charts[1]) addChartBlock(slide, ctx, charts[1], { x: 6.8, y: chartY, w: 5.2, h: chartH }, s, 'magazine');
  addFoot(slide, ctx, s.fg, 'magazine');
}
function magazineAgenda(slide, ctx, s) {
  const data = ctx.slideSpec;
  addPageHead(slide, data, s.fg, 'magazine', 0.82);
  const items = normalizeSections(data.items || data.sections || data.agenda || []).slice(0, 8);
  const cols = clampColumns(data.columnsCount || autoColumns(items.length, 4), 1, 4);
  const gap = 0.28;
  const cardW = (10.2 - gap * (cols - 1)) / cols;
  const startX = 1.55;
  const startY = data.subtitle ? 2.85 : 2.55;
  slide.addText(data.index || String(ctx.index + 1).padStart(2, '0'), { x: 0.74, y: 2.58, w: 0.65, h: 2.7, fontFace: FONTS.serifEn, fontSize: 34, bold: true, color: s.fg, transparency: 45, margin: 0, fit: 'shrink' });
  items.forEach((item, i) => {
    const x = startX + (i % cols) * (cardW + gap);
    const y = startY + Math.floor(i / cols) * 1.52;
    slide.addShape(pptx.ShapeType.line, { x, y, w: cardW, h: 0, line: { color: s.fg, transparency: 58, width: 0.7 } });
    const hasIcon = addInlineIcon(slide, item, x, y + 0.17, 0.26, s.fg, 'magazine', { fallback: defaultContentIcon(i, 'magazine'), pad: 0.045 });
    const tx = hasIcon ? x + 0.36 : x;
    slide.addText(item.title || item.label || `Part ${i + 1}`, { x: tx, y: y + 0.16, w: cardW - (hasIcon ? 0.36 : 0), h: 0.28, fontFace: FONTS.serifZh, fontSize: 14.2, bold: true, color: s.fg, margin: 0, fit: 'shrink' });
    slide.addText(item.body || item.desc || item.note || '', { x: tx, y: y + 0.56, w: cardW - (hasIcon ? 0.36 : 0), h: 0.52, fontFace: FONTS.sansZh, fontSize: 9.6, color: s.fg, transparency: 22, margin: 0.02, fit: 'shrink', valign: 'top' });
  });
  addFoot(slide, ctx, s.fg, 'magazine');
}

function magazineCaseStudy(slide, ctx, s) {
  const data = ctx.slideSpec;
  addPageHead(slide, data, s.fg, 'magazine', 0.82);
  addMediaOrChart(slide, ctx, data, { x: 0.78, y: 2.35, w: 5.5, h: 3.45 }, s, 'magazine', 'CASE');
  addCaption(slide, data.caption || data.image?.caption, 0.78, 5.92, 5.5, s.fg, 'magazine');
  slide.addText(data.caseTitle || data.label || 'Case', { x: 6.75, y: 2.35, w: 4.95, h: 0.5, fontFace: FONTS.serifZh, fontSize: 22, bold: true, color: s.fg, margin: 0, fit: 'shrink' });
  slide.addText(data.body || data.story || '', { x: 6.78, y: 3.05, w: 4.95, h: 1.15, fontFace: FONTS.sansZh, fontSize: 11, color: s.fg, transparency: 18, margin: 0.03, fit: 'shrink', valign: 'top' });
  const metrics = (data.metrics || data.items || []).slice(0, 3);
  metrics.forEach((item, i) => {
    const x = 6.78 + i * 1.72;
    slide.addShape(pptx.ShapeType.line, { x, y: 4.65, w: 1.28, h: 0, line: { color: s.fg, transparency: 55, width: 0.6 } });
    slide.addText(item.value || item.title || '', { x, y: 4.82, w: 1.45, h: 0.42, fontFace: FONTS.serifEn, fontSize: 24, bold: true, color: s.fg, margin: 0, fit: 'shrink' });
    slide.addText(item.label || item.note || '', { x, y: 5.35, w: 1.45, h: 0.38, fontFace: FONTS.sansZh, fontSize: 9.4, color: s.fg, transparency: 28, margin: 0, fit: 'shrink' });
  });
  addFoot(slide, ctx, s.fg, 'magazine');
}

function magazinePyramid(slide, ctx, s) {
  const data = ctx.slideSpec;
  addPageHead(slide, data, s.fg, 'magazine', 0.82);
  const layers = normalizeSections(data.layers || data.items || data.sections || []).slice(0, 5).reverse();
  const centerX = 6.5;
  const baseY = 5.72;
  layers.forEach((item, i) => {
    const w = 3.1 + i * 1.48;
    const h = 0.62;
    const x = centerX - w / 2;
    const y = baseY - i * 0.72;
    const fill = i === layers.length - 1 ? ctx.theme.inkTint : ctx.theme.paperTint;
    const color = i === layers.length - 1 ? ctx.theme.paper : s.fg;
    slide.addShape(pptx.ShapeType.rect, { x, y, w, h, fill: { color: fill, transparency: i === layers.length - 1 ? 0 : 8 }, line: { color: s.fg, transparency: 55, width: 0.5 } });
    slide.addText(item.title || item.label || '', { x: x + 0.18, y: y + 0.14, w: w - 0.36, h: 0.22, fontFace: FONTS.serifZh, fontSize: 13.4, bold: true, color, align: 'center', margin: 0, fit: 'shrink' });
  });
  slide.addText(data.body || data.note || '', { x: 0.78, y: 5.78, w: 3.8, h: 0.55, fontFace: FONTS.sansZh, fontSize: 10.2, color: s.fg, transparency: 25, margin: 0, fit: 'shrink' });
  addFoot(slide, ctx, s.fg, 'magazine');
}

function magazineRadial(slide, ctx, s) {
  const data = ctx.slideSpec;
  addPageHead(slide, data, s.fg, 'magazine', 0.78);
  const items = normalizeSections(data.items || data.nodes || data.sections || []).slice(0, 8);
  const cx = 6.55;
  const cy = 4.05;
  const centerBox = { x: cx - 1.12, y: cy - 0.62, w: 2.24, h: 1.24 };
  slide.addShape(pptx.ShapeType.ellipse, { ...centerBox, fill: { color: ctx.theme.paperTint, transparency: 8 }, line: { color: s.fg, transparency: 45, width: 0.8 } });
  slide.addText(data.center || data.label || data.title || '', { x: cx - 0.9, y: cy - 0.24, w: 1.8, h: 0.42, fontFace: FONTS.serifZh, fontSize: 15.5, bold: true, color: s.fg, align: 'center', margin: 0, fit: 'shrink' });
  const nodeW = 2.55;
  const nodeH = 1.08;
  items.forEach((item, i) => {
    const angle = -Math.PI / 2 + (i * 2 * Math.PI) / Math.max(items.length, 1);
    const x = cx + Math.cos(angle) * 4.05;
    const y = cy + Math.sin(angle) * 1.55;
    const boxX = clamp(x - nodeW / 2, 0.72, 10.08);
    const boxY = clamp(y - nodeH / 2, 1.95, 5.45);
    const nodeBox = { x: boxX, y: boxY, w: nodeW, h: nodeH };
    addRadialConnector(slide, centerBox, nodeBox, { color: s.fg, transparency: 74, width: 0.45 });
    slide.addShape(pptx.ShapeType.rect, { ...nodeBox, fill: { color: ctx.theme.paperTint, transparency: 28 }, line: { color: s.fg, transparency: 82, width: 0.4 } });
    const hasIcon = addInlineIcon(slide, item, boxX + 0.14, boxY + 0.13, 0.25, s.fg, 'magazine', { fallback: defaultContentIcon(i, 'magazine'), pad: 0.04 });
    const tx = hasIcon ? boxX + 0.48 : boxX + 0.18;
    const textW = boxX + nodeW - 0.18 - tx;
    slide.addText(item.title || item.label || '', { x: tx, y: boxY + 0.12, w: textW, h: 0.3, fontFace: FONTS.serifZh, fontSize: 12.2, bold: true, color: s.fg, margin: 0, fit: 'shrink' });
    slide.addText(item.body || item.desc || '', { x: tx, y: boxY + 0.48, w: textW, h: 0.42, fontFace: FONTS.sansZh, fontSize: 9.2, color: s.fg, transparency: 28, margin: 0.01, fit: 'shrink', valign: 'top' });
  });
  addFoot(slide, ctx, s.fg, 'magazine');
}
function magazineRoadmap(slide, ctx, s) {
  const data = ctx.slideSpec;
  addPageHead(slide, data, s.fg, 'magazine', 0.82);
  const steps = normalizeSections(data.steps || data.items || []).slice(0, 6);
  const x0 = 0.95;
  const y0 = 3.62;
  const stepW = 10.95 / Math.max(steps.length, 1);
  slide.addShape(pptx.ShapeType.line, { x: x0, y: y0, w: 10.8, h: 0, line: { color: s.fg, transparency: 55, width: 1 } });
  steps.forEach((item, i) => {
    const x = x0 + i * stepW;
    const y = y0 + (i % 2 === 0 ? -1.0 : 0.38);
    slide.addShape(pptx.ShapeType.ellipse, { x: x - 0.08, y: y0 - 0.08, w: 0.16, h: 0.16, fill: { color: s.fg }, line: { color: s.fg, transparency: 100 } });
    const textW = steps.length <= 4 ? 2.25 : steps.length === 5 ? 1.9 : 1.72;
    const labelX = clamp(x - 0.48, 0.65, SLIDE.w - 1.6);
    const textX = clamp(x - textW / 2, 0.65, SLIDE.w - textW - 0.35);
    slide.addText(item.label || item.date || `0${i + 1}`, { x: labelX, y: y - 0.24, w: 0.95, h: 0.2, fontFace: FONTS.mono, fontSize: 9.5, color: s.fg, transparency: 35, align: 'center', margin: 0, fit: 'shrink' });
    slide.addText(item.title || '', { x: textX, y, w: textW, h: 0.34, fontFace: FONTS.serifZh, fontSize: 12.5, bold: true, color: s.fg, align: 'center', margin: 0, fit: 'shrink' });
    slide.addText(item.body || item.desc || item.note || '', { x: textX, y: y + 0.42, w: textW, h: 0.48, fontFace: FONTS.sansZh, fontSize: 9.2, color: s.fg, transparency: 26, align: 'center', margin: 0, fit: 'shrink' });
  });
  addFoot(slide, ctx, s.fg, 'magazine');
}

function magazineSwimlane(slide, ctx, s) {
  const data = ctx.slideSpec;
  addPageHead(slide, data, s.fg, 'magazine', 0.78);
  const lanes = (data.lanes || data.sections || []).slice(0, 4);
  const stages = data.stages || data.columns || ['现在', '下一步', '后续'];
  const x0 = 1.55;
  const y0 = 2.55;
  const laneH = 0.86;
  const colW = 10.25 / Math.max(stages.length, 1);
  stages.forEach((stage, i) => slide.addText(String(stage), { x: x0 + i * colW, y: 2.22, w: colW - 0.12, h: 0.22, fontFace: FONTS.mono, fontSize: 9.5, color: s.fg, transparency: 35, margin: 0, fit: 'shrink' }));
  lanes.forEach((lane, r) => {
    const y = y0 + r * 1.02;
    slide.addText(lane.title || lane.label || `Lane ${r + 1}`, { x: 0.78, y: y + 0.2, w: 0.62, h: 0.34, fontFace: FONTS.serifZh, fontSize: READABILITY.minFontSize, bold: true, color: s.fg, margin: 0, fit: 'shrink' });
    const cells = lane.items || lane.steps || [];
    stages.forEach((stage, c) => {
      const cell = cells[c] || {};
      const x = x0 + c * colW;
      slide.addShape(pptx.ShapeType.rect, { x, y, w: colW - 0.18, h: laneH, fill: { color: ctx.theme.paperTint, transparency: 15 }, line: { color: s.fg, transparency: 72, width: 0.4 } });
      slide.addText(cellText(cell), { x: x + 0.12, y: y + 0.16, w: colW - 0.42, h: 0.46, fontFace: FONTS.sansZh, fontSize: 9.6, color: s.fg, margin: 0, fit: 'shrink', valign: 'mid' });
    });
  });
  addFoot(slide, ctx, s.fg, 'magazine');
}
function cellText(cell) {
  if (cell == null) return '';
  if (typeof cell === 'string' || typeof cell === 'number') return String(cell);
  return String(cell.title || cell.text || cell.body || cell.label || '');
}
function magazineStatementCompat(slide, ctx, s) {
  const data = ctx.slideSpec;
  const headY = pageHeadY(ctx, 1.05);
  slide.addText(data.kicker || 'Statement', { x: 0.78, y: headY, w: 5.8, h: 0.22, fontFace: FONTS.mono, fontSize: 8.2, charSpace: 1.7, color: s.fg, transparency: 35, margin: 0, fit: 'shrink' });
  slide.addText(data.title || data.quote || ctx.spec.title || '', { x: 0.78, y: headY + 0.55, w: 6.35, h: 2.28, fontFace: FONTS.serifZh, fontSize: fitTitle(data.title || data.quote || ctx.spec.title || '', 37, 27), bold: true, color: s.fg, margin: 0, fit: 'shrink' });
  slide.addShape(pptx.ShapeType.line, { x: 0.82, y: headY + 3.18, w: 2.35, h: 0, line: { color: s.fg, transparency: 55, width: 0.8 } });
  slide.addText(data.body || data.subtitle || data.cite || '', { x: 0.82, y: headY + 3.48, w: 5.65, h: 0.85, fontFace: FONTS.sansZh, fontSize: 13.8, color: s.fg, transparency: 20, margin: 0, fit: 'shrink', valign: 'top' });
  addStatementImageSlot(slide, ctx, { x: 7.35, y: headY + 0.48, w: 4.55, h: 3.85 }, s.fg, data.imageLabel || '图片占位');
  addFoot(slide, ctx, s.fg, 'magazine');
}

function magazineMatrixCompat(slide, ctx, s) {
  const data = ctx.slideSpec;
  const sections = data.sections || data.items || data.columns || [];
  magazineArticle(slide, { ...ctx, slideSpec: { ...data, sections } }, s);
}

function magazineImageHeroCompat(slide, ctx, s) {
  const data = ctx.slideSpec;
  magazineQuoteImage(slide, { ...ctx, slideSpec: { ...data, body: data.body || data.subtitle || data.quote || '', callout: data.callout || data.quote || data.subtitle || '' } }, s);
}
function magazineClosing(slide, ctx, s) {
  const data = ctx.slideSpec;
  magazineBigQuote(slide, { ...ctx, slideSpec: { ...data, quote: data.title || '谢谢', body: data.subtitle || data.body } }, s);
}

function swissCover(slide, ctx, s) {
  const data = ctx.slideSpec;
  const headY = pageHeadY(ctx, 1.12);
  slide.addText(data.kicker || 'Swiss Field Note', { x: 0.72, y: headY, w: 6, h: 0.28, fontFace: FONTS.mono, fontSize: 8.5, charSpace: 1.8, color: s.fg, transparency: 18, margin: 0 });
  slide.addText(data.title || ctx.spec.title, { x: 0.68, y: headY + 0.5, w: 11.5, h: 2.6, fontFace: FONTS.sans, fontSize: fitTitle(data.title || ctx.spec.title, 62, 40), bold: false, color: s.fg, margin: 0, fit: 'shrink' , typographyRole: 'coverTitle' });
  slide.addText(data.subtitle || ctx.spec.subtitle || '', { x: 0.75, y: headY + 4.2, w: 6.2, h: 0.55, fontFace: FONTS.sansZh, fontSize: 17, color: s.fg, transparency: 18, margin: 0, fit: 'shrink' });
  addSwissBars(slide, 10.2, headY + 3.9, 1.55, 1.1, s.fg, 55);
  addFoot(slide, ctx, s.fg, 'swiss');
}
function swissStatement(slide, ctx, s) {
  const data = ctx.slideSpec;
  const headY = pageHeadY(ctx, 1.12);
  slide.addText(data.kicker || 'Statement', { x: 0.72, y: headY, w: 5.5, h: 0.25, fontFace: FONTS.mono, fontSize: 8, charSpace: 1.8, color: s.fg, transparency: 35, margin: 0 });
  slide.addText(data.title, { x: 0.68, y: headY + 0.5, w: 6.35, h: 2.1, fontFace: FONTS.sans, fontSize: fitTitle(data.title, 43, 30), bold: false, color: s.fg, margin: 0, fit: 'shrink' });
  if (data.body || data.subtitle) {
    slide.addText(data.body || data.subtitle, { x: 0.72, y: headY + 3.22, w: 5.9, h: 0.95, fontFace: FONTS.sansZh, fontSize: 15.5, color: s.fg, transparency: 20, margin: 0, fit: 'shrink', valign: 'top' });
  }
  addStatementImageSlot(slide, ctx, { x: 7.38, y: headY + 0.74, w: 4.25, h: 3.65 }, s.fg, data.imageLabel || '图片占位');
  addFoot(slide, ctx, s.fg, 'swiss');
}
function swissKpiTower(slide, ctx, s) {
  const data = ctx.slideSpec;
  addPageHead(slide, data, s.fg, 'swiss', 0.9);
  const items = (data.items || []).slice(0, 4);
  const max = Math.max(...items.map((x) => Number(x.valueNum || String(x.value || '').replace(/[^\d.]/g, '')) || 1));
  const valueY = 2.76;
  const valueH = 0.5;
  const barBottom = 5.62;
  const barTop = 3.78;
  const maxBarH = barBottom - barTop;
  items.forEach((item, i) => {
    const x = 0.95 + i * 3.0;
    const rawValue = Number(item.valueNum || String(item.value || '').replace(/[^\d.]/g, '')) || max * 0.55;
    const barH = 0.38 + (rawValue / max) * Math.max(0.1, maxBarH - 0.38);
    const barY = barBottom - barH;
    slide.addText(item.value || '', { x, y: valueY, w: 2.3, h: valueH, fontFace: FONTS.sans, fontSize: 28, bold: true, color: i === items.length - 1 ? ctx.theme.accent : s.fg, margin: 0, fit: 'shrink' });
    slide.addShape(pptx.ShapeType.rect, { x, y: barY, w: 2.3, h: barH, fill: { color: i === items.length - 1 ? ctx.theme.accent : ctx.theme.grey1 }, line: { color: i === items.length - 1 ? ctx.theme.accent : ctx.theme.grey2, transparency: 0, width: 0.4 } });
    slide.addText(item.label || '', { x, y: 5.85, w: 2.3, h: 0.35, fontFace: FONTS.mono, fontSize: 7.5, charSpace: 1.3, color: s.fg, transparency: 25, margin: 0, fit: 'shrink' });
  });
  addFoot(slide, ctx, s.fg, 'swiss');
}

function swissDuoCompare(slide, ctx, s) {
  const data = ctx.slideSpec;
  addPageHead(slide, data, s.fg, 'swiss', 0.85);
  const cols = [data.before || data.left || {}, data.after || data.right || {}];
  cols.forEach((col, i) => {
    const x = i === 0 ? 0.78 : 6.85;
    const fill = i === 1 ? ctx.theme.accent : ctx.theme.grey1;
    const color = i === 1 ? ctx.theme.accentOn : ctx.theme.ink;
    slide.addShape(pptx.ShapeType.rect, { x, y: 2.55, w: 5.55, h: 3.25, fill: { color: fill }, line: { color: fill, transparency: 100 } });
    slide.addText(col.label || (i === 0 ? 'Before' : 'After'), { x: x + 0.35, y: 2.88, w: 4.7, h: 0.22, fontFace: FONTS.mono, fontSize: 8, charSpace: 1.3, color, transparency: 30, margin: 0 });
    slide.addText(col.title || '', { x: x + 0.35, y: 3.35, w: 4.7, h: 0.55, fontFace: FONTS.sans, fontSize: 23, color, margin: 0, fit: 'shrink' });
    addBullets(slide, compareBulletItems(col.items || []), x + 0.35, 4.22, 4.7, 1.15, color, 14, 'swiss');
  });
  addFoot(slide, ctx, s.fg, 'swiss');
}

function swissTimeline(slide, ctx, s) {
  const data = ctx.slideSpec;
  addPageHead(slide, data, s.fg, 'swiss', 0.88);
  const items = (data.items || data.steps || []).slice(0, 6);
  if (!items.length) {
    addFoot(slide, ctx, s.fg, 'swiss');
    return;
  }
  const cols = clampColumns(data.columnsCount || (items.length <= 4 ? items.length : 3), 1, 4);
  const rows = Math.ceil(items.length / cols);
  const gridX = 0.78;
  const gridW = 11.45;
  const gapX = 0.34;
  const gapY = rows > 1 ? 0.36 : 0;
  const startY = rows > 1 ? 2.48 : 2.92;
  const maxBottom = 6.35;
  const cardW = (gridW - gapX * (cols - 1)) / cols;
  const titleFont = READABILITY.minFontSize;
  const bodyFont = READABILITY.minFontSize;
  const demands = items.map((item) => {
    const title = item.title || '';
    const body = item.body || item.desc || item.note || '';
    const titleH = estimateTextHeight(title, cardW - 0.34, titleFont, { min: 0.32, max: 0.56, lineHeight: 1.24, padding: 0.02 });
    const bodyH = estimateTextHeight(body, cardW - 0.34, bodyFont, { min: body ? 0.48 : 0, empty: 0, max: rows > 1 ? 0.9 : 1.18, lineHeight: 1.34, padding: 0.08 });
    return Math.max(rows > 1 ? 1.42 : 1.78, 0.54 + titleH + 0.14 + bodyH + 0.18);
  });
  const rowHeights = distributeRowHeights(demands, rows, cols, rows > 1 ? 1.36 : 1.78, maxBottom - startY, gapY);
  items.forEach((item, i) => {
    const row = Math.floor(i / cols);
    const col = i % cols;
    const rowStart = row * cols;
    const rowCount = Math.min(cols, items.length - rowStart);
    const rowW = rowCount * cardW + Math.max(0, rowCount - 1) * gapX;
    const rowX = gridX + (gridW - rowW) / 2;
    const x = rowX + col * (cardW + gapX);
    const y = startY + rowHeights.slice(0, row).reduce((sum, h) => sum + h, 0) + row * gapY;
    const h = rowHeights[row];
    const hot = i === items.length - 1 || item.highlight;
    const axisY = y + 0.3;
    slide.addShape(pptx.ShapeType.line, { x, y: axisY, w: cardW, h: 0, line: { color: hot ? ctx.theme.accent : s.fg, transparency: hot ? 20 : 58, width: hot ? 1.0 : 0.7 } });
    slide.addShape(pptx.ShapeType.rect, { x: x + 0.02, y: axisY - 0.045, w: 0.09, h: 0.09, fill: { color: hot ? ctx.theme.accent : s.fg }, line: { color: hot ? ctx.theme.accent : s.fg, transparency: 100 } });
    slide.addText(item.label || item.year || String(i + 1).padStart(2, '0'), { x: x + 0.18, y: y + 0.04, w: Math.max(0.8, cardW - 0.36), h: 0.22, fontFace: FONTS.mono, fontSize: 8, color: hot ? ctx.theme.accent : s.fg, transparency: hot ? 0 : 30, margin: 0, fit: 'shrink' });
    const titleText = item.title || '';
    const bodyText = item.body || item.desc || item.note || '';
    const titleH = estimateTextHeight(titleText, cardW - 0.34, titleFont, { min: 0.32, max: 0.56, lineHeight: 1.24, padding: 0.02 });
    const titleY = axisY + 0.22;
    const bodyY = titleY + titleH + 0.14;
    slide.addText(titleText, { x: x + 0.17, y: titleY, w: cardW - 0.34, h: titleH, fontFace: FONTS.sansZh, fontSize: titleFont, bold: true, color: s.fg, margin: 0, fit: 'shrink', valign: 'top' });
    slide.addText(bodyText, { x: x + 0.17, y: bodyY, w: cardW - 0.34, h: Math.max(0.42, h - (bodyY - y) - 0.16), fontFace: FONTS.sansZh, fontSize: bodyFont, color: s.fg, transparency: 25, margin: 0.02, fit: 'shrink', valign: 'top' });
  });
  addFoot(slide, ctx, s.fg, 'swiss');
}

function swissMatrix(slide, ctx, s) {
  const data = ctx.slideSpec;
  addPageHead(slide, data, s.fg, 'swiss', 0.83);
  const items = (data.items || []).slice(0, 12);
  const cols = clampColumns(data.columnsCount || autoColumns(items.length, 4), 1, 4);
  const x0 = 0.78;
  const y0 = 2.55;
  const gap = 0.28;
  const w = (11.45 - gap * (cols - 1)) / cols;
  const rows = Math.ceil(items.length / cols);
  const h = rows <= 2 ? 0.92 : 0.78;
  items.forEach((item, i) => {
    const x = x0 + (i % cols) * (w + gap);
    const y = y0 + Math.floor(i / cols) * (h + 0.2);
    slide.addShape(pptx.ShapeType.rect, { x, y, w, h, fill: { color: i === items.length - 1 && data.highlightLast ? ctx.theme.accent : ctx.theme.grey1 }, line: { color: ctx.theme.grey1, transparency: 100 } });
    slide.addText(item.title || item.label || String(item), { x: x + 0.18, y: y + 0.18, w: w - 0.36, h: 0.3, fontFace: FONTS.sansZh, fontSize: 10.5, color: i === items.length - 1 && data.highlightLast ? ctx.theme.accentOn : s.fg, margin: 0, fit: 'shrink' });
  });
  if (data.heroStat) {
    slide.addText(data.heroStat.value || '', { x: 0.78, y: 5.75, w: 3.2, h: 0.65, fontFace: FONTS.sans, fontSize: 37, color: ctx.theme.accent, bold: true, margin: 0, fit: 'shrink' });
    slide.addText(data.heroStat.label || '', { x: 4.1, y: 5.95, w: 5.5, h: 0.28, fontFace: FONTS.mono, fontSize: 8, charSpace: 1.2, color: s.fg, transparency: 35, margin: 0, fit: 'shrink' });
  }
  addFoot(slide, ctx, s.fg, 'swiss');
}

function swissFourCards(slide, ctx, s) {
  const data = ctx.slideSpec;
  addPageHead(slide, data, s.fg, 'swiss', 0.85);
  const items = (data.items || []).slice(0, data.maxItems || 8);
  const count = items.length;
  if (!count) {
    addFoot(slide, ctx, s.fg, 'swiss');
    return;
  }
  const cols = clampColumns(data.columnsCount || autoCardColumns(count), 1, 4);
  const rows = Math.ceil(count / cols);
  const gapX = 0.34;
  const gapY = rows > 1 ? 0.28 : 0;
  const gridW = 11.45;
  const cardW = (gridW - gapX * (cols - 1)) / cols;
  const startX = 0.78;
  const startY = rows > 1 ? 2.42 : 2.62;
  const maxBottom = 6.35;
  const compact = rows > 1;
  const titleFont = Math.max(compact ? 13.5 : 17, READABILITY.minFontSize);
  const bodyFont = READABILITY.minFontSize;
  const minCardH = compact ? 1.58 : 2.72;
  const demands = items.map((item) => {
    const title = item.title || item.label || '';
    const body = item.desc || item.note || item.body || '';
    const titleH = estimateTextHeight(title, cardW - 0.4, titleFont, { min: compact ? 0.38 : 0.5, max: compact ? 0.72 : 0.9 });
    const bodyH = estimateTextHeight(body, cardW - 0.4, bodyFont, { min: body ? 0.42 : 0, empty: 0, max: compact ? 1.02 : 1.55 });
    return Math.max(minCardH, (compact ? 0.96 : 1.3) + titleH + bodyH + 0.32);
  });
  const rowHeights = distributeRowHeights(demands, rows, cols, minCardH, maxBottom - startY, gapY);
  items.forEach((item, i) => {
    const row = Math.floor(i / cols);
    const col = i % cols;
    const rowStart = row * cols;
    const rowCount = Math.min(cols, count - rowStart);
    const rowW = rowCount * cardW + Math.max(0, rowCount - 1) * gapX;
    const rowX = startX + (gridW - rowW) / 2;
    const x = rowX + col * (cardW + gapX);
    const y = startY + rowHeights.slice(0, row).reduce((sum, h) => sum + h, 0) + row * gapY;
    const cardH = rowHeights[row];
    const iconSize = compact ? 0.25 : 0.34;
    const titleText = item.title || item.label || '';
    const bodyText = item.desc || item.note || item.body || '';
    const titleY = compact ? y + 0.52 : y + 0.64;
    const titleH = estimateTextHeight(titleText, cardW - 0.4, titleFont, { min: compact ? 0.38 : 0.5, max: compact ? 0.68 : 0.86 });
    const descY = titleY + titleH + (compact ? 0.14 : 0.22);
    const descH = Math.max(0.34, cardH - (descY - y) - 0.24);
    slide.addShape(pptx.ShapeType.rect, { x, y, w: cardW, h: cardH, fill: { color: ctx.theme.grey1 }, line: { color: ctx.theme.grey1, transparency: 100 } });
    const hasIcon = addInlineIcon(slide, item, x + 0.2, y + 0.18, iconSize, s.fg, 'swiss', { fallback: item.icon ? null : 'layers', pad: compact ? 0.045 : 0.06 });
    if (!hasIcon) slide.addText(item.number || `0${i + 1}`, { x: x + 0.2, y: y + 0.22, w: 1.0, h: 0.22, fontFace: FONTS.mono, fontSize: 8, color: s.fg, transparency: 45, margin: 0 });
    slide.addText(titleText, { x: x + 0.2, y: titleY, w: cardW - 0.4, h: titleH, fontFace: FONTS.sansZh, fontSize: titleFont, color: s.fg, margin: 0, fit: 'shrink', valign: 'top' });
    slide.addText(bodyText, { x: x + 0.2, y: descY, w: cardW - 0.4, h: descH, fontFace: FONTS.sansZh, fontSize: bodyFont, color: s.fg, transparency: 25, margin: 0.02, fit: 'shrink', valign: 'top' });
  });
  addFoot(slide, ctx, s.fg, 'swiss');
}
function swissImageHero(slide, ctx, s) {
  const data = ctx.slideSpec;
  addMediaOrChart(slide, ctx, data, { x: 0, y: 0, w: SLIDE.w, h: 4.15 }, s, 'swiss', '21:9');
  slide.addShape(pptx.ShapeType.rect, { x: 0.78, y: 1.05, w: 4.65, h: 1.35, fill: { color: ctx.theme.paper }, line: { color: ctx.theme.paper, transparency: 100 } });
  slide.addText(data.title || '', { x: 1.08, y: 1.32, w: 4.0, h: 0.78, fontFace: FONTS.sans, fontSize: 29, color: ctx.theme.ink, margin: 0, fit: 'shrink' });
  slide.addShape(pptx.ShapeType.rect, { x: 0, y: 4.15, w: SLIDE.w, h: 3.35, fill: { color: ctx.theme.paper }, line: { color: ctx.theme.paper, transparency: 100 } });
  slide.addText(data.body || data.subtitle || '', { x: 0.78, y: 4.75, w: 5.3, h: 0.85, fontFace: FONTS.sansZh, fontSize: 14.5, color: ctx.theme.ink, margin: 0, fit: 'shrink' });
  const items = (data.items || []).slice(0, 3);
  items.forEach((item, i) => {
    const x = 6.6 + i * 2.1;
    slide.addShape(pptx.ShapeType.line, { x, y: 4.72, w: 1.75, h: 0, line: { color: ctx.theme.ink, transparency: 25, width: 0.7 } });
    slide.addText(item.label || '', { x, y: 4.9, w: 1.75, h: 0.2, fontFace: FONTS.mono, fontSize: 7, charSpace: 1, color: ctx.theme.ink, transparency: 35, margin: 0, fit: 'shrink' });
    slide.addText(item.value || '', { x, y: 5.25, w: 1.75, h: 0.45, fontFace: FONTS.sans, fontSize: 25, color: i === items.length - 1 ? ctx.theme.accent : ctx.theme.ink, margin: 0, fit: 'shrink' });
    slide.addText(item.note || '', { x, y: 5.95, w: 1.75, h: 0.35, fontFace: FONTS.sansZh, fontSize: 8.5, color: ctx.theme.ink, transparency: 35, margin: 0, fit: 'shrink' });
  });
}

function swissTextGrid(slide, ctx, s) {
  const data = ctx.slideSpec;
  addPageHead(slide, data, s.fg, 'swiss', 0.82);
  const sections = normalizeSections(data.sections || data.items || data.columns || []).slice(0, 9);
  const cols = clampColumns(data.columnsCount || autoColumns(sections.length, 3), 1, 3);
  const x0 = 0.78;
  const y0 = data.subtitle ? 2.74 : 2.42;
  const gapX = 0.4;
  const gapY = 0.3;
  const w = (11.45 - gapX * (cols - 1)) / cols;
  const rows = Math.ceil(sections.length / cols);
  const minH = rows === 1 ? 2.0 : rows === 2 ? 1.62 : 1.14;
  const maxBottom = 6.42;
  const gridBodyMax = Math.max(rows === 1 ? 2.4 : rows === 2 ? 1.48 : 0.96, (maxBottom - y0 - gapY * Math.max(0, rows - 1)) / rows - 0.78);
  const demands = sections.map((item) => {
    const title = item.title || '';
    const body = item.body || item.desc || '';
    const hasIconWidth = 0.62;
    const textW = w - hasIconWidth - 0.27;
    const titleH = estimateTextHeight(title, textW, READABILITY.minFontSize, { min: 0.36, max: 0.68, lineHeight: 1.28, padding: 0.03 });
    const bodyH = estimateTextHeight(body, textW, READABILITY.minFontSize, { min: body ? 0.52 : 0, empty: 0, max: gridBodyMax, lineHeight: 1.48, padding: 0.14 });
    return Math.max(minH, 0.3 + titleH + 0.16 + bodyH + 0.28);
  });
  const rowHeights = distributeRowHeights(demands, rows, cols, minH, maxBottom - y0, gapY);
  sections.forEach((item, i) => {
    const row = Math.floor(i / cols);
    const x = x0 + (i % cols) * (w + gapX);
    const y = y0 + rowHeights.slice(0, row).reduce((sum, h) => sum + h, 0) + row * gapY;
    const h = rowHeights[row];
    const hot = i === data.highlightIndex;
    slide.addShape(pptx.ShapeType.rect, { x, y, w, h, fill: { color: hot ? ctx.theme.accent : ctx.theme.grey1 }, line: { color: hot ? ctx.theme.accent : ctx.theme.grey1, transparency: 100 } });
    const iconColor = hot ? ctx.theme.accentOn : s.fg;
    const hasIcon = addInlineIcon(slide, item, x + 0.18, y + 0.14, 0.3, iconColor, 'swiss', { fallback: defaultContentIcon(i, 'swiss'), pad: 0.05 });
    if (!hasIcon) slide.addText(item.label || String(i + 1).padStart(2, '0'), { x: x + 0.18, y: y + 0.16, w: 0.52, h: 0.2, fontFace: FONTS.mono, fontSize: 7.2, color: hot ? ctx.theme.accentOn : s.fg, transparency: hot ? 0 : 35, margin: 0 });
    const tx = hasIcon ? x + 0.62 : x + 0.78;
    const textW = x + w - 0.27 - tx;
    const titleText = item.title || '';
    const bodyText = item.body || item.desc || '';
    const titleH = estimateTextHeight(titleText, textW, READABILITY.minFontSize, { min: 0.36, max: 0.68, lineHeight: 1.28, padding: 0.03 });
    const bodyY = y + 0.2 + titleH + 0.16;
    const bodyH = Math.max(0.48, h - (bodyY - y) - 0.28);
    slide.addText(titleText, { x: tx, y: y + 0.15, w: textW, h: titleH, fontFace: FONTS.sansZh, fontSize: READABILITY.minFontSize, bold: true, color: hot ? ctx.theme.accentOn : s.fg, margin: 0, fit: 'shrink', valign: 'top' });
    slide.addText(bodyText, { x: tx, y: bodyY, w: textW, h: bodyH, fontFace: FONTS.sansZh, fontSize: READABILITY.minFontSize, color: hot ? ctx.theme.accentOn : s.fg, transparency: hot ? 10 : 30, margin: 0.02, fit: 'shrink', valign: 'top' });
  });
  addFoot(slide, ctx, s.fg, 'swiss');
}
function swissSectionList(slide, ctx, s) {
  const data = ctx.slideSpec;
  addPageHead(slide, data, s.fg, 'swiss', 0.82);
  const sections = normalizeSections(data.sections || data.items || data.columns || []).slice(0, data.maxItems || 7);
  if (!sections.length) {
    addFoot(slide, ctx, s.fg, 'swiss');
    return;
  }
  const isCmb = ctx.spec.style === 'cmb';
  const y0 = data.subtitle ? 2.72 : 2.42;
  const bottom = 6.35;
  const gapY = 0.12;
  const rowMin = sections.length <= 4 ? 0.82 : 0.6;
  const textW = 9.22;
  const demands = sections.map((item) => {
    const title = item.title || item.label || item.name || '';
    const body = item.body || item.desc || item.note || item.summary || item.detail || item.text || '';
    const titleH = estimateTextHeight(title, textW, READABILITY.minFontSize, { min: 0.28, max: 0.46, lineHeight: 1.22, padding: 0.02 });
    const bodyH = estimateTextHeight(body, textW, READABILITY.minFontSize, { min: body ? 0.34 : 0, empty: 0, max: 0.82, lineHeight: 1.34, padding: 0.06 });
    return Math.max(rowMin, titleH + (body ? 0.08 + bodyH : 0) + 0.22);
  });
  const rowHeights = distributeRowHeights(demands, sections.length, 1, rowMin, bottom - y0, gapY);
  let y = y0;
  sections.forEach((item, i) => {
    const h = rowHeights[i];
    const hot = i === data.highlightIndex;
    const fill = hot ? ctx.theme.accent : ctx.theme.grey1;
    const color = hot ? ctx.theme.accentOn : s.fg;
    slide.addShape(pptx.ShapeType.rect, { x: 0.78, y, w: 11.45, h, fill: { color: fill, transparency: hot ? 0 : isCmb ? 18 : 0 }, line: { color: fill, transparency: 100 } });
    slide.addShape(pptx.ShapeType.rect, { x: 0.78, y, w: 0.1, h, fill: { color: hot ? ctx.theme.accentOn : ctx.theme.accent, transparency: hot ? 0 : 12 }, line: { color: hot ? ctx.theme.accentOn : ctx.theme.accent, transparency: 100 } });
    slide.addText(item.label || String(i + 1).padStart(2, '0'), { x: 1.04, y: y + 0.18, w: 0.54, h: 0.24, fontFace: FONTS.mono, fontSize: 9.4, bold: hot, color, transparency: hot ? 0 : 28, margin: 0 });
    const title = item.title || item.name || '';
    const body = item.body || item.desc || item.note || item.summary || item.detail || item.text || '';
    const titleH = estimateTextHeight(title, textW, READABILITY.minFontSize, { min: 0.3, max: 0.48, lineHeight: 1.22, padding: 0.02 });
    slide.addText(title, { x: 1.78, y: y + 0.14, w: textW, h: titleH, fontFace: FONTS.sansZh, fontSize: READABILITY.minFontSize, bold: true, color, margin: 0, valign: 'top' });
    if (body) slide.addText(body, { x: 1.78, y: y + 0.18 + titleH, w: textW, h: Math.max(0.3, h - titleH - 0.24), fontFace: FONTS.sansZh, fontSize: READABILITY.minFontSize, color, transparency: hot ? 8 : 28, margin: 0.02, valign: 'top' });
    y += h + gapY;
  });
  addFoot(slide, ctx, s.fg, 'swiss');
}

function swissDataSheet(slide, ctx, s) {
  const data = ctx.slideSpec;
  addPageHead(slide, data, s.fg, 'swiss', 0.82);
  addTableBlock(slide, ctx, data.table || data, { x: 0.78, y: 2.35, w: 8.15, h: 4.1 }, s, 'swiss');
  const notes = normalizeSections(data.notes || data.insights || []).slice(0, 4);
  notes.forEach((note, i) => {
    const y = 2.43 + i * 0.92;
    const hasIcon = addInlineIcon(slide, note, 9.42, y - 0.04, 0.28, i === 0 ? ctx.theme.accent : s.fg, 'swiss', { fallback: note.icon ? null : 'info', pad: 0.05 });
    if (!hasIcon) slide.addShape(pptx.ShapeType.rect, { x: 9.45, y, w: 0.16, h: 0.16, fill: { color: i === 0 ? ctx.theme.accent : s.fg }, line: { color: i === 0 ? ctx.theme.accent : s.fg, transparency: 100 } });
    slide.addText(note.title || note.label || '', { x: 9.8, y: y - 0.03, w: 2.2, h: 0.24, fontFace: FONTS.sansZh, fontSize: 10.2, bold: true, color: s.fg, margin: 0, fit: 'shrink' });
    slide.addText(note.body || note.desc || '', { x: 9.8, y: y + 0.28, w: 2.2, h: 0.36, fontFace: FONTS.sansZh, fontSize: 7.8, color: s.fg, transparency: 30, margin: 0, fit: 'shrink' });
  });
  addFoot(slide, ctx, s.fg, 'swiss');
}

function swissChart(slide, ctx, s) {
  const data = ctx.slideSpec;
  addPageHead(slide, data, s.fg, 'swiss', 0.82);
  addChartBlock(slide, ctx, data.chart || data, { x: 0.78, y: 2.35, w: 7.55, h: 4.05 }, s, 'swiss');
  const insights = normalizeSections(data.insights || data.notes || []).slice(0, 3);
  const panelX = 8.76;
  const iconX = 8.92;
  const textX = 9.34;
  const textW = 3.08;
  const startY = 2.42;
  const bottomY = 6.48;
  const gapY = 0.12;
  const demands = insights.map((item) => {
    const title = item.title || '';
    const body = item.body || item.desc || '';
    const titleH = estimateTextHeight(title, textW, READABILITY.minFontSize, { min: 0.28, max: 0.44, lineHeight: 1.24, padding: 0.02 });
    const bodyH = estimateTextHeight(body, textW, READABILITY.minFontSize, { min: body ? 0.48 : 0, empty: 0, max: 0.86, lineHeight: 1.34, padding: 0.08 });
    return titleH + (body ? 0.13 + bodyH : 0);
  });
  const rowHeights = insights.length
    ? distributeRowHeights(demands, insights.length, 1, 0.98, bottomY - startY, gapY)
    : [];
  let yCursor = startY;
  insights.forEach((item, i) => {
    const y = yCursor;
    const rowH = rowHeights[i] || 1.08;
    const iconColor = i === 0 ? ctx.theme.accent : s.fg;
    const hasIcon = addInlineIcon(slide, item, iconX, y + 0.04, 0.32, iconColor, 'swiss', { fallback: defaultContentIcon(i, 'swiss'), pad: 0.055 });
    if (!hasIcon) slide.addText(item.value || item.label || `0${i + 1}`, { x: iconX, y, w: 0.34, h: 0.38, fontFace: FONTS.sans, fontSize: 22, bold: true, color: i === 0 ? ctx.theme.accent : s.fg, margin: 0, fit: 'shrink' });
    const titleText = item.title || '';
    const bodyText = item.body || item.desc || '';
    const titleH = estimateTextHeight(titleText, textW, READABILITY.minFontSize, { min: 0.3, max: 0.44, lineHeight: 1.24, padding: 0.02 });
    const bodyY = y + titleH + 0.13;
    slide.addText(titleText, { x: textX, y: y + 0.02, w: textW, h: titleH, fontFace: FONTS.sansZh, fontSize: READABILITY.minFontSize, bold: true, color: s.fg, margin: 0, fit: 'shrink', valign: 'top' });
    slide.addText(bodyText, { x: textX, y: bodyY, w: textW, h: Math.max(0.48, rowH - titleH - 0.16), fontFace: FONTS.sansZh, fontSize: READABILITY.minFontSize, color: s.fg, transparency: 35, margin: 0.02, fit: 'shrink', valign: 'top' });
    yCursor += rowH + gapY;
  });
  addFoot(slide, ctx, s.fg, 'swiss');
}

function swissDashboard(slide, ctx, s) {
  const data = ctx.slideSpec;
  addPageHead(slide, data, s.fg, 'swiss', 0.72);
  const metricsY = Math.max(2.3, pageHeadSafeBottom(ctx, 0.72, 0.24));
  const chartY = Math.min(4.25, metricsY + 1.45);
  const chartH = Math.max(2.0, 6.25 - chartY);
  const metrics = (data.metrics || data.items || []).slice(0, 5);
  metrics.forEach((item, i) => {
    const x = 0.78 + i * 2.43;
    slide.addShape(pptx.ShapeType.rect, { x, y: metricsY, w: 2.05, h: 1.06, fill: { color: i === 0 ? ctx.theme.accent : ctx.theme.grey1 }, line: { color: i === 0 ? ctx.theme.accent : ctx.theme.grey1, transparency: 100 } });
    const iconColor = i === 0 ? ctx.theme.accentOn : s.fg;
    const hasIcon = addInlineIcon(slide, item, x + 0.16, metricsY + 0.14, 0.24, iconColor, 'swiss', { fallback: item.icon ? null : null, pad: 0.04 });
    const tx = hasIcon ? x + 0.48 : x + 0.16;
    slide.addText(item.label || '', { x: tx, y: metricsY + 0.18, w: x + 1.88 - tx, h: 0.18, fontFace: FONTS.mono, fontSize: 6.8, color: i === 0 ? ctx.theme.accentOn : s.fg, transparency: i === 0 ? 0 : 35, margin: 0, fit: 'shrink' });
    slide.addText(item.value || '', { x: x + 0.16, y: metricsY + 0.46, w: 1.72, h: 0.36, fontFace: FONTS.sans, fontSize: 19, bold: true, color: i === 0 ? ctx.theme.accentOn : s.fg, margin: 0, fit: 'shrink' });
  });
  const charts = data.charts || [];
  if (charts[0]) addChartBlock(slide, ctx, charts[0], { x: 0.78, y: chartY, w: 5.45, h: chartH }, s, 'swiss');
  if (charts[1]) addChartBlock(slide, ctx, charts[1], { x: 6.72, y: chartY, w: 5.45, h: chartH }, s, 'swiss');
  addFoot(slide, ctx, s.fg, 'swiss');
}
function swissAgenda(slide, ctx, s) {
  const data = ctx.slideSpec;
  addPageHead(slide, data, s.fg, 'swiss', 0.82);
  const items = normalizeSections(data.items || data.sections || data.agenda || []).slice(0, 8);
  const cols = clampColumns(data.columnsCount || autoColumns(items.length, 4), 1, 4);
  const rows = Math.max(1, Math.ceil(items.length / cols));
  const gap = 0.24;
  const cardW = (11.45 - gap * (cols - 1)) / cols;
  const y0 = data.subtitle ? 2.85 : 2.52;
  const rowGap = rows > 1 ? 0.24 : 0;
  const maxBottom = 6.42;
  const maxCardH = rows === 2 ? 1.82 : rows === 1 ? 2.15 : 1.46;
  const cardH = Math.max(1.12, Math.min(maxCardH, (maxBottom - y0 - rowGap * (rows - 1)) / rows));
  items.forEach((item, i) => {
    const x = 0.78 + (i % cols) * (cardW + gap);
    const y = y0 + Math.floor(i / cols) * (cardH + rowGap);
    const hot = i === data.highlightIndex;
    const body = item.body || item.desc || item.note || item.summary || item.detail || item.text || '';
    slide.addShape(pptx.ShapeType.rect, { x, y, w: cardW, h: cardH, fill: { color: hot ? ctx.theme.accent : ctx.theme.grey1 }, line: { color: hot ? ctx.theme.accent : ctx.theme.grey1, transparency: 100 } });
    const color = hot ? ctx.theme.accentOn : s.fg;
    slide.addText(item.label || String(i + 1).padStart(2, '0'), { x: x + 0.18, y: y + 0.16, w: 0.62, h: 0.22, fontFace: FONTS.mono, fontSize: 9.8, color, transparency: hot ? 0 : 35, margin: 0, fit: 'shrink' });
    slide.addText(item.title || item.label || `Item ${i + 1}`, { x: x + 0.18, y: y + 0.48, w: cardW - 0.36, h: 0.3, fontFace: FONTS.sansZh, fontSize: READABILITY.minFontSize, bold: true, color, margin: 0, fit: 'shrink' });
    if (body) {
      slide.addText(body, { x: x + 0.18, y: y + 0.86, w: cardW - 0.36, h: Math.max(0.28, cardH - 0.98), fontFace: FONTS.sansZh, fontSize: READABILITY.minFontSize, color, transparency: hot ? 8 : 30, margin: 0.02, valign: 'top' });
    }
  });
  addFoot(slide, ctx, s.fg, 'swiss');
}

function swissCaseStudy(slide, ctx, s) {
  const data = ctx.slideSpec;
  addMediaOrChart(slide, ctx, data, { x: 0, y: 0, w: SLIDE.w, h: 3.25 }, s, 'swiss', 'CASE');
  slide.addShape(pptx.ShapeType.rect, { x: 0.78, y: 0.92, w: 4.8, h: 1.24, fill: { color: ctx.theme.paper }, line: { color: ctx.theme.paper, transparency: 100 } });
  slide.addText(data.title || '', { x: 1.06, y: 1.15, w: 4.25, h: 0.55, fontFace: FONTS.sans, fontSize: 24, color: ctx.theme.ink, margin: 0, fit: 'shrink' });
  slide.addShape(pptx.ShapeType.rect, { x: 0, y: 3.25, w: SLIDE.w, h: 4.25, fill: { color: ctx.theme.paper }, line: { color: ctx.theme.paper, transparency: 100 } });
  slide.addText(data.caseTitle || data.label || 'Case', { x: 0.78, y: 3.85, w: 4.8, h: 0.42, fontFace: FONTS.sansZh, fontSize: 19, bold: true, color: ctx.theme.ink, margin: 0, fit: 'shrink' });
  slide.addText(data.body || data.story || data.subtitle || '', { x: 0.78, y: 4.42, w: 4.8, h: 0.78, fontFace: FONTS.sansZh, fontSize: 11, color: ctx.theme.ink, transparency: 15, margin: 0, fit: 'shrink' });
  const metrics = (data.metrics || data.items || []).slice(0, 3);
  metrics.forEach((item, i) => {
    const x = 6.25 + i * 2.0;
    slide.addShape(pptx.ShapeType.rect, { x, y: 3.9, w: 1.68, h: 1.45, fill: { color: i === 0 ? ctx.theme.accent : ctx.theme.grey1 }, line: { color: i === 0 ? ctx.theme.accent : ctx.theme.grey1, transparency: 100 } });
    slide.addText(item.value || item.title || '', { x: x + 0.15, y: 4.12, w: 1.34, h: 0.38, fontFace: FONTS.sans, fontSize: 22, bold: true, color: i === 0 ? ctx.theme.accentOn : s.fg, margin: 0, fit: 'shrink' });
    slide.addText(item.label || item.note || '', { x: x + 0.15, y: 4.67, w: 1.34, h: 0.38, fontFace: FONTS.sansZh, fontSize: 9.8, color: i === 0 ? ctx.theme.accentOn : s.fg, transparency: i === 0 ? 0 : 30, margin: 0, fit: 'shrink' });
  });
}

function swissPyramid(slide, ctx, s) {
  const data = ctx.slideSpec;
  addPageHead(slide, data, s.fg, 'swiss', 0.82);
  const layers = normalizeSections(data.layers || data.items || data.sections || []).slice(0, 5).reverse();
  const x0 = 2.05;
  const y0 = 5.75;
  layers.forEach((item, i) => {
    const w = 3.2 + i * 1.18;
    const x = x0 + (5.9 - w) / 2 + 1.8;
    const y = y0 - i * 0.72;
    const hot = i === layers.length - 1;
    slide.addShape(pptx.ShapeType.rect, { x, y, w, h: 0.58, fill: { color: hot ? ctx.theme.accent : ctx.theme.grey1 }, line: { color: hot ? ctx.theme.accent : ctx.theme.grey1, transparency: 100 } });
    slide.addText(item.title || item.label || '', { x: x + 0.18, y: y + 0.14, w: w - 0.36, h: 0.24, fontFace: FONTS.sansZh, fontSize: 12.2, bold: true, color: hot ? ctx.theme.accentOn : s.fg, align: 'center', margin: 0, fit: 'shrink' });
  });
  if (data.note || data.body) slide.addText(data.note || data.body, { x: 0.78, y: 5.75, w: 3.2, h: 0.5, fontFace: FONTS.sansZh, fontSize: 10.2, color: s.fg, transparency: 30, margin: 0, fit: 'shrink' });
  addFoot(slide, ctx, s.fg, 'swiss');
}

function swissRadial(slide, ctx, s) {
  const data = ctx.slideSpec;
  addPageHead(slide, data, s.fg, 'swiss', 0.78);
  const items = normalizeSections(data.items || data.nodes || data.sections || []).slice(0, 8);
  const cx = 6.62;
  const cy = 4.08;
  const centerBox = { x: cx - 1.08, y: cy - 0.52, w: 2.16, h: 1.04 };
  slide.addShape(pptx.ShapeType.rect, { ...centerBox, fill: { color: ctx.theme.accent }, line: { color: ctx.theme.accent, transparency: 100 } });
  slide.addText(data.center || data.label || data.title || '', { x: cx - 0.88, y: cy - 0.18, w: 1.76, h: 0.34, fontFace: FONTS.sansZh, fontSize: 13.8, bold: true, color: ctx.theme.accentOn, align: 'center', margin: 0, fit: 'shrink' });
  const nodeW = 2.35;
  const nodeH = 1.0;
  items.forEach((item, i) => {
    const angle = -Math.PI / 2 + (i * 2 * Math.PI) / Math.max(items.length, 1);
    const x = cx + Math.cos(angle) * 3.95;
    const y = cy + Math.sin(angle) * 1.58;
    const boxX = clamp(x - nodeW / 2, 0.68, 10.3);
    const boxY = clamp(y - nodeH / 2, 2.02, 5.48);
    const nodeBox = { x: boxX, y: boxY, w: nodeW, h: nodeH };
    addRadialConnector(slide, centerBox, nodeBox, { color: ctx.theme.grey3, transparency: 64, width: 0.45 });
    slide.addShape(pptx.ShapeType.rect, { ...nodeBox, fill: { color: ctx.theme.grey1 }, line: { color: ctx.theme.grey1, transparency: 100 } });
    slide.addText(item.title || item.label || '', { x: boxX + 0.16, y: boxY + 0.14, w: nodeW - 0.32, h: 0.28, fontFace: FONTS.sansZh, fontSize: 10.8, bold: true, color: s.fg, margin: 0, fit: 'shrink' });
    slide.addText(item.body || item.desc || '', { x: boxX + 0.16, y: boxY + 0.5, w: nodeW - 0.32, h: 0.38, fontFace: FONTS.sansZh, fontSize: 9.1, color: s.fg, transparency: 35, margin: 0.01, fit: 'shrink', valign: 'top' });
  });
  addFoot(slide, ctx, s.fg, 'swiss');
}
function swissRoadmap(slide, ctx, s) {
  const data = ctx.slideSpec;
  addPageHead(slide, data, s.fg, 'swiss', 0.82);
  const steps = normalizeSections(data.steps || data.items || []).slice(0, 6);
  const x0 = 0.88;
  const y0 = 3.9;
  const stepW = 11.25 / Math.max(steps.length, 1);
  slide.addShape(pptx.ShapeType.line, { x: x0, y: y0, w: 11.05, h: 0, line: { color: s.fg, transparency: 45, width: 0.8 } });
  steps.forEach((item, i) => {
    const x = x0 + i * stepW;
    const hot = i === data.highlightIndex || (data.highlightLast && i === steps.length - 1);
    slide.addShape(pptx.ShapeType.rect, { x: x - 0.06, y: y0 - 0.06, w: 0.12, h: 0.12, fill: { color: hot ? ctx.theme.accent : s.fg }, line: { color: hot ? ctx.theme.accent : s.fg, transparency: 100 } });
    const y = y0 + (i % 2 === 0 ? -1.02 : 0.35);
    const textW = steps.length <= 4 ? 2.3 : steps.length === 5 ? 1.96 : 1.78;
    const labelX = clamp(x - 0.48, 0.65, SLIDE.w - 1.6);
    const textX = clamp(x - textW / 2, 0.65, SLIDE.w - textW - 0.35);
    slide.addText(item.label || item.date || `0${i + 1}`, { x: labelX, y: y - 0.28, w: 0.95, h: 0.22, fontFace: FONTS.mono, fontSize: 9.8, color: hot ? ctx.theme.accent : s.fg, align: 'center', margin: 0, fit: 'shrink' });
    slide.addText(item.title || '', { x: textX, y, w: textW, h: 0.35, fontFace: FONTS.sansZh, fontSize: 11.4, bold: true, color: s.fg, align: 'center', margin: 0, fit: 'shrink' });
    slide.addText(item.body || item.desc || item.note || '', { x: textX, y: y + 0.44, w: textW, h: 0.48, fontFace: FONTS.sansZh, fontSize: 9.2, color: s.fg, transparency: 30, align: 'center', margin: 0, fit: 'shrink' });
  });
  addFoot(slide, ctx, s.fg, 'swiss');
}

function swissSwimlane(slide, ctx, s) {
  const data = ctx.slideSpec;
  addPageHead(slide, data, s.fg, 'swiss', 0.78);
  const lanes = (data.lanes || data.sections || []).slice(0, 4);
  const stages = data.stages || data.columns || ['现在', '下一步', '后续'];
  const x0 = 1.5;
  const y0 = 2.45;
  const laneH = 0.86;
  const colW = 10.3 / Math.max(stages.length, 1);
  stages.forEach((stage, i) => slide.addText(String(stage), { x: x0 + i * colW, y: 2.13, w: colW - 0.14, h: 0.22, fontFace: FONTS.mono, fontSize: 9.8, color: s.fg, transparency: 35, margin: 0, fit: 'shrink' }));
  lanes.forEach((lane, r) => {
    const y = y0 + r * 1.02;
    slide.addText(lane.title || lane.label || `Lane ${r + 1}`, { x: 0.78, y: y + 0.22, w: 0.58, h: 0.32, fontFace: FONTS.sansZh, fontSize: 10.6, bold: true, color: s.fg, margin: 0, fit: 'shrink' });
    const cells = lane.items || lane.steps || [];
    stages.forEach((stage, c) => {
      const cell = cells[c] || {};
      const x = x0 + c * colW;
      const hot = r === data.highlightRow || c === data.highlightColumn;
      slide.addShape(pptx.ShapeType.rect, { x, y, w: colW - 0.16, h: laneH, fill: { color: hot ? ctx.theme.accent : ctx.theme.grey1 }, line: { color: hot ? ctx.theme.accent : ctx.theme.grey1, transparency: 100 } });
      slide.addText(cellText(cell), { x: x + 0.12, y: y + 0.16, w: colW - 0.4, h: 0.46, fontFace: FONTS.sansZh, fontSize: 9.8, color: hot ? ctx.theme.accentOn : s.fg, margin: 0, fit: 'shrink', valign: 'mid' });
    });
  });
  addFoot(slide, ctx, s.fg, 'swiss');
}
function swissSectionCompat(slide, ctx, s) {
  const data = ctx.slideSpec;
  swissCover(slide, { ...ctx, slideSpec: { ...data, subtitle: data.subtitle || data.body || '' } }, s);
}

function swissBigQuoteCompat(slide, ctx, s) {
  const data = ctx.slideSpec;
  swissStatement(slide, { ...ctx, slideSpec: { ...data, title: data.quote || data.title, body: data.body || data.cite || data.subtitle || '' } }, s);
}

function swissQuoteImageCompat(slide, ctx, s) {
  const data = ctx.slideSpec;
  swissImageHero(slide, { ...ctx, slideSpec: { ...data, body: data.body || data.quote || data.callout || data.subtitle || '' } }, s);
}

function swissTextImageCompat(slide, ctx, s) {
  const data = ctx.slideSpec;
  if (data.image) {
    swissImageHero(slide, { ...ctx, slideSpec: { ...data, body: data.body || data.callout || data.subtitle || '' } }, s);
    return;
  }
  swissStatement(slide, ctx, s);
}

function swissImageGridCompat(slide, ctx, s) {
  swissMediaGrid(slide, ctx, s);
}

function swissMedia(slide, ctx, s) {
  const data = ctx.slideSpec;
  addPageHead(slide, data, s.fg, 'swiss', 0.82);
  const isCmb = ctx.spec.style === 'cmb' || data.style === 'cmb';
  const headY = pageHeadY(ctx, 0.82);
  const hasSubtitle = !!data.subtitle;
  const contentTop = Math.max(2.36, headY + (hasSubtitle ? 1.86 : 1.54));
  const contentBottom = 6.35;
  const mediaBox = { x: 0.78, y: contentTop, w: 6.25, h: Math.max(2.55, contentBottom - contentTop) };
  addMediaOrChart(slide, ctx, data, mediaBox, s, 'swiss', 'MEDIA');
  const panelX = 7.42;
  const summaryH = isCmb ? 0.78 : 0.72;
  slide.addText(data.body || data.story || data.note || '', { x: panelX, y: contentTop + 0.05, w: 4.65, h: summaryH, fontFace: FONTS.sansZh, fontSize: isCmb ? 12 : 11.7, color: s.fg, transparency: 15, margin: 0.03, fit: 'shrink', valign: 'top' });
  const sideLimit = isCmb ? 4 : 5;
  const items = normalizeSections(data.items || data.insights || data.points || []).slice(0, sideLimit);
  const itemStartY = contentTop + summaryH + (isCmb ? 0.2 : 0.18);
  const rowGap = isCmb ? 0.1 : 0.08;
  const rowMin = isCmb ? 0.9 : 0.78;
  const textWForEstimate = 4.05;
  const demands = items.map((item) => {
    const title = item.title || item.label || item.body || '';
    const body = item.body || item.desc || item.note || '';
    const titleH = estimateTextHeight(title, textWForEstimate, READABILITY.minFontSize, { min: 0.3, max: 0.44, lineHeight: 1.24, padding: 0.02 });
    const bodyH = estimateTextHeight(body, textWForEstimate, READABILITY.minFontSize, { min: body ? 0.46 : 0, empty: 0, max: 0.72, lineHeight: 1.32, padding: 0.08 });
    return titleH + (body ? 0.12 + bodyH : 0);
  });
  const rowHeights = items.length
    ? distributeRowHeights(demands, items.length, 1, rowMin, contentBottom - itemStartY, rowGap)
    : [];
  let yCursor = itemStartY;
  items.forEach((item, i) => {
    const y = yCursor;
    const rowH = rowHeights[i] || rowMin;
    slide.addShape(pptx.ShapeType.line, { x: panelX, y, w: 4.4, h: 0, line: { color: s.fg, transparency: isCmb ? 74 : 68, width: 0.5 } });
    const hasIcon = addInlineIcon(slide, item, panelX, y + (isCmb ? 0.12 : 0.1), 0.21, i === 0 ? ctx.theme.accent : s.fg, 'swiss', { fallback: defaultContentIcon(i, 'swiss'), pad: 0.04 });
    const textX = hasIcon ? panelX + 0.34 : panelX;
    const textW = hasIcon ? 4.05 : 4.4;
    const titleText = item.title || item.label || item.body || '';
    const bodyText = item.body || item.desc || item.note || '';
    const titleH = estimateTextHeight(titleText, textW, READABILITY.minFontSize, { min: 0.3, max: 0.44, lineHeight: 1.24, padding: 0.02 });
    const bodyY = y + 0.08 + titleH + 0.12;
    slide.addText(titleText, { x: textX, y: y + 0.08, w: textW, h: titleH, fontFace: FONTS.sansZh, fontSize: READABILITY.minFontSize, bold: true, color: s.fg, margin: 0, fit: 'shrink', valign: 'top' });
    slide.addText(bodyText, { x: textX, y: bodyY, w: textW, h: Math.max(isCmb ? 0.46 : 0.4, rowH - titleH - 0.2), fontFace: FONTS.sansZh, fontSize: READABILITY.minFontSize, color: s.fg, transparency: 32, margin: 0.02, fit: 'shrink', valign: 'top' });
    yCursor += rowH + rowGap;
  });
  addFoot(slide, ctx, s.fg, 'swiss');
}
function swissMediaGrid(slide, ctx, s) {
  const data = ctx.slideSpec;
  addPageHead(slide, data, s.fg, 'swiss', 0.82);
  const count = resolveMediaSlotCount(data);
  const cols = count <= 2 ? count : 3;
  const rows = Math.ceil(count / cols);
  const gapX = 0.24;
  const gapY = rows > 1 ? 0.52 : 0.3;
  const gridW = 11.45;
  const w = (gridW - gapX * (cols - 1)) / cols;
  const h = rows > 1 ? 1.5 : 2.45;
  const x0 = 0.78;
  const y0 = rows > 1 ? 2.58 : 3.0;
  const boxes = Array.from({ length: count }, (_, i) => ({ x: x0 + (i % cols) * (w + gapX), y: y0 + Math.floor(i / cols) * (h + gapY), w, h }));
  addMediaGrid(slide, ctx, data, boxes, s, 'swiss');
  addFoot(slide, ctx, s.fg, 'swiss');
}
function swissClosing(slide, ctx, s) {
  const data = ctx.slideSpec;
  swissCover(slide, { ...ctx, slideSpec: { ...data, title: data.title || '谢谢', subtitle: data.subtitle || data.body || '' } }, s);
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

function iconAlias(icon) {
  const raw = String(icon || 'dot').trim();
  const withoutPrefix = raw.startsWith('lucide:') ? raw.slice(7) : raw;
  return ICON_ALIASES[withoutPrefix] || withoutPrefix;
}

function toKebabIconName(icon) {
  return String(iconAlias(icon) || '')
    .replace(/([a-z0-9])([A-Z])/g, '$1-$2')
    .replace(/[\s_]+/g, '-')
    .replace(/[^a-zA-Z0-9-]/g, '')
    .toLowerCase();
}

function loadLucideModule() {
  if (LUCIDE_MODULE !== undefined) return LUCIDE_MODULE;
  const candidates = ['lucide', 'lucide/dist/cjs/lucide.js', 'lucide/dist/cjs/lucide.cjs'];
  for (const name of candidates) {
    try {
      const mod = require(name);
      LUCIDE_MODULE = mod?.default || mod;
      return LUCIDE_MODULE;
    } catch (_) {
      // Try the next known package entry.
    }
  }
  LUCIDE_MODULE = null;
  return LUCIDE_MODULE;
}

function toPascalIconName(name) {
  return String(name || '')
    .split('-')
    .filter(Boolean)
    .map((part) => part.charAt(0).toUpperCase() + part.slice(1))
    .join('');
}

function xmlAttrName(name) {
  return String(name || '').replace(/[A-Z]/g, (m) => `-${m.toLowerCase()}`);
}

function xmlAttrValue(value, stroke) {
  const raw = value === 'currentColor' ? `#${stroke}` : String(value ?? '');
  return raw
    .replace(/&/g, '&amp;')
    .replace(/"/g, '&quot;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;');
}

function iconNodeBody(iconNode, stroke) {
  if (!Array.isArray(iconNode)) return '';
  return iconNode.map(([tag, attrs = {}, children = []]) => {
    const attrText = Object.entries(attrs)
      .filter(([key, value]) => key !== 'key' && value !== undefined && value !== null)
      .map(([key, value]) => `${xmlAttrName(key)}="${xmlAttrValue(value, stroke)}"`)
      .join(' ');
    const childText = iconNodeBody(children, stroke);
    return childText
      ? `<${tag}${attrText ? ` ${attrText}` : ''}>${childText}</${tag}>`
      : `<${tag}${attrText ? ` ${attrText}` : ''}/>`;
  }).join('');
}

function iconNodeToSvg(iconNode, stroke) {
  const body = iconNodeBody(iconNode, stroke);
  if (!body) return null;
  return `<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 24 24" fill="none" stroke="#${stroke}" stroke-width="2" stroke-linecap="round" stroke-linejoin="round">${body}</svg>`;
}

function lucidePackageIconSvg(name, stroke) {
  const lucide = loadLucideModule();
  if (!lucide) return null;
  const pascal = toPascalIconName(name);
  const camel = pascal ? pascal.charAt(0).toLowerCase() + pascal.slice(1) : '';
  const iconSet = lucide.icons || lucide.default?.icons || lucide.default || lucide;
  const icon = iconSet?.[pascal] || iconSet?.[camel] || iconSet?.[name] || lucide[pascal] || lucide[camel] || lucide[name];
  const iconNode = icon?.iconNode || icon;
  return iconNodeToSvg(iconNode, stroke);
}

function lucideStaticIconSvg(name, stroke) {
  const file = path.join(LUCIDE_STATIC_ICON_DIR, `${name}.svg`);
  if (!fs.existsSync(file)) return null;
  let svg = fs.readFileSync(file, 'utf8')
    .replace(/<!--[^]*?-->/g, '')
    .replace(/\s(width|height)="24"/g, '')
    .replace(/stroke="currentColor"/g, `stroke="#${stroke}"`)
    .replace(/stroke="#[0-9a-fA-F]{3,6}"/g, `stroke="#${stroke}"`)
    .replace(/fill="currentColor"/g, `fill="#${stroke}"`)
    .replace(/class="[^"]*"/g, '')
    .trim();
  if (!/xmlns=/.test(svg)) svg = svg.replace('<svg', '<svg xmlns="http://www.w3.org/2000/svg"');
  return svg;
}

function lucideIconSvg(icon, color) {
  const name = toKebabIconName(icon);
  if (!name) return null;
  const stroke = normalizeHex(color || '111111');
  const cacheKey = `${name}:${stroke}`;
  if (LUCIDE_ICON_CACHE.has(cacheKey)) return LUCIDE_ICON_CACHE.get(cacheKey);
  const svg = lucidePackageIconSvg(name, stroke) || lucideStaticIconSvg(name, stroke);
  if (!svg) return null;
  const data = `data:image/svg+xml;base64,${Buffer.from(svg).toString('base64')}`;
  LUCIDE_ICON_CACHE.set(cacheKey, data);
  return data;
}
function svgIconBody(name) {
  const icon = iconAlias(name);
  const bodies = {
    checkCircle: '<circle cx="12" cy="12" r="9"/><path d="m8 12 2.6 2.6L16.5 9"/>',
    alertCircle: '<circle cx="12" cy="12" r="9"/><path d="M12 7v6"/><path d="M12 17h.01"/>',
    infoCircle: '<circle cx="12" cy="12" r="9"/><path d="M12 11v6"/><path d="M12 7h.01"/>',
    arrowRight: '<path d="M5 12h14"/><path d="m13 6 6 6-6 6"/>',
    chartBar: '<path d="M4 20V10"/><path d="M10 20V4"/><path d="M16 20v-7"/><path d="M22 20H2"/>',
    chartLine: '<path d="M3 17 8 12l4 3 7-8"/><path d="M21 20H3V4"/>',
    pieChart: '<path d="M12 3v9h9"/><path d="M21 12a9 9 0 1 1-9-9"/>',
    target: '<circle cx="12" cy="12" r="9"/><circle cx="12" cy="12" r="5"/><circle cx="12" cy="12" r="1.4"/>',
    star: '<path d="m12 3 2.7 5.5 6.1.9-4.4 4.3 1 6.1L12 17l-5.4 2.8 1-6.1-4.4-4.3 6.1-.9z"/>',
  };
  return bodies[icon] || bodies[iconAlias(icon)] || bodies.checkCircle;
}

function fallbackSvgIconData(icon, color) {
  const stroke = normalizeHex(color || '111111');
  const body = svgIconBody(icon);
  const svg = `<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 24 24"><g fill="none" stroke="#${stroke}" stroke-width="1.9" stroke-linecap="round" stroke-linejoin="round">${body}</g></svg>`;
  return `data:image/svg+xml;base64,${Buffer.from(svg).toString('base64')}`;
}

function svgIconData(icon, color) {
  return lucideIconSvg(icon, color) || fallbackSvgIconData(icon, color);
}

function iconCacheKey(icon, color) {
  return `${toKebabIconName(icon) || iconAlias(icon)}:${normalizeHex(color || '111111')}`;
}

function loadSharpModule() {
  if (SHARP_MODULE !== undefined) return SHARP_MODULE;
  try {
    SHARP_MODULE = require('sharp');
  } catch (_) {
    SHARP_MODULE = null;
  }
  return SHARP_MODULE;
}

function requireSharpForPng(scope) {
  const sharp = loadSharpModule();
  if (!sharp) {
    fail(scope + ' require sharp to be installed because PNG rasterization is the default. Run npm install in the skill directory, or explicitly set iconMode:"svg" / svgImageMode:"svg" if SVG output is intentional.');
  }
  return sharp;
}

function iconSvgBuffer(icon, color) {
  const data = svgIconData(icon, color);
  const payload = String(data || '').split(',')[1];
  return payload ? Buffer.from(payload, 'base64') : null;
}

function collectIconInputs(value, icons = new Set(), colors = new Set()) {
  if (!value || typeof value !== 'object') return { icons, colors };
  if (Array.isArray(value)) {
    value.forEach((item) => collectIconInputs(item, icons, colors));
    return { icons, colors };
  }
  if (value.icon || value.bulletIcon) icons.add(iconAlias(value.icon || value.bulletIcon));
  if (value.iconColor) colors.add(normalizeHex(value.iconColor));
  if (value.color) colors.add(normalizeHex(value.color));
  Object.values(value).forEach((item) => collectIconInputs(item, icons, colors));
  return { icons, colors };
}

function themeIconColors(theme) {
  return Object.values(theme || {})
    .filter((value) => typeof value === 'string' && /^[0-9a-fA-F]{6}$/.test(value))
    .map(normalizeHex);
}

function defaultContentIconNames() {
  return ['file-text', 'scan-search', 'shield-alert', 'arrow-right-circle', 'lightbulb', 'target', 'chart-line', 'circle-alert', 'workflow', 'users', 'database', 'settings', 'layers', 'info'];
}

async function prepareIconAssets(spec, theme) {
  ICON_RENDER_MODE = String(spec.iconMode || 'png').toLowerCase();
  if (ICON_RENDER_MODE === 'svg') return;
  requireSharpForPng('Lucide icons');
  const { icons, colors } = collectIconInputs(spec);
  [...defaultContentIconNames(), ...Object.values(ICON_ALIASES)].forEach((icon) => icons.add(icon));
  ['111111', 'FFFFFF', ...themeIconColors(theme), ...colors].forEach((color) => colors.add(normalizeHex(color)));
  const jobs = [];
  for (const icon of icons) {
    for (const color of colors) jobs.push(renderIconPngToCache(icon, color));
  }
  await Promise.all(jobs);
}

async function renderIconPngToCache(icon, color) {
  const key = iconCacheKey(icon, color);
  if (ICON_PNG_CACHE.has(key)) return;
  const sharp = requireSharpForPng('Lucide icons');
  const svgBuffer = iconSvgBuffer(icon, color);
  if (!svgBuffer) fail('Failed to build SVG source for icon: ' + icon);
  try {
    const buffer = await sharp(svgBuffer).resize(192, 192, { fit: 'contain' }).png().toBuffer();
    ICON_PNG_CACHE.set(key, `data:image/png;base64,${buffer.toString('base64')}`);
  } catch (error) {
    fail(`Failed to rasterize Lucide icon "${icon}" to PNG. ${error.message}`);
  }
}

function iconImageData(icon, color) {
  if (ICON_RENDER_MODE !== 'svg') {
    const data = ICON_PNG_CACHE.get(iconCacheKey(icon, color));
    if (data) return data;
    requireSharpForPng('Lucide icons');
    fail('Icon was not rasterized before insertion: ' + icon);
  }
  return svgIconData(icon, color);
}

function normalizeHex(color) {
  const raw = String(color || '').replace('#', '').trim();
  return /^[0-9a-fA-F]{6}$/.test(raw) ? raw.toUpperCase() : '111111';
}

function addSvgIcon(slide, icon, x, y, size, color, options = {}) {
  const bg = options.bg;
  const bgTransparency = options.bgTransparency ?? 100;
  if (bg && bgTransparency < 100) {
    slide.addShape(options.bgShape === 'circle' ? pptx.ShapeType.ellipse : pptx.ShapeType.rect, {
      x,
      y,
      w: size,
      h: size,
      fill: { color: bg, transparency: bgTransparency },
      line: { color: bg, transparency: 100 },
    });
  }
  const pad = options.pad ?? 0;
  slide.addImage({ data: iconImageData(icon, color), x: x + pad, y: y + pad, w: Math.max(0.01, size - pad * 2), h: Math.max(0.01, size - pad * 2) });
}
function defaultContentIcon(index, mode) {
  const magazine = ['file-text', 'scan-search', 'shield-alert', 'arrow-right-circle', 'lightbulb', 'target'];
  const swiss = ['chart-line', 'circle-alert', 'workflow', 'shield-alert', 'target', 'users', 'database', 'settings', 'layers'];
  const list = mode === 'swiss' ? swiss : magazine;
  return list[index % list.length];
}
function itemIcon(item, fallback = null) {
  if (!item || typeof item === 'string') return fallback;
  return item.icon || item.bulletIcon || fallback;
}

function addInlineIcon(slide, item, x, y, size, color, mode, options = {}) {
  const icon = itemIcon(item, options.fallback);
  if (!icon) return false;
  const bg = options.bg;
  const bgTransparency = options.bgTransparency ?? (mode === 'swiss' ? 100 : 88);
  addSvgIcon(slide, icon, x, y, size, item.iconColor || color, { bg, bgTransparency, bgShape: options.bgShape, pad: options.pad ?? size * 0.18 });
  return true;
}
function addBulletIcon(slide, icon, x, y, size, color, fill, transparency, mode, number) {
  const rawName = String(icon || 'dot').trim();
  const name = rawName || 'dot';
  if (!BASIC_ICON_NAMES.includes(name)) {
    addSvgIcon(slide, iconAlias(name), x - size * 0.2, y - size * 0.2, size * 1.42, color, { pad: size * 0.1 });
    return;
  }
  const stroke = { color, transparency: transparency ?? 0, width: mode === 'swiss' ? 1.0 : 0.85 };
  const solid = { color: fill || color, transparency: transparency ?? 0 };
  const cx = x + size / 2;
  const cy = y + size / 2;
  if (name === 'dot') {
    slide.addShape(pptx.ShapeType.ellipse, { x, y, w: size, h: size, fill: solid, line: { color, transparency: 100 } });
  } else if (name === 'square') {
    slide.addShape(pptx.ShapeType.rect, { x, y, w: size, h: size, fill: solid, line: { color, transparency: 100 } });
  } else if (name === 'diamond') {
    slide.addShape(pptx.ShapeType.diamond, { x, y, w: size, h: size, fill: solid, line: { color, transparency: 100 } });
  } else if (name === 'target') {
    slide.addShape(pptx.ShapeType.ellipse, { x, y, w: size, h: size, fill: { color: 'FFFFFF', transparency: 100 }, line: stroke });
    slide.addShape(pptx.ShapeType.ellipse, { x: x + size * 0.32, y: y + size * 0.32, w: size * 0.36, h: size * 0.36, fill: solid, line: { color, transparency: 100 } });
  } else if (name === 'line') {
    slide.addShape(pptx.ShapeType.line, { x, y: cy, w: size, h: 0, line: stroke });
  } else if (name === 'plus') {
    slide.addShape(pptx.ShapeType.line, { x, y: cy, w: size, h: 0, line: stroke });
    slide.addShape(pptx.ShapeType.line, { x: cx, y, w: 0, h: size, line: stroke });
  } else if (name === 'minus') {
    slide.addShape(pptx.ShapeType.line, { x, y: cy, w: size, h: 0, line: stroke });
  } else if (name === 'arrow') {
    slide.addShape(pptx.ShapeType.line, { x, y: cy, w: size * 0.72, h: 0, line: stroke });
    slide.addShape(pptx.ShapeType.triangle, { x: x + size * 0.54, y: y + size * 0.18, w: size * 0.45, h: size * 0.64, rotate: 90, fill: solid, line: { color, transparency: 100 } });
  } else if (name === 'check') {
    slide.addText('✓', { x: x - size * 0.1, y: y - size * 0.45, w: size * 1.5, h: size * 1.5, fontFace: 'Segoe UI Symbol', fontSize: size * 72, color, transparency: transparency ?? 0, bold: true, margin: 0, fit: 'shrink', align: 'center', valign: 'mid' });
  } else if (name === 'cross') {
    slide.addShape(pptx.ShapeType.line, { x, y, w: size, h: size, line: stroke });
    slide.addShape(pptx.ShapeType.line, { x, y: y + size, w: size, h: -size, line: stroke });
  } else if (name === 'alert') {
    slide.addShape(pptx.ShapeType.triangle, { x, y: y - size * 0.05, w: size, h: size * 1.1, fill: { color: 'FFFFFF', transparency: 100 }, line: stroke });
    slide.addText('!', { x, y: y - size * 0.24, w: size, h: size, fontFace: FONTS.sans, fontSize: size * 48, bold: true, color, transparency: transparency ?? 0, margin: 0, fit: 'shrink', align: 'center', valign: 'mid' });
  } else if (name === 'info') {
    slide.addShape(pptx.ShapeType.ellipse, { x, y, w: size, h: size, fill: { color: 'FFFFFF', transparency: 100 }, line: stroke });
    slide.addText('i', { x, y: y - size * 0.14, w: size, h: size, fontFace: FONTS.sans, fontSize: size * 45, bold: true, color, transparency: transparency ?? 0, margin: 0, fit: 'shrink', align: 'center', valign: 'mid' });
  } else if (name === 'star') {
    slide.addShape(pptx.ShapeType.star5, { x, y, w: size, h: size, fill: solid, line: { color, transparency: 100 } });
  } else if (name === 'number') {
    slide.addText(String(number), { x, y: y - size * 0.07, w: size, h: size, fontFace: mode === 'swiss' ? FONTS.mono : FONTS.mono, fontSize: size * 38, bold: true, color, transparency: transparency ?? 0, margin: 0, fit: 'shrink', align: 'center', valign: 'mid' });
  }
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

function renderDataBlocks(slide, ctx, s, mode) {
  const layout = ctx.slideSpec.layout || '';
  const chartLayouts = ['chart', 'dashboard', 'media', 'mediaGrid', 'gallery', 'imageGrid', 'imageHero', 'quoteImage', 'textImage', 'caseStudy'];
  const blocks = [...(ctx.slideSpec.blocks || [])];
  if (!chartLayouts.includes(layout)) {
    blocks.push(...(ctx.slideSpec.charts || []).map((chart) => ({ ...chart, type: 'chart' })));
  }
  blocks.push(...(ctx.slideSpec.tables || []).map((table) => ({ ...table, type: 'table' })));
  blocks.forEach((block) => {
    if (!block || block.enabled === false) return;
    if ((block.type === 'chart' || block.type === 'table') && !hasExplicitBox(block)) {
      fail(`slide ${ctx.index + 1} has unpositioned ${block.type} block. Put it in layout "media"/"mediaGrid", remove it, or set x/y/w/h explicitly so it is not silently skipped.`);
      return;
    }
    if (block.type === 'chart') addChartBlock(slide, ctx, block, block, s, mode);
    else if (block.type === 'table') addTableBlock(slide, ctx, block, block, s, mode);
    else if (block.type === 'text') addTextBlock(slide, block, s, mode);
    else if (block.type === 'callout') addCallout(slide, block.text || block.body || block.title || '', block.x || 0.8, block.y || 5.5, block.w || 4, block.h || 0.75, block.color || s.fg, block.fill || ctx.theme.paperTint || ctx.theme.grey1);
  });
}

function hasExplicitBox(block) {
  return ['x', 'y', 'w', 'h'].every((key) => Number.isFinite(Number(block[key])));
}

function addTextBlock(slide, block, s, mode) {
  const box = safeBox(block, { x: 0.8, y: 5.4, w: 4, h: 0.9 });
  slide.addText(block.text || block.body || '', { x: box.x, y: box.y, w: box.w, h: box.h, fontFace: mode === 'swiss' ? FONTS.sansZh : FONTS.sansZh, fontSize: block.fontSize || 9.5, bold: !!block.bold, color: block.color || s.fg, transparency: block.transparency ?? 15, margin: 0.03, fit: 'shrink', valign: 'top' });
}

function addChartBlock(slide, ctx, chart, defaults, s, mode) {
  if (!chart) return;
  const box = safeBox(chart, defaults || { x: 0.8, y: 2.4, w: 7, h: 3.5 });
  const typeName = String(chart.chartType || chart.kind || chart.type || 'bar');
  const chartType = CHART_TYPES[typeName] || CHART_TYPES.bar;
  const data = normalizeChartData(chart);
  if (!data.length) return;
  const palette = chart.colors || chart.chartColors || chartPalette(ctx, mode);
  const axisColor = mode === 'swiss' ? ctx.theme.grey3 : s.fg;
  const common = {
    x: box.x,
    y: box.y,
    w: box.w,
    h: box.h,
    chartColors: palette,
    showLegend: chart.showLegend ?? data.length > 1,
    legendPos: chart.legendPos || 'b',
    legendFontFace: mode === 'swiss' ? FONTS.sans : FONTS.sansZh,
    legendFontSize: Math.max(Number(chart.legendFontSize) || READABILITY.minChartFontSize, READABILITY.minChartFontSize),
    legendColor: s.fg,
    showTitle: !!chart.title,
    title: chart.title,
    titleFontFace: mode === 'swiss' ? FONTS.sansZh : FONTS.serifZh,
    titleFontSize: Math.max(Number(chart.titleFontSize) || READABILITY.minChartFontSize, READABILITY.minChartFontSize),
    titleColor: s.fg,
    showValue: chart.showValue ?? ['pie', 'doughnut'].includes(typeName),
    showPercent: chart.showPercent ?? ['pie', 'doughnut'].includes(typeName),
    dataLabelColor: s.fg,
    dataLabelFontFace: mode === 'swiss' ? FONTS.sans : FONTS.sansZh,
    dataLabelFontSize: Math.max(Number(chart.dataLabelFontSize) || READABILITY.minChartFontSize, READABILITY.minChartFontSize),
    catAxisLabelFontFace: mode === 'swiss' ? FONTS.sans : FONTS.sansZh,
    catAxisLabelFontSize: Math.max(Number(chart.axisFontSize) || READABILITY.minChartFontSize, READABILITY.minChartFontSize),
    catAxisLabelColor: axisColor,
    catAxisLineColor: axisColor,
    valAxisLabelFontFace: mode === 'swiss' ? FONTS.sans : FONTS.sansZh,
    valAxisLabelFontSize: Math.max(Number(chart.axisFontSize) || READABILITY.minChartFontSize, READABILITY.minChartFontSize),
    valAxisLabelColor: axisColor,
    valAxisLineColor: axisColor,
    valGridLine: { color: mode === 'swiss' ? ctx.theme.grey2 : s.fg, transparency: mode === 'swiss' ? 50 : 75, size: 0.4 },
    chartArea: { fill: { color: chart.fill || 'FFFFFF', transparency: chart.fillTransparency ?? 100 }, border: { color: chart.borderColor || s.fg, transparency: chart.borderTransparency ?? 100, pt: 0 } },
    plotArea: { fill: { color: chart.plotFill || 'FFFFFF', transparency: chart.plotTransparency ?? 100 }, border: { color: chart.borderColor || s.fg, transparency: 100, pt: 0 } },
  };
  if (typeName === 'column') common.barDir = 'col';
  if (typeName === 'bar') common.barDir = chart.barDir || 'bar';
  if (typeName === 'line') {
    common.lineSize = chart.lineSize || 2.2;
    common.lineDataSymbol = chart.lineDataSymbol || 'circle';
    common.lineDataSymbolSize = chart.lineDataSymbolSize || 4;
  }
  slide.addChart(chartType, data, { ...common, ...(chart.options || {}) });
  if (chart.caption) addCaption(slide, chart.caption, box.x, box.y + box.h + 0.08, box.w, s.fg, mode);
}

function addTableBlock(slide, ctx, table, defaults, s, mode) {
  if (!table) return;
  const box = safeBox(table, defaults || { x: 0.8, y: 2.4, w: 7, h: 3.5 });
  const rows = normalizeTableRows(table);
  if (!rows.length) return;
  const headerFill = mode === 'swiss' ? ctx.theme.accent : s.fg;
  const headerColor = mode === 'swiss' ? ctx.theme.accentOn : ctx.theme.paper;
  const bodyFill = mode === 'swiss' ? ctx.theme.grey1 : ctx.theme.paperTint;
  const borderColor = mode === 'swiss' ? ctx.theme.grey2 : s.fg;
  const styled = rows.map((row, rowIndex) => row.map((cell) => ({
    text: String(cell ?? ''),
    options: {
      bold: rowIndex === 0,
      color: rowIndex === 0 ? headerColor : s.fg,
      fill: { color: rowIndex === 0 ? headerFill : bodyFill, transparency: rowIndex === 0 ? 0 : mode === 'swiss' ? 0 : 20 },
      margin: 0.05,
      valign: 'mid',
      fontFace: rowIndex === 0 ? (mode === 'swiss' ? FONTS.sans : FONTS.mono) : FONTS.sansZh,
      fontSize: rowIndex === 0 ? Math.max(Number(table.headerFontSize) || READABILITY.minTableFontSize, READABILITY.minTableFontSize) : Math.max(Number(table.fontSize) || READABILITY.minTableFontSize, READABILITY.minTableFontSize),
    },
  })));
  slide.addTable(styled, {
    x: box.x,
    y: box.y,
    w: box.w,
    h: box.h,
    border: { type: 'solid', color: borderColor, transparency: mode === 'swiss' ? 35 : 60, pt: 0.45 },
    margin: 0.04,
    color: s.fg,
    fontFace: FONTS.sansZh,
    fontSize: Math.max(Number(table.fontSize) || READABILITY.minTableFontSize, READABILITY.minTableFontSize),
    fit: 'shrink',
  });
  if (table.caption) addCaption(slide, table.caption, box.x, box.y + box.h + 0.08, box.w, s.fg, mode);
}

function normalizeChartData(chart) {
  if (Array.isArray(chart.series)) {
    return chart.series.map((series) => ({
      name: series.name || series.label || chart.name || 'Series',
      labels: series.labels || chart.labels || chart.categories || [],
      values: (series.values || series.data || []).map(Number),
    })).filter((series) => series.labels.length && series.values.length);
  }
  if (Array.isArray(chart.values || chart.data)) {
    return [{
      name: chart.name || chart.title || 'Series',
      labels: chart.labels || chart.categories || [],
      values: (chart.values || chart.data || []).map(Number),
    }].filter((series) => series.labels.length && series.values.length);
  }
  return [];
}

function normalizeTableRows(table) {
  if (Array.isArray(table.rows) && table.rows.length) {
    return table.headers ? [table.headers, ...table.rows] : table.rows;
  }
  if (Array.isArray(table.data) && table.data.length) {
    return table.headers ? [table.headers, ...table.data] : table.data;
  }
  return [];
}

function normalizeLayoutCompatibility(spec) {
  spec.slides.forEach((slide, index) => {
    normalizeCompareSlide(slide, index);
  });
}

function normalizeCompareSlide(slide, index) {
  const layout = slide.layout || '';
  if (!['compare', 'duoCompare', 'splitCompare'].includes(layout)) return;
  const before = normalizeCompareColumn(slide.before, slide.left, slide.leftTitle || slide.beforeTitle, slide.leftLabel || slide.beforeLabel, 'left/before', index);
  const after = normalizeCompareColumn(slide.after, slide.right, slide.rightTitle || slide.afterTitle, slide.rightLabel || slide.afterLabel, 'right/after', index);
  if (before) slide.before = before;
  if (after) slide.after = after;
}

function normalizeCompareColumn(primary, alias, title, label, side, index) {
  if (primary && !Array.isArray(primary)) return normalizeCompareColumnObject(primary, title, label);
  if (Array.isArray(primary)) {
    console.warn(`Warning: slide ${index + 1} compare column "${side}" is an array; normalized it to { title, items } so body text is rendered.`);
    return normalizeCompareColumnObject({ title, label, items: primary }, title, label);
  }
  if (Array.isArray(alias)) {
    console.warn(`Warning: slide ${index + 1} compare column "${side}" is an array; normalized it to { title, items } so body text is rendered.`);
    return normalizeCompareColumnObject({ title, label, items: alias }, title, label);
  }
  if (alias && typeof alias === 'object') return normalizeCompareColumnObject(alias, title, label);
  return null;
}

function normalizeCompareColumnObject(column, title, label) {
  const next = { ...column };
  if (title && !next.title) next.title = title;
  if (label && !next.label) next.label = label;
  if (!Array.isArray(next.items)) {
    const fallbackItems = next.sections || next.points || next.bullets || next.list;
    if (Array.isArray(fallbackItems)) next.items = fallbackItems;
  }
  return next;
}

function validateSpecSlots(spec, options = {}) {
  const errors = [];
  const warnings = [];
  spec.slides.forEach((slide, index) => {
    validateTextFieldTypes(slide, index, errors);
    validateMediaSlots(slide, index, errors, warnings, options.specDir || process.cwd());
    validateTextSlots(slide, index, spec.style, errors, warnings);
    validateRenderableDataBlocks(slide, index, errors);
    validateThinContent(slide, index, errors);
  });
  if (errors.length) {
    fail(`Spec slot validation failed:\n- ${errors.join('\n- ')}`);
  }
  warnings.forEach((message) => console.warn(message));
}

const SCALAR_TEXT_FIELD_NAMES = new Set([
  'kicker', 'title', 'subtitle', 'body', 'desc', 'note', 'summary', 'detail', 'text', 'story',
  'conclusion', 'takeaway', 'footerSummary', 'nextStep', 'lead', 'callout',
  'quote', 'cite', 'source', 'caseTitle', 'summaryTitle', 'leadTitle', 'focusTitle',
  'label', 'value', 'unit', 'metric', 'name', 'caption',
]);

function validateTextFieldTypes(value, index, errors, pathName = 'slide', depth = 0) {
  if (!value || typeof value !== 'object' || depth > 5) return;
  if (Array.isArray(value)) {
    value.forEach((item, itemIndex) => validateTextFieldTypes(item, index, errors, `${pathName}[${itemIndex}]`, depth + 1));
    return;
  }
  Object.entries(value).forEach(([key, child]) => {
    const childPath = `${pathName}.${key}`;
    if (SCALAR_TEXT_FIELD_NAMES.has(key) && child !== undefined && child !== null && typeof child === 'object') {
      const kind = Array.isArray(child) ? 'array' : 'object';
      const hint = key === 'body'
        ? 'If this is a list of content blocks, put it in sections/items/columns/steps/nodes according to the slide layout; if it is bullet text inside one card, put strings in points[].'
        : 'Use a plain string/number for this field, or move structured content into the layout collection field.';
      errors.push(`slide ${index + 1} field ${childPath} must be plain text, but got ${kind}. ${hint}`);
    }
    if (child && typeof child === 'object' && !isOpaqueValidationObject(key)) {
      validateTextFieldTypes(child, index, errors, childPath, depth + 1);
    }
  });
}

function isOpaqueValidationObject(key) {
  return key === 'chart' || key === 'table' || key === 'speakerNotes';
}

function validateMediaSlots(slide, index, errors, warnings, specDir) {
  const layout = slide.layout || '';
  if (!MEDIA_SLOT_LAYOUTS.has(layout)) {
    const imageFields = unsupportedImageFields(slide);
    if (imageFields.length) {
      errors.push(`slide ${index + 1} layout "${layout}" does not render image/media field(s): ${imageFields.join(', ')}. Use a media layout such as media/mediaGrid/imageGrid/imageHero/statement/caseStudy, or remove these fields.`);
    }
    return;
  }
  const images = normalizeMediaImages(slide);
  const charts = normalizeMediaCharts(slide);
  const slotCount = isMediaGridLayout(layout) ? resolveMediaSlotCount(slide) : 1;
  const explicitCount = explicitMediaCount(slide);
  if (layout === 'statement' && images.length > 1) errors.push(`slide ${index + 1} uses statement layout with ${images.length} images, but statement supports exactly one image slot; use mediaGrid/imageGrid or split into another slide.`);
  if (layout !== 'statement' && images.length > 6) errors.push(`slide ${index + 1} has ${images.length} images, but media layouts support at most 6 image slots; split into another slide.`);
  if (layout === 'statement' && charts.length) errors.push(`slide ${index + 1} uses statement layout with chart data, but statement reserves the media area for one image or image placeholder; use chart/media layout instead.`);
  if (layout !== 'statement' && charts.length > 6) errors.push(`slide ${index + 1} has ${charts.length} charts, but media layouts support at most 6 media slots; split into another slide.`);
  if (layout !== 'statement' && Math.max(images.length, charts.length) > slotCount) {
    errors.push(`slide ${index + 1} has ${Math.max(images.length, charts.length)} media assets but only ${slotCount} slot(s).`);
  }
  if (explicitCount && images.length && explicitCount !== images.length && charts.length <= images.length && !slide.allowEmptyMediaSlots) {
    errors.push(`slide ${index + 1} declares mediaCount=${explicitCount} but provides ${images.length} image(s). Remove mediaCount or set it to ${images.length}; set allowEmptyMediaSlots:true only when blank image placeholders are intentional.`);
  }
  images.forEach((image, imageIndex) => {
    if (!resolveImage(specDir, image)) {
      errors.push(`slide ${index + 1} image ${imageIndex + 1} path is missing or unsupported; provide a valid image path, remove the image entry, or set allowEmptyMediaSlots:true with no image path when a blank placeholder is intentional.`);
    }
  });
  if (isVisualMediaLayout(layout) && !images.length && !charts.length && !slide.allowEmptyMediaSlots) {
    errors.push(`slide ${index + 1} uses layout "${layout}" with media/image slot(s) but provides no images or charts. Use a text-only layout such as textGrid/article/fourCards/agenda/radial, provide image/chart data, or set allowEmptyMediaSlots:true only when a blank placeholder is intentional.`);
  }
}

function validateTextSlots(slide, index, style, errors, warnings) {
  const layout = slide.layout || '';
  const sideMax = style === 'swiss' ? 5 : style === 'cmb' ? 4 : 3;
  const cmbTextWeaveMin = style === 'cmb' ? 0 : 1;
  const cmbTextGridMax = style === 'cmb' ? 6 : 9;
  const narrativeFields = ['body', 'desc', 'note', 'summary', 'detail', 'text', 'story'];
  const metricNarrativeFields = ['body', 'desc', 'summary', 'detail', 'text', 'story'];
  const textLayoutSuggestion = 'Use textGrid, article, sectionList, fourCards, agenda, or radial for title + body content.';
  const metricLayoutSuggestion = 'Use textGrid, article, sectionList, or media when each item needs explanatory body text.';
  const titleOnlyMatrixRule = style === 'magazine'
    ? { keys: ['items'], max: 12, min: 1, label: 'matrix cells' }
    : { keys: ['items'], max: 12, min: 1, label: 'matrix cells', itemTextKeys: ['title', 'label'], unusedItemFields: [...narrativeFields, 'name', 'value'], suggestion: textLayoutSuggestion };
  const numberMetricRule = style === 'magazine'
    ? { keys: ['items'], max: 6, min: 1, label: 'number cards', itemTextKeys: ['label', 'value', 'note', 'unit'], unusedItemFields: ['title', ...metricNarrativeFields], suggestion: metricLayoutSuggestion }
    : { keys: ['items'], max: 6, min: 1, label: 'number cards', itemTextKeys: ['label', 'value', 'note', 'unit'], unusedItemFields: ['title', ...metricNarrativeFields], suggestion: metricLayoutSuggestion };
  const dashboardMetricRule = style === 'magazine'
    ? { keys: ['metrics', 'items'], max: 4, min: 1, label: 'dashboard metrics', itemTextKeys: ['label', 'value', 'note'], unusedItemFields: ['title', ...metricNarrativeFields], suggestion: metricLayoutSuggestion }
    : { keys: ['metrics', 'items'], max: style === 'swiss' ? 5 : 4, min: 1, label: 'dashboard metrics', itemTextKeys: ['label', 'value'], unusedItemFields: ['title', 'note', ...metricNarrativeFields], suggestion: metricLayoutSuggestion };
  const rules = {
    cover: [],
    section: [],
    bigQuote: [],
    quoteImage: [],
    compare: [],
    duoCompare: [],
    splitCompare: [],
    textImage: [],
    statement: [],
    closing: [],
    bigNumbers: [numberMetricRule],
    kpiTower: [{ ...numberMetricRule, max: 4, label: 'KPI cards' }],
    pipeline: [{ keys: ['steps', 'items'], max: 6, min: 1, label: 'pipeline steps' }],
    timeline: [{ keys: ['items', 'steps'], max: 6, min: 1, label: 'timeline steps' }],
    matrix: [titleOnlyMatrixRule],
    fourCards: [{ keys: ['items'], max: 8, min: 1, label: 'cards' }],
    article: [{ keys: ['sections', 'items', 'columns'], max: 6, min: 1, label: 'article sections' }],
    briefing: [{ keys: ['sections', 'items', 'columns', 'points', 'agenda'], max: 6, min: 2, label: 'briefing text blocks' }],
    executiveBrief: [{ keys: ['sections', 'items', 'columns', 'points', 'agenda'], max: 6, min: 2, label: 'briefing text blocks' }],
    contentBrief: [{ keys: ['sections', 'items', 'columns', 'points', 'agenda'], max: 6, min: 2, label: 'briefing text blocks' }],
    textGrid: [{ keys: ['sections', 'items', 'columns'], max: cmbTextGridMax, min: cmbTextWeaveMin, label: style === 'cmb' ? 'CMB text weave cards (lead + right card)' : 'text grid cells' }],
    textWeave: [{ keys: ['sections', 'items', 'columns', 'points', 'agenda'], max: 6, min: cmbTextWeaveMin, label: style === 'cmb' ? 'CMB text weave cards (lead + right card)' : 'text weave blocks' }],
    contentSynthesis: [{ keys: ['sections', 'items', 'columns', 'points', 'agenda'], max: 6, min: cmbTextWeaveMin, label: style === 'cmb' ? 'CMB text weave cards (lead + right card)' : 'text weave blocks' }],
    denseText: [{ keys: ['sections', 'items', 'columns', 'points', 'agenda'], max: 6, min: cmbTextWeaveMin, label: style === 'cmb' ? 'CMB text weave cards (lead + right card)' : 'dense text blocks' }],
    sectionList: [{ keys: ['sections', 'items', 'columns'], max: 7, min: 1, label: 'section list items' }],
    agenda: [{ keys: ['items', 'sections', 'agenda'], max: 8, min: 1, label: 'agenda items' }],
    pyramid: [{ keys: ['layers', 'items', 'sections'], max: 5, min: 1, label: 'pyramid layers' }],
    radial: [{ keys: ['items', 'nodes', 'sections'], max: 8, min: 1, label: 'radial nodes' }],
    roadmap: [{ keys: ['steps', 'items'], max: 6, min: 1, label: 'roadmap steps' }],
    swimlane: [{ keys: ['lanes', 'sections'], max: 4, min: 1, label: 'swimlanes' }],
    media: [{ keys: ['items', 'insights', 'points'], max: sideMax, min: 0, label: 'side points' }],
    mediaGrid: [{ keys: ['captions', 'items', 'sections'], max: 6, min: 0, label: 'media captions', itemTextKeys: ['caption', 'title', 'label'], unusedItemFields: [...narrativeFields, 'points', 'bullets', 'list', 'value'], suggestion: 'Use caption/title/label for mediaGrid captions. Use media/textImage/caseStudy or split into a text slide when each image needs body text.' }],
    gallery: [{ keys: ['captions', 'items', 'sections'], max: 6, min: 0, label: 'media captions', itemTextKeys: ['caption', 'title', 'label'], unusedItemFields: [...narrativeFields, 'points', 'bullets', 'list', 'value'], suggestion: 'Use caption/title/label for gallery captions. Use media/textImage/caseStudy or split into a text slide when each image needs body text.' }],
    imageGrid: [{ keys: ['captions', 'items', 'sections'], max: 6, min: 0, label: 'media captions', itemTextKeys: ['caption', 'title', 'label'], unusedItemFields: [...narrativeFields, 'points', 'bullets', 'list', 'value'], suggestion: 'Use caption/title/label for imageGrid captions. Use media/textImage/caseStudy or split into a text slide when each image needs body text.' }],
    imageHero: [{ keys: ['items'], max: 3, min: 0, label: 'image hero metrics', itemTextKeys: ['label', 'value', 'note'], unusedItemFields: ['title', ...metricNarrativeFields], suggestion: metricLayoutSuggestion }],
    caseStudy: [{ keys: ['metrics', 'items'], max: 3, min: 0, label: 'case metrics', itemTextKeys: ['label', 'title', 'value', 'note'], unusedItemFields: metricNarrativeFields, suggestion: metricLayoutSuggestion }],
    dataSheet: [{ keys: ['notes', 'insights'], max: style === 'swiss' ? 4 : 3, min: 0, label: 'side notes' }],
    chart: [{ keys: ['insights', 'notes'], max: style === 'swiss' ? 3 : 4, min: 0, label: 'chart insights' }],
    dashboard: [dashboardMetricRule],
  };
  const layoutRules = rules[layout] || [];
  validateIgnoredSlotFields(slide, index, layout, layoutRules, errors, Object.prototype.hasOwnProperty.call(rules, layout));
  layoutRules.forEach((rule) => validateSlotCollection(slide, index, rule, errors, warnings));
  if (layout === 'compare' || layout === 'duoCompare' || layout === 'splitCompare') {
    validateSlotCollection(slide.before || {}, index, { keys: ['items'], max: 6, min: 1, label: 'before items', prefix: 'before.' }, errors, warnings);
    validateSlotCollection(slide.after || {}, index, { keys: ['items'], max: 6, min: 1, label: 'after items', prefix: 'after.' }, errors, warnings);
  }
  validateCmbTextWeaveStructure(slide, index, style, layout, errors);
  validateCmbBriefingCapacity(slide, index, style, layout, errors);
  if (layout === 'dashboard') validateChartSlots(slide, index, 2, errors, warnings);
  if (layout === 'chart') validateChartDataSlot(slide, index, errors, warnings);
  if (layout === 'dataSheet') validateTableSlot(slide, index, errors, warnings);
}

function validateSlotCollection(source, index, rule, errors, warnings) {
  const present = rule.keys.filter((key) => source[key] !== undefined && source[key] !== null);
  if (!present.length) {
    if (rule.min > 0 && !source.allowSparseContent) errors.push(`slide ${index + 1} has no ${rule.label}; the layout may look empty. Add the required content field(s), change layout, or set allowSparseContent:true only for intentional sparse draft slides.`);
    return [];
  }
  if (present.length > 1) {
    errors.push(`slide ${index + 1} provides multiple fields for ${rule.label}: ${present.map((key) => `${rule.prefix || ''}${key}`).join(', ')}. Keep only one field so content is not silently ignored.`);
  }
  const key = present[0];
  const items = normalizeSlotItemsForValidation(source[key]);
  if (!items) {
    errors.push(slotCollectionFormatError(source[key], index, rule, key));
    return [];
  }
  if (rule.max && items.length > rule.max) {
    errors.push(`slide ${index + 1} has ${items.length} ${rule.label}, but layout "${source.layout || 'nested'}" renders at most ${rule.max}; split content or change layout.`);
  }
  if (rule.min && items.length < rule.min) {
    if (!source.allowSparseContent) errors.push(`slide ${index + 1} has ${items.length} ${rule.label}; expected at least ${rule.min}. Add content, change layout, or set allowSparseContent:true only for intentional sparse draft slides.`);
  }
  items.forEach((item, itemIndex) => {
    if (!slotItemHasDisplayTextForRule(item, rule)) {
      const keys = rule.itemTextKeys || DISPLAY_ITEM_TEXT_KEYS;
      errors.push(`slide ${index + 1} ${rule.prefix || ''}${key}[${itemIndex}] has no field rendered by ${rule.label}. Use one of: ${keys.join(', ')}.`);
    }
    validateUnusedSlotItemFields(item, index, rule, key, itemIndex, errors);
  });
  return items;
}

function validateCmbTextWeaveStructure(slide, index, style, layout, errors) {
  if (style !== 'cmb') return;
  if (!['textGrid', 'fourCards', 'textWeave', 'contentSynthesis', 'denseText'].includes(layout)) return;
  const count = cmbTextItems(slide).length;
  if (count < 2) {
    errors.push(`slide ${index + 1} layout "${layout}" needs at least 2 CMB text weave cards: 1 lead card and at least 1 right-side card; got ${count}. Add another sections/items/columns/points/agenda entry or change layout.`);
  }
}

function validateCmbBriefingCapacity(slide, index, style, layout, errors) {
  if (style !== 'cmb') return;
  if (!['article', 'sectionList', 'briefing', 'executiveBrief', 'contentBrief'].includes(layout)) return;
  const items = normalizeSections(slide.sections || slide.items || slide.columns || slide.points || slide.agenda || []);
  const hasLead = !!(slide.summary || slide.body || slide.lead);
  const conclusionText = slide.conclusion || slide.takeaway || slide.footerSummary || slide.nextStep;
  const restCount = hasLead ? items.length : Math.max(0, items.length - 1);
  const maxRest = conclusionText ? 4 : 5;
  if (restCount > maxRest) {
    const conclusionNote = conclusionText ? ' when conclusion/takeaway is present' : '';
    errors.push(`slide ${index + 1} layout "${layout}" renders at most ${maxRest} middle briefing text block(s)${conclusionNote}; got ${restCount}. Split content, reduce sections/items, remove conclusion, or use textWeave/textGrid.`);
  }
}

function validateIgnoredSlotFields(slide, index, layout, layoutRules, errors, isKnownLayout = false) {
  const valid = new Set(layoutRules.flatMap((rule) => rule.keys));
  const common = ['items', 'sections', 'columns', 'steps', 'nodes', 'layers', 'metrics', 'notes', 'insights', 'agenda', 'lanes', 'captions', 'points'];
  const ignored = common.filter((key) => slide[key] !== undefined && slide[key] !== null && !valid.has(key));
  if (ignored.length && (layoutRules.length || isKnownLayout)) {
    const targets = Array.from(valid);
    const suggestion = targets.length
      ? `Rename them to one of: ${targets.join(', ')}.`
      : 'Use a collection layout such as textGrid, article, sectionList, agenda, roadmap, timeline, media, dashboard, or dataSheet.';
    errors.push(`slide ${index + 1} layout "${layout}" does not render collection field(s): ${ignored.join(', ')}. ${suggestion}`);
  }
  if (layout === 'dashboard' && slide.chart && !Array.isArray(slide.charts)) {
    errors.push(`slide ${index + 1} dashboard uses charts[]; field chart will not be rendered. Rename chart to charts: [chart].`);
  }
}

function validateChartSlots(slide, index, max, errors, warnings) {
  const charts = normalizeMediaCharts(slide);
  if (charts.length > max) errors.push(`slide ${index + 1} has ${charts.length} charts, but dashboard renders at most ${max}.`);
  if (!charts.length && !slide.allowMissingChart) errors.push(`slide ${index + 1} dashboard has no charts; chart area will be empty. Provide charts[] or set allowMissingChart:true only for intentional draft output.`);
}

function validateChartDataSlot(slide, index, errors, warnings) {
  const chart = slide.chart || slide;
  if (!normalizeChartData(chart).length && !slide.allowMissingChart) errors.push(`slide ${index + 1} chart layout has no chart data; a NO DATA box will be rendered. Provide chart data or set allowMissingChart:true only for intentional draft output.`);
}

function validateTableSlot(slide, index, errors, warnings) {
  if (!slide.table) {
    if (!slide.allowMissingTable) errors.push(`slide ${index + 1} dataSheet has no table. Provide table data or set allowMissingTable:true only for intentional draft output.`);
    return;
  }
  if (slide.table.headers && !Array.isArray(slide.table.headers)) errors.push(`slide ${index + 1} table.headers must be an array.`);
  const rows = normalizeTableRows(slide.table);
  if (!rows.length && !slide.allowMissingTable) errors.push(`slide ${index + 1} table has no rows. Provide table.rows or set allowMissingTable:true only for intentional draft output.`);
}

function slotCollectionFormatError(value, index, rule, key) {
  const field = `${rule.prefix || ''}${key}`;
  if (value && typeof value === 'object' && !Array.isArray(value)) {
    const badKeys = Object.entries(value)
      .filter(([, body]) => body !== undefined && body !== null && typeof body === 'object')
      .map(([title]) => title);
    if (badKeys.length) {
      return `slide ${index + 1} field ${field} uses an object map with structured value(s): ${badKeys.slice(0, 5).join(', ')}. Use ${field}: [{ title, body }] for structured items, or ${field}: { "Title": "plain text" } for a simple object map.`;
    }
  }
  return `slide ${index + 1} field ${field} has unsupported format; use an array of strings/objects or an object map with plain-text values.`;
}

function normalizeSlotItemsForValidation(value) {
  if (value === undefined || value === null) return [];
  if (typeof value === 'string' || typeof value === 'number') return [{ body: String(value) }];
  if (Array.isArray(value)) return value;
  if (typeof value === 'object') {
    const entries = Object.entries(value);
    if (entries.some(([, body]) => body !== undefined && body !== null && typeof body === 'object')) return null;
    return entries.map(([title, body]) => ({ title, body: String(body ?? '') }));
  }
  return null;
}

const DISPLAY_ITEM_TEXT_KEYS = ['text', 'title', 'label', 'body', 'desc', 'note', 'summary', 'detail', 'value', 'unit', 'metric', 'name'];

function slotItemHasDisplayText(item) {
  if (typeof item === 'string' || typeof item === 'number') return String(item).trim().length > 0;
  if (!item || typeof item !== 'object') return false;
  return DISPLAY_ITEM_TEXT_KEYS.some((key) => String(item[key] ?? '').trim().length > 0);
}

function slotItemHasDisplayTextForRule(item, rule) {
  if (typeof item === 'string' || typeof item === 'number') return String(item).trim().length > 0;
  if (!item || typeof item !== 'object') return false;
  const keys = rule.itemTextKeys || DISPLAY_ITEM_TEXT_KEYS;
  return keys.some((key) => String(item[key] ?? '').trim().length > 0);
}

function validateUnusedSlotItemFields(item, index, rule, key, itemIndex, errors) {
  if (!item || typeof item !== 'object' || Array.isArray(item) || !rule.unusedItemFields?.length) return;
  const ignored = rule.unusedItemFields.filter((field) => String(item[field] ?? '').trim().length > 0);
  if (!ignored.length) return;
  const suggestion = rule.suggestion || 'Use a layout that renders these item fields, or remove the unused fields.';
  errors.push(`slide ${index + 1} ${rule.prefix || ''}${key}[${itemIndex}] includes field(s) not rendered by ${rule.label}: ${ignored.join(', ')}. ${suggestion}`);
}

function unsupportedImageFields(slide) {
  const fields = [];
  ['image', 'images', 'gallery'].forEach((key) => {
    if (slide[key] !== undefined && slide[key] !== null) fields.push(key);
  });
  if (Array.isArray(slide.media) && normalizeMediaImages({ media: slide.media }).length) fields.push('media');
  ['mediaCount', 'imageSlots', 'slotCount'].forEach((key) => {
    if (slide[key] !== undefined && slide[key] !== null) fields.push(key);
  });
  return fields;
}

const MEDIA_SLOT_LAYOUTS = new Set(['statement', 'media', 'mediaGrid', 'gallery', 'imageGrid', 'imageHero', 'quoteImage', 'textImage', 'caseStudy']);
const VISUAL_MEDIA_LAYOUTS = MEDIA_SLOT_LAYOUTS;
const CHART_DATA_LAYOUTS = new Set(['chart', 'dashboard']);
const TABLE_DATA_LAYOUTS = new Set(['dataSheet']);

function isVisualMediaLayout(layout) {
  return VISUAL_MEDIA_LAYOUTS.has(layout || '');
}

function hasUserImages(data) {
  return normalizeMediaImages(data || {}).length > 0;
}

function chartCount(data) {
  return normalizeMediaCharts(data || {}).length;
}

function hasChartData(data) {
  return chartCount(data) > 0;
}

function hasTableData(data) {
  if (!data) return false;
  if (data.table && normalizeTableRows(data.table).length) return true;
  return Array.isArray(data.tables) && data.tables.some((table) => normalizeTableRows(table).length);
}

function hasCollectionOutside(slide, allowedKeys) {
  const allowed = new Set(allowedKeys || []);
  const keys = ['items', 'sections', 'columns', 'steps', 'nodes', 'layers', 'metrics', 'notes', 'insights', 'agenda', 'lanes', 'captions', 'points'];
  return keys.some((key) => slide?.[key] !== undefined && slide?.[key] !== null && !allowed.has(key));
}

function layoutAllowedByContent(slide, layout) {
  if (!layout) return false;
  if (layout === 'chart') return hasChartData(slide) && !hasCollectionOutside(slide, ['insights', 'notes']);
  if (layout === 'dashboard') return chartCount(slide) >= 2 && normalizeSections(slide.metrics || slide.items || []).length > 0 && !hasCollectionOutside(slide, ['metrics', 'items']);
  if (layout === 'dataSheet') return hasTableData(slide) && !hasCollectionOutside(slide, ['notes', 'insights']);
  if (isVisualMediaLayout(layout)) return hasUserImages(slide) || (hasChartData(slide) && !hasCollectionOutside(slide, ['items', 'insights', 'points', 'captions', 'metrics']));
  return true;
}

function filterLayoutCandidates(slide, candidates) {
  return (candidates || []).filter((layout) => layoutAllowedByContent(slide, layout));
}

function replacementCandidatesForLayout(layout) {
  const replacements = {
    textGrid: ['sectionList', 'fourCards', 'matrix', 'radial', 'chart', 'dataSheet'],
    article: ['sectionList', 'textGrid', 'fourCards', 'radial', 'matrix', 'chart', 'dataSheet'],
    fourCards: ['sectionList', 'textGrid', 'matrix', 'radial', 'chart', 'dataSheet'],
    matrix: ['sectionList', 'textGrid', 'fourCards', 'radial', 'chart', 'dataSheet'],
    statement: ['sectionList', 'agenda', 'fourCards', 'textGrid', 'radial', 'chart', 'dataSheet', 'caseStudy'],
    compare: ['swimlane', 'matrix', 'agenda', 'chart', 'dataSheet'],
    timeline: ['roadmap', 'swimlane', 'agenda', 'chart', 'dataSheet'],
    pipeline: ['roadmap', 'swimlane', 'agenda', 'chart', 'dataSheet'],
  };
  return replacements[layout] || ['sectionList', 'fourCards', 'textGrid', 'radial', 'roadmap', 'swimlane', 'chart', 'dataSheet'];
}

function suggestedLayoutsForSlide(slide, currentLayout, limit = 4) {
  const candidates = filterLayoutCandidates(slide, replacementCandidatesForLayout(currentLayout))
    .filter((layout) => layout !== currentLayout);
  const ordered = [];
  if (slide.layoutAlt && layoutAllowedByContent(slide, slide.layoutAlt) && slide.layoutAlt !== currentLayout) ordered.push(slide.layoutAlt);
  candidates.forEach((layout) => {
    if (!ordered.includes(layout)) ordered.push(layout);
  });
  return ordered.slice(0, limit);
}
function explicitMediaCount(data) {
  const n = Number(data.mediaCount || data.imageSlots || data.slotCount);
  return Number.isFinite(n) && n > 0 ? Math.floor(n) : 0;
}

function isMediaGridLayout(layout) {
  return ['mediaGrid', 'gallery', 'imageGrid'].includes(layout || '');
}

function resolveMediaSlotCount(data) {
  const explicit = explicitMediaCount(data);
  if (explicit) return clamp(explicit, 1, 6);
  const images = normalizeMediaImages(data);
  const charts = normalizeMediaCharts(data);
  const captions = normalizeSlotItemsForValidation(data.captions || []);
  const declaredMedia = Array.isArray(data.media) ? data.media.length : 0;
  return clamp(Math.max(images.length, charts.length, captions ? captions.length : 0, declaredMedia, 1), 1, 6);
}

function validateRenderableDataBlocks(slide, index, errors) {
  const layout = slide.layout || '';
  const chartLayouts = ['chart', 'dashboard', 'media', 'mediaGrid', 'gallery', 'imageGrid', 'imageHero', 'quoteImage', 'textImage', 'caseStudy'];
  const blocks = [...(slide.blocks || [])];
  if (!chartLayouts.includes(layout)) {
    blocks.push(...(slide.charts || []).map((chart) => ({ ...chart, type: 'chart' })));
  }
  blocks.push(...(slide.tables || []).map((table) => ({ ...table, type: 'table' })));
  blocks.forEach((block) => {
    if (!block || block.enabled === false) return;
    if ((block.type === 'chart' || block.type === 'table') && !hasExplicitBox(block)) {
      errors.push(`slide ${index + 1} has unpositioned ${block.type} block. Put it in layout "media"/"mediaGrid", remove it, or set x/y/w/h explicitly so it is not silently skipped.`);
    }
  });
}

function validateThinContent(slide, index, errors) {
  if (slide.allowSparseContent) return;
  const layout = slide.layout || '';
  if (['matrix', 'bigNumbers', 'kpiTower', 'dashboard', 'imageHero', 'caseStudy'].includes(layout)) return;
  const candidates = normalizeSections(slide.sections || slide.items || slide.columns || slide.nodes || slide.layers || slide.steps || slide.milestones || slide.agenda || []);
  if (candidates.length < 3) return;
  const titleOnly = candidates.filter((item) => {
    const title = item.title || item.label || item.name;
    const body = item.body || item.desc || item.note || item.text || item.summary || item.detail || (Array.isArray(item.points) && item.points.length) || (Array.isArray(item.bullets) && item.bullets.length) || (Array.isArray(item.list) && item.list.length);
    return title && !body;
  }).length;
  if (titleOnly >= Math.ceil(candidates.length * 0.6)) {
    errors.push(`slide ${index + 1} has ${titleOnly}/${candidates.length} title-only items. Add body/desc/note/points for each point, change layout, or set allowSparseContent:true only for intentional sparse draft slides.`);
  }
}

function warnLayoutVariety(spec) {
  let runLayout = null;
  let runStart = 0;
  let runLength = 0;
  const flush = () => {
    const slide = spec.slides[Math.min(spec.slides.length - 1, runStart + runLength - 1)] || {};
    const suggestions = suggestedLayoutsForSlide(slide, runLayout);
    const suffix = suggestions.length
      ? ` Suggested text/data-compatible alternatives: ${suggestions.map((layout) => `"${layout}"`).join(', ')}.`
      : ' No image/media-slot layout is suggested unless this page provides images or chart data.';
    if (runLength >= 3) {
      console.warn(
        `Warning: slides ${runStart + 1}-${runStart + runLength} use layout "${runLayout}" consecutively. Change at least one page to an equivalent layout by editing slide.layout in JSON so the deck does not feel repetitive.${suffix}`
      );
    } else if (runLength === 2) {
      console.warn(
        `Notice: slides ${runStart + 1}-${runStart + 2} both use layout "${runLayout}". If the content is not intentionally paired, consider alternating layouts for visual variety.${suffix}`
      );
    }
  };
  spec.slides.forEach((slide, index) => {
    const layout = slide.layout || (spec.style === 'magazine' ? 'textImage' : 'statement');
    if (layout === runLayout) {
      runLength += 1;
    } else {
      flush();
      runLayout = layout;
      runStart = index;
      runLength = 1;
    }
  });
  flush();
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
