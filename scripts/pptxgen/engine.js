const fs = require('node:fs');
const path = require('node:path');
const process = require('node:process');
const pptxgen = require('pptxgenjs');
const {
  SLIDE,
  THEMES,
  FONTS,
  READABILITY,
  BASIC_ICON_NAMES,
  ICON_ALIASES,
  LUCIDE_STATIC_ICON_DIR,
  isSupportedStyle,
  defaultThemeForStyle,
} = require('./config');
const { fail } = require('./errors');

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
const IMAGE_ASPECT_CACHE = new Map();
let LUCIDE_MODULE = undefined;

function normalizeSpec(spec, options = {}) {
  spec.style = spec.style || 'magazine';
  if (!isSupportedStyle(spec.style)) fail(`Unsupported style: ${spec.style}`);
  spec.theme = spec.theme || defaultThemeForStyle(spec.style);
  if (!THEMES[spec.style][spec.theme]) fail(`Unsupported theme "${spec.theme}" for style "${spec.style}"`);
  if (!Array.isArray(spec.slides) || spec.slides.length === 0) fail('Spec must include a non-empty slides array.');
  const shouldDiversify = options.diversifyLayouts || spec.diversifyLayouts === true;
  if (shouldDiversify && !options.writeNormalizedSpec && !spec.allowUnsyncedLayoutDiversify) {
    fail('Layout diversification changes slide.layout. Use --write-normalized-spec path/to/normalized.json so the JSON matches the generated PPTX, or remove --diversify-layouts.');
  }
  const layoutChanges = diversifyRepeatedLayouts(spec, { mutate: shouldDiversify });
  if (layoutChanges.length && shouldDiversify) spec.__layoutDiversified = true;
  validateSpecSlots(spec, { specDir: options.specDir || process.cwd() });
  warnThinContent(spec);
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
    headFontFace: spec.style === 'magazine' ? FONTS.serifZh : FONTS.sans,
    bodyFontFace: spec.style === 'magazine' ? FONTS.sansZh : FONTS.sansZh,
    lang: 'zh-CN',
  };
  pptx.defineLayout({ name: 'CUSTOM_WIDE', width: SLIDE.w, height: SLIDE.h });
  pptx.layout = 'CUSTOM_WIDE';

  const theme = THEMES[spec.style][spec.theme];
  spec.slides.forEach((slideSpec, index) => {
    const slide = pptx.addSlide();
    enforceReadableSlideText(slide);
    const ctx = { spec, slideSpec, theme, specDir, index, total: spec.slides.length };
    renderByStyle(spec.style, slide, ctx);
  });

  fs.mkdirSync(path.dirname(outPath), { recursive: true });
  await pptx.writeFile({ fileName: outPath });
  console.log(`Wrote ${outPath}`);
  console.log(`Slides: ${spec.slides.length}`);
  console.log(`Style: ${spec.style} / ${theme.name}`);
}

function enforceReadableSlideText(slide) {
  if (slide.__readabilityPatched) return;
  const originalAddText = slide.addText.bind(slide);
  slide.addText = (text, options = {}) => {
    const readableText = Array.isArray(text)
      ? text.map((run) => ({ ...run, options: readableTextOptions(run.options || {}, run.text) }))
      : text;
    return originalAddText(readableText, readableTextOptions(options, text));
  };
  slide.__readabilityPatched = true;
}

function readableTextOptions(options, text) {
  if (!options || options.noReadabilityScale) return options;
  const next = { ...options };
  if (typeof next.fontSize === 'number' && next.fontSize < READABILITY.minFontSize && !isSymbolText(next, text)) {
    next.fontSize = READABILITY.minFontSize;
  }
  return next;
}

function isSymbolText(options, text) {
  const raw = Array.isArray(text) ? '' : String(text || '').trim();
  return options.fontFace === 'Segoe UI Symbol' || (raw.length <= 1 && options.align === 'center' && options.valign === 'mid');
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
    textImage: magazineTextImage,
    article: magazineArticle,
    statement: magazineStatementCompat,
    kpiTower: magazineBigNumbers,
    duoCompare: magazineCompare,
    timeline: magazinePipeline,
    matrix: magazineMatrixCompat,
    fourCards: magazineMatrixCompat,
    imageHero: magazineImageHeroCompat,
    textGrid: magazineMatrixCompat,
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
    textImage: swissTextImageCompat,
    article: swissTextGrid,
    statement: swissStatement,
    kpiTower: swissKpiTower,
    duoCompare: swissDuoCompare,
    timeline: swissTimeline,
    matrix: swissMatrix,
    fourCards: swissFourCards,
    imageHero: swissImageHero,
    textGrid: swissTextGrid,
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
    timeline: swissTimeline,
    pipeline: swissTimeline,
    roadmap: swissRoadmap,
    textGrid: swissTextGrid,
    article: swissTextGrid,
    fourCards: swissFourCards,
    matrix: swissMatrix,
    agenda: swissAgenda,
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
  slide.addShape(pptx.ShapeType.rect, { x: 0, y: 0, w: SLIDE.w, h: headerH, fill: { color: 'FFFFFF', transparency: 0 }, line: { color: 'FFFFFF', transparency: 100 } });
  slide.addShape(pptx.ShapeType.line, { x: 0, y: headerH, w: SLIDE.w, h: 0, line: { color: theme.accent, transparency: 7, width: 1.2 } });
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
<rect x="0" y="0" width="${w}" height="${headerH}" fill="#FFFFFF"/>
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
  slide.addText(left, { x: SLIDE.marginX + logoW + 0.28, y: 0.34, w: 5.5, h: 0.2, fontFace: FONTS.sans, fontSize: 7.6, bold: true, charSpace: 0.8, color: ctx.theme.ink, transparency: 8, margin: 0, fit: 'shrink' });
  slide.addText(right, { x: SLIDE.w - SLIDE.marginX - 2.2, y: 0.34, w: 2.2, h: 0.2, fontFace: FONTS.sans, fontSize: 7.4, bold: true, charSpace: 0.8, color: ctx.theme.accent, align: 'right', margin: 0, fit: 'shrink' });
  slide.addShape(pptx.ShapeType.rect, { x: 0, y: headerH - 0.035, w: 1.85, h: 0.035, fill: { color: ctx.theme.accent, transparency: 0 }, line: { color: ctx.theme.accent, transparency: 100 } });
}

function cmbCover(slide, ctx, s) {
  const data = ctx.slideSpec;
  const headY = pageHeadY(ctx, 1.08);
  addCmbLogoMark(slide, ctx, { x: 10.62, y: 1.02, w: 1.72, h: 1.72 });
  slide.addText(data.kicker || 'CHINA MERCHANTS BANK', { x: 0.78, y: headY, w: 6.8, h: 0.24, fontFace: FONTS.sans, fontSize: 8.6, bold: true, charSpace: 1.5, color: s.fg, transparency: 18, margin: 0, fit: 'shrink' });
  slide.addText(data.title || ctx.spec.title, { x: 0.78, y: headY + 0.6, w: 10.2, h: 2.15, fontFace: FONTS.sansZh, fontSize: fitTitle(data.title || ctx.spec.title, 47, 31), bold: true, color: s.fg, margin: 0, fit: 'shrink' });
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
  slide.addText(data.kicker || 'Executive Summary', { x: 1.05, y: headY, w: 5.8, h: 0.24, fontFace: FONTS.sans, fontSize: 8, bold: true, charSpace: 1.1, color: ctx.theme.accent, margin: 0, fit: 'shrink' });
  slide.addText(data.title || '', { x: 1.02, y: headY + 0.48, w: 6.85, h: 1.58, fontFace: FONTS.sansZh, fontSize: fitTitle(data.title || '', 34, 26), bold: true, color: s.fg, margin: 0, fit: 'shrink' });
  if (data.body || data.subtitle) slide.addText(data.body || data.subtitle, { x: 1.05, y: headY + 2.42, w: 6.45, h: 1.1, fontFace: FONTS.sansZh, fontSize: 14.2, color: s.fg, transparency: 14, margin: 0, fit: 'shrink', valign: 'top' });
  addStatementImageSlot(slide, ctx, { x: 8.05, y: headY + 1.72, w: 3.85, h: 3.0 }, ctx.theme.accent, data.imageLabel || 'IMAGE SLOT');
  addFoot(slide, ctx, s.fg, 'swiss');
}

function cmbClosing(slide, ctx, s) {
  cmbCover(slide, { ...ctx, slideSpec: { ...ctx.slideSpec, kicker: ctx.slideSpec.kicker || 'THANK YOU', title: ctx.slideSpec.title || '谢谢观看', subtitle: ctx.slideSpec.subtitle || 'CHINA MERCHANTS BANK' } }, s);
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
    const svg = readSvgWithOpacity(imagePath, options.opacity);
    slide.addImage({ data: svgDataUri(svg), ...placedBox });
  } else {
    slide.addImage({ path: imagePath, ...placedBox });
  }
  return placedBox;
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
  const font = mode === 'swiss' ? 'JetBrains Mono' : FONTS.mono;
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
    fontFace: mode === 'swiss' ? 'JetBrains Mono' : FONTS.mono,
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
    addBullets(slide, col.items || [], x + 0.25, 3.82, 4.9, 1.65, s.fg, i === 0 ? 35 : 10, 'magazine');
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
  const cols = clampColumns(data.columnsCount || autoColumns(sections.length, 3), 1, 3);
  const x0 = 0.78;
  const y0 = data.subtitle ? 2.78 : 2.45;
  const gap = 0.34;
  const colW = (11.75 - gap * (cols - 1)) / cols;
  sections.forEach((section, i) => {
    const x = x0 + (i % cols) * (colW + gap);
    const y = y0 + Math.floor(i / cols) * 1.95;
    slide.addShape(pptx.ShapeType.line, { x, y, w: colW, h: 0, line: { color: s.fg, transparency: 65, width: 0.6 } });
    const hasIcon = addInlineIcon(slide, section, x, y + 0.15, 0.26, s.fg, 'magazine', { fallback: defaultContentIcon(i, 'magazine'), pad: 0.045 });
    const titleX = hasIcon ? x + 0.36 : x;
    slide.addText(section.title || section.label || '', { x: titleX, y: y + 0.16, w: colW - (hasIcon ? 0.36 : 0), h: 0.32, fontFace: FONTS.serifZh, fontSize: 15, bold: true, color: s.fg, margin: 0, fit: 'shrink' });
    const body = section.body || section.desc || (section.items || []).map((item) => typeof item === 'string' ? item : item.text || item.title || '').join('\n');
    slide.addText(body, { x: titleX, y: y + 0.58, w: colW - (hasIcon ? 0.36 : 0), h: 1.05, fontFace: FONTS.sansZh, fontSize: 9.5, color: s.fg, transparency: 18, margin: 0.03, fit: 'shrink', valign: 'top', breakLine: false });
  });
  if (data.callout) addCallout(slide, data.callout, 8.3, 5.88, 3.65, 0.68, s.fg, ctx.theme.paperTint);
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
    const labelX = clamp(x - 0.48, 0.65, SLIDE.w - 1.6);
    const titleX = clamp(x - 0.68, 0.65, SLIDE.w - 1.85);
    const descX = clamp(x - 0.72, 0.65, SLIDE.w - 1.95);
    slide.addText(item.label || item.date || `0${i + 1}`, { x: labelX, y: y - 0.24, w: 0.95, h: 0.2, fontFace: FONTS.mono, fontSize: 9.5, color: s.fg, transparency: 35, align: 'center', margin: 0, fit: 'shrink' });
    slide.addText(item.title || '', { x: titleX, y, w: 1.35, h: 0.34, fontFace: FONTS.serifZh, fontSize: 12.5, bold: true, color: s.fg, align: 'center', margin: 0, fit: 'shrink' });
    slide.addText(item.body || item.desc || item.note || '', { x: descX, y: y + 0.42, w: 1.45, h: 0.45, fontFace: FONTS.sansZh, fontSize: 9.2, color: s.fg, transparency: 26, align: 'center', margin: 0, fit: 'shrink' });
  });
  addFoot(slide, ctx, s.fg, 'magazine');
}

function magazineSwimlane(slide, ctx, s) {
  const data = ctx.slideSpec;
  addPageHead(slide, data, s.fg, 'magazine', 0.78);
  const lanes = (data.lanes || data.sections || []).slice(0, 4);
  const stages = data.stages || data.columns || ['Now', 'Next', 'Later'];
  const x0 = 1.55;
  const y0 = 2.55;
  const laneH = 0.86;
  const colW = 10.25 / Math.max(stages.length, 1);
  stages.forEach((stage, i) => slide.addText(String(stage), { x: x0 + i * colW, y: 2.22, w: colW - 0.12, h: 0.22, fontFace: FONTS.mono, fontSize: 9.5, color: s.fg, transparency: 35, margin: 0, fit: 'shrink' }));
  lanes.forEach((lane, r) => {
    const y = y0 + r * 1.02;
    slide.addText(lane.title || lane.label || `Lane ${r + 1}`, { x: 0.78, y: y + 0.2, w: 0.62, h: 0.34, fontFace: FONTS.serifZh, fontSize: 11.5, bold: true, color: s.fg, margin: 0, fit: 'shrink' });
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
  addStatementImageSlot(slide, ctx, { x: 7.35, y: headY + 0.48, w: 4.55, h: 3.85 }, s.fg, data.imageLabel || 'IMAGE SLOT');
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
  slide.addText(data.kicker || 'Swiss Field Note', { x: 0.72, y: headY, w: 6, h: 0.28, fontFace: 'JetBrains Mono', fontSize: 8.5, charSpace: 1.8, color: s.fg, transparency: 18, margin: 0 });
  slide.addText(data.title || ctx.spec.title, { x: 0.68, y: headY + 0.5, w: 11.5, h: 2.6, fontFace: FONTS.sans, fontSize: fitTitle(data.title || ctx.spec.title, 62, 40), bold: false, color: s.fg, margin: 0, fit: 'shrink' });
  slide.addText(data.subtitle || ctx.spec.subtitle || '', { x: 0.75, y: headY + 4.2, w: 6.2, h: 0.55, fontFace: FONTS.sansZh, fontSize: 17, color: s.fg, transparency: 18, margin: 0, fit: 'shrink' });
  addSwissBars(slide, 10.2, headY + 3.9, 1.55, 1.1, s.fg, 55);
  addFoot(slide, ctx, s.fg, 'swiss');
}
function swissStatement(slide, ctx, s) {
  const data = ctx.slideSpec;
  const headY = pageHeadY(ctx, 1.12);
  slide.addText(data.kicker || 'Statement', { x: 0.72, y: headY, w: 5.5, h: 0.25, fontFace: 'JetBrains Mono', fontSize: 8, charSpace: 1.8, color: s.fg, transparency: 35, margin: 0 });
  slide.addText(data.title, { x: 0.68, y: headY + 0.5, w: 6.35, h: 2.1, fontFace: FONTS.sans, fontSize: fitTitle(data.title, 43, 30), bold: false, color: s.fg, margin: 0, fit: 'shrink' });
  if (data.body || data.subtitle) {
    slide.addText(data.body || data.subtitle, { x: 0.72, y: headY + 3.22, w: 5.9, h: 0.95, fontFace: FONTS.sansZh, fontSize: 15.5, color: s.fg, transparency: 20, margin: 0, fit: 'shrink', valign: 'top' });
  }
  addStatementImageSlot(slide, ctx, { x: 7.38, y: headY + 0.74, w: 4.25, h: 3.65 }, s.fg, data.imageLabel || 'IMAGE SLOT');
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
    slide.addText(item.label || '', { x, y: 5.85, w: 2.3, h: 0.35, fontFace: 'JetBrains Mono', fontSize: 7.5, charSpace: 1.3, color: s.fg, transparency: 25, margin: 0, fit: 'shrink' });
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
    slide.addText(col.label || (i === 0 ? 'Before' : 'After'), { x: x + 0.35, y: 2.88, w: 4.7, h: 0.22, fontFace: 'JetBrains Mono', fontSize: 8, charSpace: 1.3, color, transparency: 30, margin: 0 });
    slide.addText(col.title || '', { x: x + 0.35, y: 3.35, w: 4.7, h: 0.55, fontFace: FONTS.sans, fontSize: 23, color, margin: 0, fit: 'shrink' });
    addBullets(slide, col.items || [], x + 0.35, 4.22, 4.7, 1.15, color, 14, 'swiss');
  });
  addFoot(slide, ctx, s.fg, 'swiss');
}

function swissTimeline(slide, ctx, s) {
  const data = ctx.slideSpec;
  addPageHead(slide, data, s.fg, 'swiss', 0.88);
  const items = (data.items || data.steps || []).slice(0, 6);
  const startX = 0.8;
  const endX = 12.0;
  const axisY = 4.25;
  slide.addShape(pptx.ShapeType.line, { x: startX, y: axisY, w: endX - startX, h: 0, line: { color: s.fg, transparency: 55, width: 0.7 } });
  items.forEach((item, i) => {
    const x = startX + (i * (endX - startX)) / Math.max(items.length - 1, 1);
    const labelX = clamp(x - 0.55, 0.32, SLIDE.w - 1.42);
    const titleX = clamp(x - 0.72, 0.32, SLIDE.w - 1.77);
    const descX = clamp(x - 0.82, 0.32, SLIDE.w - 1.97);
    slide.addShape(pptx.ShapeType.rect, { x: x - 0.04, y: axisY - 0.04, w: 0.08, h: 0.08, fill: { color: i === items.length - 1 ? ctx.theme.accent : s.fg }, line: { color: i === items.length - 1 ? ctx.theme.accent : s.fg } });
    slide.addText(item.label || item.year || String(i + 1).padStart(2, '0'), { x: labelX, y: 3.48, w: 1.1, h: 0.22, fontFace: 'JetBrains Mono', fontSize: 8, color: s.fg, transparency: 30, align: 'center', margin: 0, fit: 'shrink' });
    slide.addText(item.title || '', { x: titleX, y: 4.55, w: 1.45, h: 0.45, fontFace: FONTS.sansZh, fontSize: 10.5, bold: true, color: s.fg, align: 'center', margin: 0, fit: 'shrink' });
    slide.addText(item.desc || item.note || '', { x: descX, y: 5.08, w: 1.65, h: 0.48, fontFace: FONTS.sansZh, fontSize: 8.5, color: s.fg, transparency: 25, align: 'center', margin: 0, fit: 'shrink' });
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
    slide.addText(data.heroStat.label || '', { x: 4.1, y: 5.95, w: 5.5, h: 0.28, fontFace: 'JetBrains Mono', fontSize: 8, charSpace: 1.2, color: s.fg, transparency: 35, margin: 0, fit: 'shrink' });
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
  const titleFont = compact ? 13.5 : 17;
  const bodyFont = compact ? 8.6 : 10.2;
  const minCardH = compact ? 1.35 : 2.55;
  const demands = items.map((item) => {
    const title = item.title || item.label || '';
    const body = item.desc || item.note || item.body || '';
    const titleH = estimateTextHeight(title, cardW - 0.4, titleFont, { min: compact ? 0.34 : 0.48, max: compact ? 0.62 : 0.86 });
    const bodyH = estimateTextHeight(body, cardW - 0.4, bodyFont, { min: body ? 0.32 : 0, empty: 0, max: compact ? 0.82 : 1.35 });
    return Math.max(minCardH, (compact ? 0.88 : 1.25) + titleH + bodyH + 0.28);
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
    const titleH = estimateTextHeight(titleText, cardW - 0.4, titleFont, { min: compact ? 0.34 : 0.48, max: compact ? 0.58 : 0.82 });
    const descY = titleY + titleH + (compact ? 0.12 : 0.22);
    const descH = Math.max(0.2, cardH - (descY - y) - 0.22);
    slide.addShape(pptx.ShapeType.rect, { x, y, w: cardW, h: cardH, fill: { color: ctx.theme.grey1 }, line: { color: ctx.theme.grey1, transparency: 100 } });
    const hasIcon = addInlineIcon(slide, item, x + 0.2, y + 0.18, iconSize, s.fg, 'swiss', { fallback: item.icon ? null : 'layers', pad: compact ? 0.045 : 0.06 });
    if (!hasIcon) slide.addText(item.number || `0${i + 1}`, { x: x + 0.2, y: y + 0.22, w: 1.0, h: 0.22, fontFace: 'JetBrains Mono', fontSize: 8, color: s.fg, transparency: 45, margin: 0 });
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
    slide.addText(item.label || '', { x, y: 4.9, w: 1.75, h: 0.2, fontFace: 'JetBrains Mono', fontSize: 7, charSpace: 1, color: ctx.theme.ink, transparency: 35, margin: 0, fit: 'shrink' });
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
  const minH = rows <= 2 ? 1.2 : 1.05;
  const maxBottom = 6.35;
  const demands = sections.map((item) => {
    const title = item.title || '';
    const body = item.body || item.desc || '';
    const hasIconWidth = 0.62;
    const textW = w - hasIconWidth - 0.27;
    const titleH = estimateTextHeight(title, textW, 11.5, { min: 0.28, max: 0.52 });
    const bodyH = estimateTextHeight(body, textW, 8.5, { min: body ? 0.34 : 0, empty: 0, max: 0.78 });
    return Math.max(minH, 0.28 + titleH + 0.14 + bodyH + 0.24);
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
    if (!hasIcon) slide.addText(item.label || String(i + 1).padStart(2, '0'), { x: x + 0.18, y: y + 0.16, w: 0.52, h: 0.2, fontFace: 'JetBrains Mono', fontSize: 7.2, color: hot ? ctx.theme.accentOn : s.fg, transparency: hot ? 0 : 35, margin: 0 });
    const tx = hasIcon ? x + 0.62 : x + 0.78;
    const textW = x + w - 0.27 - tx;
    const titleText = item.title || '';
    const bodyText = item.body || item.desc || '';
    const titleH = estimateTextHeight(titleText, textW, 11.5, { min: 0.28, max: 0.5 });
    const bodyY = y + 0.18 + titleH + 0.14;
    const bodyH = Math.max(0.2, h - (bodyY - y) - 0.22);
    slide.addText(titleText, { x: tx, y: y + 0.15, w: textW, h: titleH, fontFace: FONTS.sansZh, fontSize: 11.5, bold: true, color: hot ? ctx.theme.accentOn : s.fg, margin: 0, fit: 'shrink', valign: 'top' });
    slide.addText(bodyText, { x: tx, y: bodyY, w: textW, h: bodyH, fontFace: FONTS.sansZh, fontSize: 8.5, color: hot ? ctx.theme.accentOn : s.fg, transparency: hot ? 10 : 30, margin: 0.02, fit: 'shrink', valign: 'top' });
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
  addChartBlock(slide, ctx, data.chart || data, { x: 0.78, y: 2.35, w: 8.3, h: 4.05 }, s, 'swiss');
  const insights = normalizeSections(data.insights || data.notes || []).slice(0, 3);
  insights.forEach((item, i) => {
    const y = 2.48 + i * 1.16;
    const iconColor = i === 0 ? ctx.theme.accent : s.fg;
    const hasIcon = addInlineIcon(slide, item, 9.55, y + 0.02, 0.34, iconColor, 'swiss', { fallback: defaultContentIcon(i, 'swiss'), pad: 0.055 });
    if (!hasIcon) slide.addText(item.value || item.label || `0${i + 1}`, { x: 9.55, y, w: 1.05, h: 0.38, fontFace: FONTS.sans, fontSize: 22, bold: true, color: i === 0 ? ctx.theme.accent : s.fg, margin: 0, fit: 'shrink' });
    slide.addText(item.title || '', { x: 10.35, y: y + 0.03, w: 1.85, h: 0.24, fontFace: FONTS.sansZh, fontSize: 9.6, bold: true, color: s.fg, margin: 0, fit: 'shrink' });
    slide.addText(item.body || item.desc || '', { x: 10.35, y: y + 0.34, w: 1.85, h: 0.34, fontFace: FONTS.sansZh, fontSize: 7.6, color: s.fg, transparency: 35, margin: 0, fit: 'shrink' });
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
    slide.addText(item.label || '', { x: tx, y: metricsY + 0.18, w: x + 1.88 - tx, h: 0.18, fontFace: 'JetBrains Mono', fontSize: 6.8, color: i === 0 ? ctx.theme.accentOn : s.fg, transparency: i === 0 ? 0 : 35, margin: 0, fit: 'shrink' });
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
  const gap = 0.24;
  const cardW = (11.45 - gap * (cols - 1)) / cols;
  const y0 = data.subtitle ? 2.85 : 2.52;
  items.forEach((item, i) => {
    const x = 0.78 + (i % cols) * (cardW + gap);
    const y = y0 + Math.floor(i / cols) * 1.42;
    const hot = i === data.highlightIndex;
    slide.addShape(pptx.ShapeType.rect, { x, y, w: cardW, h: 1.08, fill: { color: hot ? ctx.theme.accent : ctx.theme.grey1 }, line: { color: hot ? ctx.theme.accent : ctx.theme.grey1, transparency: 100 } });
    const color = hot ? ctx.theme.accentOn : s.fg;
    slide.addText(item.label || String(i + 1).padStart(2, '0'), { x: x + 0.18, y: y + 0.16, w: 0.62, h: 0.22, fontFace: 'JetBrains Mono', fontSize: 9.8, color, transparency: hot ? 0 : 35, margin: 0, fit: 'shrink' });
    slide.addText(item.title || '', { x: x + 0.18, y: y + 0.48, w: cardW - 0.36, h: 0.3, fontFace: FONTS.sansZh, fontSize: 12.6, bold: true, color, margin: 0, fit: 'shrink' });
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
    const labelX = clamp(x - 0.48, 0.65, SLIDE.w - 1.6);
    const titleX = clamp(x - 0.68, 0.65, SLIDE.w - 1.9);
    const descX = clamp(x - 0.76, 0.65, SLIDE.w - 2.0);
    slide.addText(item.label || item.date || `0${i + 1}`, { x: labelX, y: y - 0.28, w: 0.95, h: 0.22, fontFace: 'JetBrains Mono', fontSize: 9.8, color: hot ? ctx.theme.accent : s.fg, align: 'center', margin: 0, fit: 'shrink' });
    slide.addText(item.title || '', { x: titleX, y, w: 1.38, h: 0.35, fontFace: FONTS.sansZh, fontSize: 11.4, bold: true, color: s.fg, align: 'center', margin: 0, fit: 'shrink' });
    slide.addText(item.body || item.desc || item.note || '', { x: descX, y: y + 0.44, w: 1.52, h: 0.42, fontFace: FONTS.sansZh, fontSize: 9.2, color: s.fg, transparency: 30, align: 'center', margin: 0, fit: 'shrink' });
  });
  addFoot(slide, ctx, s.fg, 'swiss');
}

function swissSwimlane(slide, ctx, s) {
  const data = ctx.slideSpec;
  addPageHead(slide, data, s.fg, 'swiss', 0.78);
  const lanes = (data.lanes || data.sections || []).slice(0, 4);
  const stages = data.stages || data.columns || ['Now', 'Next', 'Later'];
  const x0 = 1.5;
  const y0 = 2.45;
  const laneH = 0.86;
  const colW = 10.3 / Math.max(stages.length, 1);
  stages.forEach((stage, i) => slide.addText(String(stage), { x: x0 + i * colW, y: 2.13, w: colW - 0.14, h: 0.22, fontFace: 'JetBrains Mono', fontSize: 9.8, color: s.fg, transparency: 35, margin: 0, fit: 'shrink' }));
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
  const contentTop = Math.max(2.58, headY + (hasSubtitle ? 2.04 : 1.72));
  const contentBottom = 6.35;
  const mediaBox = { x: 0.78, y: contentTop, w: 6.25, h: Math.max(2.55, contentBottom - contentTop) };
  addMediaOrChart(slide, ctx, data, mediaBox, s, 'swiss', 'MEDIA');
  const panelX = 7.42;
  const summaryH = isCmb ? 1.08 : 0.92;
  slide.addText(data.body || data.story || data.note || '', { x: panelX, y: contentTop + 0.05, w: 4.65, h: summaryH, fontFace: FONTS.sansZh, fontSize: isCmb ? 12 : 11.7, color: s.fg, transparency: 15, margin: 0.03, fit: 'shrink', valign: 'top' });
  const sideLimit = isCmb ? 4 : 5;
  const items = normalizeSections(data.items || data.insights || data.points || []).slice(0, sideLimit);
  const itemStartY = contentTop + (isCmb ? 1.42 : 1.18);
  const rowMin = isCmb ? 0.58 : 0.42;
  const rowMax = isCmb ? 0.8 : 0.56;
  const rowH = items.length ? clamp((contentBottom - itemStartY) / items.length, rowMin, rowMax) : (isCmb ? 0.66 : 0.52);
  items.forEach((item, i) => {
    const y = itemStartY + i * rowH;
    slide.addShape(pptx.ShapeType.line, { x: panelX, y, w: 4.4, h: 0, line: { color: s.fg, transparency: isCmb ? 74 : 68, width: 0.5 } });
    const hasIcon = addInlineIcon(slide, item, panelX, y + (isCmb ? 0.12 : 0.1), 0.21, i === 0 ? ctx.theme.accent : s.fg, 'swiss', { fallback: defaultContentIcon(i, 'swiss'), pad: 0.04 });
    const textX = hasIcon ? panelX + 0.34 : panelX;
    const textW = hasIcon ? 4.05 : 4.4;
    slide.addText(item.title || item.label || item.body || '', { x: textX, y: y + 0.08, w: textW, h: isCmb ? 0.28 : 0.23, fontFace: FONTS.sansZh, fontSize: isCmb ? 10.8 : 10.4, bold: true, color: s.fg, margin: 0, fit: 'shrink' });
    slide.addText(item.body || item.desc || item.note || '', { x: textX, y: y + (isCmb ? 0.42 : 0.32), w: textW, h: Math.max(isCmb ? 0.22 : 0.16, rowH - (isCmb ? 0.48 : 0.34)), fontFace: FONTS.sansZh, fontSize: isCmb ? 8.8 : 8.4, color: s.fg, transparency: 32, margin: 0.01, fit: 'shrink', valign: 'top' });
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
  slide.addText(data.kicker || '', { x: 0.78, y, w: 5.8, h: 0.25, fontFace: mode === 'swiss' ? 'JetBrains Mono' : FONTS.mono, fontSize: mode === 'swiss' ? 8 : 8.3, charSpace: mode === 'swiss' ? 1.6 : 2, color, transparency: 35, margin: 0, fit: 'shrink' });
  slide.addText(data.title || '', { x: 0.75, y: y + 0.45, w: 10.8, h: 0.9, fontFace: mode === 'swiss' ? FONTS.sans : FONTS.serifZh, fontSize: fitTitle(data.title || '', mode === 'swiss' ? 34 : 36, mode === 'swiss' ? 25 : 27), bold: mode !== 'swiss', color, margin: 0, fit: 'shrink' });
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
  slide.addImage({ data: svgIconData(icon, color), x: x + pad, y: y + pad, w: Math.max(0.01, size - pad * 2), h: Math.max(0.01, size - pad * 2) });
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
    slide.addText(String(number), { x, y: y - size * 0.07, w: size, h: size, fontFace: mode === 'swiss' ? 'JetBrains Mono' : FONTS.mono, fontSize: size * 38, bold: true, color, transparency: transparency ?? 0, margin: 0, fit: 'shrink', align: 'center', valign: 'mid' });
  }
}

function addImagePlaceholder(slide, x, y, w, h, color, label = 'IMAGE SLOT') {
  slide.addShape(pptx.ShapeType.rect, { x, y, w, h, fill: { color: 'FFFFFF', transparency: 100 }, line: { color, transparency: 45, width: 0.8, dash: 'dash' } });
  slide.addText(label, { x: x + 0.25, y: y + h / 2 - 0.12, w: Math.max(0.5, w - 0.5), h: 0.24, fontFace: FONTS.mono, fontSize: 8, charSpace: 1.8, color, transparency: 45, align: 'center', margin: 0 });
}

function addImageOrPlaceholder(slide, ctx, image, x, y, w, h, color, label) {
  const imgPath = resolveImage(ctx.specDir, image);
  if (imgPath) {
    slide.addImage({ path: imgPath, x, y, w, h });
    return;
  }
  addImagePlaceholder(slide, x, y, w, h, color, label || 'IMAGE SLOT');
}
function addStatementImageSlot(slide, ctx, box, color, label = 'IMAGE SLOT') {
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
  addImagePlaceholder(slide, box.x, box.y, box.w, box.h, s.fg, label || 'IMAGE SLOT');
  return 'placeholder';
}

function addMediaGrid(slide, ctx, data, boxes, s, mode) {
  const images = normalizeMediaImages(data);
  const charts = normalizeMediaCharts(data);
  const captions = normalizeSections(data.captions || data.items || data.sections || []);
  boxes.forEach((box, i) => {
    if (images[i]) {
      addImageOrPlaceholder(slide, ctx, images[i], box.x, box.y, box.w, box.h, s.fg, 'IMAGE SLOT');
    } else if (charts[i]) {
      addChartBlock(slide, ctx, { ...charts[i], x: box.x, y: box.y, w: box.w, h: box.h }, box, s, mode);
    } else {
      addImagePlaceholder(slide, box.x, box.y, box.w, box.h, s.fg, 'IMAGE SLOT');
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
  const skillRoot = path.resolve(__dirname, '..');
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
  slide.addText(text, { x, y, w, h: 0.25, fontFace: mode === 'swiss' ? 'JetBrains Mono' : FONTS.mono, fontSize: 7.5, charSpace: 1.2, color, transparency: 35, margin: 0, fit: 'shrink' });
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
      console.warn(`Warning: slide ${ctx.index + 1} skipped unpositioned ${block.type} block to avoid overlap. Put it in layout "media"/"mediaGrid" or set x/y/w/h explicitly.`);
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
    titleFontSize: chart.titleFontSize || 10,
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

function validateSpecSlots(spec, options = {}) {
  const errors = [];
  const warnings = [];
  spec.slides.forEach((slide, index) => {
    validateMediaSlots(slide, index, errors, warnings, options.specDir || process.cwd());
    validateTextSlots(slide, index, spec.style, errors, warnings);
  });
  warnings.forEach((message) => console.warn(message));
  if (errors.length) {
    fail(`Spec slot validation failed:\n- ${errors.join('\n- ')}`);
  }
}

function validateMediaSlots(slide, index, errors, warnings, specDir) {
  const layout = slide.layout || '';
  const mediaLayouts = ['statement', 'media', 'mediaGrid', 'gallery', 'imageGrid', 'imageHero', 'quoteImage', 'textImage', 'caseStudy'];
  if (!mediaLayouts.includes(layout)) return;
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
      warnings.push(`Warning: slide ${index + 1} image ${imageIndex + 1} path is missing or unsupported; an image placeholder will be rendered.`);
    }
  });
  if (isMediaGridLayout(layout) && !images.length && !charts.length && slotCount > 1) {
    warnings.push(`Warning: slide ${index + 1} has ${slotCount} media slot(s) but no images/charts; placeholders will be rendered.`);
  }
}

function validateTextSlots(slide, index, style, errors, warnings) {
  const layout = slide.layout || '';
  const sideMax = style === 'swiss' ? 5 : style === 'cmb' ? 4 : 3;
  const rules = {
    bigNumbers: [{ keys: ['items'], max: 6, min: 1, label: 'number cards' }],
    kpiTower: [{ keys: ['items'], max: 4, min: 1, label: 'KPI cards' }],
    pipeline: [{ keys: ['steps', 'items'], max: 6, min: 1, label: 'pipeline steps' }],
    timeline: [{ keys: ['items', 'steps'], max: 6, min: 1, label: 'timeline steps' }],
    matrix: [{ keys: ['items'], max: 12, min: 1, label: 'matrix cells' }],
    fourCards: [{ keys: ['items'], max: 8, min: 1, label: 'cards' }],
    article: [{ keys: ['sections', 'items', 'columns'], max: 6, min: 1, label: 'article sections' }],
    textGrid: [{ keys: ['sections', 'items', 'columns'], max: 9, min: 1, label: 'text grid cells' }],
    agenda: [{ keys: ['items', 'sections', 'agenda'], max: 8, min: 1, label: 'agenda items' }],
    pyramid: [{ keys: ['layers', 'items', 'sections'], max: 5, min: 1, label: 'pyramid layers' }],
    radial: [{ keys: ['items', 'nodes', 'sections'], max: 8, min: 1, label: 'radial nodes' }],
    roadmap: [{ keys: ['steps', 'items'], max: 6, min: 1, label: 'roadmap steps' }],
    swimlane: [{ keys: ['lanes', 'sections'], max: 4, min: 1, label: 'swimlanes' }],
    media: [{ keys: ['items', 'insights', 'points'], max: sideMax, min: 0, label: 'side points' }],
    mediaGrid: [{ keys: ['captions', 'items', 'sections'], max: 6, min: 0, label: 'media captions' }],
    gallery: [{ keys: ['captions', 'items', 'sections'], max: 6, min: 0, label: 'media captions' }],
    imageGrid: [{ keys: ['captions', 'items', 'sections'], max: 6, min: 0, label: 'media captions' }],
    imageHero: [{ keys: ['items'], max: 3, min: 0, label: 'image hero metrics' }],
    caseStudy: [{ keys: ['metrics', 'items'], max: 3, min: 0, label: 'case metrics' }],
    dataSheet: [{ keys: ['notes', 'insights'], max: style === 'swiss' ? 4 : 3, min: 0, label: 'side notes' }],
    chart: [{ keys: ['insights', 'notes'], max: style === 'swiss' ? 3 : 4, min: 0, label: 'chart insights' }],
    dashboard: [{ keys: ['metrics', 'items'], max: style === 'swiss' ? 5 : 4, min: 1, label: 'dashboard metrics' }],
  };
  const layoutRules = rules[layout] || [];
  validateIgnoredSlotFields(slide, index, layout, layoutRules, errors);
  layoutRules.forEach((rule) => validateSlotCollection(slide, index, rule, errors, warnings));
  if (layout === 'compare' || layout === 'duoCompare') {
    validateSlotCollection(slide.before || {}, index, { keys: ['items'], max: 6, min: 1, label: 'before items', prefix: 'before.' }, errors, warnings);
    validateSlotCollection(slide.after || {}, index, { keys: ['items'], max: 6, min: 1, label: 'after items', prefix: 'after.' }, errors, warnings);
  }
  if (layout === 'dashboard') validateChartSlots(slide, index, 2, errors, warnings);
  if (layout === 'chart') validateChartDataSlot(slide, index, errors, warnings);
  if (layout === 'dataSheet') validateTableSlot(slide, index, errors, warnings);
}

function validateSlotCollection(source, index, rule, errors, warnings) {
  const present = rule.keys.filter((key) => source[key] !== undefined && source[key] !== null);
  if (!present.length) {
    if (rule.min > 0) warnings.push(`Warning: slide ${index + 1} has no ${rule.label}; the layout may look empty.`);
    return [];
  }
  if (present.length > 1) {
    warnings.push(`Warning: slide ${index + 1} provides multiple fields for ${rule.label}: ${present.map((key) => `${rule.prefix || ''}${key}`).join(', ')}. The generator uses the first non-empty field only.`);
  }
  const key = present[0];
  const items = normalizeSlotItemsForValidation(source[key]);
  if (!items) {
    errors.push(`slide ${index + 1} field ${rule.prefix || ''}${key} has unsupported format; use an array of strings/objects or an object map.`);
    return [];
  }
  if (rule.max && items.length > rule.max) {
    errors.push(`slide ${index + 1} has ${items.length} ${rule.label}, but layout "${source.layout || 'nested'}" renders at most ${rule.max}; split content or change layout.`);
  }
  if (rule.min && items.length < rule.min) {
    warnings.push(`Warning: slide ${index + 1} has ${items.length} ${rule.label}; expected at least ${rule.min}.`);
  }
  items.forEach((item, itemIndex) => {
    if (!slotItemHasDisplayText(item)) {
      errors.push(`slide ${index + 1} ${rule.prefix || ''}${key}[${itemIndex}] has no displayable text/value field. Use text/title/label/body/desc/note/summary/value.`);
    }
  });
  return items;
}

function validateIgnoredSlotFields(slide, index, layout, layoutRules, errors) {
  const valid = new Set(layoutRules.flatMap((rule) => rule.keys));
  const common = ['items', 'sections', 'columns', 'steps', 'nodes', 'layers', 'metrics', 'notes', 'insights', 'agenda', 'lanes', 'captions', 'points'];
  const ignored = common.filter((key) => slide[key] !== undefined && slide[key] !== null && !valid.has(key));
  if (ignored.length && layoutRules.length) {
    errors.push(`slide ${index + 1} layout "${layout}" does not render field(s): ${ignored.join(', ')}. Rename them to one of: ${Array.from(valid).join(', ')}.`);
  }
  if (layout === 'dashboard' && slide.chart && !Array.isArray(slide.charts)) {
    errors.push(`slide ${index + 1} dashboard uses charts[]; field chart will not be rendered. Rename chart to charts: [chart].`);
  }
}

function validateChartSlots(slide, index, max, errors, warnings) {
  const charts = normalizeMediaCharts(slide);
  if (charts.length > max) errors.push(`slide ${index + 1} has ${charts.length} charts, but dashboard renders at most ${max}.`);
  if (!charts.length) warnings.push(`Warning: slide ${index + 1} dashboard has no charts; chart area will be empty.`);
}

function validateChartDataSlot(slide, index, errors, warnings) {
  const chart = slide.chart || slide;
  if (!normalizeChartData(chart).length) warnings.push(`Warning: slide ${index + 1} chart layout has no chart data; a NO DATA box will be rendered.`);
}

function validateTableSlot(slide, index, errors, warnings) {
  if (!slide.table) {
    warnings.push(`Warning: slide ${index + 1} dataSheet has no table.`);
    return;
  }
  if (slide.table.headers && !Array.isArray(slide.table.headers)) errors.push(`slide ${index + 1} table.headers must be an array.`);
  const rows = normalizeTableRows(slide.table);
  if (!rows.length) warnings.push(`Warning: slide ${index + 1} table has no rows.`);
}

function normalizeSlotItemsForValidation(value) {
  if (value === undefined || value === null) return [];
  if (typeof value === 'string' || typeof value === 'number') return [{ body: String(value) }];
  if (Array.isArray(value)) return value;
  if (typeof value === 'object') return Object.entries(value).map(([title, body]) => ({ title, body: String(body ?? '') }));
  return null;
}

function slotItemHasDisplayText(item) {
  if (typeof item === 'string' || typeof item === 'number') return String(item).trim().length > 0;
  if (!item || typeof item !== 'object') return false;
  return ['text', 'title', 'label', 'body', 'desc', 'note', 'summary', 'detail', 'value', 'unit', 'metric', 'name'].some((key) => String(item[key] ?? '').trim().length > 0);
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
function warnThinContent(spec) {
  spec.slides.forEach((slide, index) => {
    const candidates = normalizeSections(slide.sections || slide.items || slide.columns || slide.nodes || slide.layers || slide.steps || slide.milestones || slide.agenda || []);
    if (candidates.length < 3) return;
    const titleOnly = candidates.filter((item) => {
      const title = item.title || item.label || item.name;
      const body = item.body || item.desc || item.note || item.text || item.summary || item.detail;
      return title && !body;
    }).length;
    if (titleOnly >= Math.ceil(candidates.length * 0.6)) {
      console.warn(`Warning: slide ${index + 1} has ${titleOnly}/${candidates.length} title-only items. Add body/desc/note for each point so the page is not just headings.`);
    }
  });
}
function diversifyRepeatedLayouts(spec, options = {}) {
  if (spec.preserveLayouts || spec.lockLayouts || spec.disableLayoutDiversify) return [];
  const mutate = options.mutate === true;
  const replacements = {
    textGrid: ['agenda', 'fourCards', 'matrix', 'radial'],
    article: ['agenda', 'fourCards', 'caseStudy', 'radial'],
    fourCards: ['textGrid', 'agenda', 'matrix', 'radial'],
    matrix: ['textGrid', 'agenda', 'fourCards', 'radial'],
    statement: ['agenda', 'caseStudy', 'radial'],
    compare: ['swimlane', 'caseStudy', 'matrix'],
    timeline: ['roadmap', 'swimlane', 'agenda'],
    pipeline: ['roadmap', 'swimlane', 'agenda'],
  };
  let previous = null;
  let runLength = 0;
  const diversifyCounts = new Map();
  const changes = [];
  spec.slides.forEach((slide, index) => {
    const layout = slide.layout || (spec.style === 'magazine' ? 'textImage' : 'statement');
    if (slide.lockLayout || slide.preserveLayout) {
      previous = layout;
      runLength = 1;
      return;
    }
    if (layout === previous) {
      runLength += 1;
    } else {
      previous = layout;
      runLength = 1;
      return;
    }
    if (runLength < 2) return;
    const replacementCount = diversifyCounts.get(layout) || 0;
    const nextLayout = chooseDiversifiedLayout(slide, layout, replacementCount, replacements[layout] || ['agenda', 'caseStudy', 'radial', 'roadmap', 'swimlane']);
    if (nextLayout && nextLayout !== layout) {
      changes.push({ slide: index + 1, from: layout, to: nextLayout });
      diversifyCounts.set(layout, replacementCount + 1);
      previous = nextLayout;
      runLength = 1;
      if (mutate) {
        slide.originalLayout = slide.originalLayout || layout;
        slide.layout = nextLayout;
        console.warn(`Info: slide ${index + 1} layout changed from "${layout}" to "${nextLayout}". A normalized spec will be written so JSON and PPTX stay in sync.`);
      } else {
        console.warn(`Suggestion: slide ${index + 1} repeats layout "${layout}"; consider changing it to "${nextLayout}" in JSON. The generator did not auto-change it, so JSON and PPTX stay consistent.`);
      }
    }
  });
  return changes;
}

function chooseDiversifiedLayout(slide, currentLayout, replacementCount, candidates) {
  if (slide.layoutAlt) return slide.layoutAlt;
  if ((slide.lanes || slide.stages) && currentLayout !== 'swimlane') return 'swimlane';
  if ((slide.steps || slide.milestones) && currentLayout !== 'roadmap') return 'roadmap';
  if ((slide.image || slide.metrics) && currentLayout !== 'caseStudy') return 'caseStudy';
  const items = normalizeSections(slide.sections || slide.items || slide.columns || slide.nodes || []);
  const count = items.length;
  const offset = replacementCount % candidates.length;
  if (['textGrid', 'article', 'fourCards', 'matrix'].includes(currentLayout) && count >= 3) return candidates[offset];
  if (count >= 7 && currentLayout !== 'matrix') return 'matrix';
  if (count >= 5 && currentLayout !== 'agenda') return 'agenda';
  if (count >= 3 && count <= 8 && currentLayout !== 'fourCards') return 'fourCards';
  return candidates[offset];
}
function warnLayoutVariety(spec) {
  let runLayout = null;
  let runStart = 0;
  let runLength = 0;
  spec.slides.forEach((slide, index) => {
    const layout = slide.layout || (spec.style === 'magazine' ? 'textImage' : 'statement');
    if (layout === runLayout) {
      runLength += 1;
    } else {
      if (runLength >= 3) {
        console.warn(`Warning: slides ${runStart + 1}-${runStart + runLength} use layout "${runLayout}" consecutively; consider alternating layouts for visual variety.`);
      }
      runLayout = layout;
      runStart = index;
      runLength = 1;
    }
  });
  if (runLength >= 3) {
    console.warn(`Warning: slides ${runStart + 1}-${runStart + runLength} use layout "${runLayout}" consecutively; consider alternating layouts for visual variety.`);
  }
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

