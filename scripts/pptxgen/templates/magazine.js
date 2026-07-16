'use strict';

module.exports = function createMagazineTemplate(api) {
  const {
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
  } = api;


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
    slide.addText(data.body || data.summary || data.story || data.note || '', { x: 0.78, y: 2.55, w: 5.35, h: 1.18, fontFace: FONTS.sansZh, fontSize: 13.2, color: s.fg, transparency: 16, margin: 0.03, fit: 'shrink', valign: 'top' });
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
    const gapY = rows > 1 ? 0.84 : 0.3;
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
    slide.addText(data.body || data.cite || data.source || '', { x: 0.82, y: 4.55, w: 7.2, h: 0.6, fontFace: FONTS.serifZh, fontSize: 16, color: s.fg, transparency: 30, margin: 0, fit: 'shrink' });
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
    slide.addText(data.body || data.summary || data.story || '', { x: 6.78, y: 3.05, w: 4.95, h: 1.15, fontFace: FONTS.sansZh, fontSize: 11, color: s.fg, transparency: 18, margin: 0.03, fit: 'shrink', valign: 'top' });
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


  return {
    render: renderMagazine,
  };
};
