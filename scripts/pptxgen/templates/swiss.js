'use strict';

module.exports = function createSwissTemplate(api) {
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
    if (data.callout) addCallout(slide, data.callout, 0.72, headY + 4.35, 5.9, 0.58, s.fg, ctx.theme.grey1 || ctx.theme.paperTint);
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
    slide.addText(data.body || data.summary || data.story || data.subtitle || '', { x: 0.78, y: 4.42, w: 4.8, h: 0.78, fontFace: FONTS.sansZh, fontSize: 11, color: ctx.theme.ink, transparency: 15, margin: 0, fit: 'shrink' });
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
    swissStatement(slide, { ...ctx, slideSpec: { ...data, title: data.quote || data.title, body: data.body || data.cite || data.source || data.subtitle || '' } }, s);
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
    slide.addText(data.body || data.summary || data.story || data.note || '', { x: panelX, y: contentTop + 0.05, w: 4.65, h: summaryH, fontFace: FONTS.sansZh, fontSize: isCmb ? 12 : 11.7, color: s.fg, transparency: 15, margin: 0.03, fit: 'shrink', valign: 'top' });
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


  return {
    render: renderSwiss,
  };
};
