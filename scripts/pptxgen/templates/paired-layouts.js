'use strict';

module.exports = function createPairedLayoutRenderers(api, style) {
  const {
    pptx,
    FONTS,
    READABILITY,
    addPageHead,
    addCallout,
    addFoot,
    addImageOrPlaceholder,
    normalizeMediaImages,
    normalizeSections,
    estimateTextHeight,
  } = api;

  function renderPairedStatementText(slide, ctx, state) {
    const data = ctx.slideSpec;
    addPageHead(slide, data, state.fg, style, 0.82);
    if (data.body) {
      slide.addText(data.body, {
        x: 0.82, y: 2.55, w: 8.9, h: 1.15,
        fontFace: FONTS.sansZh, fontSize: 16, color: state.fg,
        transparency: 14, margin: 0.02, valign: 'top',
      });
    }
    if (data.callout) addCallout(slide, data.callout, 0.82, 4.15, 8.9, 1.05, state.fg, ctx.theme.paperTint || ctx.theme.grey1 || ctx.theme.paper);
    addFoot(slide, ctx, state.fg, style);
  }

  function renderPairedStatementMedia(slide, ctx, state) {
    const data = ctx.slideSpec;
    addPageHead(slide, data, state.fg, style, 0.82);
    if (data.body) {
      slide.addText(data.body, {
        x: 0.82, y: 2.55, w: 5.7, h: 1.15,
        fontFace: FONTS.sansZh, fontSize: 16, color: state.fg,
        transparency: 14, margin: 0.02, valign: 'top',
      });
    }
    if (data.callout) addCallout(slide, data.callout, 0.82, 4.15, 5.7, 1.05, state.fg, ctx.theme.paperTint || ctx.theme.grey1 || ctx.theme.paper);
    const images = normalizeMediaImages(data).slice(0, 1);
    addImagePanel(slide, ctx, state, images, { x: 7.25, y: 2.18, w: 5.05, h: 3.85 });
    addFoot(slide, ctx, state.fg, style);
  }

  function renderPairedQuoteText(slide, ctx, state) {
    const data = ctx.slideSpec;
    addPageHead(slide, { ...data, title: data.title || data.quote }, state.fg, style, 0.82);
    const quote = data.quote || data.title || '';
    if (quote) {
      slide.addText(quote, {
        x: 0.82, y: 2.28, w: 10.9, h: 1.45,
        fontFace: FONTS.sansZh, fontSize: 20, bold: true, color: state.fg,
        margin: 0.02, valign: 'mid',
      });
    }
    const explanation = data.body || data.cite || data.source || '';
    if (explanation) {
      slide.addText(explanation, {
        x: 0.86, y: 4.0, w: 7.45, h: 0.62,
        fontFace: FONTS.sansZh, fontSize: 14, color: state.fg,
        transparency: 20, margin: 0.02, valign: 'top',
      });
    }
    if (data.callout) addCallout(slide, data.callout, 8.65, 3.9, 3.65, 0.9, state.fg, ctx.theme.paperTint || ctx.theme.grey1 || ctx.theme.paper);
    const attribution = [data.caption, data.source, data.cite].filter(Boolean).join(' · ');
    if (attribution) {
      slide.addText(attribution, {
        x: 0.86, y: 5.18, w: 8.6, h: 0.32,
        fontFace: FONTS.sansZh, fontSize: 12, color: state.fg,
        transparency: 34, margin: 0,
      });
    }
    addFoot(slide, ctx, state.fg, style);
  }

  function renderPairedQuoteMedia(slide, ctx, state) {
    const data = ctx.slideSpec;
    addPageHead(slide, { ...data, title: data.title || data.quote }, state.fg, style, 0.82);
    const quote = data.quote || data.title || '';
    if (quote) {
      slide.addText(quote, {
        x: 0.82, y: 2.28, w: 6.0, h: 1.38,
        fontFace: FONTS.sansZh, fontSize: 20, bold: true, color: state.fg,
        margin: 0.02, valign: 'mid',
      });
    }
    const explanation = data.body || data.cite || data.source || '';
    if (explanation) {
      slide.addText(explanation, {
        x: 0.86, y: 3.92, w: 5.95, h: 0.62,
        fontFace: FONTS.sansZh, fontSize: 14, color: state.fg,
        transparency: 20, margin: 0.02, valign: 'top',
      });
    }
    if (data.callout) addCallout(slide, data.callout, 0.86, 4.82, 5.95, 0.82, state.fg, ctx.theme.paperTint || ctx.theme.grey1 || ctx.theme.paper);
    const images = normalizeMediaImages(data).slice(0, 1);
    addImagePanel(slide, ctx, state, images, { x: 7.35, y: 2.18, w: 4.95, h: 3.55 });
    const attribution = [data.caption, data.source, data.cite].filter(Boolean).join(' · ');
    if (attribution) {
      slide.addText(attribution, {
        x: 7.35, y: 5.88, w: 4.95, h: 0.34,
        fontFace: FONTS.sansZh, fontSize: 12, color: state.fg,
        transparency: 34, margin: 0,
      });
    }
    addFoot(slide, ctx, state.fg, style);
  }

  function renderPairedText(slide, ctx, state) {
    const data = ctx.slideSpec;
    addPageHead(slide, data, state.fg, style, 0.82);
    const body = data.body || data.note || '';
    const items = normalizeSections(data.items || []).slice(0, Number(data.maxItems) || 8);
    const lead = data.caseTitle || '';
    const stages = normalizeStages(data.stages);
    let top = lead || body ? 2.82 : 2.28;
    if (lead) {
      slide.addText(lead, { x: 0.78, y: 2.12, w: 4.2, h: 0.34, fontFace: FONTS.sansZh, fontSize: 16, bold: true, color: state.fg, margin: 0 });
    }
    if (body) {
      slide.addText(body, {
        x: lead ? 5.05 : 0.78, y: 2.18, w: lead ? 7.45 : 11.72, h: 0.48,
        fontFace: FONTS.sansZh, fontSize: 14, color: state.fg,
        transparency: 18, margin: 0.02, fit: 'shrink', valign: 'top',
      });
    }
    if (stages.length) {
      addStageStrip(slide, stages, { x: 0.78, y: top, w: 11.72, h: 0.34 }, state);
      top += 0.5;
    }
    addContentCards(slide, ctx, state, items, { x: 0.78, y: top, w: 11.72, h: 6.28 - top }, data, 3);
    addFoot(slide, ctx, state.fg, style);
  }

  function renderPairedMedia(slide, ctx, state) {
    const data = ctx.slideSpec;
    addPageHead(slide, data, state.fg, style, 0.82);
    const images = normalizeMediaImages(data).slice(0, 6);
    const items = normalizeSections(data.items || []).slice(0, Number(data.maxItems) || 8);
    const body = data.body || data.note || data.center || '';
    const lead = data.caseTitle || '';
    const stages = normalizeStages(data.stages);
    addImagePanel(slide, ctx, state, images, { x: 0.78, y: 2.24, w: 4.18, h: 3.98 });
    const bodyTop = 2.24;
    let cardsTop = lead || body ? 2.92 : bodyTop;
    if (lead) {
      slide.addText(lead, { x: 5.36, y: bodyTop, w: 2.35, h: 0.34, fontFace: FONTS.sansZh, fontSize: 16, bold: true, color: state.fg, margin: 0 });
    }
    if (body) {
      slide.addText(body, {
        x: lead ? 7.9 : 5.36, y: bodyTop, w: lead ? 4.6 : 7.14, h: 0.52,
        fontFace: FONTS.sansZh, fontSize: 14, color: state.fg,
        transparency: 16, margin: 0.02, fit: 'shrink', valign: 'top',
      });
    }
    if (stages.length) {
      addStageStrip(slide, stages, { x: 5.36, y: cardsTop, w: 7.14, h: 0.34 }, state);
      cardsTop += 0.5;
    }
    addContentCards(slide, ctx, state, items, { x: 5.36, y: cardsTop, w: 7.14, h: 6.28 - cardsTop }, data, 2);
    addFoot(slide, ctx, state.fg, style);
  }

  function addImagePanel(slide, ctx, state, images, box) {
    const count = Math.max(1, images.length);
    const cols = count === 1 ? 1 : 2;
    const rows = Math.ceil(count / cols);
    const gap = 0.12;
    const cellW = (box.w - gap * (cols - 1)) / cols;
    const cellH = (box.h - gap * (rows - 1)) / rows;
    images.forEach((image, index) => {
      const x = box.x + (index % cols) * (cellW + gap);
      const y = box.y + Math.floor(index / cols) * (cellH + gap);
      addImageOrPlaceholder(slide, ctx, image, x, y, cellW, cellH, state.fg, '图片');
    });
  }

  function addContentCards(slide, ctx, state, items, box, data = {}, maxColumns = 2) {
    if (!items.length) return;
    const requestedColumns = Number(data.columnsCount);
    const cols = Number.isFinite(requestedColumns)
      ? Math.max(1, Math.min(maxColumns, Math.round(requestedColumns)))
      : Math.min(maxColumns, items.length <= 3 ? 1 : 2);
    const rows = Math.ceil(items.length / cols);
    const gapX = 0.22;
    const gapY = 0.16;
    const cardW = (box.w - gapX * (cols - 1)) / cols;
    const cardH = (box.h - gapY * (rows - 1)) / rows;
    items.forEach((item, index) => {
      const x = box.x + (index % cols) * (cardW + gapX);
      const y = box.y + Math.floor(index / cols) * (cardH + gapY);
      addContentCard(slide, ctx, state, item, index, { x, y, w: cardW, h: cardH }, resolveHighlightIndex(data, items));
    });
  }

  function addContentCard(slide, ctx, state, item, index, box, highlightIndex) {
    const colors = cardColors(ctx, state, index, highlightIndex);
    slide.addShape(pptx.ShapeType.rect, {
      ...box,
      fill: { color: colors.fill },
      line: { color: colors.line, transparency: colors.lineTransparency, width: 0.65 },
    });
    const title = item.title || item.label || item.name || item.value || `要点 ${index + 1}`;
    const body = itemBody(item);
    const titleH = estimateTextHeight(title, box.w - 0.4, 16, { min: 0.34, max: 0.62, lineHeight: 1.25, padding: 0.02 });
    const titleY = box.y + 0.18;
    slide.addText(title, {
      x: box.x + 0.2, y: titleY, w: box.w - 0.4, h: titleH,
      fontFace: FONTS.sansZh, fontSize: Math.max(16, READABILITY.minFontSize), bold: true,
      color: colors.text, margin: 0, fit: 'shrink', valign: 'top',
    });
    if (!body) return;
    const bodyY = titleY + titleH + 0.1;
    slide.addText(body, {
      x: box.x + 0.2, y: bodyY, w: box.w - 0.4, h: Math.max(0.32, box.y + box.h - bodyY - 0.18),
      fontFace: FONTS.sansZh, fontSize: Math.max(12, READABILITY.minFontSize),
      color: colors.text, transparency: colors.bodyTransparency,
      margin: 0.02, fit: 'shrink', valign: 'top', breakLine: false,
    });
  }

  function addStageStrip(slide, stages, box, state) {
    const cellW = box.w / stages.length;
    stages.forEach((stage, index) => {
      slide.addText(stage, {
        x: box.x + index * cellW, y: box.y, w: cellW - 0.08, h: box.h,
        fontFace: FONTS.sansZh, fontSize: 12, bold: true, color: state.fg,
        transparency: 18, align: 'center', margin: 0,
      });
    });
  }

  function normalizeStages(stages) {
    return Array.isArray(stages) ? stages.map((stage) => String(stage || '').trim()).filter(Boolean).slice(0, 8) : [];
  }

  function resolveHighlightIndex(data, items) {
    if (Number.isInteger(data.highlightIndex)) return Math.max(0, Math.min(items.length - 1, data.highlightIndex));
    if (data.highlightLast) return Math.max(0, items.length - 1);
    return 0;
  }

  function itemBody(item) {
    const direct = item.body || item.desc || item.note || item.summary || item.detail || item.text;
    if (direct) return String(direct);
    if (!Array.isArray(item.items)) return '';
    return item.items.map((entry) => {
      if (typeof entry === 'string' || typeof entry === 'number') return String(entry);
      return entry?.text || entry?.title || entry?.label || '';
    }).filter(Boolean).join(' · ');
  }

  function cardColors(ctx, state, index, highlightIndex) {
    const hot = index === highlightIndex;
    if (style === 'cmb') {
      return {
        fill: hot ? ctx.theme.accent : ctx.theme.grey1,
        line: hot ? ctx.theme.accent : ctx.theme.grey3,
        lineTransparency: hot ? 100 : 72,
        text: hot ? ctx.theme.accentOn : state.fg,
        bodyTransparency: hot ? 8 : 25,
      };
    }
    if (style === 'swiss') {
      return {
        fill: hot ? ctx.theme.accent : ctx.theme.grey1,
        line: hot ? ctx.theme.accent : state.fg,
        lineTransparency: hot ? 100 : 78,
        text: hot ? ctx.theme.accentOn : state.fg,
        bodyTransparency: hot ? 6 : 24,
      };
    }
    return {
      fill: hot ? ctx.theme.ink : ctx.theme.paperTint,
      line: state.fg,
      lineTransparency: hot ? 100 : 76,
      text: hot ? ctx.theme.paper : state.fg,
      bodyTransparency: hot ? 8 : 22,
    };
  }

  return { renderPairedStatementText, renderPairedStatementMedia, renderPairedQuoteText, renderPairedQuoteMedia, renderPairedText, renderPairedMedia };
};
