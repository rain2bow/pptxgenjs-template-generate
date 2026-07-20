'use strict';

module.exports = function createPairedLayoutRenderers(api, style) {
  const {
    pptx,
    FONTS,
    READABILITY,
    addPageHead,
    addFoot,
    addImageOrPlaceholder,
    normalizeMediaImages,
    normalizeSections,
    estimateTextHeight,
  } = api;

  function renderPairedText(slide, ctx, state) {
    const data = ctx.slideSpec;
    addPageHead(slide, data, state.fg, style, 0.82);
    const body = data.body || data.summary || data.story || data.note || '';
    const items = normalizeSections(data.items || []).slice(0, 8);
    const top = body ? 2.82 : 2.28;
    if (body) {
      slide.addText(body, {
        x: 0.78, y: 2.18, w: 11.72, h: 0.48,
        fontFace: FONTS.sansZh, fontSize: 14, color: state.fg,
        transparency: 18, margin: 0.02, fit: 'shrink', valign: 'top',
      });
    }
    addContentCards(slide, ctx, state, items, { x: 0.78, y: top, w: 11.72, h: 6.28 - top });
    addFoot(slide, ctx, state.fg, style);
  }

  function renderPairedMedia(slide, ctx, state) {
    const data = ctx.slideSpec;
    addPageHead(slide, data, state.fg, style, 0.82);
    const images = normalizeMediaImages(data).slice(0, 6);
    const items = normalizeSections(data.items || []).slice(0, 8);
    const body = data.body || data.summary || data.story || data.note || data.center || '';
    addImagePanel(slide, ctx, state, images, { x: 0.78, y: 2.24, w: 4.18, h: 3.98 });
    const bodyTop = 2.24;
    const cardsTop = body ? 2.92 : bodyTop;
    if (body) {
      slide.addText(body, {
        x: 5.36, y: bodyTop, w: 7.14, h: 0.52,
        fontFace: FONTS.sansZh, fontSize: 14, color: state.fg,
        transparency: 16, margin: 0.02, fit: 'shrink', valign: 'top',
      });
    }
    addContentCards(slide, ctx, state, items, { x: 5.36, y: cardsTop, w: 7.14, h: 6.28 - cardsTop });
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

  function addContentCards(slide, ctx, state, items, box) {
    if (!items.length) return;
    const cols = items.length <= 3 ? 1 : 2;
    const rows = Math.ceil(items.length / cols);
    const gapX = 0.22;
    const gapY = 0.16;
    const cardW = (box.w - gapX * (cols - 1)) / cols;
    const cardH = (box.h - gapY * (rows - 1)) / rows;
    items.forEach((item, index) => {
      const x = box.x + (index % cols) * (cardW + gapX);
      const y = box.y + Math.floor(index / cols) * (cardH + gapY);
      addContentCard(slide, ctx, state, item, index, { x, y, w: cardW, h: cardH });
    });
  }

  function addContentCard(slide, ctx, state, item, index, box) {
    const colors = cardColors(ctx, state, index);
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

  function itemBody(item) {
    const direct = item.body || item.desc || item.note || item.summary || item.detail || item.text;
    if (direct) return String(direct);
    if (!Array.isArray(item.items)) return '';
    return item.items.map((entry) => {
      if (typeof entry === 'string' || typeof entry === 'number') return String(entry);
      return entry?.text || entry?.title || entry?.label || '';
    }).filter(Boolean).join(' · ');
  }

  function cardColors(ctx, state, index) {
    if (style === 'cmb') {
      const hot = index === 0;
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
        fill: index === 0 ? ctx.theme.accent : ctx.theme.grey1,
        line: index === 0 ? ctx.theme.accent : state.fg,
        lineTransparency: index === 0 ? 100 : 78,
        text: index === 0 ? ctx.theme.accentOn : state.fg,
        bodyTransparency: index === 0 ? 6 : 24,
      };
    }
    return {
      fill: index === 0 ? ctx.theme.ink : ctx.theme.paperTint,
      line: state.fg,
      lineTransparency: index === 0 ? 100 : 76,
      text: index === 0 ? ctx.theme.paper : state.fg,
      bodyTransparency: index === 0 ? 8 : 22,
    };
  }

  return { renderPairedText, renderPairedMedia };
};
