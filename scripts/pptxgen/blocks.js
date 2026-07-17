'use strict';

module.exports = function createBlockTools(deps) {
  const {
    pptx,
    FONTS,
    TYPOGRAPHY,
    READABILITY,
    CHART_TYPES,
    fail,
    safeBox,
    chartPalette,
  } = deps;

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
    if (chart.caption) addBlockCaption(slide, chart.caption, box.x, box.y + box.h + 0.08, box.w, s.fg, mode);
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
    if (table.caption) addBlockCaption(slide, table.caption, box.x, box.y + box.h + 0.08, box.w, s.fg, mode);
  }

  function addBlockCaption(slide, text, x, y, w, color, mode) {
    if (!text) return;
    slide.addText(String(text), {
      x,
      y,
      w,
      h: 0.42,
      fontFace: mode === 'swiss' ? FONTS.sansZh : FONTS.sansZh,
      fontSize: TYPOGRAPHY.dense,
      color,
      transparency: 20,
      margin: 0,
      valign: 'top',
    });
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


  return {
    renderDataBlocks,
    hasExplicitBox,
    addTextBlock,
    addChartBlock,
    addTableBlock,
    normalizeChartData,
    normalizeTableRows,
  };
};
