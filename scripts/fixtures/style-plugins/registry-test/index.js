'use strict';

module.exports = {
  id: 'registry-test',
  name: '注册机制测试风格',
  description: '仅用于验证新增目录即可注册 style，不属于内置生产模板。',
  defaultTheme: 'default',
  themes: {
    default: {
      name: '注册测试主题',
      paper: 'F7F8FA',
      paperTint: 'E9EDF2',
      ink: '20242A',
      grey1: 'E9EDF2',
      grey2: 'C7CED8',
      grey3: '66707D',
      accent: '2F6B5F',
      accentOn: 'FFFFFF',
      chartColors: ['2F6B5F', '527F9A', '88929E', 'C7CED8'],
    },
  },
  createTemplate(api) {
    const { pptx, SLIDE, FONTS, normalizeSections } = api;
    return {
      render(slide, ctx) {
        const data = ctx.slideSpec;
        const source = ctx.sourceSlideSpec || data;
        slide.background = { color: ctx.theme.paper };
        slide.addShape(pptx.ShapeType.rect, { x: 0, y: 0, w: 0.22, h: SLIDE.h, fill: { color: ctx.theme.accent }, line: { color: ctx.theme.accent, transparency: 100 } });
        slide.addText(source.kicker || 'CUSTOM STYLE', { x: 0.72, y: 0.62, w: 4.8, h: 0.28, fontFace: FONTS.zh, fontSize: 12, color: ctx.theme.accent, margin: 0 });
        slide.addText(source.title || ctx.spec.title || '', { x: 0.72, y: 1.16, w: 11.6, h: 0.78, fontFace: FONTS.zh, fontSize: 28, bold: true, color: ctx.theme.ink, margin: 0 });
        const body = source.body || source.subtitle || '';
        if (body) slide.addText(body, { x: 0.76, y: 2.18, w: 10.8, h: 0.72, fontFace: FONTS.zh, fontSize: 14, color: ctx.theme.ink, transparency: 15, margin: 0.02 });
        const items = normalizeSections(source.items || []).slice(0, 8);
        items.forEach((item, index) => {
          const y = 3.08 + index * 0.43;
          slide.addText(item.title || item.label || item.value || String(item), { x: 0.78, y, w: 3.0, h: 0.3, fontFace: FONTS.zh, fontSize: 14, bold: true, color: ctx.theme.ink, margin: 0 });
          slide.addText(item.body || item.desc || item.note || '', { x: 3.95, y, w: 7.6, h: 0.3, fontFace: FONTS.zh, fontSize: 12, color: ctx.theme.ink, transparency: 20, margin: 0 });
        });
        slide.addText(`STYLE / ${ctx.spec.style}`, { x: 10.2, y: 6.92, w: 2.3, h: 0.2, fontFace: FONTS.en, fontSize: 9, color: ctx.theme.grey3, align: 'right', margin: 0 });
      },
    };
  },
};
