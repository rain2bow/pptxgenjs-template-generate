'use strict';

module.exports = function createIconTools(deps) {
  const {
    fs,
    path,
    process,
    pptx,
    FONTS,
    BASIC_ICON_NAMES,
    ICON_ALIASES,
    LUCIDE_STATIC_ICON_DIR,
    fail,
  } = deps;

  const LUCIDE_ICON_CACHE = new Map();

  const ICON_PNG_CACHE = new Map();

  let LUCIDE_MODULE = undefined;

  let SHARP_MODULE = undefined;

  let ICON_RENDER_MODE = 'png';

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


  return {
    iconAlias,
    prepareIconAssets,
    normalizeHex,
    addSvgIcon,
    defaultContentIcon,
    itemIcon,
    addInlineIcon,
    addBulletIcon,
  };
};
