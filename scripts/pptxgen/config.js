const path = require('node:path');

const SLIDE = {
  w: 13.333,
  h: 7.5,
  marginX: 0.62,
  marginTop: 0.42,
  marginBottom: 0.45,
};

const THEMES = {
  magazine: {
    ink: {
      name: '墨水经典',
      ink: '0A0A0B',
      paper: 'F1EFEA',
      paperTint: 'E8E5DE',
      inkTint: '18181A',
    },
    indigo: {
      name: '靛蓝瓷',
      ink: '0A1F3D',
      paper: 'F1F3F5',
      paperTint: 'E4E8EC',
      inkTint: '152A4A',
    },
    forest: {
      name: '森林墨',
      ink: '1A2E1F',
      paper: 'F5F1E8',
      paperTint: 'ECE7DA',
      inkTint: '253D2C',
    },
    kraft: {
      name: '牛皮纸',
      ink: '2A1E13',
      paper: 'EEDFC7',
      paperTint: 'E0D0B6',
      inkTint: '3A2A1D',
    },
    dune: {
      name: '沙丘',
      ink: '1F1A14',
      paper: 'F0E6D2',
      paperTint: 'E3D7BF',
      inkTint: '2D2620',
    },
    cmb: {
      name: '招商银行红',
      ink: '111111',
      paper: 'FFFFFF',
      paperTint: 'F4F1EF',
      inkTint: 'C8102E',
    },
  },
  swiss: {
    ikb: {
      name: '克莱因蓝',
      paper: 'FAFAF8',
      ink: '0A0A0A',
      grey1: 'F0F0EE',
      grey2: 'D4D4D2',
      grey3: '737373',
      accent: '002FA7',
      accentOn: 'FFFFFF',
    },
    lemon: {
      name: '柠檬黄',
      paper: 'FAFAF8',
      ink: '0A0A0A',
      grey1: 'F0F0EE',
      grey2: 'D4D4D2',
      grey3: '737373',
      accent: 'FFD500',
      accentOn: '0A0A0A',
    },
    green: {
      name: '柠檬绿',
      paper: 'FAFAF8',
      ink: '0A0A0A',
      grey1: 'F0F0EE',
      grey2: 'D4D4D2',
      grey3: '737373',
      accent: 'C5E803',
      accentOn: '0A0A0A',
    },
    orange: {
      name: '安全橙',
      paper: 'FAFAF8',
      ink: '0A0A0A',
      grey1: 'F0F0EE',
      grey2: 'D4D4D2',
      grey3: '737373',
      accent: 'FF6B35',
      accentOn: 'FFFFFF',
    },
    cmb: {
      name: '招商银行红',
      paper: 'FFFFFF',
      ink: '111111',
      grey1: 'F6F3F2',
      grey2: 'DED8D6',
      grey3: '6F6663',
      accent: 'C8102E',
      accentOn: 'FFFFFF',
    },
  },
  cmb: {
    classic: {
      name: '招商银行经典红',
      paper: 'FFFFFF',
      paperTint: 'F8F1F1',
      ink: '111111',
      grey1: 'F6F3F2',
      grey2: 'DED8D6',
      grey3: '6F6663',
      accent: 'C8102E',
      accent2: '8A1538',
      accentSoft: 'F7DDE2',
      accentOn: 'FFFFFF',
    },
    pearl: {
      name: '招商银行珍珠白',
      paper: 'FFFFFF',
      paperTint: 'F7F4F2',
      ink: '171313',
      grey1: 'F2EFED',
      grey2: 'D8D0CD',
      grey3: '6E625F',
      accent: 'B5122B',
      accent2: '5C111C',
      accentSoft: 'F5E4E7',
      accentOn: 'FFFFFF',
    },
    graphite: {
      name: '招商银行石墨红',
      paper: 'FBFAF9',
      paperTint: 'F0ECEA',
      ink: '1C1A19',
      grey1: 'ECE7E5',
      grey2: 'CFC6C2',
      grey3: '5D5552',
      accent: 'C8102E',
      accent2: '3A2629',
      accentSoft: 'F4D9DE',
      accentOn: 'FFFFFF',
    },
  },
};

const FONTS = {
  serifZh: 'Noto Serif SC',
  serifEn: 'Playfair Display',
  sansZh: 'Noto Sans SC',
  sans: 'Inter',
  mono: 'IBM Plex Mono',
};

const READABILITY = {
  minFontSize: 9.8,
  minChartFontSize: 9,
  minTableFontSize: 9.2,
};

const BASIC_ICON_NAMES = ['dot', 'square', 'diamond', 'plus', 'minus', 'cross', 'number'];
const ICON_ALIASES = {
  check: 'check-circle',
  checkCircle: 'check-circle',
  success: 'check-circle',
  done: 'check-circle',
  alert: 'circle-alert',
  alertCircle: 'circle-alert',
  warning: 'triangle-alert',
  danger: 'circle-x',
  error: 'circle-x',
  info: 'info',
  infoCircle: 'info',
  arrow: 'arrow-right',
  arrowRight: 'arrow-right',
  trend: 'trending-up',
  trendUp: 'trending-up',
  chart: 'chart-column',
  chartBar: 'chart-column',
  bar: 'chart-column',
  chartLine: 'chart-line',
  line: 'chart-line',
  pie: 'chart-pie',
  pieChart: 'chart-pie',
  fileText: 'file-text',
  idea: 'lightbulb',
  process: 'workflow',
  risk: 'shield-alert',
  cost: 'badge-dollar-sign',
  money: 'badge-dollar-sign',
  people: 'users',
  team: 'users',
  config: 'settings',
};
const BULLET_ICON_NAMES = [...BASIC_ICON_NAMES, 'lucide:<name>', '<any lucide icon name>'];

const STYLE_ORDER = ['magazine', 'swiss', 'cmb'];
const DEFAULT_THEMES = { magazine: 'ink', swiss: 'ikb', cmb: 'classic' };
const LUCIDE_STATIC_ICON_DIR = path.join(__dirname, '..', '..', 'node_modules', 'lucide-static', 'icons');

function isSupportedStyle(style) {
  return STYLE_ORDER.includes(style);
}

function defaultThemeForStyle(style) {
  return DEFAULT_THEMES[style] || DEFAULT_THEMES.magazine;
}

module.exports = {
  SLIDE,
  THEMES,
  FONTS,
  READABILITY,
  BASIC_ICON_NAMES,
  ICON_ALIASES,
  BULLET_ICON_NAMES,
  STYLE_ORDER,
  DEFAULT_THEMES,
  LUCIDE_STATIC_ICON_DIR,
  isSupportedStyle,
  defaultThemeForStyle,
};
