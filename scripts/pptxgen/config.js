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
      ink: '101113',
      paper: 'F3F0EA',
      paperTint: 'E3DED5',
      inkTint: '2B2D31',
      accent: '4D5561',
      accentOn: 'F8F5EF',
      chartColors: ['101113', '4D5561', '8B8F96', 'C0BAB0', 'E3DED5'],
    },
    indigo: {
      name: '靛蓝瓷',
      ink: '102942',
      paper: 'F4F7F9',
      paperTint: 'E2EAF0',
      inkTint: '2F5F86',
      accent: '4B78A0',
      accentOn: 'FFFFFF',
      chartColors: ['102942', '4B78A0', '7FA6C2', 'B8C9D8', 'E2EAF0'],
    },
    forest: {
      name: '森林雾绿',
      ink: '203528',
      paper: 'F4F5ED',
      paperTint: 'E4E8D9',
      inkTint: '58745C',
      accent: '6F8A68',
      accentOn: 'FFFFFF',
      chartColors: ['203528', '58745C', '8EA081', 'C2C9B5', 'E4E8D9'],
    },
    kraft: {
      name: '陶土纸',
      ink: '3E3029',
      paper: 'F4E8DA',
      paperTint: 'E6D2C0',
      inkTint: 'A76C55',
      accent: 'A76C55',
      accentOn: 'FFF7EF',
      chartColors: ['3E3029', 'A76C55', 'C98D72', 'D7B9A2', 'E6D2C0'],
    },
    dune: {
      name: '沙丘雾蓝',
      ink: '393733',
      paper: 'F2ECE2',
      paperTint: 'DED8CC',
      inkTint: '7E8791',
      accent: '7E8791',
      accentOn: 'FFFFFF',
      chartColors: ['393733', '7E8791', '9EA8AE', 'C2B9AB', 'DED8CC'],
    },
    cmb: {
      name: '招商银行红',
      ink: '111111',
      paper: 'FFFFFF',
      paperTint: 'F4F1EF',
      inkTint: 'C8102E',
      accent: 'C8102E',
      accentOn: 'FFFFFF',
      chartColors: ['C8102E', '8A1538', '6B6258', 'A39A8F', 'F4F1EF'],
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
      headerBg: 'FFFFFF',
      headerInk: '111111',
      headerLine: 'C8102E',
      chartColors: ['C8102E', '8A1538', '6F6663', 'DED8D6', 'F6F3F2'],
    },
    pearl: {
      name: '招商银行暖玫瑰',
      paper: 'FDF7F2',
      paperTint: 'F6ECE4',
      ink: '4B3836',
      grey1: 'F1E3DA',
      grey2: 'D7C3B8',
      grey3: '8B746C',
      accent: 'A85560',
      accent2: 'C9828A',
      accentSoft: 'E8C5C8',
      accentOn: 'FFF8F4',
      headerBg: 'FFF7F0',
      headerInk: '4B3836',
      headerLine: 'D8B6AD',
      chartColors: ['A85560', 'C9828A', '8B746C', 'D7C3B8', 'F1E3DA'],
      emphasisSolid: true,
    },
    graphite: {
      name: '招商银行雾灰紫',
      paper: 'F0EDEA',
      paperTint: 'E3DED9',
      ink: '3E3937',
      grey1: 'DAD4CE',
      grey2: 'B8AEA8',
      grey3: '746B66',
      accent: '786070',
      accent2: '9D7A82',
      accentSoft: 'D2C4CB',
      accentOn: 'F8F4F1',
      headerBg: 'E9E4DF',
      headerInk: '3E3937',
      headerLine: 'B8AEA8',
      chartColors: ['786070', '9D7A82', '746B66', 'B8AEA8', 'DAD4CE'],
      emphasisSolid: true,
    },
  },
};

const FONTS = {
  zh: 'Microsoft YaHei',
  en: 'Times New Roman',
  serifZh: 'Microsoft YaHei',
  serifEn: 'Times New Roman',
  sansZh: 'Microsoft YaHei',
  sans: 'Microsoft YaHei',
  mono: 'Microsoft YaHei',
};

const TYPOGRAPHY = {
  coverTitle: 36,
  pageTitle: 28,
  itemTitle: 16,
  body: 14,
  dense: 12,
};

const READABILITY = {
  minFontSize: TYPOGRAPHY.dense,
  minChartFontSize: TYPOGRAPHY.dense,
  minTableFontSize: TYPOGRAPHY.dense,
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
  TYPOGRAPHY,
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
