'use strict';

const fs = require('node:fs');
const path = require('node:path');
const { THEMES: BUILTIN_THEMES } = require('./config');

const BUILTIN_STYLES = Object.freeze([
  builtin('cmb', '招商银行商务风', '白色品牌页眉、招商红与灰白金融版式，适合银行经营、管理层汇报和正式商务场景。', 'classic', './templates/cmb'),
  builtin('swiss', '瑞士国际主义', '高亮色、直角网格、清晰信息层级，适合产品、战略、数据分析和方法论汇报。', 'ikb', './templates/swiss'),
  builtin('magazine', '电子杂志风', '衬线标题、编辑式图文节奏和纸感留白，适合叙事、观点、案例和图片材料较多的内容。', 'ink', './templates/magazine'),
]);

let cachedRegistry = null;

function builtin(id, name, description, defaultTheme, modulePath) {
  return Object.freeze({
    id,
    name,
    description,
    defaultTheme,
    themes: BUILTIN_THEMES[id],
    modulePath: require.resolve(modulePath),
    builtin: true,
  });
}

function getStyleRegistry(options = {}) {
  if (!options.roots && cachedRegistry) return cachedRegistry;
  const registry = createStyleRegistry(options);
  if (!options.roots) cachedRegistry = registry;
  return registry;
}

function createStyleRegistry(options = {}) {
  const styles = new Map(BUILTIN_STYLES.map((style) => [style.id, style]));
  const errors = [];
  styleRoots(options.roots).forEach((root) => discoverRoot(root, styles, errors));
  return Object.freeze({ styles, errors });
}

function styleRoots(explicitRoots) {
  if (Array.isArray(explicitRoots)) return explicitRoots.map((root) => path.resolve(root));
  const roots = [path.join(__dirname, 'templates', 'styles')];
  const extra = String(process.env.PPTXGEN_STYLE_PATHS || '').split(path.delimiter).map((value) => value.trim()).filter(Boolean);
  extra.forEach((root) => roots.push(path.resolve(root)));
  return Array.from(new Set(roots));
}

function discoverRoot(root, styles, errors) {
  if (!fs.existsSync(root)) return;
  fs.readdirSync(root, { withFileTypes: true })
    .filter((entry) => entry.isDirectory())
    .sort((a, b) => a.name.localeCompare(b.name))
    .forEach((entry) => {
      const modulePath = path.join(root, entry.name, 'index.js');
      if (!fs.existsSync(modulePath)) return;
      try {
        const definition = normalizePlugin(require(modulePath), modulePath);
        if (styles.has(definition.id)) throw new Error(`style id "${definition.id}" is already registered`);
        styles.set(definition.id, definition);
      } catch (error) {
        errors.push({ modulePath, message: error.message });
      }
    });
}

function normalizePlugin(raw, modulePath) {
  const plugin = raw?.default || raw;
  if (!plugin || typeof plugin !== 'object') throw new Error('plugin must export a style definition object');
  const id = String(plugin.id || '').trim();
  if (!/^[a-z][a-z0-9-]*$/.test(id)) throw new Error('plugin id must match /^[a-z][a-z0-9-]*$/');
  if (typeof plugin.createTemplate !== 'function') throw new Error(`style "${id}" must export createTemplate(api)`);
  if (!plugin.themes || typeof plugin.themes !== 'object' || Array.isArray(plugin.themes) || !Object.keys(plugin.themes).length) {
    throw new Error(`style "${id}" must export a non-empty themes object`);
  }
  const defaultTheme = String(plugin.defaultTheme || '').trim();
  if (!defaultTheme || !plugin.themes[defaultTheme]) throw new Error(`style "${id}" defaultTheme must name one key in themes`);
  return Object.freeze({
    id,
    name: String(plugin.name || id),
    description: String(plugin.description || '自定义 PPTXGenJS 风格。'),
    defaultTheme,
    themes: Object.freeze({ ...plugin.themes }),
    createTemplate: plugin.createTemplate,
    sampleSpec: typeof plugin.sampleSpec === 'function' ? plugin.sampleSpec : null,
    modulePath,
    builtin: false,
  });
}

function listStyles() {
  reportDiscoveryErrors();
  return Array.from(getStyleRegistry().styles.values());
}

function styleDefinition(id) {
  reportDiscoveryErrors();
  return getStyleRegistry().styles.get(id) || null;
}

function isRegisteredStyle(id) {
  return Boolean(styleDefinition(id));
}

function defaultThemeForRegisteredStyle(id) {
  return styleDefinition(id)?.defaultTheme || null;
}

function themesForStyle(id) {
  return styleDefinition(id)?.themes || null;
}

function allRegisteredThemes() {
  return Object.fromEntries(listStyles().map((style) => [style.id, style.themes]));
}

function createRegisteredTemplate(id, api) {
  const definition = styleDefinition(id);
  if (!definition) return null;
  const factory = definition.builtin ? require(definition.modulePath) : definition.createTemplate;
  const template = factory(api);
  if (!template || typeof template.render !== 'function') {
    throw new Error(`style "${id}" createTemplate(api) must return { render(slide, ctx) }`);
  }
  return template;
}

let reportedErrors = false;
function reportDiscoveryErrors() {
  const errors = getStyleRegistry().errors;
  if (reportedErrors || !errors.length) return;
  reportedErrors = true;
  errors.forEach((error) => console.warn(`Warning: skipped invalid style plugin ${error.modulePath}: ${error.message}`));
}

module.exports = {
  BUILTIN_STYLES,
  createStyleRegistry,
  getStyleRegistry,
  listStyles,
  styleDefinition,
  isRegisteredStyle,
  defaultThemeForRegisteredStyle,
  themesForStyle,
  allRegisteredThemes,
  createRegisteredTemplate,
};
