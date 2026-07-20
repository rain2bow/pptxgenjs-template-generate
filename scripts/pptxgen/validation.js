'use strict';

module.exports = function createValidationTools(deps) {
  const {
    fail,
    normalizeMediaImages,
    normalizeMediaCharts,
    normalizeChartData,
    normalizeTableRows,
    normalizeSections,
    resolveImage,
    clamp,
  } = deps;

  const VALIDATION_FORMAT_HINT = 'Check the JSON examples in SKILL.md / README.md, or run: node scripts/generate-pptx.js --layout-examples <style> --out outputs/layout-examples.md';

  const SCALAR_TEXT_FIELD_NAMES = new Set([
    'kicker', 'title', 'subtitle', 'body', 'desc', 'note', 'summary', 'detail', 'text', 'story',
    'conclusion', 'takeaway', 'footerSummary', 'nextStep', 'lead', 'callout',
    'quote', 'cite', 'source', 'caseTitle', 'summaryTitle', 'leadTitle', 'focusTitle',
    'conclusionTitle', 'takeawayTitle', 'footerSummaryTitle', 'nextStepTitle',
    'label', 'value', 'unit', 'metric', 'name', 'caption',
  ]);

  const SLIDE_CONTENT_SCALAR_FIELDS = [
    'kicker', 'title', 'subtitle', 'body', 'desc', 'note', 'summary', 'detail', 'text', 'story',
    'conclusion', 'takeaway', 'footerSummary', 'nextStep', 'lead', 'callout',
    'quote', 'cite', 'source', 'caseTitle', 'summaryTitle', 'leadTitle', 'focusTitle',
    'conclusionTitle', 'takeawayTitle', 'footerSummaryTitle', 'nextStepTitle',
    'label', 'value', 'unit', 'metric', 'name', 'caption',
  ];

  const DISPLAY_ITEM_TEXT_KEYS = ['text', 'title', 'label', 'body', 'desc', 'note', 'summary', 'detail', 'value', 'unit', 'metric', 'name'];

  const MEDIA_SLOT_LAYOUTS = new Set(['statement', 'media', 'mediaGrid', 'gallery', 'imageGrid', 'imageHero', 'quoteImage', 'textImage', 'caseStudy', 'pairedMedia', 'pairedStatementMedia', 'pairedQuoteMedia']);

  const VISUAL_MEDIA_LAYOUTS = MEDIA_SLOT_LAYOUTS;

  const CHART_DATA_LAYOUTS = new Set(['chart', 'dashboard']);

  const TABLE_DATA_LAYOUTS = new Set(['dataSheet']);

  function normalizeLayoutCompatibility(spec) {
    spec.slides.forEach((slide, index) => {
      normalizeCompareSlide(slide, index);
    });
  }

  function normalizeCompareSlide(slide, index) {
    const layout = slide.layout || '';
    if (!['compare', 'duoCompare', 'splitCompare'].includes(layout)) return;
    const before = normalizeCompareColumn(slide.before, slide.left, slide.leftTitle || slide.beforeTitle, slide.leftLabel || slide.beforeLabel, 'left/before', index);
    const after = normalizeCompareColumn(slide.after, slide.right, slide.rightTitle || slide.afterTitle, slide.rightLabel || slide.afterLabel, 'right/after', index);
    if (before) slide.before = before;
    if (after) slide.after = after;
  }

  function normalizeCompareColumn(primary, alias, title, label, side, index) {
    if (primary && !Array.isArray(primary)) return normalizeCompareColumnObject(primary, title, label);
    if (Array.isArray(primary)) {
      console.warn(`Warning: slide ${index + 1} compare column "${side}" is an array; normalized it to { title, items } so body text is rendered.`);
      return normalizeCompareColumnObject({ title, label, items: primary }, title, label);
    }
    if (Array.isArray(alias)) {
      console.warn(`Warning: slide ${index + 1} compare column "${side}" is an array; normalized it to { title, items } so body text is rendered.`);
      return normalizeCompareColumnObject({ title, label, items: alias }, title, label);
    }
    if (alias && typeof alias === 'object') return normalizeCompareColumnObject(alias, title, label);
    return null;
  }

  function normalizeCompareColumnObject(column, title, label) {
    const next = { ...column };
    if (title && !next.title) next.title = title;
    if (label && !next.label) next.label = label;
    if (!Array.isArray(next.items)) {
      const fallbackItems = next.sections || next.points || next.bullets || next.list;
      if (Array.isArray(fallbackItems)) next.items = fallbackItems;
    }
    return next;
  }

  function validateSpecSlots(spec, options = {}) {
    const errors = [];
    const warnings = [];
    spec.slides.forEach((slide, index) => {
      validateTextFieldTypes(slide, index, errors);
      validateSlideScalarFields(slide, index, spec.style, errors);
      validateRequiredSlideFields(slide, index, spec.style, errors);
      validateMediaSlots(slide, index, spec.style, errors, warnings, options.specDir || process.cwd());
      validateTextSlots(slide, index, spec.style, errors, warnings);
      validateRenderableDataBlocks(slide, index, spec.style, errors);
      validateThinContent(slide, index, spec.style, errors);
    });
    if (errors.length) {
      fail(`Spec slot validation failed:\n- ${errors.join('\n- ')}\n\n${VALIDATION_FORMAT_HINT}`);
    }
    warnings.forEach((message) => console.warn(message));
  }

  function validateTextFieldTypes(value, index, errors, pathName = 'slide', depth = 0) {
    if (!value || typeof value !== 'object' || depth > 5) return;
    if (Array.isArray(value)) {
      value.forEach((item, itemIndex) => validateTextFieldTypes(item, index, errors, `${pathName}[${itemIndex}]`, depth + 1));
      return;
    }
    Object.entries(value).forEach(([key, child]) => {
      const childPath = `${pathName}.${key}`;
      if (SCALAR_TEXT_FIELD_NAMES.has(key) && child !== undefined && child !== null && typeof child === 'object') {
        const kind = Array.isArray(child) ? 'array' : 'object';
        const hint = key === 'body'
          ? 'If this is a list of content blocks, put it in sections/items/columns/steps/nodes according to the slide layout; if it is bullet text inside one card, put strings in points[].'
          : 'Use a plain string/number for this field, or move structured content into the layout collection field.';
        errors.push(`slide ${index + 1} field ${childPath} must be plain text, but got ${kind}. ${hint}`);
      }
      if (child && typeof child === 'object' && !isOpaqueValidationObject(key)) {
        validateTextFieldTypes(child, index, errors, childPath, depth + 1);
      }
    });
  }

  function isOpaqueValidationObject(key) {
    return key === 'chart' || key === 'table' || key === 'speakerNotes';
  }

  function validateSlideScalarFields(slide, index, style, errors) {
    const layout = slide.layout || defaultLayoutForStyle(style);
    const allowed = allowedScalarFieldsForLayout(style, layout);
    const filled = SLIDE_CONTENT_SCALAR_FIELDS.filter((key) => Object.prototype.hasOwnProperty.call(slide, key) && hasMeaningfulValue(slide[key]));
    const ignored = filled.filter((key) => !allowed.has(key));
    if (ignored.length) {
      errors.push(`slide ${index + 1} layout "${layout}" does not render field(s): ${ignored.join(', ')}. Remove them, rename them to fields supported by this layout, or change slide.layout.`);
    }
  }

  function validateRequiredSlideFields(slide, index, style, errors) {
    const layout = slide.layout || defaultLayoutForStyle(style);
    const groups = requiredScalarGroupsForLayout(style, layout);
    groups.forEach((group) => {
      if (group.skipWhen?.(slide)) return;
      if (group.allowSparse && slide.allowSparseContent) return;
      if (!group.keys.some((key) => hasMeaningfulValue(slide[key]))) {
        errors.push(`slide ${index + 1} layout "${layout}" is missing required field: ${group.keys.join(' or ')}. ${group.reason || 'Fill the required content field or change slide.layout.'}`);
      }
    });
  }

  function defaultLayoutForStyle(style) {
    return style === 'magazine' ? 'textImage' : 'statement';
  }

  function hasMeaningfulValue(value) {
    if (value === undefined || value === null) return false;
    if (typeof value === 'string') return value.trim().length > 0;
    if (typeof value === 'number' || typeof value === 'boolean') return true;
    if (Array.isArray(value)) return value.length > 0;
    if (typeof value === 'object') return Object.keys(value).length > 0;
    return false;
  }

  function allowedScalarFieldsForLayout(style, layout) {
    const common = ['kicker', 'title', 'subtitle'];
    const set = new Set(common);
    const add = (...keys) => keys.forEach((key) => set.add(key));
    const isCmb = style === 'cmb';
    if (['cover', 'section', 'closing'].includes(layout)) add('body');
    else if (layout === 'statement') add('body', 'callout');
    else if (layout === 'bigQuote') add('quote', 'body', 'cite', 'source');
    else if (layout === 'quoteImage') add('quote', 'body', 'cite', 'source', 'callout', 'caption');
    else if (layout === 'textImage') add('body', 'callout', 'caption');
    else if (layout === 'media') add('body', 'summary', 'story', 'note', 'caption');
    else if (layout === 'pairedStatementText' || layout === 'pairedStatementMedia') add('body', 'callout');
    else if (layout === 'pairedQuoteText' || layout === 'pairedQuoteMedia') add('body', 'quote', 'cite', 'source', 'callout', 'caption');
    else if (layout === 'pairedMedia' || layout === 'pairedText') add('body', 'note', 'caseTitle');
    else if (layout === 'caseStudy') add('caseTitle', 'label', 'body', 'summary', 'story', 'caption');
    else if (layout === 'imageHero') add('body');
    else if (layout === 'dataSheet' && style === 'magazine') add('body');
    else if (layout === 'pyramid') add('body', 'note');
    else if (['article', 'sectionList', 'briefing', 'executiveBrief', 'contentBrief'].includes(layout) && isCmb) {
      add('summary', 'body', 'lead', 'summaryTitle', 'leadTitle', 'focusTitle', 'conclusion', 'takeaway', 'footerSummary', 'nextStep', 'conclusionTitle', 'takeawayTitle', 'footerSummaryTitle', 'nextStepTitle');
    } else if (['article', 'briefing', 'executiveBrief', 'contentBrief'].includes(layout) && style === 'magazine') {
      add('callout');
    }
    return set;
  }

  function requiredScalarGroupsForLayout(style, layout) {
    const titleKeys = ['bigQuote', 'quoteImage'].includes(layout) ? ['title', 'quote'] : ['title'];
    const groups = [{ keys: titleKeys, reason: 'Every slide should have a visible title or primary headline.' }];
    if (layout === 'statement') groups.push({ keys: ['body', 'subtitle'], allowSparse: true, reason: 'Statement slides need supporting text unless this is an intentional sparse draft.' });
    if (layout === 'textImage') groups.push({ keys: ['body'], reason: 'textImage renders its main paragraph from body.' });
    if (layout === 'bigQuote') groups.push({ keys: ['quote', 'title'], reason: 'bigQuote needs quote or title as the large quote text.' });
    if (layout === 'media') groups.push({ keys: ['body', 'summary', 'story', 'note', 'items', 'insights', 'points'], allowSparse: true, reason: 'media needs a summary/body or side point collection in addition to the media area.' });
    if (layout === 'caseStudy') groups.push({ keys: ['body', 'summary', 'story', 'subtitle'], allowSparse: true, reason: 'caseStudy needs case narrative text.' });
    if (layout === 'imageHero') groups.push({ keys: ['body', 'subtitle'], allowSparse: true, reason: 'imageHero needs a short body/subtitle under the hero media.' });
    return groups;
  }

  function validateMediaSlots(slide, index, style, errors, warnings, specDir) {
    const layout = slide.layout || defaultLayoutForStyle(style);
    if (!MEDIA_SLOT_LAYOUTS.has(layout)) {
      const imageFields = unsupportedImageFields(slide);
      if (imageFields.length) {
        errors.push(`slide ${index + 1} layout "${layout}" does not render image/media field(s): ${imageFields.join(', ')}. Use a media layout such as media/mediaGrid/imageGrid/imageHero/statement/caseStudy, or remove these fields.`);
      }
      return;
    }
    const images = normalizeMediaImages(slide);
    const charts = normalizeMediaCharts(slide);
    const slotCount = isMediaGridLayout(layout) ? resolveMediaSlotCount(slide) : 1;
    const explicitCount = explicitMediaCount(slide);
    if (layout === 'statement' && images.length > 1) errors.push(`slide ${index + 1} uses statement layout with ${images.length} images, but statement supports exactly one image slot; use mediaGrid/imageGrid or split into another slide.`);
    if (layout !== 'statement' && images.length > 6) errors.push(`slide ${index + 1} has ${images.length} images, but media layouts support at most 6 image slots; split into another slide.`);
    if (layout === 'statement' && charts.length) errors.push(`slide ${index + 1} uses statement layout with chart data, but statement reserves the media area for one image; use chart/media layout instead.`);
    if (layout !== 'statement' && charts.length > 6) errors.push(`slide ${index + 1} has ${charts.length} charts, but media layouts support at most 6 media slots; split into another slide.`);
    if (layout !== 'statement' && Math.max(images.length, charts.length) > slotCount) {
      errors.push(`slide ${index + 1} has ${Math.max(images.length, charts.length)} media assets but only ${slotCount} slot(s).`);
    }
    const assetCount = Math.max(images.length, charts.length);
    if (explicitCount && explicitCount !== assetCount) {
      errors.push(`slide ${index + 1} declares mediaCount/imageSlots/slotCount=${explicitCount} but provides ${assetCount} image/chart asset(s). Remove the explicit slot count or provide matching media assets; blank placeholders are not allowed.`);
    }
    images.forEach((image, imageIndex) => {
      if (!resolveImage(specDir, image)) {
        errors.push(`slide ${index + 1} image ${imageIndex + 1} path is missing or unsupported; provide a valid image path or remove the image entry.`);
      }
    });
    if (isVisualMediaLayout(layout) && !images.length && !charts.length) {
      errors.push(`slide ${index + 1} uses layout "${layout}" with media/image slot(s) but provides no images or charts. Use a text-only layout such as textGrid/article/fourCards/agenda/radial, or provide image/chart data. Blank media placeholders are not allowed.`);
    }
  }

  function validateTextSlots(slide, index, style, errors, warnings) {
    const layout = slide.layout || defaultLayoutForStyle(style);
    const sideMax = style === 'swiss' ? 5 : style === 'cmb' ? 4 : 3;
    const cmbTextWeaveMin = style === 'cmb' ? 0 : 1;
    const cmbTextGridMax = style === 'cmb' ? 6 : 9;
    const narrativeFields = ['body', 'desc', 'note', 'summary', 'detail', 'text', 'story'];
    const metricNarrativeFields = ['body', 'desc', 'summary', 'detail', 'text', 'story'];
    const textLayoutSuggestion = 'Use textGrid, article, sectionList, fourCards, agenda, or radial for title + body content.';
    const metricLayoutSuggestion = 'Use textGrid, article, sectionList, or media when each item needs explanatory body text.';
    const titleOnlyMatrixRule = style === 'magazine'
      ? { keys: ['items'], max: 12, min: 1, label: 'matrix cells' }
      : { keys: ['items'], max: 12, min: 1, label: 'matrix cells', itemTextKeys: ['title', 'label'], unusedItemFields: [...narrativeFields, 'name', 'value'], suggestion: textLayoutSuggestion };
    const textCardRule = (keys, max, min, label, extra = {}) => ({
      keys,
      max,
      min,
      label,
      itemTextKeys: ['title', 'label', 'name', 'body', 'desc', 'note', 'summary', 'detail', 'text', 'points', 'bullets', 'list'],
      requiredItemTitleKeys: ['title', 'label', 'name', 'heading'],
      requiredItemBodyKeys: ['body', 'desc', 'note', 'summary', 'detail', 'text', 'points', 'bullets', 'list'],
      suggestion: textLayoutSuggestion,
      ...extra,
    });
    const processStepRule = (keys, max, min, label, extra = {}) => textCardRule(keys, max, min, label, {
      requiredItemTitleKeys: ['title', 'label', 'year'],
      ...extra,
    });
    const numberMetricRule = style === 'magazine'
      ? { keys: ['items'], max: 6, min: 1, label: 'number cards', itemTextKeys: ['label', 'value', 'note', 'unit'], unusedItemFields: ['title', ...metricNarrativeFields], suggestion: metricLayoutSuggestion }
      : { keys: ['items'], max: 6, min: 1, label: 'number cards', itemTextKeys: ['label', 'value', 'note', 'unit'], unusedItemFields: ['title', ...metricNarrativeFields], suggestion: metricLayoutSuggestion };
    const dashboardMetricRule = style === 'magazine'
      ? { keys: ['metrics', 'items'], max: 4, min: 1, label: 'dashboard metrics', itemTextKeys: ['label', 'value', 'note'], unusedItemFields: ['title', ...metricNarrativeFields], suggestion: metricLayoutSuggestion }
      : { keys: ['metrics', 'items'], max: style === 'swiss' ? 5 : 4, min: 1, label: 'dashboard metrics', itemTextKeys: ['label', 'value'], unusedItemFields: ['title', 'note', ...metricNarrativeFields], suggestion: metricLayoutSuggestion };
    const rules = {
      cover: [],
      section: [],
      bigQuote: [],
      quoteImage: [],
      compare: [],
      duoCompare: [],
      splitCompare: [],
      textImage: [],
      pairedStatementText: [],
      pairedStatementMedia: [],
      pairedQuoteText: [],
      pairedQuoteMedia: [],
      pairedText: [{ keys: ['items'], max: 8, min: 1, label: 'paired text items', itemTextKeys: ['title', 'label', 'name', 'body', 'desc', 'note', 'summary', 'detail', 'text', 'value', 'items'] }],
      pairedMedia: [{ keys: ['items'], max: 8, min: 1, label: 'paired media text items', itemTextKeys: ['title', 'label', 'name', 'body', 'desc', 'note', 'summary', 'detail', 'text', 'value', 'items'] }],
      statement: [],
      closing: [],
      bigNumbers: [numberMetricRule],
      kpiTower: [{ ...numberMetricRule, max: 4, label: 'KPI cards' }],
      pipeline: [processStepRule(['steps', 'items'], 6, 1, 'pipeline steps')],
      timeline: [processStepRule(['items', 'steps'], 6, 1, 'timeline steps')],
      matrix: [titleOnlyMatrixRule],
      fourCards: [textCardRule(['items'], 8, 1, 'cards')],
      article: [textCardRule(['sections', 'items', 'columns'], 6, 1, 'article sections')],
      briefing: [textCardRule(['sections', 'items', 'columns', 'points', 'agenda'], 6, 2, 'briefing text blocks')],
      executiveBrief: [textCardRule(['sections', 'items', 'columns', 'points', 'agenda'], 6, 2, 'briefing text blocks')],
      contentBrief: [textCardRule(['sections', 'items', 'columns', 'points', 'agenda'], 6, 2, 'briefing text blocks')],
      textGrid: [textCardRule(['sections', 'items', 'columns'], cmbTextGridMax, cmbTextWeaveMin, style === 'cmb' ? 'CMB text weave cards (lead + right card)' : 'text grid cells')],
      textWeave: [textCardRule(['sections', 'items', 'columns', 'points', 'agenda'], 6, cmbTextWeaveMin, style === 'cmb' ? 'CMB text weave cards (lead + right card)' : 'text weave blocks')],
      contentSynthesis: [textCardRule(['sections', 'items', 'columns', 'points', 'agenda'], 6, cmbTextWeaveMin, style === 'cmb' ? 'CMB text weave cards (lead + right card)' : 'text weave blocks')],
      denseText: [textCardRule(['sections', 'items', 'columns', 'points', 'agenda'], 6, cmbTextWeaveMin, style === 'cmb' ? 'CMB text weave cards (lead + right card)' : 'dense text blocks')],
      sectionList: [textCardRule(['sections', 'items', 'columns'], 7, 1, 'section list items')],
      agenda: [{ keys: ['items', 'sections', 'agenda'], max: 8, min: 1, label: 'agenda items' }],
      pyramid: [textCardRule(['layers', 'items', 'sections'], 5, 1, 'pyramid layers')],
      radial: [textCardRule(['items', 'nodes', 'sections'], 8, 1, 'radial nodes')],
      roadmap: [processStepRule(['steps', 'items'], 6, 1, 'roadmap steps')],
      swimlane: [textCardRule(['lanes', 'sections'], 4, 1, 'swimlanes', { requiredItemTitleKeys: ['title', 'label', 'name'] })],
      media: [textCardRule(['items', 'insights', 'points'], sideMax, 1, 'side points', { suggestion: 'Add 1-4 side points with title/body, or use textImage/statement when the slide only needs one paragraph beside media.' })],
      mediaGrid: [{ keys: ['captions', 'items', 'sections'], max: 6, min: 1, label: 'media captions', itemTextKeys: ['caption', 'title', 'label'], unusedItemFields: [...narrativeFields, 'points', 'bullets', 'list', 'value'], suggestion: 'Use caption/title/label for mediaGrid captions. Use media/textImage/caseStudy or split into a text slide when each image needs body text.' }],
      gallery: [{ keys: ['captions', 'items', 'sections'], max: 6, min: 1, label: 'media captions', itemTextKeys: ['caption', 'title', 'label'], unusedItemFields: [...narrativeFields, 'points', 'bullets', 'list', 'value'], suggestion: 'Use caption/title/label for gallery captions. Use media/textImage/caseStudy or split into a text slide when each image needs body text.' }],
      imageGrid: [{ keys: ['captions', 'items', 'sections'], max: 6, min: 1, label: 'media captions', itemTextKeys: ['caption', 'title', 'label'], unusedItemFields: [...narrativeFields, 'points', 'bullets', 'list', 'value'], suggestion: 'Use caption/title/label for imageGrid captions. Use media/textImage/caseStudy or split into a text slide when each image needs body text.' }],
      imageHero: [{ keys: ['items'], max: 3, min: 1, label: 'image hero metrics', itemTextKeys: ['label', 'value', 'note'], unusedItemFields: ['title', ...metricNarrativeFields], suggestion: metricLayoutSuggestion }],
      caseStudy: [{ keys: ['metrics', 'items'], max: 3, min: 1, label: 'case metrics', itemTextKeys: ['label', 'title', 'value', 'note'], unusedItemFields: metricNarrativeFields, suggestion: metricLayoutSuggestion }],
      dataSheet: [{ keys: ['notes', 'insights'], max: style === 'swiss' ? 4 : 3, min: 0, label: 'side notes' }],
      chart: [{ keys: ['insights', 'notes'], max: style === 'swiss' ? 3 : 4, min: 0, label: 'chart insights' }],
      dashboard: [dashboardMetricRule],
    };
    const layoutRules = rules[layout] || [];
    validateIgnoredSlotFields(slide, index, layout, layoutRules, errors, Object.prototype.hasOwnProperty.call(rules, layout));
    layoutRules.forEach((rule) => validateSlotCollection(slide, index, rule, errors, warnings));
    if (layout === 'compare' || layout === 'duoCompare' || layout === 'splitCompare') {
      validateSlotCollection(slide.before || {}, index, { keys: ['items'], max: 6, min: 1, label: 'before items', prefix: 'before.' }, errors, warnings);
      validateSlotCollection(slide.after || {}, index, { keys: ['items'], max: 6, min: 1, label: 'after items', prefix: 'after.' }, errors, warnings);
    }
    validateCmbTextWeaveStructure(slide, index, style, layout, errors);
    validateCmbBriefingCapacity(slide, index, style, layout, errors);
    if (layout === 'dashboard') validateChartSlots(slide, index, 2, errors, warnings);
    if (layout === 'chart') validateChartDataSlot(slide, index, errors, warnings);
    if (layout === 'dataSheet') validateTableSlot(slide, index, errors, warnings);
  }

  function validateSlotCollection(source, index, rule, errors, warnings) {
    const present = rule.keys.filter((key) => source[key] !== undefined && source[key] !== null);
    const suggestion = rule.suggestion ? ` ${rule.suggestion}` : '';
    if (!present.length) {
      if (rule.min > 0 && !source.allowSparseContent) errors.push(`slide ${index + 1} has no ${rule.label}; the layout may look empty. Add the required content field(s), change layout, or set allowSparseContent:true only for intentional sparse draft slides.${suggestion}`);
      return [];
    }
    if (present.length > 1) {
      errors.push(`slide ${index + 1} provides multiple fields for ${rule.label}: ${present.map((key) => `${rule.prefix || ''}${key}`).join(', ')}. Keep only one field so content is not silently ignored.`);
    }
    const key = present[0];
    const items = normalizeSlotItemsForValidation(source[key]);
    if (!items) {
      errors.push(slotCollectionFormatError(source[key], index, rule, key));
      return [];
    }
    if (rule.max && items.length > rule.max) {
      errors.push(`slide ${index + 1} has ${items.length} ${rule.label}, but layout "${source.layout || 'nested'}" renders at most ${rule.max}; split content or change layout.`);
    }
    if (rule.min && items.length < rule.min) {
      if (!source.allowSparseContent) errors.push(`slide ${index + 1} has ${items.length} ${rule.label}; expected at least ${rule.min}. Add content, change layout, or set allowSparseContent:true only for intentional sparse draft slides.${suggestion}`);
    }
    items.forEach((item, itemIndex) => {
      if (!slotItemHasDisplayTextForRule(item, rule)) {
        const keys = rule.itemTextKeys || DISPLAY_ITEM_TEXT_KEYS;
        errors.push(`slide ${index + 1} ${rule.prefix || ''}${key}[${itemIndex}] has no field rendered by ${rule.label}. Use one of: ${keys.join(', ')}.`);
      }
      validateRequiredSlotItemFields(item, index, rule, key, itemIndex, errors);
      validateUnusedSlotItemFields(item, index, rule, key, itemIndex, errors);
    });
    return items;
  }

  function validateCmbTextWeaveStructure(slide, index, style, layout, errors) {
    if (style !== 'cmb') return;
    if (!['textGrid', 'fourCards', 'textWeave', 'contentSynthesis', 'denseText'].includes(layout)) return;
    const count = validationCmbTextItems(slide).length;
    if (count < 2) {
      errors.push(`slide ${index + 1} layout "${layout}" needs at least 2 CMB text weave cards: 1 lead card and at least 1 right-side card; got ${count}. Add another sections/items/columns/points/agenda entry or change layout.`);
    }
  }

  function validateCmbBriefingCapacity(slide, index, style, layout, errors) {
    if (style !== 'cmb') return;
    if (!['article', 'sectionList', 'briefing', 'executiveBrief', 'contentBrief'].includes(layout)) return;
    const items = normalizeSections(slide.sections || slide.items || slide.columns || slide.points || slide.agenda || []);
    const hasLead = !!(slide.summary || slide.body || slide.lead);
    const conclusionText = slide.conclusion || slide.takeaway || slide.footerSummary || slide.nextStep;
    const restCount = hasLead ? items.length : Math.max(0, items.length - 1);
    const maxRest = conclusionText ? 4 : 5;
    if (restCount > maxRest) {
      const conclusionNote = conclusionText ? ' when conclusion/takeaway is present' : '';
      errors.push(`slide ${index + 1} layout "${layout}" renders at most ${maxRest} middle briefing text block(s)${conclusionNote}; got ${restCount}. Split content, reduce sections/items, remove conclusion, or use textWeave/textGrid.`);
    }
  }

  function validateIgnoredSlotFields(slide, index, layout, layoutRules, errors, isKnownLayout = false) {
    const valid = new Set(layoutRules.flatMap((rule) => rule.keys));
    const common = ['items', 'sections', 'columns', 'steps', 'nodes', 'layers', 'metrics', 'notes', 'insights', 'agenda', 'lanes', 'captions', 'points'];
    const ignored = common.filter((key) => slide[key] !== undefined && slide[key] !== null && !valid.has(key));
    if (ignored.length && (layoutRules.length || isKnownLayout)) {
      const targets = Array.from(valid);
      const suggestion = targets.length
        ? `Rename them to one of: ${targets.join(', ')}.`
        : 'Use a collection layout such as textGrid, article, sectionList, agenda, roadmap, timeline, media, dashboard, or dataSheet.';
      errors.push(`slide ${index + 1} layout "${layout}" does not render collection field(s): ${ignored.join(', ')}. ${suggestion}`);
    }
    if (layout === 'dashboard' && slide.chart && !Array.isArray(slide.charts)) {
      errors.push(`slide ${index + 1} dashboard uses charts[]; field chart will not be rendered. Rename chart to charts: [chart].`);
    }
  }

  function validateChartSlots(slide, index, max, errors, warnings) {
    const charts = normalizeMediaCharts(slide);
    if (charts.length > max) errors.push(`slide ${index + 1} has ${charts.length} charts, but dashboard renders at most ${max}.`);
    if (!charts.length && !slide.allowMissingChart) errors.push(`slide ${index + 1} dashboard has no charts; chart area will be empty. Provide charts[] or set allowMissingChart:true only for intentional draft output.`);
  }

  function validateChartDataSlot(slide, index, errors, warnings) {
    const chart = slide.chart || slide;
    if (!normalizeChartData(chart).length && !slide.allowMissingChart) errors.push(`slide ${index + 1} chart layout has no chart data; a NO DATA box will be rendered. Provide chart data or set allowMissingChart:true only for intentional draft output.`);
  }

  function validateTableSlot(slide, index, errors, warnings) {
    if (!slide.table) {
      if (!slide.allowMissingTable) errors.push(`slide ${index + 1} dataSheet has no table. Provide table data or set allowMissingTable:true only for intentional draft output.`);
      return;
    }
    if (slide.table.headers && !Array.isArray(slide.table.headers)) errors.push(`slide ${index + 1} table.headers must be an array.`);
    const rows = normalizeTableRows(slide.table);
    if (!rows.length && !slide.allowMissingTable) errors.push(`slide ${index + 1} table has no rows. Provide table.rows or set allowMissingTable:true only for intentional draft output.`);
  }

  function slotCollectionFormatError(value, index, rule, key) {
    const field = `${rule.prefix || ''}${key}`;
    if (value && typeof value === 'object' && !Array.isArray(value)) {
      const badKeys = Object.entries(value)
        .filter(([, body]) => body !== undefined && body !== null && typeof body === 'object')
        .map(([title]) => title);
      if (badKeys.length) {
        return `slide ${index + 1} field ${field} uses an object map with structured value(s): ${badKeys.slice(0, 5).join(', ')}. Use ${field}: [{ title, body }] for structured items, or ${field}: { "Title": "plain text" } for a simple object map.`;
      }
    }
    return `slide ${index + 1} field ${field} has unsupported format; use an array of strings/objects or an object map with plain-text values.`;
  }

  function normalizeSlotItemsForValidation(value) {
    if (value === undefined || value === null) return [];
    if (typeof value === 'string' || typeof value === 'number') return [{ body: String(value) }];
    if (Array.isArray(value)) return value;
    if (typeof value === 'object') {
      const entries = Object.entries(value);
      if (entries.some(([, body]) => body !== undefined && body !== null && typeof body === 'object')) return null;
      return entries.map(([title, body]) => ({ title, body: String(body ?? '') }));
    }
    return null;
  }

  function slotItemHasDisplayText(item) {
    if (typeof item === 'string' || typeof item === 'number') return String(item).trim().length > 0;
    if (!item || typeof item !== 'object') return false;
    return DISPLAY_ITEM_TEXT_KEYS.some((key) => String(item[key] ?? '').trim().length > 0);
  }

  function slotItemHasDisplayTextForRule(item, rule) {
    if (typeof item === 'string' || typeof item === 'number') return String(item).trim().length > 0;
    if (!item || typeof item !== 'object') return false;
    const keys = rule.itemTextKeys || DISPLAY_ITEM_TEXT_KEYS;
    return keys.some((key) => String(item[key] ?? '').trim().length > 0);
  }

  function validateUnusedSlotItemFields(item, index, rule, key, itemIndex, errors) {
    if (!item || typeof item !== 'object' || Array.isArray(item) || !rule.unusedItemFields?.length) return;
    const ignored = rule.unusedItemFields.filter((field) => String(item[field] ?? '').trim().length > 0);
    if (!ignored.length) return;
    const suggestion = rule.suggestion || 'Use a layout that renders these item fields, or remove the unused fields.';
    errors.push(`slide ${index + 1} ${rule.prefix || ''}${key}[${itemIndex}] includes field(s) not rendered by ${rule.label}: ${ignored.join(', ')}. ${suggestion}`);
  }

  function validateRequiredSlotItemFields(item, index, rule, key, itemIndex, errors) {
    if (!item || typeof item !== 'object' || Array.isArray(item)) return;
    if (rule.requiredItemTitleKeys?.length && !hasAnyTextField(item, rule.requiredItemTitleKeys)) {
      errors.push(`slide ${index + 1} ${rule.prefix || ''}${key}[${itemIndex}] is missing a title field required by ${rule.label}. Use one of: ${rule.requiredItemTitleKeys.join(', ')}.`);
    }
    if (rule.requiredItemBodyKeys?.length && !hasAnyTextField(item, rule.requiredItemBodyKeys)) {
      errors.push(`slide ${index + 1} ${rule.prefix || ''}${key}[${itemIndex}] is missing body/detail text required by ${rule.label}. Use one of: ${rule.requiredItemBodyKeys.join(', ')}.`);
    }
  }

  function hasAnyTextField(item, keys) {
    return keys.some((field) => {
      const value = item[field];
      if (Array.isArray(value)) return value.length > 0;
      return String(value ?? '').trim().length > 0;
    });
  }

  function unsupportedImageFields(slide) {
    const fields = [];
    ['image', 'images', 'gallery'].forEach((key) => {
      if (slide[key] !== undefined && slide[key] !== null) fields.push(key);
    });
    if (Array.isArray(slide.media) && normalizeMediaImages({ media: slide.media }).length) fields.push('media');
    collectUnsupportedNestedImageFields(slide, 'slide', fields);
    ['mediaCount', 'imageSlots', 'slotCount'].forEach((key) => {
      if (slide[key] !== undefined && slide[key] !== null) fields.push(key);
    });
    return Array.from(new Set(fields));
  }

  function collectUnsupportedNestedImageFields(value, pathName, fields, depth = 0) {
    if (!value || depth > 5) return;
    if (typeof value === 'string') {
      if (looksLikeImagePath(value)) fields.push(pathName);
      return;
    }
    if (Array.isArray(value)) {
      value.forEach((item, itemIndex) => collectUnsupportedNestedImageFields(item, `${pathName}[${itemIndex}]`, fields, depth + 1));
      return;
    }
    if (typeof value !== 'object') return;
    Object.entries(value).forEach(([key, child]) => {
      if (isTopLevelCheckedImageKey(pathName, key) || isNonContentValidationObject(key)) return;
      const childPath = `${pathName}.${key}`;
      if (isImageLikeKey(key) && hasMeaningfulValue(child)) {
        fields.push(childPath);
        return;
      }
      collectUnsupportedNestedImageFields(child, childPath, fields, depth + 1);
    });
  }

  function isTopLevelCheckedImageKey(pathName, key) {
    return pathName === 'slide' && ['image', 'images', 'gallery', 'media', 'mediaCount', 'imageSlots', 'slotCount'].includes(key);
  }

  function isNonContentValidationObject(key) {
    return ['chart', 'charts', 'table', 'tables', 'speakerNotes', 'speaker_notes', 'presenterNotes', 'presenter_notes'].includes(key);
  }

  function isImageLikeKey(key) {
    return ['image', 'images', 'gallery', 'media', 'photo', 'photos', 'picture', 'pictures', 'src', 'path'].includes(key);
  }

  function looksLikeImagePath(value) {
    return /\.(png|jpe?g|webp|gif|bmp|tiff?|svg)(?:[?#].*)?$/i.test(String(value).trim());
  }

  function isVisualMediaLayout(layout) {
    return VISUAL_MEDIA_LAYOUTS.has(layout || '');
  }

  function hasUserImages(data) {
    return normalizeMediaImages(data || {}).length > 0;
  }

  function chartCount(data) {
    return normalizeMediaCharts(data || {}).length;
  }

  function hasChartData(data) {
    return chartCount(data) > 0;
  }

  function hasTableData(data) {
    if (!data) return false;
    if (data.table && normalizeTableRows(data.table).length) return true;
    return Array.isArray(data.tables) && data.tables.some((table) => normalizeTableRows(table).length);
  }

  function hasCollectionOutside(slide, allowedKeys) {
    const allowed = new Set(allowedKeys || []);
    const keys = ['items', 'sections', 'columns', 'steps', 'nodes', 'layers', 'metrics', 'notes', 'insights', 'agenda', 'lanes', 'captions', 'points'];
    return keys.some((key) => slide?.[key] !== undefined && slide?.[key] !== null && !allowed.has(key));
  }

  function layoutAllowedByContent(slide, layout) {
    if (!layout) return false;
    if (layout === 'chart') return hasChartData(slide) && !hasCollectionOutside(slide, ['insights', 'notes']);
    if (layout === 'dashboard') return chartCount(slide) >= 2 && normalizeSections(slide.metrics || slide.items || []).length > 0 && !hasCollectionOutside(slide, ['metrics', 'items']);
    if (layout === 'dataSheet') return hasTableData(slide) && !hasCollectionOutside(slide, ['notes', 'insights']);
    if (isVisualMediaLayout(layout)) return hasUserImages(slide) || (hasChartData(slide) && !hasCollectionOutside(slide, ['items', 'insights', 'points', 'captions', 'metrics']));
    return true;
  }

  function filterLayoutCandidates(slide, candidates) {
    return (candidates || []).filter((layout) => layoutAllowedByContent(slide, layout));
  }

  function replacementCandidatesForLayout(layout) {
    const replacements = {
      textGrid: ['sectionList', 'fourCards', 'matrix', 'radial', 'chart', 'dataSheet'],
      article: ['sectionList', 'textGrid', 'fourCards', 'radial', 'matrix', 'chart', 'dataSheet'],
      fourCards: ['sectionList', 'textGrid', 'matrix', 'radial', 'chart', 'dataSheet'],
      matrix: ['sectionList', 'textGrid', 'fourCards', 'radial', 'chart', 'dataSheet'],
      statement: ['sectionList', 'agenda', 'fourCards', 'textGrid', 'radial', 'chart', 'dataSheet', 'caseStudy'],
      compare: ['swimlane', 'matrix', 'agenda', 'chart', 'dataSheet'],
      timeline: ['roadmap', 'swimlane', 'agenda', 'chart', 'dataSheet'],
      pipeline: ['roadmap', 'swimlane', 'agenda', 'chart', 'dataSheet'],
    };
    return replacements[layout] || ['sectionList', 'fourCards', 'textGrid', 'radial', 'roadmap', 'swimlane', 'chart', 'dataSheet'];
  }

  function suggestedLayoutsForSlide(slide, currentLayout, limit = 4) {
    const candidates = filterLayoutCandidates(slide, replacementCandidatesForLayout(currentLayout))
      .filter((layout) => layout !== currentLayout);
    const ordered = [];
    if (slide.layoutAlt && layoutAllowedByContent(slide, slide.layoutAlt) && slide.layoutAlt !== currentLayout) ordered.push(slide.layoutAlt);
    candidates.forEach((layout) => {
      if (!ordered.includes(layout)) ordered.push(layout);
    });
    return ordered.slice(0, limit);
  }

  function explicitMediaCount(data) {
    const n = Number(data.mediaCount || data.imageSlots || data.slotCount);
    return Number.isFinite(n) && n > 0 ? Math.floor(n) : 0;
  }

  function isMediaGridLayout(layout) {
    return ['mediaGrid', 'gallery', 'imageGrid', 'pairedMedia'].includes(layout || '');
  }

  function resolveMediaSlotCount(data) {
    const explicit = explicitMediaCount(data);
    if (explicit) return clamp(explicit, 1, 6);
    const images = normalizeMediaImages(data);
    const charts = normalizeMediaCharts(data);
    const captions = normalizeSlotItemsForValidation(data.captions || data.items || data.sections || []);
    const declaredMedia = Array.isArray(data.media) ? data.media.length : 0;
    return clamp(Math.max(images.length, charts.length, captions ? captions.length : 0, declaredMedia, 1), 1, 6);
  }

  function validateRenderableDataBlocks(slide, index, style, errors) {
    const layout = slide.layout || defaultLayoutForStyle(style);
    const chartLayouts = ['chart', 'dashboard', 'media', 'mediaGrid', 'gallery', 'imageGrid', 'imageHero', 'quoteImage', 'textImage', 'caseStudy'];
    const blocks = [...(slide.blocks || [])];
    if (!chartLayouts.includes(layout)) {
      blocks.push(...(slide.charts || []).map((chart) => ({ ...chart, type: 'chart' })));
    }
    blocks.push(...(slide.tables || []).map((table) => ({ ...table, type: 'table' })));
    blocks.forEach((block) => {
      if (!block || block.enabled === false) return;
      if ((block.type === 'chart' || block.type === 'table') && !hasExplicitBox(block)) {
        errors.push(`slide ${index + 1} has unpositioned ${block.type} block. Put it in layout "media"/"mediaGrid", remove it, or set x/y/w/h explicitly so it is not silently skipped.`);
      }
    });
  }

  function validateThinContent(slide, index, style, errors) {
    if (slide.allowSparseContent) return;
    const layout = slide.layout || defaultLayoutForStyle(style);
    if (['matrix', 'bigNumbers', 'kpiTower', 'dashboard', 'imageHero', 'caseStudy'].includes(layout) || slide.__canonicalLayout === 'image-matrix') return;
    const candidates = normalizeSections(slide.sections || slide.items || slide.columns || slide.nodes || slide.layers || slide.steps || slide.milestones || slide.agenda || []);
    if (candidates.length < 3) return;
    const titleOnly = candidates.filter((item) => {
      const title = item.title || item.label || item.name;
      const body = item.body || item.desc || item.note || item.text || item.summary || item.detail || (Array.isArray(item.points) && item.points.length) || (Array.isArray(item.bullets) && item.bullets.length) || (Array.isArray(item.list) && item.list.length);
      return title && !body;
    }).length;
    if (titleOnly >= Math.ceil(candidates.length * 0.6)) {
      errors.push(`slide ${index + 1} has ${titleOnly}/${candidates.length} title-only items. Add body/desc/note/points for each point, change layout, or set allowSparseContent:true only for intentional sparse draft slides.`);
    }
  }

  function warnLayoutVariety(spec) {
    let runLayout = null;
    let runStart = 0;
    let runLength = 0;
    const flush = () => {
      const slide = spec.slides[Math.min(spec.slides.length - 1, runStart + runLength - 1)] || {};
      const suggestions = suggestedLayoutsForSlide(slide, runLayout);
      const suffix = suggestions.length
        ? ` Suggested text/data-compatible alternatives: ${suggestions.map((layout) => `"${layout}"`).join(', ')}.`
        : ' No image/media-slot layout is suggested unless this page provides images or chart data.';
      if (runLength >= 3) {
        console.warn(
          `Warning: slides ${runStart + 1}-${runStart + runLength} use layout "${runLayout}" consecutively. Change at least one page to an equivalent layout by editing slide.layout in JSON so the deck does not feel repetitive.${suffix}`
        );
      } else if (runLength === 2) {
        console.warn(
          `Notice: slides ${runStart + 1}-${runStart + 2} both use layout "${runLayout}". If the content is not intentionally paired, consider alternating layouts for visual variety.${suffix}`
        );
      }
    };
    spec.slides.forEach((slide, index) => {
      const layout = slide.__canonicalLayout || slide.layout || (spec.style === 'magazine' ? 'textImage' : 'statement');
      if (layout === runLayout) {
        runLength += 1;
      } else {
        flush();
        runLayout = layout;
        runStart = index;
        runLength = 1;
      }
    });
    flush();
  }

  function validationCmbTextItems(slide) {
    return normalizeSections(slide.sections || slide.items || slide.columns || slide.points || slide.agenda || []);
  }


  return {
    normalizeLayoutCompatibility,
    validateSpecSlots,
    warnLayoutVariety,
    resolveMediaSlotCount,
    normalizeSlotItemsForValidation,
  };
};
