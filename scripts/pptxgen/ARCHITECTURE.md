# pptxgen 生成器架构说明

本文档说明 `scripts/pptxgen/` 下各个 JS 文件的核心职责、调用链和主要实现点。目标是让后续新增 `style`、新增 `layout`、调整图表/图片/文本逻辑时，不必从头阅读过长的 `engine.js`。

## 总体调用链

普通命令：

```bash
node scripts/generate-pptx.js --spec path/to/deck.json --out outputs/deck.pptx
```

实际执行路径：

```text
scripts/generate-pptx.js
  -> scripts/pptxgen/cli.js
    -> spec-io.loadSpecFile() or samples.sampleSpec()
    -> engine.normalizeSpec()
    -> spec-io.writeNormalizedSpec()          # only when requested
    -> engine.buildDeck()
      -> renderByStyle()
        -> templates/magazine.js / templates/swiss.js / templates/cmb.js
          -> concrete layout renderer
          -> renderDataBlocks()
```

模板脚本兼容路径：

```js
const { buildDeck } = require('../scripts/generate-pptx.js');
```

`generate-pptx.js` 现在只是兼容入口，会继续导出 `buildDeck`、`sampleSpec` 和 JSON spec 工具。

## 文件职责

### `generate-pptx.js`

薄入口文件。

- 当作为 CLI 运行时，调用 `scripts/pptxgen/cli.js` 的 `main()`。
- 当被 `assets/template-*.js` 通过 `require()` 引入时，转发 `scripts/pptxgen/index.js` 的公共导出。
- 不放业务逻辑，不放设计配置，不放 layout 实现。

### `index.js`

公共 API 汇总。

- 导出 `buildDeck`、`normalizeSpec`、`sampleSpec`。
- 展开导出 `spec-io.js` 中的 JSON 解析和写出工具。
- 展开导出 `layout-schema.js` 的 canonical layout 协议工具。
- 展开导出 `layout-examples.js` 的 style 介绍和 JSON 示例工具。
- 保持外部模板脚本只依赖一个稳定入口。

### `cli.js`

命令行流程编排。

核心逻辑：

1. `parseArgs(argv)` 读取命令参数。
2. 如果传入 `--style-guide`，输出 style 介绍供交互选择。
3. 如果传入 `--layout-examples <style>`，输出该 style 的 31 种 canonical JSON 示例 Markdown。
4. 判断是 `--sample` 还是 `--spec`。
5. 解析 `specDir` 和 `outPath`。
6. 从 `samples.sampleSpec()` 或 `spec-io.loadSpecFile()` 得到 spec 对象。
7. 调用 `engine.normalizeSpec()` 做 canonical schema、style/theme、槽位和布局多样性校验。
8. 如果传入 `--write-normalized-spec`，调用 `writeNormalizedSpec()`。
9. 调用 `buildDeck()` 输出 PPTX。

这里不直接碰 PowerPoint 渲染细节。

### `config.js`

设计配置和全局常量。

核心内容：

- `SLIDE`：宽屏页面尺寸、默认边距。
- `THEMES`：所有 style/theme 的颜色配置。
- `FONTS`：字体名。
- `READABILITY`：最小可读字号等可读性下限。
- `BASIC_ICON_NAMES`、`ICON_ALIASES`：图标别名和 fallback 名称。
- `STYLE_ORDER`：当前支持的 style 列表。
- `DEFAULT_THEMES`：每个 style 的默认 theme。
- `LUCIDE_STATIC_ICON_DIR`：兼容旧 `lucide-static` 的 SVG 路径。
- `isSupportedStyle()`：判断 style 是否可用。
- `defaultThemeForStyle()`：取默认 theme。

新增 theme 时，主要改 `THEMES`。

新增 style 时，至少改：

- `STYLE_ORDER`
- `DEFAULT_THEMES`
- `THEMES[style]`
- `engine.js` 的 `getTemplateRenderers()`
- `scripts/pptxgen/templates/<style>.js`
- `samples.js` 的 `sampleSpec(style)`

### `spec-io.js`

JSON spec 输入输出。

核心能力：

- `parseArgs(argv)`：解析 CLI 参数。
- `loadSpecFile(specPath)`：读取 UTF-8 JSON 文件并调用 `parseSpecJson()`。
- `parseSpecJson(raw, sourceName)`：依次尝试严格 JSON、Markdown fenced JSON、宽松 JSON。
- `extractJsonPayload(raw)`：从 Markdown 或混合文本中提取 `{...}` / `[...]`。
- `normalizeLooseJson(raw)`：组合注释移除、智能引号修复、尾逗号移除。
- `normalizeSmartJsonQuotes(text)`：把中英文弯引号、单引号等尽量规整成 JSON 双引号。
- `stripJsonComments(text)`：移除 `//` 和 `/* */` 注释，同时尽量不破坏字符串。
- `removeTrailingCommas(text)`：移除对象/数组末尾多余逗号。
- `writeNormalizedSpec(spec, outPath)`：去掉内部字段后写出严格 UTF-8 JSON。

注意：

- 该模块不支持未加引号的 JSON key，例如 `{slides: []}` 仍应报错。
- 宽松修复只处理常见生成器/LLM 输出问题，不替代严格 JSON 生成。

### `layout-schema.js`

跨 style 的 canonical layout 与字段协议。

- `LAYOUTS`：31 个 `deck-*`、`text-*`、`image-*`、`data-*` 布局定义。
- `validateCanonicalSpec()`：拒绝旧 layout、旧集合字段和媒体/数据类型不匹配。
- `createRendererSpec()` / `createRendererSlide()`：只在内部把 canonical spec 适配给成熟 renderer；外部 JSON 不暴露旧字段。

### `layout-examples.js`

style 选择说明和全布局 JSON 示例生成器。

- `styleGuideMarkdown()`：输出三种 style 的用途说明。
- `layoutExamplesMarkdown(style)`：输出指定 style 的 31 个完整 slide JSON 示例，不包含字数限制。
- `writeLayoutExamples(style, outPath)`：以 UTF-8 写入 Markdown。

### `samples.js`

内置 sample spec。

核心能力：

- `sampleSpec(style = 'swiss')` 返回对应 style 的完整示例 spec。
- 当前支持 `swiss`、`magazine`、`cmb`。
- `cli.js --sample --sample-style xxx` 会走这里。

新增 style 时，应给这里加一个能覆盖核心版式的最小可用 sample。

### `errors.js`

统一失败出口。

- `fail(message)` 打印错误并用退出码 `2` 退出。
- CLI 参数错误、spec 结构错误、槽位校验错误都可以使用它。

### `validation.js`

Spec、layout 和槽位校验模块。

- `normalizeLayoutCompatibility()`：兼容旧 JSON 字段结构，例如 compare 左右列。
- `validateSpecSlots()`：统一入口，检查字段类型、必填字段、图片/图表/表格槽位、文本集合数量、静默忽略字段、内容过空等问题。
- `warnLayoutVariety()`：提示连续重复 layout。
- `resolveMediaSlotCount()`：根据图片、图表、caption 和显式 `mediaCount` 计算媒体槽位数量。

新增 layout 字段规则、缺字段 error/warning、静默忽略字段检查时，优先改这里。

### `media.js`

图片和媒体槽位模块。

- `prepareImageAspectAssets()`：生成前预读取本地图片尺寸。
- `prepareSvgImageAssets()`：SVG 图片/logo 默认转 PNG，避免兼容性问题。
- `resolveImage()`：按 spec 目录、当前目录、skill assets、skill root 解析图片路径。
- `addImageAsset()`：按原始比例居中插入图片，避免 PNG/JPG 变形。
- `addImagePlaceholder()` / `addImageOrPlaceholder()` / `addStatementImageSlot()`：媒体占位符和图片槽位。
- `normalizeMediaImages()` / `normalizeMediaCharts()`：统一解析 image/images/gallery/media/charts。
- `addMediaOrChart()` / `addMediaGrid()`：媒体槽位中插入图片、图表或占位符。

Linux 路径、图片比例、SVG/PNG/JPG/WebP、空媒体槽位等问题优先看这里。

### `icons.js`

Lucide 图标和项目符号图标模块。

- `prepareIconAssets()`：生成前把图标 rasterize 到 PNG 缓存。
- `iconAlias()`：图标别名映射。
- `addSvgIcon()` / `addInlineIcon()` / `addBulletIcon()`：模板中使用的图标绘制入口。
- `normalizeHex()`：颜色标准化工具，供模板背景和图标复用。

图标库、图标兼容性、SVG 转 PNG、lucide 包兼容问题优先看这里。

### `blocks.js`

自由数据块、图表和表格模块。

- `renderDataBlocks()`：渲染 `blocks[]`、显式定位的 chart/table/text/callout。
- `addChartBlock()`：PptxGenJS 原生图表。
- `addTableBlock()`：PptxGenJS 表格。
- `normalizeChartData()` / `normalizeTableRows()`：图表和表格数据归一化。

新增图表类型、表格样式、自由 block 类型时，优先改这里。

## `engine.js` 核心结构

`engine.js` 保留生成主流程、spec 校验、图片/图表/文本/图标等共用能力。具体模板设计和 layout 渲染放在 `templates/` 目录；新增或修改某个 style 时，优先只改对应模板文件。

### 1. 顶部依赖和运行时缓存

核心对象：

- `pptxgen`：PowerPoint 生成库。
- `pptx`：当前生成进程使用的 pptxgen 实例。
- `CHART_TYPES`：把 spec 中的 `bar`、`line`、`doughnut` 等映射到 pptxgenjs chart type。
- `LUCIDE_ICON_CACHE`：缓存已经生成/读取过的 icon SVG。
- `IMAGE_ASPECT_CACHE`：缓存图片宽高比，避免重复读取文件。
- `LUCIDE_MODULE`：懒加载 `lucide` 包。

### 2. Spec 归一化和 deck 生成

核心函数：

- `normalizeSpec(spec, options)`
- `buildDeck(spec, specDir, outPath)`
- `enforceReadableSlideText(slide)`
- `readableTextOptions(options, text)`
- `warnTextExceedsBox(text, options)`
- `estimatedBoxTextCapacity(options, explicitFontSize)`

`normalizeSpec()` 做这些事：

- 填默认 `style`。
- 校验 style 是否在 `config.STYLE_ORDER` 中。
- 填默认 theme。
- 校验 `slides` 非空。
- 保持用户 JSON 中的 `slide.layout` 不变；页面类型只由 JSON 字段决定。
- 调用 `validateSpecSlots()`。
- 调用 canonical schema、槽位完整性、内容完整性和 `warnLayoutVariety()` 检查。
- 标记 `spec.__normalized = true`。

`buildDeck()` 做这些事：

- 如 spec 尚未归一化，先调用 `normalizeSpec()`。
- 设置 pptx 元数据和自定义宽屏尺寸。
- 为每页创建 slide。
- 给 slide 打 `enforceReadableSlideText()` 补丁，抬高过小字号，并在最终 `addText()` 前按文本框尺寸估算最大可容纳字数；超出时只输出 warning，不截断、不改写原文，提示修改 JSON 后重新生成。
- 构造 `ctx = { spec, slideSpec, theme, specDir, index, total }`。
- 调用 `renderByStyle(spec.style, slide, ctx)`。
- 写出 PPTX。

### 3. Style 分发

核心函数：

- `renderByStyle(style, slide, ctx)`
- `templates/magazine.js`
- `templates/swiss.js`
- `templates/cmb.js`

`renderByStyle()` 是 style 分发入口。新增 style 时，要新增 `templates/<style>.js`，并在 `getTemplateRenderers()` 注册 renderer。

模板文件之间不互相 `require`。即使某些 layout 代码相似，也应复制到对应模板文件内，避免修改一个模板时影响另一个模板。

每个 `renderXxx()` 的基本模式一致：

1. 读取 `ctx.slideSpec.layout`。
2. 根据 layout 判断背景色、前景色、accent 页面状态。
3. 绘制背景和页眉/页脚 chrome。
4. 建立 layout name 到具体 layout renderer 的 map。
5. 调用对应 renderer。
6. 调用 `renderDataBlocks()` 插入后置 block/chart/table。

CMB 内部保留了与 Swiss 相似的 layout 代码副本，但运行时不依赖 Swiss 模板文件。这样修改 CMB 不会影响 Swiss，修改 Swiss 也不会影响 CMB。

### 4. 背景、页眉、品牌和图片基础能力

核心函数：

- CMB 背景：`addCmbBackground()`、`cmbBackgroundSvg()`
- CMB 页眉：`addCmbChrome()`
- CMB logo：`addCmbLogoMark()`、`addCmbLogoWatermark()`、`resolveCmbLogoMark()`
- 通用背景：`addDecorativeBackground()`、`decorativeBackgroundSvg()`
- 通用 chrome：`addChrome()`、`addBrandLogo()`、`resolveBrandLogo()`、`addFoot()`
- 图片插入：`addImageAsset()`、`fitImageBoxToAspect()`、`imageAspectRatio()`
- SVG/PNG 尺寸：`svgAspectRatio()`、`pngAspectRatio()`
- SVG 工具：`readSvgWithOpacity()`、`svgDataUri()`、`svgEsc()`
- 页眉安全区：`hasBrandHeader()`、`pageHeadY()`、`pageHeadSafeBottom()`

实现重点：

- `addImageAsset()` 会调用 `fitImageBoxToAspect()`，保证 logo 和图片不被压扁。
- SVG logo 可通过 `readSvgWithOpacity()` 注入透明度。
- CMB 页眉保持白色 band，内容区再做渐变和水印。
- `pageHeadSafeBottom()` 用于避免标题/内容撞到品牌页眉区域。

### 5. Magazine layout 渲染器

典型函数：

- `magazineCover()`
- `magazineSection()`
- `magazineBigNumbers()`
- `magazineQuoteImage()`
- `magazineMedia()`
- `magazineMediaGrid()`
- `magazinePipeline()`
- `magazineCompare()`
- `magazineArticle()`
- `magazineDataSheet()`
- `magazineChart()`
- `magazineDashboard()`
- `magazineAgenda()`
- `magazineCaseStudy()`
- `magazinePyramid()`
- `magazineRadial()`
- `magazineRoadmap()`
- `magazineSwimlane()`

Magazine 的视觉特征：

- 更偏 editorial / magazine。
- 大标题常用 serif 中文字体。
- 背景有纸感、细线、页眉页脚。
- 适合叙事、观点、数据大字报。

兼容函数：

- `magazineStatementCompat()`
- `magazineMatrixCompat()`
- `magazineImageHeroCompat()`

这些用于接收 Swiss/CMB 常见 layout 名称，让切换 style 时不至于直接丢内容。

### 6. Swiss layout 渲染器

典型函数：

- `swissCover()`
- `swissStatement()`
- `swissKpiTower()`
- `swissDuoCompare()`
- `swissTimeline()`
- `swissMatrix()`
- `swissFourCards()`
- `swissImageHero()`
- `swissTextGrid()`
- `swissDataSheet()`
- `swissChart()`
- `swissDashboard()`
- `swissAgenda()`
- `swissCaseStudy()`
- `swissPyramid()`
- `swissRadial()`
- `swissRoadmap()`
- `swissSwimlane()`

Swiss 的视觉特征：

- 直角、网格、强 accent 色。
- 偏数据、方法论、结构化表达。
- 多数 CMB 内容页也复用 Swiss 的结构 renderer，再套 CMB 背景和品牌 chrome。

兼容函数：

- `swissSectionCompat()`
- `swissBigQuoteCompat()`
- `swissQuoteImageCompat()`
- `swissTextImageCompat()`
- `swissImageGridCompat()`

### 7. CMB layout 渲染器

独立实现：

- `cmbCover()`
- `cmbSection()`
- `cmbStatement()`
- `cmbClosing()`
- `cmbBriefing()`: CMB 高密度纯文本总分总页面，顶部 summary，中部 2-4 个分析卡片，底部 takeaway。
- `cmbTextWeave()`: CMB 非均匀纯文本卡片页面，支持 1-6 个文本块，避免普通三栏网格在长文本下显得杂乱。
- `addCmbTextCard()` / `addBulletedCardBody()`: CMB 卡片文本渲染；只有 JSON 中显式提供 `points` / `bullets` / `list` 数组时才按项目符号分点渲染，不再自动把长正文拆句。

CMB 专用映射：

- `article`、`sectionList`、`briefing`、`executiveBrief`、`contentBrief` -> `cmbBriefing()`
- `agenda` -> `cmbAgenda()`
- `textGrid`、`fourCards`、`textWeave`、`contentSynthesis`、`denseText` -> `cmbTextWeave()`

其中 `fourCards` 是兼容旧 JSON 的历史名称。在 CMB 下它不会固定渲染 4 张卡片，而是复用 `cmbTextWeave()`，按输入条目数自适应为 1-6 张文本卡片。

CMB 继续复用 Swiss renderer 的 layout：

- `dashboard`
- `dataSheet`
- `chart`
- `kpiTower`
- `bigNumbers`
- `media`
- `mediaGrid`
- `gallery`
- `imageGrid`
- `compare`
- `duoCompare`
- `timeline`
- `pipeline`
- `roadmap`
- `matrix`
- `caseStudy`
- `pyramid`
- `radial`
- `swimlane`
- `imageHero`
- `quoteImage`
- `bigQuote`
- `textImage`

实现重点：

- CMB 是独立 `style: "cmb"`，不是 Swiss theme。
- 页眉使用 `logos/cmb-logo-lockup.png`。
- 页眉外装饰和水印使用 `logos/cmb-logo-mark.svg`。
- 所有 logo 通过 `addImageAsset()` 保持比例。
- 内容页右上角水印透明度由 `addCmbLogoWatermark()` 控制。
- 高密度文本优先换用 `briefing` / `textWeave`；过长正文只 warning，不截断。

### 8. 文本、图标和 bullet

核心函数：

- `addPageHead()`
- `addCallout()`
- `addBullets()`
- `normalizeBulletItem()`
- `iconAlias()`
- `toKebabIconName()`
- `loadLucideModule()`
- `lucidePackageIconSvg()`
- `lucideStaticIconSvg()`
- `lucideIconSvg()`
- `fallbackSvgIconData()`
- `svgIconData()`
- `addSvgIcon()`
- `defaultContentIcon()`
- `itemIcon()`
- `addInlineIcon()`
- `addBulletIcon()`

实现重点：

- 优先使用 `lucide` 包。
- 若旧环境安装过 `lucide-static`，也兼容读取静态 SVG。
- 没有可用 icon 时回退到内置基础图形。
- `ICON_ALIASES` 在 `config.js` 中维护，允许把 `risk`、`trend`、`people` 等业务词映射到 lucide icon。

### 9. 图片、媒体槽位和 caption

核心函数：

- `addImagePlaceholder()`
- `addImageOrPlaceholder()`
- `addStatementImageSlot()`
- `normalizeMediaImages()`
- `normalizeMediaCharts()`
- `addMediaOrChart()`
- `addMediaGrid()`
- `imageCaption()`
- `resolveImage()`
- `addCaption()`

实现重点：

- 用户提供图片时优先插入图片。
- 没有图片但显式提供 chart 时插入原生 PPT 图表。
- 没有图片和图表时校验失败；应提供素材或改用纯文本 layout。
- `mediaGrid`、`gallery`、`imageGrid` 的槽位数由图片数、图表数、caption 数或 `mediaCount` 决定。
- `resolveImage()` 负责把 spec 相对路径、技能内置 asset 路径解析成可用文件路径。

### 10. 图表、表格和后置数据块

核心函数：

- `renderDataBlocks()`
- `hasExplicitBox()`
- `addTextBlock()`
- `addChartBlock()`
- `addTableBlock()`
- `normalizeChartData()`
- `normalizeTableRows()`
- `chartPalette()`

实现重点：

- `blocks`、`charts`、`tables` 作为后置组件，只在显式提供 `x/y/w/h` 时插入。
- 缺少显式坐标的后置组件会被跳过并打印 warning，避免和正文重叠。
- layout 自带的 chart/table 区域不属于后置 blocks，直接由 layout renderer 控制。
- `normalizeChartData()` 兼容 `values`、`series` 等输入。
- `normalizeTableRows()` 兼容 `rows`、`data`、`headers` 等输入。

### 11. Spec 槽位校验

核心函数：

- `validateSpecSlots()`
- `validateMediaSlots()`
- `validateTextSlots()`
- `validateSlotCollection()`
- `validateIgnoredSlotFields()`
- `validateChartSlots()`
- `validateChartDataSlot()`
- `validateTableSlot()`
- `normalizeSlotItemsForValidation()`
- `slotItemHasDisplayText()`
- `explicitMediaCount()`
- `isMediaGridLayout()`
- `resolveMediaSlotCount()`

实现重点：

- 在真正渲染前检查字段和槽位数量。
- 防止传入 5 个 item，但 layout 只渲染 4 个导致内容丢失。
- 防止字段名不匹配；canonical JSON 的顶层集合统一为 `items`，内部适配字段不会暴露给用户。
- 防止媒体数量和 `mediaCount` 不匹配。
- `validateTextSlots()` 中的 `rules` 是每个 layout 支持数量和字段名的核心规则表。

修改 layout 支持数量时，通常需要同时改：

- 对应 layout renderer
- `validateTextSlots()` 中该 layout 的规则
- 必要时更新 `assets/template-cmb-all-layouts.js`
- 必要时更新 README/SKILL 说明

### 12. 内容质量和 layout 多样性提示

核心函数：

- `warnThinContent()`
- `warnLayoutVariety()`
- `suggestedLayoutsForSlide()`

实现重点：

- `warnThinContent()` 检查大量 title-only item，提示补 body/desc/note。
- `warnLayoutVariety()` 发现连续 3 页以上同 layout 时给 warning。
- 生成器不会自动修改 `slide.layout`；如需调整页面类型，必须手动编辑 JSON。

### 13. 自适应排版和几何工具

核心函数：

- `autoCardColumns()`
- `autoColumns()`
- `clampColumns()`
- `normalizeSections()`
- `safeBox()`
- `fitTitle()`
- `textVisualLength()`
- `estimateTextHeight()`
- `distributeRowHeights()`
- `clamp()`
- `rectCenter()`
- `rectEdgePoint()`
- `connectorBetweenRects()`
- `addRadialConnector()`

实现重点：

- `autoColumns()` / `autoCardColumns()` 让多分点 layout 根据数量自适应列数。
- `estimateTextHeight()` 用粗略字符宽度估算文本高度，帮助卡片高度分配。
- `distributeRowHeights()` 在固定总高度内按文本需求分配行高。
- `connectorBetweenRects()` 从矩形边缘到矩形边缘生成连线，避免 radial 连线覆盖中心卡片。
- `safeBox()` 限制后置 block 不越过页面安全区。

## 新增 style 的推荐步骤

1. 在 `config.js` 中新增 `THEMES[newStyle]`。
2. 把新 style 加入 `STYLE_ORDER`。
3. 在 `DEFAULT_THEMES` 中设置默认 theme。
4. 新增 `scripts/pptxgen/templates/newStyle.js`，导出 `createNewStyleTemplate(api)`。
5. 在 `engine.js` 的 `getTemplateRenderers()` 中 require 并注册新 style。
6. 复用已有 layout renderer，或新增独立 layout renderer。
7. 在 `samples.js` 中增加 `sampleSpec(newStyle)` 分支。
8. 增加或更新 `assets/template-new-style.js`。
9. 运行：

```bash
node --check scripts/generate-pptx.js
node --check scripts/pptxgen/cli.js
node --check scripts/pptxgen/config.js
node --check scripts/pptxgen/engine.js
node --check scripts/pptxgen/spec-io.js
node --check scripts/pptxgen/samples.js
node scripts/generate-pptx.js --sample --sample-style newStyle --out outputs/sample-new-style.pptx
node scripts/validate-pptx-native.js outputs/sample-new-style.pptx
node scripts/validate-pptx-layout.js outputs/sample-new-style.pptx
```

## 新增 layout 的推荐步骤

1. 选择目标 style：`magazine`、`swiss`、`cmb`，或决定是否三者都支持。
2. 在对应模板文件中新增具体 renderer，例如 `templates/swiss.js` 里的 `swissNewLayout()`。
3. 在同一模板文件的 `renderXxx()` layout map 中注册 `newLayout`。
4. 如果跨 style 兼容，给其他 style 注册兼容 renderer 或映射到相近 layout。
5. 在 `validateTextSlots()` 的 `rules` 中新增该 layout 的字段名、最大数量、最小数量。
6. 如果有媒体区，检查 `validateMediaSlots()` 的 `mediaLayouts` 是否需要加入该 layout。
7. 如果需要图表/表格，复用 `addChartBlock()`、`addTableBlock()` 或明确 layout 自带槽位。
8. 更新 sample/template，至少覆盖一个正常输入。
9. 运行 sample 和 layout validation。

## 维护建议

- 普通生成任务不要阅读 `engine.js`；先看 `SKILL.md`、用户 spec、必要的 `assets/template-*.js`。
- 修改主题或配色优先看 `config.js`。
- 修改 JSON 解析问题优先看 `spec-io.js`。
- 修改命令行参数优先看 `cli.js` 和 `spec-io.parseArgs()`。
- 修改单个 style 的 layout 或渲染视觉问题，优先看 `scripts/pptxgen/templates/<style>.js`。
- 修改跨模板共用能力，例如图片、图表、校验、字体可读性，再看 `engine.js`。
