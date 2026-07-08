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
        -> renderMagazine() / renderSwiss() / renderCmb()
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
- 展开导出 `text-capacity.js` 中的字数容量指南工具。
- 保持外部模板脚本只依赖一个稳定入口。

### `cli.js`

命令行流程编排。

核心逻辑：

1. `parseArgs(argv)` 读取命令参数。
2. 如果传入 `--capacity-guide <style>`，直接输出对应 style 的 layout 字数容量指南并结束。
3. 判断是 `--sample` 还是 `--spec`。
4. 解析 `specDir` 和 `outPath`。
5. 从 `samples.sampleSpec()` 或 `spec-io.loadSpecFile()` 得到 spec 对象。
6. 调用 `engine.normalizeSpec()` 做 style/theme、槽位、文本容量、布局多样性校验。
7. 如果传入 `--write-normalized-spec`，调用 `writeNormalizedSpec()`。
8. 调用 `buildDeck()` 输出 PPTX。

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
- `engine.js` 的 `renderByStyle()` 和对应 `renderXxx()`
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

### `text-capacity.js`

Layout 字数容量指南和 JSON 预检 warning。

核心能力：

- `layoutCapacityGuide(style)`：返回机器可读的 style/layout/field 建议字数范围。
- `layoutCapacityMarkdown(style)`：返回用户/模型可读的 Markdown 表格，用于生成 JSON 前参考。
- `writeLayoutCapacityGuide(style, outPath)`：根据后缀写出 `.md` 或 `.json`。
- `warnSpecTextCapacity(spec)`：在 `normalizeSpec()` 阶段检查 JSON 字段是否超过建议范围；超出时 warning，要求修改 JSON 后重新生成。

修改 layout 文本框尺寸、字段名或数量时，应同步更新这里的范围，否则模型生成 JSON 前拿到的容量提示会失真。

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

## `engine.js` 核心结构

`engine.js` 仍然较长，因为它承载了真正的 PowerPoint 原生渲染逻辑。阅读时建议按下面的分区看。

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
- 校验是否允许自动改 layout。
- 调用 `diversifyRepeatedLayouts()`。
- 调用 `validateSpecSlots()`。
- 调用 `warnThinContent()`、`warnSpecTextCapacity()` 和 `warnLayoutVariety()`。
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
- `renderMagazine(slide, ctx)`
- `renderSwiss(slide, ctx)`
- `renderCmb(slide, ctx)`

`renderByStyle()` 是 style 分发入口。新增 style 时，要在这里注册新 renderer。

每个 `renderXxx()` 的基本模式一致：

1. 读取 `ctx.slideSpec.layout`。
2. 根据 layout 判断背景色、前景色、accent 页面状态。
3. 绘制背景和页眉/页脚 chrome。
4. 建立 layout name 到具体 layout renderer 的 map。
5. 调用对应 renderer。
6. 调用 `renderDataBlocks()` 插入后置 block/chart/table。

CMB 当前复用了大量 Swiss layout renderer，只对 CMB 独有背景、页眉、封面、statement、closing 做独立实现。

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
- 没有图片和图表时显示图片占位符。
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
- 防止字段名不匹配，例如 layout 只读 `sections`，用户却填了会被忽略的字段。
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
- `diversifyRepeatedLayouts()`
- `chooseDiversifiedLayout()`
- `warnLayoutVariety()`

实现重点：

- `warnThinContent()` 检查大量 title-only item，提示补 body/desc/note。
- `warnLayoutVariety()` 发现连续 3 页以上同 layout 时给 warning。
- `diversifyRepeatedLayouts()` 默认只给建议，不改 JSON。
- 只有用户使用 `--diversify-layouts --write-normalized-spec` 时，才会真实改 slide.layout，并写出 normalized JSON。

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
4. 在 `engine.js` 中新增 `renderNewStyle(slide, ctx)`。
5. 在 `renderByStyle()` 的 `renderers` map 中注册新 style。
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
2. 在 `engine.js` 中新增具体 renderer，例如 `swissNewLayout()`。
3. 在对应 `renderXxx()` 的 `renderers` map 中注册 `newLayout`。
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
- 修改 layout 或渲染视觉问题才看 `engine.js`。
- 若继续拆分 `engine.js`，建议下一步按 `styles/`、`renderers/`、`media/`、`validation/`、`geometry/` 分目录拆，而不是再引入一个更大的工具文件。
