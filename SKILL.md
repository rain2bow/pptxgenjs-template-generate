---
name: pptxgenjs-template-generate
description: 使用 PptxGenJS 基于结构化 JSON 规格生成可编辑的 PowerPoint .pptx 演示文稿。当 Codex 需要创建商务演示、招商银行风格演示、杂志风或瑞士风演示、文本密集型可编辑幻灯片、包含图片/图表/表格/图标/演讲者备注的幻灯片、容量警告、DOCX 内容解析导入，或用户友好的 Markdown 大纲时使用。
---

# PPTXGenJS 模板生成

该技能使用 pptxgenjs 创建可编辑的 PowerPoint 文件。幻灯片必须由 PowerPoint 原生文本框、形状、线条、图片、表格、图表、演讲者备注和图标构建。除非用户明确要求生成不可编辑的图片型演示文稿，否则不要将整页幻灯片栅格化为扁平图片。

## 核心规则

- 保持输出文件在 PowerPoint、WPS 和 Keynote 中可编辑。
- 所有 JSON、Markdown 和脚本输出均使用 UTF-8 编码。
- 使用 `JSON.stringify(data, null, 2)` 或等效的结构化写入方式编写规格文件。不要手动拼接 JSON 字符串。
- JSON 字符串中的英文直双引号必须转义为反斜杠加双引号。
- 普通演示文稿生成时，不要读取生成器源码文件。只有在调试或修改生成器时才读取脚本。
- 不要忽略生成器错误和警告。缺图、缺表、缺 chart data、必需集合缺失、字段冲突、标题过多正文不足会作为错误阻断生成；文本容量和布局审美提示保留为 warning，应尽量修复后重新生成。

## 字体排版

所有样式统一使用同一套字体系统：

- 中文以及中英文混排文本：Microsoft YaHei。
- 纯英文文本：Times New Roman。
- 演示文稿级别的封面、章节页和结尾页标题：36 pt。
- 内容页顶部标题：28 pt。
- 内容层级：要点标题 16 pt，普通正文 14 pt，密集文本或小区域文本 12 pt。
- PPTX 生成后，图表文本会统一规范为 Microsoft YaHei 和 12 pt。
- 生成器在 `scripts/pptxgen/engine.js` 中强制执行这些层级。不要添加布局专属字体族。

## 样式

支持的样式：

- `magazine`：编辑/杂志式布局，纸张质感背景，大字号，适合图片页和数据页。
- `swiss`：基于网格的商务布局，扁平几何风格，克制的强调色。
- `cmb`：招商银行风格，包含 CMB 红色主题和品牌标识支持。

支持的 CMB 主题：`classic`、`pearl`、`graphite`。
支持的 Swiss 主题：`ikb`、`lemon`、`green`、`orange`、`cmb`。
支持的 Magazine 主题：`ink`、`indigo`、`forest`、`kraft`、`dune`、`cmb`。

## 工作流程

1. 选择样式和主题。如果用户没有指定样式，招商银行或银行类需求使用 `cmb`，数据/产品/战略类演示使用 `swiss`，编辑叙事类演示使用 `magazine`。
2. 先根据用户输入创建标题级规划文件 `deck.plan.json`。该文件只写结构，不写正文：顶层写 `style`、`theme`、`title`、`slides`；每页写 `layout`、`title`；有分点、卡片、步骤或节点时，只写对应集合项的 `title` 或 `label`，不要写 `body` 正文。如果选择媒体槽位 layout，必须在 plan 中已经提供 `image`、`images`、`media`、`gallery`、`chart` 或 `charts`，只写 `mediaCount`、`imageSlots` 或 `allowEmptyMediaSlots` 不算有效素材。
3. 根据 `deck.plan.json` 生成真实容量指南。该指南只包含已规划页面，并按每页 layout、卡片数量、分点数量估算每个可填字段的容量：

   ```bash
   node scripts/generate-pptx.js --capacity-guide --spec outputs/deck.plan.json --out outputs/deck-capacity-guide.md
   ```

   该命令会先校验 `deck.plan.json`。如果页面 layout 不存在、集合字段与 layout 不匹配、集合项缺少标题字段、或标题级 plan 中提前写入正文类字段，会直接失败，先修正 plan 再继续。

4. 按 `deck-capacity-guide.md` 扩写完整 JSON 正文。容量指南会在每页表格后给出 `Example slide JSON` 作为字段结构参考。不要新增 guide 中没有列出的正文槽位；如果需要更换 layout 或改变分点数量，先更新 `deck.plan.json` 并重新生成容量指南。
5. 在最终生成前，将完整 JSON 规格转换为用户友好的 Markdown 大纲：

   ```bash
   node scripts/spec-to-md.js --spec path/to/deck.json --out path/to/deck-outline.md
   ```

6. 生成 PPTX：

   ```bash
   node scripts/generate-pptx.js --spec path/to/deck.json --out outputs/deck.pptx
   ```

   最终生成也会再次校验槽位完整性；例如正文卡片类 layout 的集合项不能只有 `body` 而没有 `title`，也不能只有 `title` 而没有正文说明。

7. 如果生成失败，优先修复缺图、缺表、缺 chart data、必需集合缺失、字段冲突、标题过多正文不足等问题；如果 warning 提到文本容量或重复布局，也应尽量更新 JSON 并重新生成。
8. 验证原生/可编辑结构和布局安全性：

   ```bash
   node scripts/validate-pptx-native.js outputs/deck.pptx
   node scripts/validate-pptx-layout.js outputs/deck.pptx
   ```
## 安装依赖

在技能目录中安装一次 Node 依赖：

```bash
npm install
```

必需运行时包包括 `pptxgenjs`、`jszip`、`sharp` 和 `lucide`。

- `sharp` 是必需依赖。如果缺失，应安装它，而不是退回到直接插入 SVG。
- 为了提升 Office 和 LibreOffice 兼容性，SVG 图标和 SVG 标志会被栅格化为 PNG。
- 内置 CMB 标志路径可在 JSON 中引用为 `logos/cmb-logo-lockup.png` 和 `logos/cmb-logo-mark.svg`。

## 常用命令

生成内置示例：

```bash
node scripts/generate-pptx.js --sample --sample-style cmb --out outputs/sample-cmb.pptx
node scripts/generate-pptx.js --sample --sample-style swiss --out outputs/sample-swiss.pptx
node scripts/generate-pptx.js --sample --sample-style magazine --out outputs/sample-magazine.pptx
```

生成 CMB 全布局检查文件和最大文本容量检查文件：

```bash
npm run sample:cmb:layouts
npm run sample:cmb:max-text
```

布局只能通过 JSON 中每页的 `layout` 字段手动修改；生成器不会自动改写 layout。

从 JSON 生成 Markdown 大纲：

```bash
node scripts/spec-to-md.js --spec path/to/deck.json --out outputs/deck-outline.md
```


## DOCX 导入工作流

当用户提供 `.docx` 并要求基于文档生成 PPTX 时，使用 DOCX 专用入口，不要先手工复制 Word 内容：

```bash
node scripts/docx-to-pptx.js \
  --docx path/to/source.docx \
  --write-extracted outputs/from-docx.extracted.json \
  --write-md outputs/from-docx.extracted.md
```

规则：

- `--write-extracted` 必须在复杂文档或含图片文档中使用，便于检查文本 block、图片 block、段落/run 相对位置、inline/anchor 信息、`wp:extent` 显示尺寸和 `a:srcRect` 裁剪参数。
- 严禁把 DOCX 解析 block 顺序直接当成 PPT 页面结构；block 顺序只能作为阅读材料和图片引用依据。
- PPT JSON 中引用图片时，必须使用 extracted 结果中的 `path`。无裁剪图片会保留原始文件，带裁剪图片会输出按源像素裁剪后的 PNG。
- 图片会提取到输出目录旁的 `*-docx-assets/` 目录。DOCX 导入阶段只应用 Word 中的图片裁剪，不按 Word 显示尺寸重采样；PPTX 生成阶段会根据实际图片比例放入页面槽位。
- 本命令不会生成 `from-docx.spec.json`，也不会按 DOCX 顺序规则转换 PPT。必须由模型或人工阅读 `from-docx.extracted.json` / `from-docx.extracted.md`，根据语义重新规划页面并编写 PPT JSON。
- DOCX 导入只保留文本、图片顺序参考、图片视觉裁剪后的资产；不要信任 Word 排版本身，不要把 DOCX 布局机械映射为 PPT 布局。生成 PPT 前必须先基于语义写 JSON，再检查 spec 和 PPTX。
- Linux 上传入绝对路径或带空格路径时，直接把路径作为 CLI 参数传入；不要把路径拆成多个参数，也不要通过 shell 拼接图片路径。

## JSON 规格结构

最小规格字段：

- 顶层字段：`style`、`theme`、`title`、`subtitle`、`author`、`company`、`slides`。
- 幻灯片级字段：`layout`、`kicker`、`title`、`subtitle`、`body`、`summary`、`conclusion`。
- `summary` 在 `media`、`caseStudy` 和 CMB `briefing` 中会作为正文/摘要显示；如果同页同时写 `body` 和 `summary`，优先渲染 `body`。
- 集合字段：`sections`、`items`、`columns`、`steps`、`nodes`、`layers`、`lanes`、`metrics`。
- 媒体和数据字段：`image`、`images`、`media`、`captions`、`chart`、`charts`、`table`。
- 备注字段：`speakerNotes`、`speaker_notes`、`presenterNotes`、`presenter_notes`。
- 草稿放行：只有在明确接受草稿输出时，才使用 `allowSparseContent`、`allowMissingChart`、`allowMissingTable`。媒体槽位 layout 必须提供真实图片或图表；不要用 `allowEmptyMediaSlots` 生成空图片占位符。

## 布局选择

先按 `style` 选择布局。同名 layout 在不同 style 下会尽量兼容字段，但视觉结构和容量不一定完全相同；不要假设直接切换 style 后仍适合原来的文本密度。写 JSON 前始终运行对应 style 的 `--capacity-guide`。

生成器会按 layout 校验已填写字段是否实际渲染，也会校验主要内容字段是否缺失。不要向某个 layout 填写容量指南或样例中没有的正文槽位；如果确实需要这些字段，应改 `layout` 或按样例字段名重写。

通用基础页面：

- `cover`：封面页。
- `section`：章节分隔页。
- `closing`：结尾页。

### CMB 样式

`cmb` 适合招商银行、银行、金融经营汇报和正式商务汇报。优先使用这些布局：

- `briefing`、`executiveBrief`、`contentBrief`：高密度文本摘要页，包含顶部摘要、中部分析卡片和底部结论/要点。
- `textWeave`、`contentSynthesis`、`denseText`：非对称文本卡片页，支持 2 到 6 个要点，左侧 1 个主卡片，右侧至少 1 个卡片，适合长文本拆块。
- `article`、`sectionList`：在 CMB 下渲染为 briefing 变体，适合正文型内容，不要当成杂志式长文章排版。
- `textGrid`、`fourCards`：在 CMB 下渲染为 textWeave 变体，卡片数量会自适应；`fourCards` 是兼容旧 JSON 的历史名称，不表示固定 4 个卡片，CMB 下需要至少 2 个文本卡片、最多渲染 6 个文本卡片。
- `agenda`：CMB 专用目录页，支持 1 到 8 条目录；1 到 4 条为单列，5 到 8 条为双列。目录项正文应简短，不要用来承载正文段落。
- `dashboard`、`chart`、`dataSheet`、`bigNumbers`、`kpiTower`：经营数据、图表、表格和指标页。
- `media`、`mediaGrid`、`gallery`、`imageGrid`、`imageHero`、`caseStudy`：图片、证据材料或图表区域页。

CMB 兼容但应谨慎使用：

- `matrix`：标题型矩阵，只保留标题/标签，不适合带长 body 的条目。
- `radial`、`pyramid`、`roadmap`、`timeline`、`pipeline`、`swimlane`：结构关系、路径、流程和分层页面，适合短文本节点。
- `statement`：一个图片槽位加大型陈述，不是纯文本页；没有图片或图表时不要默认选择。

CMB `briefing` 规则：

- 带有 `conclusion`、`takeaway`、`footerSummary` 或 `nextStep` 时，最多支持 4 个中部文本块。
- 没有底部要点区时，最多支持 5 个中部文本块。
- 更多要点请使用 `textWeave` 或拆分为多页。

### Swiss 样式

`swiss` 适合网格化商务汇报、产品方案、战略分析和偏理性的图表页。优先使用这些布局：

- `statement`、`kpiTower`、`bigNumbers`、`dashboard`：核心结论、KPI 和数据看板。
- `textGrid`、`article`、`sectionList`、`fourCards`：中等密度文本页；字段兼容 CMB，但容量通常低于 CMB 高密度文本布局。
- `compare`、`duoCompare`、`splitCompare`：左右对比页。
- `timeline`、`pipeline`、`roadmap`、`swimlane`：时间线、流程、路线和执行矩阵。
- `media`、`mediaGrid`、`imageGrid`、`imageHero`、`caseStudy`：图文和案例页。
- `chart`、`dataSheet`：单图洞察和表格页。

Swiss 中的 `agenda`、`radial`、`pyramid`、`matrix` 更适合短标题或短说明，不适合替代 CMB 的高密度正文页。

### Magazine 样式

`magazine` 适合编辑叙事、视觉化展示、图片材料较多的汇报。优先使用这些布局：

- `section`、`bigQuote`、`quoteImage`、`textImage`：章节、引用、大图和图文叙事页。
- `article`：杂志式多栏正文页，适合中等长度段落；不要塞入过多卡片式要点。
- `compare`、`pipeline`、`imageGrid`、`mediaGrid`、`gallery`：视觉化对比、流程和图片网格。
- `agenda`、`caseStudy`、`pyramid`、`radial`、`roadmap`、`swimlane`：可用作叙事结构页，但保持短文本。

Magazine 不适合大量密集正文；如果每页有很多文字，优先切到 `cmb` 的 `briefing` / `textWeave`，或拆分更多页面。

### 媒体规则

- 如果用户没有提供图片，不要选择包含媒体槽位的布局，包括 `statement`；除非图表数据会填充媒体区域。
- 只有媒体布局可以填写图片字段 `image`、`images`、`gallery`，或 `media` 中的图片路径；纯文本布局带这些字段会报错，因为它们不会被渲染。
- 媒体槽位 layout 必须提供图片或图表；缺少图片、图表或媒体内容会阻断生成。不要用空白图片占位符替代真实素材。
- 如果已知用户图片数量，应选择槽位数量匹配的布局。
- `mediaGrid`、`gallery`、`imageGrid` 的 `captions` / `items` / `sections` 只渲染 `caption` / `title` / `label`；如果需要每张图配正文，改用 `media`、`textImage`、`caseStudy` 或拆成图文页。

### 分点规则

- 需要项目符号分点时，在 JSON 中显式写入 `points`、`bullets` 或 `list` 数组。
- 数组长度大于 1 时，CMB 卡片正文使用 PPT 原生项目符号；只有 1 条时渲染为普通正文，不强制项目符号。项目符号使用 PptxGenJS 的 `bullet: true`，生成 PPTX 原生 `<a:buChar>`，不是手写文本字符。
- 不要要求生成器把一段长正文自动切成多个分点；应在生成 JSON 前由模型按语义拆分。
- CMB 卡片中的 `points`、`bullets`、`list` 会按实际估算行数检查容量：每个分点至少占 1 行，较长分点会按卡片宽度换行；因此不能只看总字数是否低于上限。

## 文本容量

为演示文稿编写完整 JSON 正文前，始终先创建 `deck.plan.json`，再运行 `--capacity-guide --spec deck.plan.json`。不要只按 style 生成全量容量指南，因为全量指南会列出大量本次不使用的 layout，且无法知道每页实际卡片数量。

`deck.plan.json` 示例：

```json
{
  "style": "cmb",
  "theme": "classic",
  "slides": [
    {
      "layout": "textWeave",
      "title": "客户经营能力拆解",
      "sections": [
        { "title": "客户分层" },
        { "title": "渠道协同" },
        { "title": "风险预警" }
      ]
    }
  ]
}
```

容量范围基于 Microsoft YaHei / Times New Roman 以及 36、28、16、14、12 pt 字号层级校准。plan-specific guide 会按实际 layout 和集合项数量输出 `sections[0].body`、`items[2].points[]` 等可填写字段。容量下限按 `max * 60%` 给出，只表示建议内容密度；低于下限不会 warning，只有超过上限才需要缩短或拆分。

生成过程会执行两类检查：

- 来自 `scripts/pptxgen/text-capacity.js` 的 JSON 字段级容量警告。
- 来自 `scripts/pptxgen/engine.js` 的最终文本框容量警告。

如果检查发出文本容量 warning，应缩短相关 JSON 字段、拆分内容、扩大布局空间或选择其他布局。缺图、缺表、缺 chart data、必需集合缺失、字段冲突、标题过多正文不足会作为错误阻断生成；生成器不会静默重写或截断用户文本。
## 演讲者备注

每页幻灯片使用 `speakerNotes` 写入显式备注。接受的别名包括 `speaker_notes`、`presenterNotes` 和 `presenter_notes`。

不要使用 `notes` 作为演讲者备注，因为某些布局会将 `notes` 渲染为可见的幻灯片内容。

## 需要读取的文件

普通演示文稿生成时，只读取：

- `SKILL.md`
- 用户提供的源材料
- 演示文稿 JSON 规格
- 规格中引用的必要素材

除非调试或修改生成器，否则不要读取以下源码文件：

- `scripts/generate-pptx.js`
- `scripts/pptxgen/*.js`
- `scripts/validate-pptx-native.js`
- `scripts/validate-pptx-layout.js`

## 质量检查

交付前：

1. 确认生成的 PPTX 存在，且幻灯片页数符合预期。
2. 运行原生结构验证器和布局验证器。
3. 检查警告是否已处理，或是否为有意接受。
4. 对字体排版变更，检查 PPTX XML 或生成结果，确认只使用批准的字体和字号层级。
5. 对媒体页面，确认路径可解析，且图片没有变形。
6. 对图表页面，确认图表数据存在，且文本使用统一字体系统。
7. 对密集型或面向客户的演示文稿，尽可能进行可视化打开检查。
