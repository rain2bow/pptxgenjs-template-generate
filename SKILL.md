---
name: pptxgenjs-template-generate
description: 使用 pptxgenjs 原生生成可编辑 .pptx 演示文稿，支持 magazine、swiss、cmb 三套风格，可从结构化 JSON 插入文本、图片、图标、表格和原生图表，并进行槽位、布局和可编辑性校验。当用户需要直接生成 PPTX、制作可编辑 PowerPoint、从长文/大纲/素材生成演示文稿、使用招商银行/CMB 商务模板或切换 PPT 风格时使用。
---

# PPTXGenJS PPT Skill

此技能使用 `pptxgenjs` 直接生成可编辑 `.pptx`。所有页面必须用 PowerPoint 原生文本框、形状、图片、线条、表格和图表结构搭建，便于后续在 PowerPoint/WPS/Keynote 中继续编辑。

**硬约束**：默认输出必须保留可编辑性，不得把整页渲染成单张位图或用不可编辑背景图伪装成 PPTX。只有用户明确放弃可编辑性并单独要求图片版演示文稿时，才可以另走转换路线。

**视觉一致性的定义**：同一 deck 内必须保持风格、主题 token、字体分工、字号比例、网格/留白、图片槽位、页眉页脚、卡片和线条规则一致。不要承诺像素级完全一致，因为 PowerPoint/WPS/Keynote 的字体渲染和图形引擎存在差异。

## 视觉系统

### 风格 A · 电子杂志 × 电子墨水

- 衬线大标题：中文优先用 `Noto Serif SC` / `SimSun` 兜底，英文用 `Playfair Display` / `Georgia`。
- 正文使用无衬线：`Noto Sans SC` / `Microsoft YaHei UI`。
- 元数据使用等宽：`IBM Plex Mono` / `Consolas`。
- 视觉关键词：纸感底色、墨色文字、杂志页眉页脚、大字号标题、图片网格、数据大字报、克制线条。
- 主题色只能从生成器内置的 `magazine` 预设中选：`ink`、`indigo`、`forest`、`kraft`、`dune` 或 `cmb`；不要任意混搭 hex。

### 风格 B · 瑞士国际主义

- 全程无衬线：`Inter` / `Helvetica Neue` / `Noto Sans SC` / `Microsoft YaHei UI`。
- 只用一个高饱和 accent：`ikb`、`lemon`、`green`、`orange` 或 `cmb`；不要在同页混用多个 accent。
- 视觉关键词：直角、纯色、无阴影、无渐变、发丝线、12/16 列网格、大标题 200/300 轻字重、小字更粗。
- 页面必须保持左上内容轴，不要把正文页标题居中。

## 核心工作流

1. 先确认风格 A 或 B。用户未指定时：人文、观点、故事默认 A；科技、数据、方法论、产品默认 B。
2. 确认受众、场景、页数、素材、图片需求、主题色和硬约束。信息足够时可以合理假设并在交付说明里写明。
3. 把内容整理成结构化 JSON 规格，优先使用 `scripts/generate-pptx.js` 已支持的版式。
4. 调用脚本生成 `.pptx`：

```bash
node scripts/generate-pptx.js --spec path/to/deck.json --out path/to/deck.pptx
```

生成 PPTX 前，先把 JSON spec 转成用户友好的 Markdown 大纲给用户审阅。该 Markdown 应展示总页数、每页页面类型、标题、正文、要点、图表、表格、图片说明和演讲者备注，不要求用户阅读 JSON 字段名：

```bash
node scripts/spec-to-md.js --spec path/to/deck.json --out path/to/deck-outline.md
```


如果没有 spec，可以生成内置样例验证环境：

```bash
node scripts/generate-pptx.js --sample --out outputs/sample-deck.pptx
node scripts/generate-pptx.js --sample --sample-style magazine --out outputs/sample-magazine.pptx
node scripts/generate-pptx.js --sample --sample-style cmb --out outputs/sample-cmb.pptx
```

5. 生成后检查文件存在、可打开、页数正确。复杂项目应打开 PPTX 做视觉检查，重点看标题是否溢出、图片是否裁切错误、底部页脚是否挡内容。

## 脚本操作指南

Default: do not read `scripts/generate-pptx.js`, `scripts/pptxgen/*.js`, `scripts/validate-pptx-native.js`, or `scripts/validate-pptx-layout.js` source for ordinary deck generation. Read them only when modifying scripts, diagnosing a generator bug, or adding/changing style/layout behavior.

### 安装依赖

首次使用、迁移到新机器或部署到 Linux 服务器时，先进入技能目录安装 Node 依赖：

```bash
cd pptxgenjs-template-generate
npm install
```

依赖来自本目录 `package.json`：

- `pptxgenjs`: 生成可编辑 `.pptx` 的核心库，负责文本框、形状、图片、表格和原生图表输出。
- `lucide`: 分点、卡片、指标等位置使用的 Lucide SVG 图标库；缺失时只有少量内置 fallback 图标可用。

环境要求：

- Node.js 建议使用 18 或更高版本。
- Linux/macOS/Windows 都使用同一套命令；不要在用户项目目录单独安装依赖，除非用户把整个 skill 复制到了该项目目录。
- 正式生成前应保证 `node_modules/` 已存在；若缺失，先重新运行 `npm install`。

### 生成 PPTX

有 JSON spec 时直接运行：

```bash
node scripts/generate-pptx.js --spec path/to/deck.json --out path/to/deck.pptx
```

没有 spec、只想验证环境时运行样例：

```bash
node scripts/generate-pptx.js --sample --sample-style swiss --out outputs/sample-swiss.pptx
node scripts/generate-pptx.js --sample --sample-style magazine --out outputs/sample-magazine.pptx
node scripts/generate-pptx.js --sample --sample-style cmb --out outputs/sample-cmb.pptx
```

`--sample-style` 只能取 `swiss`、`magazine` 或 `cmb`。输出路径建议放到项目 `outputs/` 下，避免覆盖技能自带模板或用户源文件。

### 演讲者备注

每页可以用 `speakerNotes` 写入 PowerPoint speaker notes；兼容别名 `speaker_notes`、`presenterNotes`、`presenter_notes`。字段可以是字符串、字符串数组，或 `{ opening, points, closing }` 这类对象。不要用 `notes` 表示演讲者备注，因为 `notes` 已在 `dataSheet`、`chart` 等 layout 中作为页面内侧边说明渲染。

如果用户需要每页都有基础讲稿，但 JSON 没有逐页写备注，可以在 deck 根部或单页设置 `generateSpeakerNotes: true`。生成器会根据标题、正文、要点、图表、表格和图片字段生成基础备注，并写入 PPTX 的备注区；同时 `scripts/spec-to-md.js` 输出的 Markdown 也会显示这些备注，便于人工检查。

示例：

```json
{
  "generateSpeakerNotes": true,
  "slides": [
    {
      "layout": "textGrid",
      "title": "客户经营重点",
      "sections": [{"title": "增长", "body": "提升高价值客户渗透"}],
      "speakerNotes": ["先说明本页结论。", "再逐项解释三个经营抓手。"]
    }
  ]
}
```


### 内置模板入口

技能自带三套可直接运行的 JS 模板，用于快速生成完整示例 deck，运行位置为技能目录：

```bash
node assets/template-magazine.js
node assets/template-swiss.js
node assets/template-cmb.js
node assets/template-cmb-all-layouts.js
```

- `assets/template-magazine.js`：电子杂志 / 电子墨水风格。
- `assets/template-swiss.js`：瑞士国际主义风格。
- `assets/template-cmb.js`：独立招商银行品牌风格模板，使用 `style: "cmb"`，内置页眉白底 PNG logo 和页眉外透明 SVG 纯图 logo，适合银行、金融、经营汇报场景。用户要求招商银行、CMB、银行品牌配色或红灰白商务汇报时，优先参考此模板或 `--sample-style cmb`。`assets/template-cmb-all-layouts.js` 会生成覆盖 CMB 当前全部支持 layout 的 31 页检查文件，用于人工检查排版。

### 校验 PPTX

生成后按顺序运行：

```bash
node scripts/validate-pptx-native.js path/to/deck.pptx
node scripts/validate-pptx-layout.js path/to/deck.pptx
```

`validate-pptx-native.js` 用于确认包含可编辑的原生 PPTX 结构；`validate-pptx-layout.js` 用于检查明显的文本覆盖、元素冲突和底部安全区问题。校验通过后仍应打开 PPTX 做人工视觉核对。

### 需要读取哪些文件

Read `SKILL.md`, user materials, the deck JSON spec, and only the needed `assets/template-*.js` example. For ordinary generation do not read `scripts/*.js` or `scripts/pptxgen/*.js`; use the documented commands instead.

仅在这些情况下读取脚本源码：
- 要新增或修改版式、图表、图标、布局算法。
- 生成或校验命令报错，且报错信息、JSON spec、SKILL.md 无法定位问题。
- 用户明确询问脚本内部实现。

普通内容生成、改文案、换主题、换版式、插图表、插 icon 时，按本文档命令执行即可，不需要读取生成器源码。主题、layout 和渲染规则以 `scripts/pptxgen/config.js`、`scripts/pptxgen/engine.js` 与 `scripts/pptxgen/ARCHITECTURE.md` 为准。
## JSON Spec

最小结构：

```json
{
  "title": "一人公司：被 AI 折叠的组织",
  "subtitle": "一个关于 AI、组织和个体的新叙事",
  "author": "Presentation Team",
  "style": "magazine",
  "theme": "ink",
  "slides": [
    {
      "layout": "cover",
      "kicker": "A Talk · 2026",
      "title": "一人公司",
      "subtitle": "被 AI 折叠的组织"
    },
    {
      "layout": "bigNumbers",
      "kicker": "Proof",
      "title": "过去 64 天",
      "items": [
        {"label": "Duration", "value": "64", "unit": "天", "note": "从 0 到现在"},
        {"label": "Lines", "value": "110K+", "note": "代码规模"}
      ]
    }
  ]
}
```

字段约定：

- `style`: `magazine`、`swiss` 或 `cmb`。
- `theme`: 风格 A 可用 `ink`、`indigo`、`forest`、`kraft`、`dune`、`cmb`；风格 B 可用 `ikb`、`lemon`、`green`、`orange`、`cmb`；招商银行独立风格可用 `classic`、`pearl`、`graphite`。
- `slides[].layout`: 使用下方支持的版式名。
- 图片/logo 路径解析顺序：绝对路径、相对 spec 文件、相对当前工作目录、相对技能 `assets/`、相对技能根目录。内置素材可直接写 `logos/cmb-logo-lockup.png`、`logos/cmb-logo-mark.svg` 或对应的 `assets/logos/...` 路径，不需要复制到用户项目目录；脚本会按实际文件插入。所有 logo 必须保持素材原始比例，不能为了填满指定 `w/h` 而横向或纵向压扁；生成器会在给定框内等比居中放置。
- 文本太长时先改写或拆页，不要压到很小字号。生成器会按最终文本框 `w/h/fontSize/margin` 估算可容纳字数，超出时只打印 `Warning: text may overflow box`，不会截断或改写原文；看到 warning 后应降低该段字数、放大卡片或拆页。
JSON 引号与编码规则：所有 spec 文件统一使用 UTF-8。生成 JSON 时优先由程序对象调用 `JSON.stringify(data, null, 2)` 写出，不要手写拼接字符串。普通中文内容里的引号优先使用 `「」` 或中文弯引号；如果必须在 JSON 字符串中使用英文直引号 `"`，必须写成 `\"`。生成器读取 `--spec` 时会自动尝试修复常见问题：Markdown fenced code block、UTF-8 BOM、`//`/`/* */` 注释、尾逗号、用作 JSON 结构分隔符的中英文弯引号。若自动修复成功，会打印 warning；需要落盘为严格 JSON 时，继续使用 `--write-normalized-spec path/to/normalized.json`。无法可靠判断的未转义英文直引号仍会报错，此时必须人工改成 `\"` 或改用 `「」`。

## 支持版式

风格 A 推荐版式：

- `cover`: 开场封面，hero dark 杂志感。
- `section`: 章节幕封。
- `bigNumbers`: 3-6 个数据大字报。
- `quoteImage`: 左文右图，适合故事、身份反差、案例。
- `imageGrid`: 1-6 图证据网格；图片数决定槽位数，缺图时显示 `IMAGE SLOT` 占位符。
- `media`: 正文/洞察 + 统一媒体区；有用户图片优先放图，没有图片且显式提供 `chart`/`charts` 时放可编辑图表，否则显示 `IMAGE SLOT` 占位符。
- `mediaGrid` / `gallery`: 1-6 个图片/图表媒体位，自适应一张或多张图。
- `pipeline`: 3-6 步流程。
- `bigQuote`: 大引用/金句。
- `compare`: Before / After 并列对比。
- `textImage`: 大段正文 + 辅助图。
- `article`: 2-3 列高密度文字页。
- `dataSheet`: 表格 + 右侧解释页。
- `chart`: 单主图表 + 右侧洞察页。
- `dashboard`: KPI 数字条 + 双图表页。
- `closing`: 收束页。

风格 B 推荐版式：

- `cover`: IKB/accent 满屏封面。
- `statement`: 极简大字论点。
- `kpiTower`: 3-4 个 KPI 高度对比。
- `duoCompare`: Before / After 双栏对照。
- `timeline`: 4-6 步横向时间线。
- `matrix`: 8-12 项矩阵 + 总数据。
- `fourCards`: 1-8 项自适应卡片网格，5-6 项自动两行。
- `imageHero`: 顶部 21:9 主图 + 下方说明/KPI；无用户图片且无显式 `chart` 时显示 `IMAGE SLOT` 占位符。
- `media`: 统一媒体区 + 侧边说明；有用户图片优先放图，没有图片且显式提供 `chart`/`charts` 时放可编辑图表，否则显示 `IMAGE SLOT` 占位符。
- `mediaGrid` / `gallery`: 1-6 个图片/图表媒体位，自适应一张或多张图。
- `sectionList`: text-only vertical section list for 3-7 `sections` / `items` / `columns`; preserves `title` + `body` and is the preferred visual alternative when repeated `article` pages need variety without image slots.
- `textGrid`: 6-9 项三列信息网格。
- `dataSheet`: Swiss 表格 + 侧边备注页。
- `chart`: 大图表 + 右侧指标洞察页。
- `dashboard`: KPI strip + 双图表仪表盘页。
- `closing`: 瑞士风收束页。


通用丰富版式（两套模板同名兼容）：

- `agenda`: 目录/议程/章节导航页，适合 3-8 个主题入口。
- `caseStudy`: 案例页，图片/证据 + 故事正文 + 关键指标组合；无图且无显式 `chart` 时媒体区显示 `IMAGE SLOT` 占位符。
- `pyramid`: 分层结构/能力栈/战略金字塔。
- `radial`: 中心概念 + 周边节点关系图。
- `roadmap`: 路线图/阶段计划/里程碑页。
- `swimlane`: 泳道矩阵，适合角色 × 阶段、团队 × 任务、模块 × 时间。
跨模板兼容：`magazine`、`swiss` 与 `cmb` 都必须接受上述两套 layout 名称。切换模板时优先只改顶层 `style` / `theme`，不要批量改每页 `layout`；生成器会把另一套模板的 layout 映射为当前风格中语义最接近的页面：`bigNumbers` ↔ `kpiTower`、`compare` ↔ `duoCompare`、`pipeline` ↔ `timeline`、`article` ↔ `textGrid`、`quoteImage`/`textImage` ↔ `imageHero`、`bigQuote`/`section` ↔ `statement`/`cover`。如果某页切换后信息密度明显不合适，再人工换成同风格推荐版式。
媒体区规则：两个模板都使用同名 `media` / `mediaGrid` / `gallery` layout 留出统一放图区域。用户提供 `image` / `images` / `gallery` 时优先插入用户图片；`mediaGrid` / `gallery` / `imageGrid` 未显式设置 `mediaCount` 时，槽位数自动等于图片数、显式图表数或 caption 数，不再用默认 4 格。若显式设置 `mediaCount`，必须与用户图片数一致；例如 3 张图就用 3 个槽位，不要生成 4 个槽位。没有用户图片且显式提供 `chart` / `charts` 时才用 PowerPoint 原生图表填充；没有图片和显式图表时显示 `IMAGE SLOT` 占位符，表示这里可以放图。

Media selection rule: without user-provided images, do not select image/media-slot layouts unless the slide has explicit chart data that will fill the media region. If there are no images and no charts, use text-only layouts and do not create empty image placeholders unless the user explicitly asks for placeholders.
槽位校验：生成器会在生成前检查每页图片槽位、文本槽位和字段格式。超过布局最大数量的 `items` / `sections` / `steps` / `charts` 会直接报错，避免内容被静默截断；同一组同义字段（如 `sections`、`items`、`columns`）不要同时填写，否则只有第一个字段会被使用并打印警告；分点对象必须至少包含 `text` / `title` / `label` / `body` / `desc` / `note` / `summary` / `value` 之一，避免格式不匹配导致内容不显示。
内容完整性：每个分点、卡片、栏目不能只写标题，至少补 `body` / `desc` / `note` / `summary` 之一。生成器会对大多数分点只有标题的页面打印警告；生成大纲和 spec 时必须把“标题 + 一句解释/证据/结论”作为最小单元。
自适应分栏：多分点页面必须根据实际条目数排版，不要为了模板默认 6/9/12 个位置而补空内容。`article`、`textGrid`、`matrix`、`fourCards`、`bigNumbers` 会自动选择列数：4 条默认 2×2，`fourCards` 的 5-6 条默认 3×2，7-8 条默认 4×2；`textGrid`/`matrix` 的 7-9 条默认 3 列多行；确需固定列数时才在页面 spec 中显式设置 `columnsCount`。
页面多样性：规划 deck 时主动避免连续 3 页以上使用同一 layout 或同一视觉节奏。连续说明页应在 `statement`、`textGrid`、`article`、`fourCards`、`matrix`、`chart`、`media`、`mediaGrid`、`imageHero`、`compare`、`timeline`、`agenda`、`caseStudy`、`pyramid`、`radial`、`roadmap`、`swimlane` 之间轮换；如果内容语义相同但页面相邻，优先换成等价版式、调整条目数量、加入图表/表格/图片页或章节页。只有用户明确要求统一模板页时，才允许长段连续重复同一 layout。生成器默认只对连续重复 layout 输出替换建议，不会擅自修改 `slides[].layout`，这样输入 JSON 与生成 PPTX 保持一致。若确实要自动改 layout，必须同时使用 
Layout gating rule: if a slide has no user-provided image fields (`image`, `images`, `gallery`, or image media items), do not choose image/media-slot layouts (`media`, `mediaGrid`, `gallery`, `imageGrid`, `imageHero`, `quoteImage`, `textImage`, `caseStudy`) just for visual variety. Use text/structure layouts instead. If the slide has explicit `chart` / `charts` or `table` data, use `chart`, `dashboard`, or `dataSheet` as the preferred variation. The generator layout suggestions follow the same filtering rule.
`--diversify-layouts --write-normalized-spec path/to/normalized.json`，并以后续 normalized JSON 作为真实源文件；否则可能出现“PPT 已换 layout、原 JSON 未变、内容字段不匹配”的问题。

## 生成准则

- 保留可编辑性：必须使用 PPT 原生文本、形状、线条、图片，不把整页渲成一张大图。
- 不写 PowerPoint 动画；需要动态效果时用版式节奏、层级和图形关系表达。
- 背景只能使用 PPT 原生形状、线条、图片或纹理元素：风格 A 可用低透明等高线/弧线/纸感纹理；风格 B 可用发丝网格、点阵或 ASCII 感字符矩阵。
- 必须避免元素冲突：先画背景/纹理，再画色块/图片，最后画文字和页眉页脚。任何装饰、色块或图片都不能在 z-order 上晚于它覆盖区域内的文字。
- 所有布局都使用固定安全区：左右至少 `0.6in`，顶部至少 `0.35in`，底部正文最低点不超过 `6.85in`，页脚/页码单独放在底部安全区。
- 色块内文字必须在同一函数里紧跟色块之后绘制，并留出内边距：大色块左右 ≥ `0.25in`，卡片左右 ≥ `0.18in`，上下 ≥ `0.12in`。
- 不允许通过负坐标、负 margin、任意绝对偏移来“救布局”。内容过多时先删减、拆页或换版式。
- 图片是内容证据，不是背景装饰。信息图和产品界面图用标准比例：21:9、16:10、16:9、4:3、3:2、1:1。
- 风格 A 可以使用轻微纸感背景和大衬线标题；风格 B 禁止渐变、阴影、圆角、多个 accent。
招商银行风格：优先使用独立 `style: "cmb"`，主题可选 `classic`、`pearl`、`graphite`，并设置 `logoHeader` / `logoMark`。推荐复用 `assets/template-cmb.js` 或运行 `--sample-style cmb`。不要把它仅当作 `swiss` 的一个主题变体。招商银行模板每页页眉必须整条保持白色，并使用白底 PNG 完整 logo：`logos/cmb-logo-lockup.png`；页眉以外的封面、收尾页或装饰性品牌标识必须使用透明背景 SVG 纯图 logo：`logos/cmb-logo-mark.svg`。封面保留清晰 SVG 纯图 logo，且与内容页右上角 SVG logo 使用相同大小和位置；所有内容页右上角的原圆形装饰必须替换为不透明度 20% 的 SVG 纯图 logo 水印。两者都会自动解析到技能内置 `assets/logos/`，不要提示用户手动复制素材。所有 logo 都必须等比显示，不要只在红色页面上放一个局部白底 logo，也不要用 SVG 纯图 logo 作为页眉品牌标识。
所有模板都必须保持页首信息位置稳定：同一 deck 内每页 kicker/顶部小标题默认使用统一 headY 基线，避免翻页时顶部文字上下跳动。纯色强调页不能是无层次纯色块，应叠加低透明浅色层、暗色层、线性/弧线纹理或其它 PPT 原生形状做柔化。
- 不要使用 emoji 作为页面视觉元素；图标需求优先用简单线条/形状，或由用户提供图标资产。
- 投影可读性优先：页眉、页脚、标签、旁注、图表坐标轴、表格正文等小字必须保持可阅读，生成器会把普通文本、图表文字和表格文字的最小字号抬到约 12pt，并按文本框容量提示溢出风险。内容放不下时拆页、换版式或删减，不要继续压小字；极特殊文本框可显式设置 `noTextLimitWarning` / `noTextLimit` / `allowOverflowText` 跳过提示。
- 中文大标题要主动断行。风格 B 的中文超大标题通常比英文小一档。
- 底部页码、页脚和注释信息不要遮挡主体，至少保留 0.35 英寸安全区。

## 资源导览

- `scripts/generate-pptx.js`: thin CLI and compatibility export entry; keeps existing template `require` calls working.
- `scripts/pptxgen/cli.js`: command-line orchestration.
- `scripts/pptxgen/config.js`: style/theme registry, default themes, fonts, slide constants, icon aliases, and `READABILITY.minFontSize`; add new style/theme configuration here first.
- `scripts/pptxgen/engine.js`: PPTX rendering runtime, layout renderers, media/chart/table insertion, slot validation, and readability logic; add or change layouts here.
- `scripts/pptxgen/spec-io.js`: JSON spec loading, loose parsing, quote/comment/trailing-comma repair, and normalized spec output.
- `scripts/pptxgen/spec-md.js`: JSON spec to user-facing Markdown outline renderer.
- `scripts/pptxgen/speaker-notes.js`: speaker notes normalization and auto-generation helpers used by both PPTX output and Markdown outline.
- `scripts/spec-to-md.js`: CLI for generating the Markdown outline after JSON spec creation and before PPTX generation.
- `scripts/check-media-slot-warnings.js`: regression check for media-slot layout warnings when images/charts are missing.
- `scripts/pptxgen/samples.js`: built-in `--sample` specs; add a sample when adding a new style.
- `scripts/pptxgen/ARCHITECTURE.md`: read this first when modifying generator internals; it maps modules, engine sections, and style/layout extension points.
- `assets/template-magazine.js`: 电子杂志 / 电子墨水完整示例模板，可直接运行生成 `assets/outputs/deck-magazine.pptx`。
- `assets/template-swiss.js`: 瑞士国际主义完整示例模板，可直接运行生成 `assets/outputs/deck-swiss.pptx`。
- `assets/template-cmb.js`: 招商银行独立品牌风格完整示例模板，可直接运行生成 `assets/outputs/deck-cmb.pptx`；该模板使用 `style: "cmb"`、`theme: "classic"`、`logoHeader: "logos/cmb-logo-lockup.png"`、`logoMark: "logos/cmb-logo-mark.svg"`，logo 会自动从技能内置 `assets/logos/` 解析。`assets/template-cmb-all-layouts.js` 可直接生成 `assets/outputs/deck-cmb-all-layouts.pptx`，包含 CMB 当前全部 31 个支持 layout 的排版检查页。
- 主题色只能从生成器内置的 `magazine` 预设中选：`ink`、`indigo`、`forest`、`kraft`、`dune` 或 `cmb`；不要任意混搭 hex。
- 只用一个高饱和 accent：`ikb`、`lemon`、`green`、`orange` 或 `cmb`；不要在同页混用多个 accent。

## 质量检查

交付前至少做这些检查：

1. `node scripts/generate-pptx.js --sample --out outputs/sample-deck.pptx` 能成功运行。
2. 目标 PPTX 文件存在且大小大于 0。
3. 页数与 spec 的 `slides.length` 一致。
4. 风格 A 没有丢失衬线大标题、杂志页眉页脚、纸感/电子墨水纹理和标准图片槽位；风格 B 没有出现圆角、阴影、渐变、多 accent，标题保持左上内容轴。
5. 图片路径缺失时脚本会用占位框提示，正式交付前应替换为真实图片或删掉图片页。
6. 运行 `node scripts/validate-pptx-native.js path/to/deck.pptx`，确认包含可编辑的原生 PPTX 结构。
7. 运行 `node scripts/validate-pptx-layout.js path/to/deck.pptx`，检查文字是否被后绘制的色块/图片覆盖、文本框是否冲突、是否进入底部安全区。
8. 修改媒体槽位或 layout 选择逻辑后，运行 `npm run check:media-slots`，确认所有带图片/媒体槽位的 layout 缺少图片/图表时都会 warning，不能静默生成。
9. 打开 PPTX 做视觉核对：每页必须符合当前 `style` 的版式语法、主题 token 和安全区规则。若某页看起来像普通 PowerPoint 模板，或出现文字被色块覆盖、元素互相压住，必须回退重做。

## 新增高密度版式与数据组件

当原有页面文字密度不足时，优先使用这些版式，而不是把正文硬塞进展示型页面：

风格 A `magazine` 新增：
- `article`: 2-3 列高密度文字页，字段用 `sections` / `items`，每项支持 `title`、`body`。
- `dataSheet`: 表格 + 右侧解释页，字段用 `table.headers`、`table.rows`、`notes`。
- `chart`: 单主图表 + 右侧洞察页，字段用 `chart`、`insights`。
- `dashboard`: KPI 数字条 + 双图表页，字段用 `metrics`、`charts`。

风格 B `swiss` 新增：
- `textGrid`: 6-9 项三列信息网格，字段用 `sections`，可用 `highlightIndex` 强调一项。
- `dataSheet`: Swiss 表格 + 侧边备注页。
- `chart`: 大图表 + 右侧指标洞察页。
- `dashboard`: KPI strip + 双图表仪表盘页。

任意页面都可以额外插入可编辑数据组件：

```json
{
  "layout": "statement",
  "title": "页面标题",
  "blocks": [
    {
      "type": "chart",
      "chartType": "line",
      "x": 6.8,
      "y": 3.2,
      "w": 4.8,
      "h": 2.4,
      "labels": ["Jan", "Feb", "Mar"],
      "series": [{"name": "Actual", "values": [12, 18, 25]}]
    },
    {
      "type": "table",
      "x": 0.8,
      "y": 4.5,
      "w": 5.6,
      "h": 1.8,
      "headers": ["Item", "Value"],
      "rows": [["A", "42%"], ["B", "31%"]]
    }
  ]
}
```

也可以使用简写：
- `charts`: 等价于多个 `{ "type": "chart" }` block，适合插入到普通页面。
- `tables`: 等价于多个 `{ "type": "table" }` block。

图表类型支持：`bar`、`column`、`line`、`pie`、`doughnut`、`area`、`radar`、`scatter`。图表使用 pptxgenjs 原生 `addChart`，表格使用原生 `addTable`，输出后仍可在 PowerPoint/WPS 中编辑。所有作为 `blocks`、`charts`、`tables` 后置插入的组件必须设置合理的 `x/y/w/h`；缺少显式坐标的后置图表/表格会被跳过并打印警告，避免与正文重叠。若希望自动留区放图表，使用 `media`、`mediaGrid`、`chart` 或 `dashboard` layout。
## 分点图标库

分点列表里的 `items` 支持对象写法，在每个分点前放置可编辑缩放的 SVG 图标。图标来源优先使用 `lucide` 图标库，因此可以直接使用 Lucide 的几百个图标名；脚本会从 `lucide` 包读取对应图标节点，按当前主题色生成 SVG 后插入 PPTX。若旧环境仍安装了 `lucide-static`，脚本也会兼容回退读取 `node_modules/lucide-static/icons/<icon>.svg`。

```json
{
  "layout": "compare",
  "after": {
    "title": "新模式",
    "items": [
      {"icon": "check-circle", "text": "结果已验证"},
      {"icon": "arrow-right-circle", "text": "流程可继续推进"},
      {"icon": "sparkles", "text": "关键亮点"}
    ]
  }
}
```

图标名规则：
- 推荐直接使用 Lucide 文件名，如 `activity`、`database`、`chart-line`、`chart-column`、`chart-pie`、`circle-alert`、`shield-alert`、`badge-dollar-sign`、`users`、`workflow`、`scan-search`。
- 也兼容少量语义别名：`check`、`alert`、`warning`、`info`、`arrow`、`trend`、`bar`、`line`、`pie`、`risk`、`cost`、`people`、`team`、`process`。
- 基础原生符号仍可用：`dot`、`square`、`diamond`、`plus`、`minus`、`cross`、`number`。

字段说明：
- `icon`: Lucide 图标名；未填时 bullet 默认 `dot`，多栏/卡片页只有设置 `icon` 或该布局有默认图标时才显示。
- `text`: 分点文字；也兼容 `title` 或 `label`。
- `iconColor`: 可选，覆盖图标颜色。
- `bold` / `color` / `transparency`: 可选，控制该条文字样式。

插入范围：`compare` 的 bullet、`article` 多栏、`textGrid` 多栏、`fourCards` 卡片、`dataSheet` 侧边 notes、`chart` insights、`dashboard` metrics 都会尽量在文本前插入图标。`article` 和 `textGrid` 这类多分栏页面即使没有逐项填写 `icon`，也会按分栏顺序自动补齐默认 Lucide 图标；手动填写 `icon` 时优先使用用户指定图标。不要使用 emoji 作为分点图标；需要丰富图标时直接查 Lucide 名称并写入 `icon` 字段。内容过多时仍然应拆页，不要依赖图标列表承载长段落。






