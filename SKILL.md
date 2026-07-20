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

1. 必须先选择 style。运行以下命令读取三种模板说明：

   ```bash
   node scripts/generate-pptx.js --style-guide
   ```

   如果用户尚未明确指定 style，使用 `askUserQuestion` 展示 `cmb`、`swiss`、`magazine` 的介绍并让用户选择。只有当运行环境没有该交互工具时，才用普通文本提问；不要自行跳过选择阶段。

2. 选择 style 后，生成该 style 的全布局 JSON 示例 Markdown。该文件包含 47 种 canonical layout 的完整字段示例，不包含字数限制：

   ```bash
   node scripts/generate-pptx.js --layout-examples cmb --out outputs/cmb-layout-examples.md
   ```

   将 `cmb` 替换为用户选择的 `swiss` 或 `magazine`。只需读取这份 Markdown，不要读取模板 JS 源码。

3. 阅读用户内容和布局示例，直接编写完整 `deck.json`。不再创建 `deck.plan.json`，也不再生成 capacity-guide。生成前检查每页 layout、`items` 数量、`images` 数量、`charts` 数量和 `table` 是否匹配。

4. 将完整 JSON 规格转换为用户友好的 Markdown 大纲：

   ```bash
   node scripts/spec-to-md.js --spec path/to/deck.json --out path/to/deck-outline.md
   ```

5. 生成 PPTX：

   ```bash
   node scripts/generate-pptx.js --spec path/to/deck.json --out outputs/deck.pptx
   ```

   最终生成也会再次校验槽位完整性；例如正文卡片类 layout 的集合项不能只有 `body` 而没有 `title`，也不能只有 `title` 而没有正文说明。

6. 如果生成失败，优先修复缺图、缺表、缺 chart data、必需集合缺失、字段冲突、标题过多正文不足等问题；如果最终文本框 warning 提到实际内容可能溢出，应缩短、拆页或更换 layout 后重新生成。
7. 验证原生/可编辑结构和布局安全性：

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

生成三种 style 的布局 JSON 示例：

```bash
npm run layouts:cmb
npm run layouts:swiss
npm run layouts:magazine
```

生成三种 style 共用的全布局排版检查文件：

```bash
node assets/template-cmb-all-layouts.js --style cmb
node assets/template-cmb-all-layouts.js --style swiss
node assets/template-cmb-all-layouts.js --style magazine
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
- 幻灯片级通用字段：`layout`、`kicker`、`title`、`subtitle`、`body`、`items`。
- 所有卡片、步骤、节点、指标、图片说明和侧边说明统一写入 `items`；不要再使用 `sections`、`columns`、`steps`、`nodes`、`layers`、`lanes`、`metrics`、`captions` 等旧顶层集合字段。
- 图片统一写入 `images` 数组，即使只有一张图片；不要使用旧字段 `image`、`gallery` 或 `media`。
- 图表统一写入 `charts` 数组，即使只有一张图表；不要使用旧字段 `chart`。
- 表格使用 `table`。对比页使用 `before.items` 和 `after.items`。
- 备注字段：`speakerNotes`、`speaker_notes`、`presenterNotes`、`presenter_notes`。
- 旧 layout 名称和旧顶层内容字段会直接报错，并提示对应 canonical 名称，不会静默兼容。

## 布局选择

三种 style 使用完全相同的 canonical layout 名称和顶层字段。切换 `style` 只改变视觉实现，不改变 JSON 结构。

名称前缀直接表示页面类型：`deck-*` 为演示结构页，`text-*` 为纯文本页，`image-*` 必须提供图片，`data-*` 使用图表或表格数据。

通用基础页面：

- `deck-cover`、`deck-section`、`deck-closing`

纯文本页面：

- `text-statement`、`text-quote`、`text-article`、`text-briefing`、`text-feature`、`text-list`
- `text-grid`、`text-cards`、`text-weave`、`text-agenda`
- `text-timeline`、`text-pipeline`、`text-roadmap`
- `text-matrix`、`text-radial`、`text-pyramid`、`text-swimlane`
- `text-hero`、`text-case-study`

图片页面：

- 每个 `text-<name>` 都有字段兼容的 `image-<name>`：`statement`、`quote`、`article`、`briefing`、`feature`、`list`、`grid`、`cards`、`weave`、`agenda`、`timeline`、`pipeline`、`roadmap`、`matrix`、`radial`、`pyramid`、`swimlane`、`hero`、`case-study`

数据页面：

- `data-numbers`、`data-kpis`、`data-compare`
- `data-chart`、`data-dashboard`、`data-table`

### Style 使用建议

- `cmb`：银行、金融经营和正式商务汇报；高密度正文优先 `text-briefing` / `text-weave`。
- `swiss`：结构化商务、产品、战略和数据汇报；优先 `text-grid` / `data-dashboard`。
- `magazine`：编辑叙事、观点和图片材料；优先 `text-article` / `image-quote` / `image-grid`。

### 媒体规则

- `image-*` 必须提供 `images`；`text-*` 带 `images` 或 `image-*` 缺少图片时，错误会直接给出字段兼容的对应 layout 名称。
- `image-statement`、`image-quote` 支持 1 张图片；其他 `image-*` 对应布局支持 1 到 6 张图片。
- `text-<name>` 与 `image-<name>` 除 `images` 外字段完全一致，可通过修改 `layout` 并添加或删除 `images` 直接切换。
- 图片 layout 不接受 `charts`；图表请使用 `data-chart` 或 `data-dashboard`。

### 分点规则

- 需要项目符号分点时，在 JSON 中显式写入 `points`、`bullets` 或 `list` 数组。
- 数组长度大于 1 时，CMB 卡片正文使用 PPT 原生项目符号；只有 1 条时渲染为普通正文，不强制项目符号。项目符号使用 PptxGenJS 的 `bullet: true`，生成 PPTX 原生 `<a:buChar>`，不是手写文本字符。
- 不要要求生成器把一段长正文自动切成多个分点；应在生成 JSON 前由模型按语义拆分。
- CMB 卡片中的 `points`、`bullets`、`list` 会按实际估算行数检查容量：每个分点至少占 1 行，较长分点会按卡片宽度换行；因此不能只看总字数是否低于上限。

## 文本溢出检查

本流程不再生成字数容量指南。最终生成时，`engine.js` 会根据实际文本框宽高、字号、换行和项目符号行数估算文本是否可能溢出。

如果出现溢出 warning，应缩短相关字段、减少同页条目、拆分页面或手动更换 layout。生成器不会静默截断、改写正文或自动修改 layout。
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
