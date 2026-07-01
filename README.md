# guizang-pptxgenjs-ppt-skill

当前分支：`cmb-independent-style`

本目录是一个基于 `pptxgenjs` 的可编辑 PowerPoint 生成技能。它不通过 HTML 截图、整页图片或 PDF 转换交付，而是直接生成 PowerPoint 原生文本框、形状、线条、图片、表格和图表，方便后续在 PowerPoint/WPS 中继续编辑。

## 当前分支重点

`cmb-independent-style` 分支在原有 `magazine` 和 `swiss` 风格基础上，新增并持续完善了独立的招商银行风格：

- 新增 `style: "cmb"`，不是把 CMB 当作 Swiss 的一个 theme。
- 内置招商银行页眉白底 PNG logo 和页眉外透明 SVG 纯图 logo。
- CMB 页面保持白色页眉，内容页右上角使用低透明度 SVG logo 水印。
- CMB 支持当前全部通用 layout，并提供全布局检查文件生成入口。
- 修复多处 layout 切换、文本溢出、连线错误、媒体槽位、图标和 logo 解析问题。

## 支持的风格

- `magazine`：电子杂志 / 电子墨水风格，适合观点、叙事、报告型页面。
- `swiss`：瑞士国际主义风格，适合科技、数据、方法论和产品说明。
- `cmb`：招商银行品牌风格，适合银行、金融、经营汇报和商务汇报。

默认主题：

- `magazine`: `ink`
- `swiss`: `ikb`
- `cmb`: `classic`

## 安装依赖

在技能目录下执行：

```bash
npm install
```

主要依赖：

- `pptxgenjs`: 生成可编辑 `.pptx`
- `lucide`: 为分点、卡片、指标等位置提供图标

建议使用 Node.js 18 或更高版本。

## 生成 PPTX

使用 JSON spec 生成：

```bash
node scripts/generate-pptx.js --spec path/to/deck.json --out outputs/deck.pptx
```

生成内置示例：

```bash
npm run sample:magazine
npm run sample:swiss
npm run sample:cmb
```

生成 CMB 全布局检查文件：

```bash
npm run sample:cmb:layouts
```

输出文件：

```text
assets/outputs/deck-cmb-all-layouts.pptx
```

## 校验 PPTX

生成后建议按顺序执行：

```bash
node scripts/validate-pptx-native.js path/to/deck.pptx
node scripts/validate-pptx-layout.js path/to/deck.pptx
```

含义：

- `validate-pptx-native.js`: 检查是否为 PowerPoint 原生结构，避免整页截图伪装成 PPTX。
- `validate-pptx-layout.js`: 检查明显的布局风险，如重叠、底部安全区、文本框冲突等。

## Spec 基本结构

最小示例：

```json
{
  "title": "招商银行经营汇报",
  "subtitle": "Business Review",
  "style": "cmb",
  "theme": "classic",
  "logoHeader": "logos/cmb-logo-lockup.png",
  "logoMark": "logos/cmb-logo-mark.svg",
  "slides": [
    {
      "layout": "cover",
      "kicker": "CHINA MERCHANTS BANK / 2026",
      "title": "业务增长与数字化经营汇报",
      "subtitle": "围绕客户经营、风险控制与效率提升"
    },
    {
      "layout": "media",
      "kicker": "Customer Operation",
      "title": "客户经营从单点触达转向分层运营",
      "body": "通过客群分层、权益匹配与渠道协同，提升客户转化与长期价值。",
      "chart": {
        "chartType": "line",
        "title": "客户活跃趋势",
        "labels": ["Q1", "Q2", "Q3", "Q4"],
        "values": [42, 51, 63, 78]
      },
      "items": [
        { "icon": "users", "title": "客户分层", "body": "按资产、行为与生命周期拆分运营策略。" },
        { "icon": "workflow", "title": "渠道协同", "body": "联动 App、网点、远程服务与客户经理。" }
      ]
    }
  ]
}
```

## Layout 能力摘要

当前 CMB 模板可通过 `assets/template-cmb-all-layouts.js` 生成完整检查页，覆盖 31 个 layout。常用 layout 包括：

- 封面与章节：`cover`、`section`、`closing`
- 结论页：`statement`
- 数据页：`kpiTower`、`bigNumbers`、`dashboard`、`chart`、`dataSheet`
- 图文页：`media`、`mediaGrid`、`gallery`、`imageGrid`、`imageHero`、`quoteImage`、`textImage`
- 结构页：`compare`、`duoCompare`、`timeline`、`pipeline`、`roadmap`、`textGrid`、`article`、`fourCards`、`matrix`、`agenda`、`caseStudy`、`pyramid`、`radial`、`swimlane`

布局支持数量由 `scripts/generate-pptx.js` 内的 `validateTextSlots()` 规则表和各 layout 渲染函数共同决定。生成前会检查分点、图片、图表和表格槽位是否匹配。

## 当前分支已修复的重点问题

- CMB logo 自动从技能内置 `assets/logos/` 解析，不再提示用户手动复制。
- CMB logo 保持原始比例，不横向或纵向压扁。
- CMB 页眉统一白底，内容页右上角使用 SVG 纯图 logo 水印。
- CMB 背景网格与实际网格元素对齐。
- `statement` 页右侧改为图片槽位；无图片时显示 `IMAGE SLOT`。
- `media` / `mediaGrid` / `gallery` / `imageGrid` 图片槽位按用户图片数量自适应。
- 无用户图片且无显式图表时，不再默认塞图表，而是显示图片占位符。
- `fourCards` 和 `textGrid` 会根据文本量调整卡片高度，减少文本跑出卡片。
- CMB `media` 右侧说明区行距放宽，最多支持 4 个侧边分点。
- `radial` 连线改为从中心卡片边缘连到节点卡片边缘，避免覆盖中心卡片。
- `radial` 负斜率连线使用稳定方向，修复右上/左下连接方向错误。
- `radial` 外圈卡片增大，能容纳更多标题和正文。
- 生成 JSON 支持宽松解析和规范化输出，减少中英文引号、注释、尾逗号导致的解析失败。
- 自动变更重复 layout 时要求写出 normalized JSON，避免 PPTX 和 JSON 不一致。

## 图片、图表和占位符规则

- 用户提供 `image` / `images` / `gallery` 时优先插入用户图片。
- `mediaGrid` / `gallery` / `imageGrid` 默认槽位数跟随图片、图表或 caption 数量。
- 显式设置 `mediaCount` 时必须与图片数量匹配，除非明确允许空槽。
- 没有用户图片但提供 `chart` / `charts` 时，使用 PowerPoint 原生图表。
- 没有图片和显式图表时，显示 `IMAGE SLOT` 占位符。
- `statement` 只支持 1 个图片槽位，不支持 chart；需要多图请用 `mediaGrid` / `imageGrid`。

## 图标规则

分点对象可写 `icon` 字段：

```json
{ "icon": "shield-alert", "title": "风险预警", "body": "识别异常行为并前置处置。" }
```

图标优先来自 `lucide` 包。旧环境如果安装过 `lucide-static`，生成器也兼容读取。

## 推荐测试流程

修改模板或生成器后建议执行：

```bash
node --check scripts/generate-pptx.js
npm run sample:cmb:layouts
node scripts/validate-pptx-native.js assets/outputs/deck-cmb-all-layouts.pptx
node scripts/validate-pptx-layout.js assets/outputs/deck-cmb-all-layouts.pptx
```

如修改了复杂 layout，再补专项 JSON 到 `outputs/` 下生成测试 PPTX。`outputs/` 和 `assets/outputs/` 是生成物目录，不建议提交。

## 当前已知约束

- 这是原生 PPTX 生成器，不保证与 HTML/CSS 像素级一致。
- PowerPoint 字体渲染和浏览器不同，长文本仍建议拆页或减少分点。
- `layout` 自动多样化默认只给建议，不会直接改输入 JSON。需要真实改 layout 时必须配合 `--write-normalized-spec`。
- CMB 是独立风格，切换模板时优先只改 `style` / `theme`，不要批量改每页字段。

