# guizang-pptxgenjs-ppt-skill

基于 `pptxgenjs` 的可编辑 PowerPoint 生成技能。它直接生成 PowerPoint 原生文本框、形状、线条、图片、表格和图表，不依赖 HTML 截图、整页图片或 PDF 转换，便于后续在 PowerPoint/WPS 中继续编辑。

当前 `main` 分支包含三套可切换风格、统一的 JSON spec 输入、内置样例、图片/图表槽位校验、布局多样性提示、CMB 招商银行独立品牌风格，以及 PPTX 原生结构和布局风险校验脚本。

## 安装

建议使用 Node.js 18 或更高版本。

```bash
npm install
```

主要依赖：

- `pptxgenjs`：生成可编辑 `.pptx`
- `lucide`：为分点、卡片、指标等元素提供图标

## 生成 PPTX

使用 JSON spec 生成：

```bash
node scripts/generate-pptx.js --spec path/to/deck.json --out outputs/deck.pptx
```

生成内置样例：

```bash
npm run sample:magazine
npm run sample:swiss
npm run sample:cmb
```

生成 CMB 全 layout 检查文件：

```bash
npm run sample:cmb:layouts
```

输出位置：

```text
assets/outputs/deck-cmb-all-layouts.pptx
```

## 校验

生成后建议执行：

```bash
node scripts/validate-pptx-native.js path/to/deck.pptx
node scripts/validate-pptx-layout.js path/to/deck.pptx
```

- `validate-pptx-native.js`：检查是否为 PowerPoint 原生结构，避免整页截图伪装成 PPTX。
- `validate-pptx-layout.js`：检查明显布局风险，例如文本覆盖、元素冲突和底部安全区问题。

修改技能结构后可运行：

```bash
python C:\Users\lizey\.codex\skills\.system\skill-creator\scripts\quick_validate.py C:\Users\lizey\Desktop\guizang-ppt-skill-main\guizang-pptxgenjs-ppt-skill
```

## 支持风格

- `magazine`：电子杂志 / 电子墨水风格，适合观点、叙事、报告型页面。
- `swiss`：瑞士国际主义风格，适合科技、数据、方法论和产品说明。
- `cmb`：招商银行独立品牌风格，适合银行、金融、经营汇报和商务汇报。

默认主题：

- `magazine`: `ink`
- `swiss`: `ikb`
- `cmb`: `classic`

CMB 风格使用技能内置 logo 资源：

- 页眉完整白底 PNG logo：`logos/cmb-logo-lockup.png`
- 页眉外透明 SVG 纯图 logo：`logos/cmb-logo-mark.svg`

生成器会自动从 `assets/logos/` 解析这些资源，不需要把 logo 手动复制到项目目录。所有 logo 都按原比例显示，避免被压扁。

## JSON Spec 示例

所有 spec 文件建议保存为 UTF-8。不要手写拼接 JSON 字符串，优先由程序对象调用 `JSON.stringify(data, null, 2)` 写出。

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

JSON 引号与编码规则：

- 文件统一使用 UTF-8。
- 普通中文内容里的引号优先使用 `「」` 或中文弯引号。
- 如果 JSON 字符串中必须使用英文直引号 `"`，必须写成 `\"`。
- `--spec` 读取时会尝试修复常见问题：Markdown fenced code block、UTF-8 BOM、注释、尾逗号、部分中英文弯引号。
- 若需要把修复后的内容落盘为严格 JSON，使用 `--write-normalized-spec path/to/normalized.json`。

## 常用 Layout

三套风格支持统一的常用 layout 名称，切换模板时优先只修改顶层 `style` / `theme`。

- 封面与章节：`cover`、`section`、`closing`
- 结论页：`statement`、`bigQuote`
- 数据页：`kpiTower`、`bigNumbers`、`dashboard`、`chart`、`dataSheet`
- 图文页：`media`、`mediaGrid`、`gallery`、`imageGrid`、`imageHero`、`quoteImage`、`textImage`
- 结构页：`compare`、`duoCompare`、`timeline`、`pipeline`、`roadmap`、`textGrid`、`article`、`fourCards`、`matrix`、`agenda`、`caseStudy`、`pyramid`、`radial`、`swimlane`

布局支持的条目数量由 `scripts/generate-pptx.js` 中的槽位校验规则和对应渲染函数共同决定。生成前会检查文本、图片、图表和表格槽位是否匹配，避免多槽位、少槽位或字段名不匹配导致内容丢失。

## 图片、图表和占位符

- 用户提供 `image` / `images` / `gallery` 时优先插入用户图片。
- `mediaGrid` / `gallery` / `imageGrid` 未显式设置 `mediaCount` 时，槽位数自动等于图片数、显式图表数或 caption 数。
- 显式设置 `mediaCount` 时必须与图片数量匹配，除非明确允许空槽。
- 没有用户图片但提供 `chart` / `charts` 时，使用 PowerPoint 原生图表。
- 没有图片和显式图表时，显示 `IMAGE SLOT` 占位符，不再默认填充图表。
- `statement` 只支持 1 个图片槽位，不支持 chart；多图请使用 `mediaGrid` / `imageGrid`。

图表类型支持：

```text
bar, column, line, pie, doughnut, area, radar, scatter
```

## 布局多样性

规划 deck 时应避免连续 3 页以上使用同一 layout 或同一视觉节奏。生成器默认只输出重复 layout 的替换建议，不会擅自修改输入 JSON，保证 JSON 与 PPTX 一致。

如果确实要自动改 layout，必须同时使用：

```bash
node scripts/generate-pptx.js --spec path/to/deck.json --out outputs/deck.pptx --diversify-layouts --write-normalized-spec outputs/deck.normalized.json
```

后续应以 normalized JSON 作为真实源文件。

## 推荐开发检查流程

修改模板或生成器后建议执行：

```bash
node --check scripts/generate-pptx.js
npm run sample:magazine
npm run sample:swiss
npm run sample:cmb
npm run sample:cmb:layouts
node scripts/validate-pptx-native.js assets/outputs/deck-cmb-all-layouts.pptx
node scripts/validate-pptx-layout.js assets/outputs/deck-cmb-all-layouts.pptx
```

`outputs/` 和 `assets/outputs/` 是生成物目录，不建议提交。

## 已知边界

- 这是原生 PPTX 生成器，不保证与 HTML/CSS 像素级一致。
- PowerPoint 与 WPS 的字体渲染可能不同，长文本仍建议拆页或减少分点。
- 后置插入的 `blocks`、`charts`、`tables` 若缺少显式 `x/y/w/h` 会被跳过并打印 warning，避免与正文重叠。
- 自动 layout 多样化会改变页面类型，必须配合 `--write-normalized-spec` 使用。
