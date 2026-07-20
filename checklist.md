# 新分支开发规划验收 Checklist

- 验收分支：`layout-schema-v2`
- 对照文件：`新分支开发规划.md`
- 验收日期：2026-07-20
- 总体结果：**9 项完成，0 项部分完成，0 项未开始**

## 开发过程约束

- [x] 按 1 → 2 → 3 → 4 的顺序实施。
- [x] 每阶段完成测试后单独提交，再进入下一阶段。
- [x] 阶段提交分别为 `1ddcb93`、`04330d4`、`50eb989`、`d43e841`；最终验证记录为 `bd425f1`。

## 1. 名称澄清

### 1.1 按纯文本、图片、图表区分 layout 名称

- [x] **完成**
- 对外名称统一使用 `deck-*`、`text-*`、`image-*`、`data-*` 前缀。
- `layout-schema.js` 当前定义 47 个 canonical layouts；`check-layout-schema.js` 强制检查每个名称必须带类别前缀。
- 数据页面使用 `data-chart`、`data-dashboard`、`data-table` 等明确名称，避免与图片页混淆。

### 1.2 不同 style 使用完全一致的 layout 名称和类型

- [x] **完成**
- CMB、Swiss、Magazine 共用同一份 `layout-schema.js`，不再对外暴露 style 专属 layout 名称。
- 同一份 47 页 spec 已分别生成三种 style，并全部通过 native/layout validator。
- 模板内部仍保留旧 renderer 名称，但只由 `createRendererSlide()` 私有适配；用户 JSON、文档和样例不暴露该兼容层。

### 1.3 尽量统一 layout 内容字段

- [x] **完成**
- 页面集合统一为 `items`，图片统一为 `images[]`，图表统一为 `charts[]`，表格统一为 `table`。
- 对比页因左右语义保留 `before.items` / `after.items`。
- 旧 `sections`、`columns`、`steps`、`nodes`、`layers`、`captions` 等顶层集合字段会明确报错并提示改为 `items`。
- renderer 内部可使用 `steps`、`layers` 等旧名称，但这是私有适配结果，不属于对外 JSON 协议。

## 2. 流程优化

### 2.1 添加 style 选择阶段并使用交互工具

- [x] **完成**
- 已实现 `--style-guide`，会从 style registry 动态输出当前已注册模板的名称、介绍和适用场景。
- `SKILL.md` 已要求先读取 `--style-guide` 的动态输出，再通过 `askUserQuestion` 展示所有已注册 style，包括 registry 插件；运行环境没有交互工具时才使用普通文本提问。
- `check:paired-fields` 会检查说明中不再写死三种内置 style。

### 2.2 移除 plan/字数 MD，改为全 layout JSON 示例 MD

- [x] **完成**
- 主流程已删除 `deck.plan.json` 和 capacity-guide 阶段。
- `text-capacity.js` 与旧最大字数样例脚本已删除。
- `--capacity-guide` 会明确报错并提示改用 `--layout-examples <style>`，不会静默继续旧流程。
- 三种内置 style 均可生成包含 47 个完整 JSON 示例的 Markdown，且不包含字数限制表。

## 3. 兼容性优化

### 3.1 所有纯文本/图片 layout 双向配对且字段兼容

- [x] **完成**
- 已建立 19 组 `text-<name>` / `image-<name>` 双向配对，每个定义均包含 reciprocal `counterpart`。
- 每个 pair 在 `layout-schema.js` 中具有显式 `publicFields` 内容字段契约；`check-layout-schema.js` 会验证两侧契约和示例字段完全一致。
- `paired-layouts.js` 可在图片页渲染 1-6 张图片及最多 8 个文本项，避免切换后直接丢失常见 `items` 内容。
- statement/quote 的 text/image 两侧均使用配对 renderer；`callout`、`caption`、`source/cite` 会实际写入 PPTX。
- `stages`、`highlightIndex`、`highlightLast`、`columnsCount`、`maxItems` 已由通用配对 renderer 消费；不在契约中的内容字段会在两侧一致报错，不再静默忽略。

### 3.2 缺图/不能含图时提示准确对应 layout，图表规则不变

- [x] **完成**
- `text-grid` 带图片会提示改为 `image-grid`；`image-grid` 缺图会提示改为 `text-grid`，其他配对同理。
- `check-media-slot-warnings.js` 已覆盖 19 个 image layout 缺图和 19 个 text layout 误带图片场景。
- 图表仍只使用 `data-chart` / `data-dashboard`，没有纳入 text/image 自动互换。

## 4. 模板引入优化

### 4.1 通过纯新增方式注册和使用新 style

- [x] **完成**
- 新增 `style-registry.js`，自动扫描 `templates/styles/<style-id>/index.js`。
- 也支持通过 `PPTXGEN_STYLE_PATHS` 在 Linux/Windows 引入技能目录外插件。
- 新插件自带 themes、defaultTheme、createTemplate 和可选 sampleSpec，无需修改 `engine.js`、`config.js`、`samples.js` 或已有模板文件。
- `check:style-plugin` 已验证正常插件发现、布局示例、PPTX 生成、native/layout 检查、损坏插件隔离和 fixture 不泄漏。

### 4.2 提供创建模板的指南 MD

- [x] **完成**
- 已新增 `scripts/pptxgen/STYLE_PLUGIN_GUIDE.md`。
- 指南包含目录约定、外部插件路径、最小 JS、主题字段、渲染上下文、canonical layout 责任和验证命令。

## 已通过的总体验证

- [x] `npm run check:layout-schema`：47 layouts、19 组配对通过。
- [x] `npm run check:media-slots`：19 个缺图和 19 个误带图片场景通过。
- [x] `npm run check:style-plugin`：插件发现、生成、验证和隔离通过。
- [x] `npm run check:paired-fields`：动态 style 说明、字段对称校验、三 style 文本写入、stages 和高亮控制通过。
- [x] CMB、Swiss、Magazine 各 47 页全布局 PPTX 通过 native/layout validator。
- [x] `scripts/pptxgen/**/*.js` 全部通过 `node --check`。

## 缺口补齐结果

- [x] `SKILL.md` 的 style 交互说明已改为读取 `--style-guide` 动态结果，不再写死三种内置 style。
- [x] 19 组 text/image layout 已建立显式公共内容字段契约，并增加三 style 实际 PPTX 文本标记验证。

## 缺口复验与修复结果

原始复验脚本为 `outputs/verify-remaining-gaps.js`；正式回归已固化为 `scripts/check-paired-layout-fields.js` 和 `npm run check:paired-fields`。测试 PPTX 写入系统临时目录并在结束后清理。

### 动态 style 交互

- [x] `PPTXGEN_STYLE_PATHS` 加载 `registry-test` 后，`--style-guide` 能发现该插件。
- [x] 同一环境下检查 `SKILL.md`，不再命中写死的“展示 `cmb`、`swiss`、`magazine`”。
- 结论：registry 和 Agent 交互说明均已动态化。

### text/image 字段兼容

| 验证字段 | text layout | image layout | 结果 |
|---|---:|---:|---|
| statement `callout` | 接受并写入 | 接受并写入 | 兼容 |
| quote `caption` | 接受并写入 | 接受并写入 | 兼容 |
| article `conclusion` | 明确拒绝 | 明确拒绝 | 兼容 |
| swimlane `stages` | 接受且写入 PPTX | 接受且写入 PPTX | 兼容 |

`check:paired-fields` 在 CMB、Swiss、Magazine 三种 style 中检查 statement、quote、swimlane 的唯一文本标记确实写入 PPTX，并确认 `stages`、`highlightIndex`、`highlightLast` 已被配对 renderer 消费。
