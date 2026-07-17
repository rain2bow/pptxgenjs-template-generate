# 技术发现

## 基线

- 当前分支基于 `master` 的 `893f6c9`。
- 仓库已有 `magazine`、`swiss`、`cmb` 三种 style。
- 当前工作区存在本地输出、依赖和 IDE 文件，不属于本次开发。
- `scripts/pptxgen/config.js` 仅表现为换行符状态，内容哈希与索引一致。

## 待调查

- layout 名称的定义、支持范围和 style 分派位置。
- 各 layout 实际读取的字段及验证规则。
- 样例、容量指南、Markdown 大纲和 CLI 对 layout/schema 的依赖。
- 模板模块之间是否存在调用或共享实现依赖。

## 阶段 1 初步发现

- layout 语义分散在 `templates/*.js`、`validation.js`、`text-capacity.js`、`spec-md.js`、`samples.js` 和媒体/数据块处理逻辑中，不能只做字符串替换。
- `validation.js` 已承担字段是否渲染、集合槽位数量、媒体/图表/表格必填等中心校验，适合作为统一协议的消费者，但目前规则仍以内联集合维护。
- 当前存在历史兼容 layout 和同义 layout 映射，例如 CMB 的 `article/sectionList/briefing/executiveBrief/contentBrief` 共用一套渲染，`textGrid/fourCards/textWeave/contentSynthesis/denseText` 共用另一套渲染。
- 三个 style 通过兼容 renderer 接收其他 style 的 layout，这与规划要求的“名称完全一致、类型完全一致”冲突；最终应以共同 canonical layout 集合替代兼容别名。
- 媒体布局目前集中识别为 `statement/media/mediaGrid/gallery/imageGrid/imageHero/quoteImage/textImage/caseStudy`，但名字不能稳定表达纯文本、图片或图表类型。
- 阶段 1 应先建立中央 schema：canonical 名称、类别、主集合字段、允许字段、数量约束；模板只负责同一 layout 在不同 style 下的视觉实现。

## 阶段 1 协议决策

- 对外 layout 使用类别前缀：`deck-*`、`text-*`、`image-*`、`data-*`，从名称即可判断是否需要图片或数据。
- 页面正文集合统一为 `items`。不再让同类页面在 `sections/items/columns/points/agenda/steps/nodes/layers/lanes/metrics/captions` 之间任选。
- 媒体统一为 `images`，图表统一为 `charts`，表格保留 `table`；单槽位也使用数组，减少单复数字段分叉。
- 对比页面保留 `before.items` / `after.items`，因为左右语义不能压平为一个无结构数组。
- 模板内部允许通过一个集中适配层复用现有 renderer，但对外校验、样例、Markdown 和文档只暴露 canonical schema。
- 旧 layout/字段不静默兼容：应给出目标 canonical layout/field 的明确迁移错误，避免继续积累双协议。
- `text-swimlane` 顶层统一使用 `items`，每个泳道条目仍需要 `title + body + items[]`：前两个字段用于泳道说明，内层 `items[]` 表示各阶段单元格。
- Swiss/CMB 原 `bigQuote` 兼容函数错误复用了强制图片的 statement renderer；canonical `text-quote` 已改为真正的全宽纯文本引用页。
- `text-briefing` 不再暴露 CMB 专属的 `summary/conclusion` 组合；三种 style 统一只接收 `items`，CMB 在无独立 lead 时把首项作为总领内容。
- Swiss/CMB 的 image hero、quote 和 case study 原先从页面顶边铺图，会覆盖统一页眉；媒体框现从 `y=0.82` 开始。
- 双行 `image-grid` 的图片与 caption 区域已整体上移并保留 `0.72in` 行间距；`data-table` 高度缩短，caption 不再进入页脚安全区。

## 阶段 2 流程决策

- style 选择前置为独立步骤：`--style-guide` 输出三套风格的适用场景和视觉说明；SKILL 明确要求用户未指定风格时通过 `askUserQuestion` 选择。
- 删除 `deck.plan.json -> capacity-guide -> 完整 JSON` 的多阶段主流程，改为选择 style 后生成对应的全布局 JSON 示例 Markdown，再直接编写完整 deck JSON。
- 新增 `layout-examples.js` 作为样例单一来源，三种 style 使用同一组 31 个 canonical layout 示例；示例展示完整字段结构，不包含字数上下限表。
- `--capacity-guide` 保留为显式失败入口，避免旧命令被静默解释为其他行为。
- 删除旧 `text-capacity.js` 和最大字数样例脚本；最终生成阶段的真实文本框溢出警告仍保留。
- 三个 `assets/template-*.js` 已迁移到 canonical layout/field 协议，证明直接生成完整 JSON 的路径可运行。
