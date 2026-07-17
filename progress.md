# 开发进度

## 2026-07-17

- 读取《新分支开发规划.md》并确认四个串行优化阶段。
- 从 `master` 创建并切换到 `layout-schema-v2`。
- 建立持久化计划、技术发现和进度日志。
- 下一步：扫描布局注册、字段协议、验证器与样例生成链路，形成阶段 1 迁移清单。
- 已完成第一轮架构扫描，确认 layout 协议横跨模板、校验、容量、Markdown 和样例模块。
- 决定阶段 1 先引入中央 canonical layout schema，再迁移三个 style 的 renderer 映射和校验消费者。
- 完成阶段 1 对外协议决策：类别前缀 layout，内容 `items`，媒体 `images`，图表 `charts`，表格 `table`。
- 新增 `layout-schema.js` 并接入 `engine.js`；迁移三种内置样例和 CMB 全布局样例。
- 首次全布局测试发现泳道条目缺少正文，已补充 `body` 后待复测。
- 第二次全布局测试发现 `blocks.js` 的表格 caption 调用未定义函数；已改为模块内 `addBlockCaption()`。
- 第三次全布局测试发现三个模板的泳道 renderer 引用了未定义的 `cellText`；已统一改为局部文本归一化。
- 第四次全布局测试定位第 30 页 `text-quote` 被旧兼容函数误送入图片 statement；已为 Swiss/CMB 实现无图片的全宽引用渲染。
- 首次三 style 共用全布局测试发现 `text-briefing` 暴露了 CMB 专属 scalar；已统一为仅使用 `items`。
- 布局校验确认 Swiss/CMB 三种大图页面覆盖统一页眉，且表格 caption/双行图片网格过低；已调整媒体框和底部边界。
- 阶段 1 最终结果：同一份 31 页 canonical spec 在 `cmb`、`swiss`、`magazine` 下均成功生成。
- 三套全布局文件均通过 `validate-pptx-native.js` 和 `validate-pptx-layout.js`。
- `npm run check:layout-schema` 通过，覆盖 31 个布局、内部适配、旧 layout 和旧字段迁移错误。
- `check-media-slot-warnings.js` 已迁移到 7 个 `image-*` 布局并通过。
- 阶段 1 已完成，下一步进入 style 选择和全布局 JSON 示例 Markdown 流程。
- 新增 `--style-guide`，输出 CMB、Swiss、Magazine 三套风格说明；SKILL 要求未指定风格时调用 `askUserQuestion`。
- 新增 `--layout-examples <style>`，每种 style 均生成 31 个完整 canonical layout JSON 示例且不输出容量区间。
- 删除 plan/capacity 主流程及 `text-capacity.js`，旧 `--capacity-guide` 会明确报错并提示新命令。
- README、SKILL、架构文档和三个模板样例脚本均已迁移到“选风格 -> 布局示例 -> 完整 JSON -> 生成”的新流程。
- `check:layout-schema` 通过 31 个布局，媒体槽位回归通过 7 个图片布局，相关 JS 均通过语法检查。
- CMB、Swiss、Magazine 示例分别生成 8、13、12 页，三份文件均通过原生 PPTX 与版面校验。
- 阶段 2 已完成，下一步建立图文布局配对矩阵和跨 style 对应 renderer。
