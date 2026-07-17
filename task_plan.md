# layout-schema-v2 开发计划

## 目标

按《新分支开发规划.md》依次完成布局协议、生成流程、图文兼容和模板插件化改造。每个阶段独立测试、独立提交，禁止跨阶段混合提交。

## 阶段

| 阶段 | 状态 | 交付物 | 验证要求 |
|---|---|---|---|
| 0. 建立分支与基线 | completed | `layout-schema-v2` 分支、规划文件 | 分支与工作区状态确认 |
| 1. 统一布局名称与字段协议 | completed | 统一的跨 style 布局名；文本/图片/图表类型可辨；尽量统一内容字段 | 三种 style 样例、旧字段错误提示、原生与布局校验；完成后 commit |
| 2. 优化选择与样例流程 | in_progress | style 选择说明；移除 plan/容量 MD 主流程；按 style 生成全布局 JSON 样例 MD | CLI/文档/样例测试；完成后 commit |
| 3. 补齐图文对应布局 | pending | 纯文本与含图布局成对、除媒体字段外协议一致；缺图/错图提示对应布局 | 配对矩阵、三 style PPTX、错误信息测试；完成后 commit |
| 4. 模板无损引入机制 | pending | 新 style 通过独立模块注册，不修改已有模板逻辑；模板创建指南 | 用最小测试 style 验证注册、生成和隔离；完成后 commit |
| 5. 总体验证 | pending | 全量回归、文档整理 | 三 style 全布局、关键错误场景、git 状态确认 |

## 约束与决策

- 以单一 canonical layout/schema 为目标，不以 style 别名兼容作为最终方案。
- 若必须迁移旧 JSON，使用显式迁移或明确错误，不静默改 layout。
- 图表布局保持独立数据协议，不纳入图文互换。
- 每完成一个优化阶段，必须测试并提交后才能进入下一阶段。
- 不提交 `node_modules/`、`outputs/`、`.idea/`、`assets/outputs/` 和仅有换行状态的 `scripts/pptxgen/config.js`。

## 遇到的错误

| 错误 | 尝试次数 | 解决方案 |
|---|---:|---|
| Windows sandbox `CreateProcessAsUserW failed: 1312` | 1 | 对必要的只读检查使用已批准的提升权限重试 |
| `data-table` 带 caption 时 `addCaption is not defined` | 1 | 在 `blocks.js` 增加独立的原生文本 caption 绘制函数，避免未注入依赖 |
| `text-swimlane` 渲染时 `cellText is not defined` | 1 | 三个模板均在泳道 renderer 内显式归一化字符串/对象单元格文本 |
| PowerShell 解析内联 Node 正则失败 | 1 | 改用 `outputs/inspect-slide-pictures.js` 临时只读检查脚本，避免继续尝试同一转义方式 |
