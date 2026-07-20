# 新增 Style 模板指南

本项目支持纯新增式 style 插件。新增模板时不需要修改 `engine.js`、`config.js`、`layout-examples.js`、`samples.js` 或任何已有模板文件。

## 目录约定

在以下目录新增一个以 style id 命名的子目录：

```text
scripts/pptxgen/templates/styles/
└── my-style/
    └── index.js
```

也可以把插件放在技能目录外，并通过环境变量指定一个或多个插件根目录：

```bash
export PPTXGEN_STYLE_PATHS="/opt/pptx-styles:/data/team-styles"
```

Windows 使用系统路径分隔符 `;`。每个根目录下仍按 `<style-id>/index.js` 组织。

## 最小插件

`index.js` 必须导出一个 style 定义对象：

```js
'use strict';

module.exports = {
  id: 'my-style',
  name: '我的模板',
  description: '说明视觉特征、适用内容和使用场景。',
  defaultTheme: 'default',
  themes: {
    default: {
      name: '默认配色',
      paper: 'FFFFFF',
      paperTint: 'F3F5F7',
      ink: '20242A',
      grey1: 'EEF1F4',
      grey2: 'CDD3DA',
      grey3: '69737F',
      accent: '176B5B',
      accentOn: 'FFFFFF',
      chartColors: ['176B5B', '3D7F9D', '8B96A2', 'CDD3DA'],
    },
  },
  createTemplate(api) {
    const { pptx, FONTS } = api;
    return {
      render(slide, ctx) {
        const data = ctx.sourceSlideSpec;
        slide.background = { color: ctx.theme.paper };
        slide.addText(data.title || '', {
          x: 0.8, y: 0.9, w: 11.7, h: 0.8,
          fontFace: FONTS.zh, fontSize: 28, bold: true,
          color: ctx.theme.ink, margin: 0,
        });
      },
    };
  },
};
```

字段要求：

- `id`：必须匹配 `^[a-z][a-z0-9-]*$`，且不能与已有 style 重名。
- `name`、`description`：会自动出现在 `--style-guide` 中。
- `themes`：至少包含一个主题；`defaultTheme` 必须指向其中一个 key。
- `createTemplate(api)`：必须返回 `{ render(slide, ctx) }`。
- 可选 `sampleSpec()`：返回该 style 的完整 deck JSON；未提供时生成器会使用三页通用样例。

## 渲染上下文

- `ctx.sourceSlideSpec`：用户输入的 canonical slide JSON，应优先读取它判断 `text-*`、`image-*`、`data-*`。
- `ctx.slideSpec`：为兼容内置成熟 renderer 生成的内部适配对象，新模板通常不需要使用。
- `ctx.theme`：当前主题对象。
- `ctx.spec`：整份 deck spec。
- `ctx.index`、`ctx.total`：当前页索引和总页数。
- `ctx.specDir`：输入 JSON 所在目录，用于解析相对图片路径。

`api` 提供 PptxGenJS 实例、字体、图片、图标、图表、表格、页眉页脚和文本估算等公共能力。完整列表查看 `engine.js` 的 `createTemplateApi()`。不要 `require` 其他 style 文件；共享能力只从 `api` 获取。

## Layout 要求

插件应处理 `layout-schema.js` 中的全部 canonical layouts。三类页面必须保持以下行为：

- `text-*`：完整渲染标题、正文和 `items`，不能要求图片。
- `image-*`：除 `images` 外与同后缀 `text-*` 字段一致，必须渲染图片且保持比例。
- `data-*`：使用原生图表或表格，不要把数据静默转成普通文本。

同一 renderer 可以服务多个 layout，但应读取 `ctx.sourceSlideSpec.layout` 保留页面语义。禁止静默忽略已填写字段。

## 验证

新增目录后无需修改注册表，直接运行：

```bash
node scripts/generate-pptx.js --style-guide
node scripts/generate-pptx.js --layout-examples my-style --out outputs/my-style-layouts.md
node scripts/generate-pptx.js --sample --sample-style my-style --out outputs/my-style-sample.pptx
node scripts/validate-pptx-native.js outputs/my-style-sample.pptx
node scripts/validate-pptx-layout.js outputs/my-style-sample.pptx
```

仓库内的最小隔离测试：

```bash
npm run check:style-plugin
```

测试插件位于 `scripts/fixtures/style-plugins/registry-test/`，只用于验证自动发现、样例输出、PPTX 生成和插件隔离，不会出现在默认 style 列表中。
