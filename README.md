# Excel 高清截图

上传 `.xlsx` / `.csv`，在浏览器中从数据矢量重绘表格，导出任意倍率（1x–4x）的高清 PNG。
解决传统屏幕截图缩小后文字退化为位图、发虚模糊的问题——文字始终以矢量字形按目标分辨率光栅化。

## 使用

直接用浏览器打开 `index.html`（无需服务器），拖入或点击选择文件，调整参数后点击"导出 PNG"。

## 参数

- 工作表：多 sheet 文件可切换导出哪张表
- 倍率：1x/2x/3x/4x，越大图片越清晰（导出尺寸 = 表格尺寸 × 倍率）
- 白边：图片四周留白像素
- 背景：白色 / 透明（透明背景 PNG 适合叠加到文档/网页）
- 圆角：导出图片的圆角半径

## 开发

```bash
npm install          # 安装 exceljs(devDep 供单测与夹具) 与 @playwright/test
npm run fixtures     # 重新生成测试夹具
npm run test:unit    # Node 内置 test runner 单测
npm run test:e2e     # Playwright e2e（需 chromium，系统已装）
```

## 说明

- xlsx 解析与样式还原使用 [ExcelJS](https://github.com/exceljs/exceljs)（MIT，已 vendored 到 `vendor/`）
- CSV 支持 UTF-8/GBK 自动识别、逗号/制表符分隔、引号转义、BOM
- 还原能力：单元格文本（含常见数字/日期格式）、字体、填充色、边框、对齐、合并单元格、行高列宽
- 导出管线：DOM 表格 → SVG `foreignObject`（XMLSerializer）→ canvas 按倍率 `zoom` 布局缩放 → PNG
- 已知限制：建议使用 Chromium/Edge 内核浏览器（`zoom` 布局缩放依赖其支持）；数字格式仅还原常见类型（货币符号/引号字面量/科学计数不还原）；未设列宽默认 59px；合并区域样式以主单元格为准；深色调（负 tint）按线性近似，与 Excel 显示略有偏差；部分透明色/渐变填充不还原；加密 xlsx 不支持
- 大表格（>8000 单元格）导出耗时较长属正常现象