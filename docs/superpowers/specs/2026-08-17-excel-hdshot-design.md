# Excel 高清截图工具 — 设计文档

日期：2026-08-17
项目位置：`~/Projects/excel-hdshot`

## 1. 背景与目标

**痛点**：需要对 Excel 表格做"完整截图"（整个表格导出为一张图）。传统做法是屏幕截图或 Excel 自带导出，当表格较大被缩放显示时，文字像素过小，字体从矢量渲染退化为位图渲染，文字发虚、模糊。

**解决方案**：不走屏幕位图，而是解析 `.xlsx/.csv` 源文件，在浏览器中用 HTML 表格重建内容，再以 SVG `foreignObject` + canvas 的方式按用户指定倍率（1x–4x）**矢量光栅化**导出 PNG。文字始终是矢量字形，按目标分辨率直接绘制，任意倍率下边缘清晰，彻底解决文字变位图的问题。

**成功标准**：
- 上传 xlsx 后，表格内容+样式被还原渲染，导出 2x/3x PNG 文字清晰锐利
- 中文、合并单元格、粗体、背景色、边框正确还原
- 纯前端、无框架，打开 HTML 即可使用

## 2. 技术选型

- 纯 HTML + CSS + JS（无框架），单文件或少量文件，直接 `file://` 打开可用
- **ExcelJS**（CDN，MIT）：解析 xlsx，读取单元格值、字体、填充、边框、对齐、数字格式、合并单元格、行高列宽
  - 放弃 SheetJS 社区版（读取样式是 Pro 付费功能）
- **导出管线**：DOM 表格 → XMLSerializer 序列化为 SVG `foreignObject` → `Image` 加载 dataURL → canvas 按倍率 `drawImage` → `toBlob` 下载 PNG
  - 不引入 html2canvas，保持可控与高质量

## 3. 模块划分

| 模块 | 职责 |
|---|---|
| `index.html` | 页面结构：上传区、sheet 选择、预览区、导出面板 |
| `style.css` | 界面样式与导出表格默认样式 |
| `parser.js` | 文件读取、ExcelJS 解析、sheet 列表、数据+样式提取成中间模型 |
| `renderer.js` | 中间模型 → HTML 表格（内联样式，合并单元格 colspan/rowspan，精确还原） |
| `exporter.js` | 表格 HTML → SVG → canvas → PNG（倍率/白边/背景色/圆角参数） |
| `main.js` | 事件编排、UI 状态、下载逻辑 |

### 中间数据模型（renderer 输入）
```js
{
  rows: [ { cells: [{ text, font, fill, border, align, numFmt, rowspan, colspan, hidden }], height } ],
  widths: [px...],      // 列宽
  merges: [{r1,c1,r2,c2}]
}
```

### 样式映射（ExcelJS → CSS）
- font: name/size/bold/italic/color → `font-family/size/weight/style/color`
- fill: fgColor → `background-color`（纯色；pattern 取 fgColor）
- border: 四边 style+color → `border-*`（thin/medium/thick/dashed/dotted 等映射）
- alignment: horizontal/vertical/wrapText → `text-align/vertical-align/white-space`
- 数字格式：解析时按 numFmt 格式化数值为文本（简化：用 ExcelJS 的文本输出或自定义常见格式）

## 4. 数据流

1. 用户拖拽/选择 `.xlsx` 或 `.csv`
2. `parser.js` 读取文件（`workbook.xlsx.load(arrayBuffer)`）
3. sheet 下拉列出所有工作表，默认选第一个
4. 选中 sheet → 提取数据+样式 → 中间模型
5. `renderer.js` 生成 HTML 表格注入预览区（CSS 等比缩放显示）
6. 用户调导出参数（倍率 1x/2x/3x/4x、白边距 px、背景白/透明、圆角 px）
7. `exporter.js`：克隆预览表格（去除预览缩放样式）→ 序列化 SVG → canvas（宽高×倍率）→ PNG 下载
8. 导出前显示目标分辨率与预估大小

## 5. 关键技术点

- **SVG foreignObject 序列化**：表格需内联所有样式；图片资源（Excel 内嵌图片 v1 忽略，表格内若有 `<img>` 转 dataURL）；避免外部资源导致 canvas 污染（tainted canvas）
- **矢量清晰度**：`ctx.drawImage` 前设置 `canvas.width = width*scale`，文字随 SVG 矢量光栅化到目标分辨率，不经过预览的 CSS 缩放路径
- **合并单元格**：转为 `rowspan/colspan`，被合并覆盖的单元格输出占位隐藏格
- **列宽/行高**：ExcelJS 给出列宽（字符单位）与行高（磅），换算为 px（1 字符宽 ≈ 7px，1pt ≈ 1.333px），预览与导出一致
- **大文件**：> 8000 单元格提示性能警告，导出时提示可能耗时长
- **CSV**：无样式，用默认样式（11px Calibri/等宽、全边框可选开关）

## 6. 错误处理

- 非 xlsx/csv 文件：提示"请选择 .xlsx 或 .csv 文件"
- 解析失败（损坏文件/加密文件）：明确错误提示，不上报
- 图片污染 canvas：捕获 SecurityError，提示改用无外部图片的表格
- 空 sheet：提示"该工作表无数据"

## 7. 验证方式

- 用 Node + ExcelJS 生成测试 xlsx：中文、合并单元格、粗体/斜体、背景色、四边不同边框、数字格式、多 sheet
- Playwright（webapp-testing skill）本地起服务打开页面：
  - 上传测试文件，检查预览还原（截图对比关键样式）
  - 导出 2x PNG，校验 `imageSize` 像素尺寸 = 原尺寸×2，文字区域无过度模糊（抽查截图目视）
  - 测试 CSV 上传、空文件、错误格式的错误提示路径

## 8. 范围外（YAGNI）

- 不做表格编辑（仅预览+导出参数调整）
- 不做多 sheet 合并导出
- 不做 PDF 导出
- 不做 xlsx 内嵌图片/图表还原（v1 仅还原单元格文本与样式）
- 不支持加密 xlsx
