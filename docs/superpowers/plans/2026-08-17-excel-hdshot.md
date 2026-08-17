# Excel 高清截图工具 Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** 构建纯 H5 单页工具：上传 .xlsx/.csv，浏览器内重建表格并以 SVG+canvas 按 1x–4x 倍率矢量光栅化导出高清 PNG。

**Architecture:** ExcelJS 解析文件为中间模型（数据+样式+合并），renderer 将模型转 HTML 表格字符串注入预览，exporter 把表格 DOM 序列化为 SVG foreignObject 后画到 canvas 按倍率导出 PNG。文字始终矢量字形直接光栅化到目标分辨率。

**Tech Stack:** 纯 HTML/CSS/JS（无框架）；ExcelJS 4.4.0（本地 vendor + npm devDep）；Node 内置 test runner（单元测试）；@playwright/test（e2e）。

**Spec:** `docs/superpowers/specs/2026-08-17-excel-hdshot-design.md`

**File Structure:**
```
excel-hdshot/
  index.html          页面结构：上传区、工具栏（sheet选择/倍率/白边/背景/圆角/导出）、预览区、toast
  style.css           界面样式 + 预览容器
  js/parser.js        文件→中间模型（xlsx via ExcelJS；csv 原生解析含 GBK 回退）
  js/renderer.js      中间模型→HTML 表格字符串
  js/exporter.js      DOM 表格→SVG→canvas→PNG blob（仅浏览器）
  js/main.js          事件编排、状态、下载（仅浏览器）
  vendor/exceljs.min.js   本地 vendored CDN 文件
  scripts/make-fixtures.js  生成测试 xlsx/csv/bad.txt
  tests/unit/parser.test.js
  tests/unit/renderer.test.js
  tests/e2e/hdshot.spec.js
  tests/fixtures/      （生成物，提交入库）
  playwright.config.js
  package.json
  README.md
```

**统一接口（任务间类型一致，勿改）：**

```js
// parser.js（UMD，Node 下 module.exports = factory(require('exceljs'))，浏览器挂 window.XlsxParser）
XlsxParser.parseXlsx(arrayBuffer) -> Promise<{ sheets: Sheet[] }>
XlsxParser.parseCsv(arrayBuffer, fileName) -> { sheets: [Sheet] }

// Sheet = { name: string, rows: Row[], colWidths: (number|null)[], cellCount: number }
// Row = { height: number|null(px), cells: (Cell|null)[] }   // cells 按列下标对齐，空位为 null
// Cell = {
//   text: string, rowspan: number, colspan: number, hidden: boolean,
//   font: null | { name, size(pt), bold, italic, underline, color('#rrggbb'|null) },
//   fill: null | { color('#rrggbb') },
//   align: null | { h: 'left'|'center'|'right'|'justify'|null, v: 'top'|'middle'|'bottom'|null, wrap: boolean },
//   border: null | { top: null|{style,color}, bottom: ..., left: ..., right: ... }
// }

// renderer.js（UMD，浏览器挂 window.TableRenderer）
TableRenderer.renderSheet(sheet, opts?: { allBorders?: boolean }) -> string  // HTML 表格字符串

// exporter.js（UMD，浏览器挂 window.Exporter）
Exporter.exportPng(tableEl, opts) -> Promise<{ blob: Blob, width: number, height: number }>
// opts = { scale: number, margin: number(px), background: string(css), radius: number(px) }
```

---

### Task 1: 脚手架与环境

**Files:**
- Create: `package.json`, `.gitignore`

- [ ] **Step 1: 创建 package.json**

```json
{
  "name": "excel-hdshot",
  "private": true,
  "scripts": {
    "fixtures": "node scripts/make-fixtures.js",
    "test:unit": "node --test tests/unit/",
    "test:e2e": "playwright test",
    "test": "npm run test:unit && npm run test:e2e"
  },
  "devDependencies": {
    "exceljs": "4.4.0",
    "@playwright/test": "^1.48.0"
  }
}
```

- [ ] **Step 2: 创建 .gitignore**

```
node_modules/
test-results/
playwright-report/
```

- [ ] **Step 3: 安装依赖并 vendor ExcelJS**

```bash
npm install
mkdir -p vendor
curl -L -o vendor/exceljs.min.js https://cdn.jsdelivr.net/npm/exceljs@4.4.0/dist/exceljs.min.js
ls -la vendor/exceljs.min.js   # 期望 > 500KB
```

若网络失败：把 `index.html` 中 `vendor/exceljs.min.js` 改为 CDN URL 并在 README 注明。

- [ ] **Step 4: 验证 Node 测试可用**

```bash
node --test 2>&1 | head -5
```
Expected: 输出 node:test 用法说明，无报错。

- [ ] **Step 5: Commit**

```bash
git add package.json package-lock.json .gitignore vendor/exceljs.min.js
git commit -m "chore: 脚手架、依赖与 vendored exceljs"
```

---

### Task 2: 测试夹具生成脚本

**Files:**
- Create: `scripts/make-fixtures.js`

- [ ] **Step 1: 编写脚本**

```js
const ExcelJS = require('exceljs');
const fs = require('fs');
const path = require('path');

const outDir = path.join(__dirname, '..', 'tests', 'fixtures');

(async () => {
  fs.mkdirSync(outDir, { recursive: true });

  const wb = new ExcelJS.Workbook();

  const ws = wb.addWorksheet('样式测试');
  ws.mergeCells('A1:D1');
  const title = ws.getCell('A1');
  title.value = '销售数据汇总';
  title.font = { name: '微软雅黑', size: 16, bold: true, color: { argb: 'FFFFFFFF' } };
  title.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF4472C4' } };
  title.alignment = { horizontal: 'center', vertical: 'middle' };
  ws.getRow(1).height = 30;

  const headers = ['商品', '数量', '单价', '日期'];
  headers.forEach((h, i) => {
    const c = ws.getRow(2).getCell(i + 1);
    c.value = h;
    c.font = { bold: true };
    c.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFE7E6E6' } };
    c.alignment = { horizontal: 'center', vertical: 'middle' };
    c.border = { top: { style: 'thin' }, bottom: { style: 'thin' }, left: { style: 'thin' }, right: { style: 'thin' } };
  });

  ws.getCell('A3').value = '苹果';
  ws.getCell('B3').value = 1234.5;
  ws.getCell('B3').numFmt = '#,##0.00';
  ws.getCell('C3').value = 3.2;
  ws.getCell('D3').value = new Date(2026, 7, 1);
  ws.getCell('D3').numFmt = 'yyyy-mm-dd';

  ws.getCell('A4').value = '香蕉（较长备注说明文本测试换行显示）';
  ws.getCell('A4').alignment = { wrapText: true, vertical: 'middle' };
  ws.getCell('B4').value = 20;
  ws.getCell('B4').font = { italic: true };
  ws.getCell('C4').value = 5.5;
  ws.getCell('C4').fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFFFF2CC' } };
  ws.getCell('D4').value = new Date(2026, 7, 2);
  ws.getCell('D4').numFmt = 'yyyy-mm-dd';
  ws.getRow(4).height = 40;

  ws.mergeCells('A5:A6');
  ws.getCell('A5').value = '合并两行';
  ws.getCell('A5').alignment = { horizontal: 'center', vertical: 'middle' };
  ws.getCell('B5').value = 7;
  ws.getCell('C5').value = 1.5;
  ws.getCell('D5').value = new Date(2026, 7, 3);
  ws.getCell('D5').numFmt = 'yyyy-mm-dd';
  ws.getCell('B6').value = 8;
  ws.getCell('C6').value = 2.5;
  ws.getCell('D6').value = new Date(2026, 7, 4);
  ws.getCell('D6').numFmt = 'yyyy-mm-dd';
  ws.getCell('B5').border = { top: { style: 'dashed' } };

  ws.getColumn(1).width = 26;
  ws.getColumn(2).width = 12;
  ws.getColumn(3).width = 10;
  ws.getColumn(4).width = 14;

  wb.addWorksheet('空表');

  const ws3 = wb.addWorksheet('纯数据');
  ws3.getCell('A1').value = '姓名';
  ws3.getCell('B1').value = '年龄';
  ws3.getCell('A2').value = '张三';
  ws3.getCell('B2').value = 28;

  await wb.xlsx.writeFile(path.join(outDir, 'test.xlsx'));

  fs.writeFileSync(
    path.join(outDir, 'test.csv'),
    '\ufeff商品,数量,备注\n苹果,3,"含,逗号"\n香蕉,5,"他说""你好"""\n',
    'utf8'
  );

  fs.writeFileSync(path.join(outDir, 'bad.txt'), 'not a sheet', 'utf8');
  console.log('fixtures written to', outDir);
})();
```

- [ ] **Step 2: 运行生成**

Run: `npm run fixtures`
Expected: 输出 `fixtures written to .../tests/fixtures`

- [ ] **Step 3: 验证生成物**

```bash
ls -la tests/fixtures/   # test.xlsx, test.csv, bad.txt
```

- [ ] **Step 4: Commit**

```bash
git add scripts/make-fixtures.js tests/fixtures
git commit -m "test: 生成测试夹具（含中文/合并/样式/多sheet）"
```

---

### Task 3: parser.js — CSV 解析（含 GBK 回退）

**Files:**
- Create: `js/parser.js`
- Test: `tests/unit/parser.test.js`

- [ ] **Step 1: 编写失败测试**

```js
const test = require('node:test');
const assert = require('node:assert');
const fs = require('fs');
const path = require('path');
const parser = require('../../js/parser.js');

const fixture = (n) => fs.readFileSync(path.join(__dirname, '..', 'fixtures', n));

test('parseCsv: utf8 中文+引号内逗号与双引号转义', () => {
  const { sheets } = parser.parseCsv(fixture('test.csv'), 'test.csv');
  assert.strictEqual(sheets.length, 1);
  assert.strictEqual(sheets[0].name, 'test');
  assert.strictEqual(sheets[0].rows.length, 3);
  assert.strictEqual(sheets[0].rows[0].cells[0].text, '商品');
  assert.strictEqual(sheets[0].rows[1].cells[2].text, '含,逗号');
  assert.strictEqual(sheets[0].rows[2].cells[2].text, '他说"你好"');
});

test('parseCsv: 默认样式与 cellCount', () => {
  const { sheets } = parser.parseCsv(fixture('test.csv'), 'test.csv');
  const cell = sheets[0].rows[1].cells[1];
  assert.strictEqual(cell.font.name, 'Calibri');
  assert.strictEqual(cell.font.size, 11);
  assert.strictEqual(cell.font.bold, false);
  assert.strictEqual(cell.rowspan, 1);
  assert.strictEqual(cell.hidden, false);
  assert.strictEqual(sheets[0].cellCount, 9);
});

test('parseCsv: GBK 编码回退', () => {
  const gbk = Buffer.from([0xCA, 0xFD, 0xBE, 0xDD, 0x0A, 0xD6, 0xD0, 0xCE, 0xC4]); // "数据\n中文"
  const { sheets } = parser.parseCsv(new Uint8Array(gbk).buffer, 'gbk.csv');
  assert.strictEqual(sheets[0].rows[0].cells[0].text, '数据');
  assert.strictEqual(sheets[0].rows[1].cells[0].text, '中文');
});

test('parseCsv: tab 分隔自动识别', () => {
  const buf = new TextEncoder().encode('a\tb\n1\t2\n');
  const { sheets } = parser.parseCsv(buf, 't.csv');
  assert.strictEqual(sheets[0].rows[0].cells[1].text, 'b');
  assert.strictEqual(sheets[0].rows[1].cells[1].text, '2');
});
```

- [ ] **Step 2: 运行测试确认失败**

Run: `node --test tests/unit/parser.test.js`
Expected: FAIL — `Cannot find module '../../js/parser.js'`

- [ ] **Step 3: 实现 parser.js（UMD 骨架 + CSV 部分；parseXlsx 留待 Task 4）**

```js
(function (root, factory) {
  if (typeof module === 'object' && module.exports) {
    module.exports = factory(require('exceljs'));
  } else {
    root.XlsxParser = factory(root.ExcelJS);
  }
})(typeof self !== 'undefined' ? self : this, function (ExcelJS) {
  const DEFAULT_FONT = { name: 'Calibri', size: 11, bold: false, italic: false, underline: false, color: null };

  function decodeBuffer(buf) {
    try {
      return new TextDecoder('utf-8', { fatal: true }).decode(buf);
    } catch (e) {
      return new TextDecoder('gbk').decode(buf);
    }
  }

  function detectDelimiter(text) {
    const nl = text.indexOf('\n');
    const firstLine = text.slice(0, nl === -1 ? text.length : nl);
    const tabs = (firstLine.match(/\t/g) || []).length;
    const commas = (firstLine.match(/,/g) || []).length;
    return tabs > commas ? '\t' : ',';
  }

  function parseCsvText(text) {
    const delim = detectDelimiter(text);
    const rows = [];
    let row = [], field = '', inQuotes = false;
    for (let i = 0; i < text.length; i++) {
      const ch = text[i];
      if (inQuotes) {
        if (ch === '"') {
          if (text[i + 1] === '"') { field += '"'; i++; } else { inQuotes = false; }
        } else {
          field += ch;
        }
      } else if (ch === '"') {
        inQuotes = true;
      } else if (ch === delim) {
        row.push(field); field = '';
      } else if (ch === '\n') {
        row.push(field); rows.push(row); row = []; field = '';
      } else if (ch !== '\r') {
        field += ch;
      }
    }
    if (field !== '' || row.length) { row.push(field); rows.push(row); }
    return rows;
  }

  function makeCell(text) {
    return {
      text: String(text), rowspan: 1, colspan: 1, hidden: false,
      font: Object.assign({}, DEFAULT_FONT), fill: null, align: null, border: null
    };
  }

  function parseCsv(buf, fileName) {
    const text = decodeBuffer(buf).replace(/^\ufeff/, '');
    const rows = parseCsvText(text).map(r => ({ height: null, cells: r.map(makeCell) }));
    const cellCount = rows.reduce((n, r) => n + r.cells.length, 0);
    return {
      sheets: [{ name: fileName.replace(/\.[^.]+$/, ''), rows, colWidths: [], cellCount }]
    };
  }

  function parseXlsx() {
    throw new Error('parseXlsx 未实现');
  }

  return { parseXlsx, parseCsv };
});
```

- [ ] **Step 4: 运行测试确认通过**

Run: `node --test tests/unit/parser.test.js`
Expected: PASS（4 个测试全过）

- [ ] **Step 5: Commit**

```bash
git add js/parser.js tests/unit/parser.test.js
git commit -m "feat: CSV解析（引号转义/分隔符识别/GBK回退）"
```

---

### Task 4: parser.js — XLSX 解析（样式+合并单元格）

**Files:**
- Modify: `js/parser.js`（替换 parseXlsx 与新增辅助函数）
- Test: `tests/unit/parser.test.js`（追加测试）

- [ ] **Step 1: 追加失败测试**

```js
test('parseXlsx: sheet 列表', async () => {
  const { sheets } = await parser.parseXlsx(fixture('test.xlsx'));
  assert.deepStrictEqual(sheets.map(s => s.name), ['样式测试', '空表', '纯数据']);
});

test('parseXlsx: 合并标题单元格样式', async () => {
  const { sheets } = await parser.parseXlsx(fixture('test.xlsx'));
  const cell = sheets[0].rows[0].cells[0];
  assert.strictEqual(cell.text, '销售数据汇总');
  assert.strictEqual(cell.rowspan, 1);
  assert.strictEqual(cell.colspan, 4);
  assert.strictEqual(cell.font.bold, true);
  assert.strictEqual(cell.font.size, 16);
  assert.strictEqual(cell.font.color, '#FFFFFF');
  assert.strictEqual(cell.fill.color, '#4472C4');
  assert.strictEqual(cell.align.h, 'center');
  assert.strictEqual(cell.align.v, 'middle');
  assert.strictEqual(sheets[0].rows[0].cells[1].hidden, true);
});

test('parseXlsx: 数字格式与日期文本', async () => {
  const { sheets } = await parser.parseXlsx(fixture('test.xlsx'));
  const row3 = sheets[0].rows[2];
  assert.strictEqual(row3.cells[1].text, '1,234.50');
  assert.strictEqual(row3.cells[3].text, '2026-08-01');
});

test('parseXlsx: 垂直合并与换行、行高列宽', async () => {
  const { sheets } = await parser.parseXlsx(fixture('test.xlsx'));
  const s = sheets[0];
  assert.strictEqual(s.rows[4].cells[0].rowspan, 2);
  assert.strictEqual(s.rows[4].cells[0].text, '合并两行');
  assert.strictEqual(s.rows[5].cells[0].hidden, true);
  assert.strictEqual(s.rows[3].cells[0].align.wrap, true);
  assert.strictEqual(s.rows[0].height, 40);          // 30pt*1.3333≈40
  assert.strictEqual(s.colWidths[0], 182);           // 26chars*7
  assert.strictEqual(s.colWidths[3], 98);            // 14chars*7
});
```

- [ ] **Step 2: 运行测试确认新测试失败**

Run: `node --test tests/unit/parser.test.js`
Expected: 新 4 个测试 FAIL（`parseXlsx 未实现`）

- [ ] **Step 3: 实现 parseXlsx**

在 `js/parser.js` 中把 `function parseXlsx() {...}` 替换为：

```js
  function colToIndex(ref) {
    let n = 0;
    for (const ch of ref) n = n * 26 + ch.charCodeAt(0) - 64;
    return n - 1;
  }

  function argbToCss(argb) {
    if (!argb || argb === 'FF000000') return null;
    if (argb.length === 8 && argb.startsWith('FF')) return '#' + argb.slice(2).toUpperCase();
    return null;
  }

  function normalizeFont(f) {
    if (!f) return Object.assign({}, DEFAULT_FONT);
    return {
      name: f.name || 'Calibri',
      size: f.size || 11,
      bold: !!f.bold,
      italic: !!f.italic,
      underline: !!f.underline,
      color: argbToCss(f.color && f.color.argb)
    };
  }

  function normalizeFill(f) {
    if (!f || f.type !== 'pattern') return null;
    const c = argbToCss(f.fgColor && f.fgColor.argb);
    return c ? { color: c } : null;
  }

  function side(s) {
    if (!s || !s.style) return null;
    return { style: s.style, color: argbToCss(s.color && s.color.argb) };
  }

  function normalizeBorder(b) {
    if (!b) return null;
    const out = { top: side(b.top), bottom: side(b.bottom), left: side(b.left), right: side(b.right) };
    return (out.top || out.bottom || out.left || out.right) ? out : null;
  }

  function normalizeAlign(a) {
    if (!a) return null;
    const h = { left: 'left', center: 'center', right: 'right', justify: 'justify' }[a.horizontal] || null;
    const v = { top: 'top', middle: 'middle', bottom: 'bottom' }[a.vertical] || null;
    return (h || v || a.wrapText) ? { h, v, wrap: !!a.wrapText } : null;
  }

  function hasContent(cell, st) {
    if (cell.value !== null && cell.value !== undefined) return true;
    if (st.fill) return true;
    if (st.border && (st.border.top || st.border.bottom || st.border.left || st.border.right)) return true;
    if (st.alignment) return true;
    if (st.font) {
      const f = st.font;
      if (f.bold || f.italic || f.underline || f.name || f.size || f.color) return true;
    }
    return false;
  }

  // 注意：ExcelJS 4.4.0 的 cell.text 不按 numFmt 格式化（数字返回原始值、日期返回 Date.toString）。
  // 因此 parser 自带小型 numFmt 格式化器（已验证：B3 '#,##0.00'→'1,234.50'，D3 'yyyy-mm-dd'→'2026-08-01'）。
  const MONTHS = ['January','February','March','April','May','June','July','August','September','October','November','December'];
  const DAYS = ['Sunday','Monday','Tuesday','Wednesday','Thursday','Friday','Saturday'];

  function pad2(n) { return String(n).padStart(2, '0'); }

  function formatDateValue(d, fmt) {
    const hasTime = /(h+|s+)/i.test(fmt);
    const tokens = [
      ['yyyy', () => String(d.getFullYear())],
      ['yy', () => String(d.getFullYear()).slice(-2)],
      ['mmmm', () => MONTHS[d.getMonth()]],
      ['mmm', () => MONTHS[d.getMonth()].slice(0, 3)],
      ['mm', () => pad2(hasTime ? d.getMinutes() : d.getMonth() + 1)],
      ['m', () => hasTime ? d.getMinutes() : d.getMonth() + 1],
      ['dddd', () => DAYS[d.getDay()]],
      ['ddd', () => DAYS[d.getDay()].slice(0, 3)],
      ['dd', () => pad2(d.getDate())],
      ['d', () => d.getDate()],
      ['hh', () => pad2(d.getHours())],
      ['h', () => d.getHours()],
      ['ss', () => pad2(d.getSeconds())],
      ['s', () => d.getSeconds()],
      ['AM/PM', () => d.getHours() < 12 ? 'AM' : 'PM'],
      ['A/P', () => d.getHours() < 12 ? 'A' : 'P']
    ];
    let out = '';
    let i = 0;
    while (i < fmt.length) {
      const ch = fmt[i];
      if (ch === '"') {
        const end = fmt.indexOf('"', i + 1);
        if (end > i) { out += fmt.slice(i + 1, end); i = end + 1; continue; }
        out += ch; i++; continue;
      }
      if (ch === '\\') {
        if (i + 1 < fmt.length) out += fmt[i + 1];
        i += 2; continue;
      }
      let matched = false;
      for (const [tok, fn] of tokens) {
        if (fmt.startsWith(tok, i)) { out += fn(); i += tok.length; matched = true; break; }
      }
      if (!matched) { out += ch; i++; }
    }
    return out;
  }

  function formatNumberValue(v, fmt) {
    if (fmt == null || fmt === '' || fmt === 'General') return String(v);
    const sections = String(fmt).split(';');
    const isNeg = v < 0;
    const section = isNeg ? (sections[1] || sections[0]) : sections[0];
    const abs = Math.abs(v);
    const pct = section.includes('%');
    const pattern = section.replace(/%/g, '');
    const dot = pattern.indexOf('.');
    const intPat = dot >= 0 ? pattern.slice(0, dot) : pattern;
    const decPat = dot >= 0 ? pattern.slice(dot + 1) : '';
    const grouping = intPat.includes(',');
    const decPlaces = (decPat.match(/0/g) || []).length;
    const num = pct ? abs * 100 : abs;
    let str = num.toFixed(decPlaces);
    if (grouping) {
      const parts = str.split('.');
      str = parts[0].replace(/\B(?=(\d{3})+(?!\d))/g, ',') + (parts[1] !== undefined ? '.' + parts[1] : '');
    }
    if (pct) str += '%';
    return (isNeg ? '-' : '') + str;
  }

  function formatCellValue(v, numFmt) {
    if (v === null || v === undefined) return '';
    if (v instanceof Date) return formatDateValue(v, numFmt || 'yyyy-mm-dd');
    if (typeof v === 'number') return formatNumberValue(v, numFmt);
    if (Array.isArray(v)) return v.map(r => (r && r.text) || '').join('');
    if (v && typeof v === 'object' && v.error) return v.error;
    return String(v);
  }

  function textOf(cell) {
    return formatCellValue(cell.value, cell.numFmt);
  }

  async function parseXlsx(buf) {
    const wb = new ExcelJS.Workbook();
    await wb.xlsx.load(buf);
    const sheets = wb.worksheets.map(ws => {
      const merges = [];
      for (const spec of (ws.model.merges || [])) {
        const m = /^([A-Z]+)(\d+):([A-Z]+)(\d+)$/.exec(spec);
        if (!m) continue;
        merges.push({ r1: +m[2] - 1, c1: colToIndex(m[1]), r2: +m[4] - 1, c2: colToIndex(m[3]) });
      }
      const rowCount = ws.actualRowCount;
      const colCount = ws.actualColumnCount;
      const rows = [];
      let cellCount = 0;
      for (let r = 1; r <= rowCount; r++) {
        const xrow = ws.getRow(r);
        const cells = [];
        for (let c = 1; c <= colCount; c++) {
          const xc = xrow.getCell(c);
          const st = xc.style || {};
          if (!hasContent(xc, st)) { cells.push(null); continue; }
          const ri = r - 1, ci = c - 1;
          const master = merges.find(mm => mm.r1 === ri && mm.c1 === ci);
          const covered = merges.some(mm => mm.r1 <= ri && mm.c1 <= ci && mm.r2 >= ri && mm.c2 >= ci && !(mm.r1 === ri && mm.c1 === ci));
          cells.push({
            text: textOf(xc),
            rowspan: master ? master.r2 - master.r1 + 1 : 1,
            colspan: master ? master.c2 - master.c1 + 1 : 1,
            hidden: covered,
            font: normalizeFont(st.font),
            fill: normalizeFill(st.fill),
            align: normalizeAlign(st.alignment),
            border: normalizeBorder(st.border)
          });
          cellCount++;
        }
        const h = xrow.height;
        rows.push({ height: h ? Math.round(h * 1.3333) : null, cells });
      }
      const colWidths = [];
      for (let c = 1; c <= colCount; c++) {
        const w = ws.getColumn(c).width;
        colWidths.push(w ? Math.round(w * 7) : null);
      }
      return { name: ws.name, rows, colWidths, cellCount };
    });
    return { sheets };
  }
```

- [ ] **Step 4: 运行测试确认全部通过**

Run: `node --test tests/unit/parser.test.js`
Expected: PASS（8 个测试全过）

- [ ] **Step 5: Commit**

```bash
git add js/parser.js tests/unit/parser.test.js
git commit -m "feat: xlsx解析（字体/填充/边框/对齐/合并/行高列宽/数字格式）"
```

---

### Task 5: renderer.js — 模型转 HTML 表格

**Files:**
- Create: `js/renderer.js`
- Test: `tests/unit/renderer.test.js`

- [ ] **Step 1: 编写失败测试**

```js
const test = require('node:test');
const assert = require('node:assert');
const renderer = require('../../js/renderer.js');

const sheet = {
  name: 't',
  rows: [
    {
      height: 40,
      cells: [
        {
          text: '合并标题', rowspan: 1, colspan: 2, hidden: false,
          font: { name: '微软雅黑', size: 16, bold: true, italic: false, underline: false, color: '#FFFFFF' },
          fill: { color: '#4472C4' },
          align: { h: 'center', v: 'middle', wrap: false },
          border: null
        },
        { text: '', rowspan: 1, colspan: 1, hidden: true, font: null, fill: null, align: null, border: null }
      ]
    },
    {
      height: null,
      cells: [
        { text: '<b>&"', rowspan: 1, colspan: 1, hidden: false, font: null, fill: null, align: null, border: null },
        { text: '普通', rowspan: 1, colspan: 1, hidden: false, font: null, fill: null, align: null, border: null }
      ]
    }
  ],
  colWidths: [182, 70],
  cellCount: 3
};

test('合并单元格输出 colspan 且跳过 hidden 格', () => {
  const html = renderer.renderSheet(sheet);
  assert.ok(html.includes('colspan="2"'));
  const rowHtml = html.slice(html.indexOf('<tr'), html.indexOf('</tr>') + 5);
  assert.strictEqual((rowHtml.match(/<td/g) || []).length, 1);
});

test('样式内联输出', () => {
  const html = renderer.renderSheet(sheet);
  assert.ok(html.includes('font-family:\'微软雅黑\',sans-serif'));
  assert.ok(html.includes('font-size:16pt'));
  assert.ok(html.includes('font-weight:bold'));
  assert.ok(html.includes('color:#FFFFFF'));
  assert.ok(html.includes('background:#4472C4'));
  assert.ok(html.includes('text-align:center'));
  assert.ok(html.includes('vertical-align:middle'));
});

test('HTML 转义', () => {
  const html = renderer.renderSheet(sheet);
  assert.ok(html.includes('&lt;b&gt;&amp;&quot;'));
  assert.ok(!html.includes('<b>&"'));
});

test('行高与列宽', () => {
  const html = renderer.renderSheet(sheet);
  assert.ok(html.includes('height:40px'));
  assert.ok(html.includes('<col style="width:182px">'));
  assert.ok(html.includes('<col style="width:70px">'));
});

test('allBorders 选项（CSV）', () => {
  const csvSheet = {
    name: 'c',
    rows: [{ height: null, cells: [{ text: 'a', rowspan: 1, colspan: 1, hidden: false, font: null, fill: null, align: null, border: null }] }],
    colWidths: [],
    cellCount: 1
  };
  const html = renderer.renderSheet(csvSheet, { allBorders: true });
  assert.ok(html.includes('border-top:1px solid #999'));
  assert.ok(html.includes('border-bottom:1px solid #999'));
});

test('边框样式映射（thin/dashed）', () => {
  const bSheet = {
    name: 'b',
    rows: [{
      height: null,
      cells: [{
        text: 'x', rowspan: 1, colspan: 1, hidden: false, font: null, fill: null, align: null,
        border: { top: { style: 'thin', color: '#000000' }, bottom: { style: 'dashed', color: null }, left: null, right: null }
      }]
    }],
    colWidths: [],
    cellCount: 1
  };
  const html = renderer.renderSheet(bSheet);
  assert.ok(html.includes('border-top:1px solid #000000'));
  assert.ok(html.includes('border-bottom:1px dashed'));
});
```

- [ ] **Step 2: 运行测试确认失败**

Run: `node --test tests/unit/renderer.test.js`
Expected: FAIL — `Cannot find module '../../js/renderer.js'`

- [ ] **Step 3: 实现 renderer.js**

```js
(function (root, factory) {
  if (typeof module === 'object' && module.exports) {
    module.exports = factory();
  } else {
    root.TableRenderer = factory();
  }
})(typeof self !== 'undefined' ? self : this, function () {
  const BORDER_STYLE = {
    thin: '1px solid', hair: '0.5px solid', medium: '2px solid', thick: '3px solid',
    dashed: '1px dashed', dotted: '1px dotted', double: '3px double',
    mediumDashed: '2px dashed', mediumDashDot: '2px dashed', mediumDashDotDot: '2px dashed',
    slantDashDot: '2px dashed'
  };
  const DEFAULT_FONT = { name: 'Calibri', size: 11, bold: false, italic: false, underline: false, color: null };

  function esc(s) {
    return String(s)
      .replace(/&/g, '&amp;').replace(/</g, '&lt;')
      .replace(/>/g, '&gt;').replace(/"/g, '&quot;');
  }

  function styleString(s) {
    return Object.entries(s)
      .map(([k, v]) => k.replace(/[A-Z]/g, m => '-' + m.toLowerCase()) + ':' + v)
      .join(';');
  }

  function borderCss(sideSpec) {
    if (!sideSpec) return null;
    const st = BORDER_STYLE[sideSpec.style] || '1px solid';
    return sideSpec.color ? st + ' ' + sideSpec.color : st;
  }

  function renderCell(cell, opts) {
    const s = {};
    const f = cell.font || DEFAULT_FONT;
    if (f.name) s.fontFamily = "'" + f.name.replace(/'/g, '') + "',sans-serif";
    s.fontSize = (f.size || 11) + 'pt';
    if (f.bold) s.fontWeight = 'bold';
    if (f.italic) s.fontStyle = 'italic';
    if (f.underline) s.textDecoration = 'underline';
    if (f.color) s.color = f.color;
    if (cell.fill && cell.fill.color) s.background = cell.fill.color;
    const a = cell.align;
    s.textAlign = (a && a.h) || 'left';
    s.verticalAlign = (a && a.v) || 'middle';
    s.whiteSpace = (a && a.wrap) ? 'normal' : 'nowrap';
    s.padding = '0 4px';
    const b = cell.border || (opts.allBorders ? {
      top: { style: 'thin', color: '#999' }, bottom: { style: 'thin', color: '#999' },
      left: { style: 'thin', color: '#999' }, right: { style: 'thin', color: '#999' }
    } : null);
    if (b) {
      const t = borderCss(b.top); if (t) s.borderTop = t;
      const bt = borderCss(b.bottom); if (bt) s.borderBottom = bt;
      const l = borderCss(b.left); if (l) s.borderLeft = l;
      const r = borderCss(b.right); if (r) s.borderRight = r;
    }
    const attrs = ['rowspan', 'colspan']
      .filter(k => cell[k] > 1)
      .map(k => ' ' + k + '="' + cell[k] + '"')
      .join('');
    return '<td' + attrs + ' style="' + styleString(s) + '">' + esc(cell.text) + '</td>';
  }

  function renderSheet(sheet, opts) {
    opts = Object.assign({ allBorders: false }, opts);
    let html = '<table style="border-collapse:collapse;table-layout:fixed';
    const widths = (sheet.colWidths || []).filter(w => w != null);
    if (widths.length) {
      html += ';width:' + widths.reduce((a, b) => a + b, 0) + 'px';
    }
    html += '">';
    if (widths.length) {
      html += '<colgroup>' + (sheet.colWidths || []).map(w => w ? '<col style="width:' + w + 'px">' : '<col>').join('') + '</colgroup>';
    }
    html += '<tbody>';
    for (const row of sheet.rows) {
      html += '<tr' + (row.height ? ' style="height:' + row.height + 'px"' : '') + '>';
      for (const cell of row.cells) {
        if (!cell || cell.hidden) continue;
        html += renderCell(cell, opts);
      }
      html += '</tr>';
    }
    return html + '</tbody></table>';
  }

  return { renderSheet };
});
```

- [ ] **Step 4: 运行测试确认通过**

Run: `node --test tests/unit/renderer.test.js`
Expected: PASS（6 个测试全过）

- [ ] **Step 5: Commit**

```bash
git add js/renderer.js tests/unit/renderer.test.js
git commit -m "feat: 模型转HTML表格（内联样式/合并/转义/边框映射）"
```

---

### Task 6: index.html + style.css 静态骨架

**Files:**
- Create: `index.html`, `style.css`

- [ ] **Step 1: 编写 index.html**

```html
<!DOCTYPE html>
<html lang="zh-CN">
<head>
  <meta charset="UTF-8">
  <meta name="viewport" content="width=device-width, initial-scale=1">
  <title>Excel 高清截图</title>
  <link rel="stylesheet" href="style.css">
</head>
<body>
  <header>
    <h1>Excel 高清截图</h1>
    <p>上传 .xlsx / .csv，从数据矢量重绘，导出任意倍率的清晰 PNG</p>
  </header>
  <main>
    <section id="drop-zone">
      <p class="big">拖入文件，或点击选择</p>
      <p class="hint">支持 .xlsx / .csv</p>
      <input type="file" id="file-input" accept=".xlsx,.csv" hidden>
    </section>
    <section id="workspace" hidden>
      <div class="toolbar">
        <label>工作表
          <select id="sheet-select"></select>
        </label>
        <label>倍率
          <select id="scale-select">
            <option value="1">1x</option>
            <option value="2" selected>2x</option>
            <option value="3">3x</option>
            <option value="4">4x</option>
          </select>
        </label>
        <label>白边
          <input type="number" id="margin-input" value="16" min="0" max="200"> px
        </label>
        <label>背景
          <select id="bg-select">
            <option value="#ffffff">白色</option>
            <option value="transparent">透明</option>
          </select>
        </label>
        <label>圆角
          <input type="number" id="radius-input" value="0" min="0" max="60"> px
        </label>
        <button id="export-btn">导出 PNG</button>
        <span id="dims-info"></span>
      </div>
      <div id="preview"></div>
    </section>
  </main>
  <div id="toast"></div>
  <script src="vendor/exceljs.min.js"></script>
  <script src="js/parser.js"></script>
  <script src="js/renderer.js"></script>
  <script src="js/exporter.js"></script>
  <script src="js/main.js"></script>
</body>
</html>
```

- [ ] **Step 2: 编写 style.css**

```css
* { box-sizing: border-box; }
body {
  margin: 0;
  font-family: "PingFang SC", "Microsoft YaHei", system-ui, sans-serif;
  background: #f2f3f5;
  color: #1f2329;
}
header { padding: 24px 32px 8px; }
header h1 { margin: 0 0 4px; font-size: 22px; }
header p { margin: 0; color: #646a73; font-size: 13px; }
main { padding: 16px 32px 40px; }
#drop-zone {
  border: 2px dashed #c4c9d1;
  border-radius: 12px;
  background: #fff;
  text-align: center;
  padding: 64px 16px;
  cursor: pointer;
}
#drop-zone.over { border-color: #3370ff; background: #f0f5ff; }
#drop-zone .big { font-size: 18px; margin: 0 0 8px; }
#drop-zone .hint { color: #8f959e; margin: 0; font-size: 13px; }
#workspace .toolbar {
  display: flex;
  align-items: center;
  gap: 12px;
  flex-wrap: wrap;
  background: #fff;
  border: 1px solid #e4e6ea;
  border-radius: 10px;
  padding: 10px 16px;
  margin-bottom: 12px;
  font-size: 13px;
}
.toolbar label { display: flex; align-items: center; gap: 6px; }
.toolbar select, .toolbar input {
  padding: 4px 8px;
  border: 1px solid #d5d8dd;
  border-radius: 6px;
  font-size: 13px;
}
.toolbar input[type="number"] { width: 64px; }
#export-btn {
  padding: 7px 18px;
  border: none;
  border-radius: 8px;
  background: #3370ff;
  color: #fff;
  font-size: 14px;
  cursor: pointer;
}
#export-btn:disabled { background: #9ab4ff; cursor: wait; }
#dims-info { color: #646a73; }
#preview {
  overflow: auto;
  background:
    linear-gradient(45deg, #eee 25%, transparent 25%, transparent 75%, #eee 75%),
    linear-gradient(45deg, #eee 25%, #fff 25%, #fff 75%, #eee 75%);
  background-size: 16px 16px;
  background-position: 0 0, 8px 8px;
  border: 1px solid #e4e6ea;
  border-radius: 10px;
  padding: 16px;
}
#preview table { background: #fff; }
#toast {
  position: fixed;
  left: 50%;
  bottom: 32px;
  transform: translateX(-50%) translateY(20px);
  background: #1f2329;
  color: #fff;
  padding: 10px 20px;
  border-radius: 8px;
  font-size: 14px;
  opacity: 0;
  pointer-events: none;
  transition: all 0.25s;
  max-width: 80vw;
}
#toast.show { opacity: 1; transform: translateX(-50%) translateY(0); }
#toast.err { background: #d83931; }
```

- [ ] **Step 3: 浏览器打开冒烟验证**

Run: `npx playwright screenshot file://$(pwd)/index.html /tmp/opencode/hdshot-home.png`
Expected: 截图显示页头、虚线上传区，无报错。

- [ ] **Step 4: Commit**

```bash
git add index.html style.css
git commit -m "feat: 页面骨架与样式"
```

---

### Task 7: exporter.js — SVG+canvas 高清导出

**Files:**
- Create: `js/exporter.js`, `tests/e2e/hdshot.spec.js`, `playwright.config.js`
- Modify: `index.html`（已含 `<script src="js/exporter.js">`，无需改）

- [ ] **Step 1: 编写 e2e 失败测试（exportPng 直接调用，不依赖 main.js）**

`playwright.config.js`:

```js
const { defineConfig } = require('@playwright/test');
module.exports = defineConfig({
  testDir: './tests/e2e',
  timeout: 30000,
  use: { browserName: 'chromium' }
});
```

`tests/e2e/hdshot.spec.js`:

```js
const { test, expect } = require('@playwright/test');
const path = require('path');

const pageUrl = 'file://' + path.join(__dirname, '..', '..', 'index.html');

test('exportPng 按倍率输出矢量渲染 PNG', async ({ page }) => {
  await page.goto(pageUrl);
  const { tw, th } = await page.evaluate(() => {
    document.body.insertAdjacentHTML('beforeend',
      '<table id="t" style="border-collapse:collapse">' +
      '<tbody><tr><td style="border:1px solid #000;padding:0 4px;font-size:11pt">高清文字测试</td>' +
      '<td style="background:#4472C4;color:#fff;padding:0 4px;font-size:11pt">填充</td></tr></tbody></table>');
    const t = document.getElementById('t');
    return { tw: t.offsetWidth, th: t.offsetHeight };
  });
  const res = await page.evaluate(async () => {
    const r = await Exporter.exportPng(document.getElementById('t'),
      { scale: 3, margin: 0, background: '#ffffff', radius: 0 });
    const u8 = new Uint8Array(await r.blob.arrayBuffer());
    return { w: r.width, h: r.height, magic: Array.from(u8.slice(0, 4)), bytes: u8.length };
  });
  expect(res.magic).toEqual([0x89, 0x50, 0x4e, 0x47]);
  expect(res.w).toBe(Math.round(tw * 3));
  expect(res.h).toBe(Math.round(th * 3));
  expect(res.bytes).toBeGreaterThan(100);
});

test('exportPng 白边与倍率计入尺寸', async ({ page }) => {
  await page.goto(pageUrl);
  const res = await page.evaluate(async () => {
    document.body.insertAdjacentHTML('beforeend',
      '<table id="t2" style="border-collapse:collapse;width:100px;table-layout:fixed">' +
      '<tbody><tr><td style="padding:0 4px;font-size:11pt">a</td></tr></tbody></table>');
    const r = await Exporter.exportPng(document.getElementById('t2'),
      { scale: 2, margin: 10, background: 'transparent', radius: 0 });
    return { w: r.width, h: r.height };
  });
  expect(res.w).toBe(240); // (100 + 10*2) * 2
});
```

- [ ] **Step 2: 运行确认失败**

Run: `npx playwright test tests/e2e/hdshot.spec.js --reporter=line`
Expected: FAIL — `Exporter is not defined`

- [ ] **Step 3: 实现 exporter.js**

```js
(function (root, factory) {
  if (typeof module === 'object' && module.exports) {
    module.exports = factory();
  } else {
    root.Exporter = factory();
  }
})(typeof self !== 'undefined' ? self : this, function () {
  function buildSvgDataUrl(wrapEl, w, h) {
    const svg = '<svg xmlns="http://www.w3.org/2000/svg" width="' + w + '" height="' + h + '">' +
      '<foreignObject width="100%" height="100%">' +
      '<div xmlns="http://www.w3.org/1999/xhtml">' + wrapEl.outerHTML + '</div>' +
      '</foreignObject></svg>';
    return 'data:image/svg+xml;charset=utf-8,' + encodeURIComponent(svg);
  }

  function loadImage(src) {
    return new Promise((resolve, reject) => {
      const img = new Image();
      img.onload = () => resolve(img);
      img.onerror = () => reject(new Error('图片加载失败'));
      img.src = src;
    });
  }

  function measureAndWrap(tableEl, opts) {
    const wrap = document.createElement('div');
    wrap.style.cssText = 'display:inline-block;padding:' + opts.margin + 'px;background:' +
      opts.background + ';border-radius:' + opts.radius + 'px;overflow:hidden;' +
      'position:fixed;left:-99999px;top:0;';
    wrap.appendChild(tableEl.cloneNode(true));
    document.body.appendChild(wrap);
    const w = wrap.offsetWidth;
    const h = wrap.offsetHeight;
    return { wrap, w, h };
  }

  async function exportPng(tableEl, opts) {
    opts = Object.assign({ scale: 2, margin: 16, background: '#ffffff', radius: 0 }, opts);
    const { wrap, w, h } = measureAndWrap(tableEl, opts);
    try {
      const url = buildSvgDataUrl(wrap, w, h);
      const img = await loadImage(url);
      const canvas = document.createElement('canvas');
      canvas.width = Math.round(w * opts.scale);
      canvas.height = Math.round(h * opts.scale);
      const ctx = canvas.getContext('2d');
      ctx.drawImage(img, 0, 0, canvas.width, canvas.height);
      const blob = await new Promise((res, rej) =>
        canvas.toBlob(b => b ? res(b) : rej(new Error('导出失败')), 'image/png'));
      return { blob, width: canvas.width, height: canvas.height };
    } catch (e) {
      throw new Error('渲染失败：表格可能包含外部图片资源，无法导出');
    } finally {
      wrap.remove();
    }
  }

  return { exportPng };
});
```

- [ ] **Step 4: 运行确认通过**

Run: `npx playwright test tests/e2e/hdshot.spec.js --reporter=line`
Expected: PASS（2 个测试全过）

- [ ] **Step 5: Commit**

```bash
git add js/exporter.js tests/e2e/hdshot.spec.js playwright.config.js
git commit -m "feat: SVG foreignObject+canvas 矢量倍率导出PNG"
```

---

### Task 8: main.js — 交互编排

**Files:**
- Create: `js/main.js`
- Test: `tests/e2e/hdshot.spec.js`（追加测试）

- [ ] **Step 1: 追加 e2e 失败测试（完整上传→预览→下载链路）**

```js
const xlsx = path.join(__dirname, '..', 'fixtures', 'test.xlsx');
const csv = path.join(__dirname, '..', 'fixtures', 'test.csv');

function pngSize(buf) {
  if (buf.length < 24 || buf.readUInt32BE(0) !== 0x89504e47) throw new Error('not a png');
  return { w: buf.readUInt32BE(16), h: buf.readUInt32BE(20) };
}

test('上传xlsx→预览还原→切换sheet', async ({ page }) => {
  await page.goto(pageUrl);
  await page.setInputFiles('#file-input', xlsx);
  await expect(page.locator('#workspace')).toBeVisible();
  await expect(page.locator('#preview td').first()).toHaveText('销售数据汇总');
  await expect(page.locator('#preview tr')).toHaveCount(6);
  await page.selectOption('#sheet-select', '2');
  await expect(page.locator('#preview td').first()).toHaveText('姓名');
});

test('导出2x PNG 尺寸=预览表尺寸×2 且可下载', async ({ page }) => {
  await page.goto(pageUrl);
  await page.setInputFiles('#file-input', xlsx);
  await expect(page.locator('#preview table')).toBeVisible();
  const size = await page.evaluate(() => {
    const t = document.querySelector('#preview table');
    return { w: t.offsetWidth, h: t.offsetHeight };
  });
  const [download] = await Promise.all([
    page.waitForEvent('download'),
    page.click('#export-btn')
  ]);
  expect(download.suggestedFilename()).toMatch(/@2x\.png$/);
  const buf = fs.readFileSync(await download.path());
  const { w, h } = pngSize(buf);
  expect(Math.abs(w - (size.w + 32) * 2)).toBeLessThanOrEqual(8);
  expect(Math.abs(h - (size.h + 32) * 2)).toBeLessThanOrEqual(8);
});

test('CSV 上传渲染', async ({ page }) => {
  await page.goto(pageUrl);
  await page.setInputFiles('#file-input', csv);
  await expect(page.locator('#preview tr')).toHaveCount(3);
  await expect(page.locator('#preview td').nth(2)).toHaveText('含,逗号');
});
```

（文件顶部补充 `const fs = require('fs');`）

- [ ] **Step 2: 运行确认失败**

Run: `npx playwright test tests/e2e/hdshot.spec.js --reporter=line`
Expected: 新增 3 个测试 FAIL（上传后 workspace 不显示）

- [ ] **Step 3: 实现 main.js**

```js
(function () {
  var $ = function (id) { return document.getElementById(id); };
  var fileInput = $('file-input'), dropZone = $('drop-zone'), workspace = $('workspace');
  var sheetSelect = $('sheet-select'), scaleSelect = $('scale-select'), marginInput = $('margin-input');
  var bgSelect = $('bg-select'), radiusInput = $('radius-input'), exportBtn = $('export-btn');
  var preview = $('preview'), dimsInfo = $('dims-info'), toastEl = $('toast');

  var state = { sheets: null, tableEl: null, fileName: 'sheet', isCsv: false };

  function esc(s) {
    return String(s).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');
  }

  function toast(msg, isErr) {
    toastEl.textContent = msg;
    toastEl.className = 'show' + (isErr ? ' err' : '');
    clearTimeout(toast._t);
    toast._t = setTimeout(function () { toastEl.className = ''; }, 3200);
  }

  function currentOpts() {
    return {
      scale: +scaleSelect.value,
      margin: +marginInput.value,
      background: bgSelect.value,
      radius: +radiusInput.value
    };
  }

  function updateDims() {
    if (!state.tableEl) return;
    var o = currentOpts();
    var w = Math.round((state.tableEl.offsetWidth + o.margin * 2) * o.scale);
    var h = Math.round((state.tableEl.offsetHeight + o.margin * 2) * o.scale);
    dimsInfo.textContent = w + ' × ' + h + ' px';
  }

  function renderCurrentSheet() {
    var sheet = state.sheets[+sheetSelect.value];
    if (!sheet.rows.length) {
      toast('该工作表无数据', true);
      preview.innerHTML = '';
      state.tableEl = null;
      return;
    }
    preview.innerHTML = TableRenderer.renderSheet(sheet, { allBorders: state.isCsv });
    state.tableEl = preview.firstElementChild;
    updateDims();
  }

  async function handleFile(file) {
    if (!file) return;
    if (!/\.(xlsx|csv)$/i.test(file.name)) {
      toast('请选择 .xlsx 或 .csv 文件', true);
      return;
    }
    try {
      var buf = await file.arrayBuffer();
      var isXlsx = /\.xlsx$/i.test(file.name);
      var parsed = isXlsx ? await XlsxParser.parseXlsx(buf) : XlsxParser.parseCsv(buf, file.name);
      state.sheets = parsed.sheets;
      state.fileName = file.name.replace(/\.[^.]+$/, '');
      state.isCsv = !isXlsx;
      var totalCells = state.sheets.reduce(function (n, s) { return n + s.cellCount; }, 0);
      if (totalCells > 8000) toast('表格较大，导出可能耗时较长');
      sheetSelect.innerHTML = state.sheets.map(function (s, i) {
        return '<option value="' + i + '">' + esc(s.name) + '</option>';
      }).join('');
      sheetSelect.disabled = state.sheets.length < 2;
      workspace.hidden = false;
      renderCurrentSheet();
    } catch (e) {
      toast('解析失败：文件可能损坏或为加密文件', true);
    }
  }

  exportBtn.addEventListener('click', async function () {
    if (!state.tableEl) return;
    try {
      exportBtn.disabled = true;
      var o = currentOpts();
      var res = await Exporter.exportPng(state.tableEl, o);
      var a = document.createElement('a');
      a.href = URL.createObjectURL(res.blob);
      a.download = state.fileName + '@' + o.scale + 'x.png';
      a.click();
      setTimeout(function () { URL.revokeObjectURL(a.href); }, 5000);
      toast('已导出 ' + res.width + '×' + res.height + ' PNG');
    } catch (e) {
      toast(e.message, true);
    } finally {
      exportBtn.disabled = false;
    }
  });

  fileInput.addEventListener('change', function (e) { handleFile(e.target.files[0]); });
  dropZone.addEventListener('click', function () { fileInput.click(); });
  dropZone.addEventListener('dragover', function (e) { e.preventDefault(); dropZone.classList.add('over'); });
  dropZone.addEventListener('dragleave', function () { dropZone.classList.remove('over'); });
  dropZone.addEventListener('drop', function (e) {
    e.preventDefault();
    dropZone.classList.remove('over');
    handleFile(e.dataTransfer.files[0]);
  });
  sheetSelect.addEventListener('change', renderCurrentSheet);
  [scaleSelect, marginInput, bgSelect, radiusInput].forEach(function (el) {
    el.addEventListener('change', updateDims);
  });
})();
```

- [ ] **Step 4: 运行确认通过**

Run: `npx playwright test tests/e2e/hdshot.spec.js --reporter=line`
Expected: PASS（5 个测试全过）

- [ ] **Step 5: Commit**

```bash
git add js/main.js tests/e2e/hdshot.spec.js
git commit -m "feat: 上传/预览/sheet切换/参数调整/下载全链路"
```

---

### Task 9: 错误路径 e2e + README + 全量验证

**Files:**
- Modify: `tests/e2e/hdshot.spec.js`（追加错误路径测试）
- Create: `README.md`

- [ ] **Step 1: 追加错误路径失败测试**

```js
test('非法文件类型提示', async ({ page }) => {
  await page.goto(pageUrl);
  await page.setInputFiles('#file-input', path.join(__dirname, '..', 'fixtures', 'bad.txt'));
  await expect(page.locator('#toast')).toContainText('请选择 .xlsx 或 .csv 文件');
  await expect(page.locator('#workspace')).toBeHidden();
});

test('空表提示', async ({ page }) => {
  await page.goto(pageUrl);
  await page.setInputFiles('#file-input', xlsx);
  await page.selectOption('#sheet-select', '1');
  await expect(page.locator('#toast')).toContainText('该工作表无数据');
});
```

- [ ] **Step 2: 运行确认通过（main.js 已实现这些路径，直接应通过）**

Run: `npx playwright test tests/e2e/hdshot.spec.js --reporter=line`
Expected: PASS（7 个测试全过）；若有失败则修复 main.js 对应分支后重跑。

- [ ] **Step 3: 编写 README.md**

```markdown
# Excel 高清截图

上传 `.xlsx` / `.csv`，在浏览器中从数据矢量重绘表格，导出任意倍率（1x–4x）的高清 PNG。
解决传统屏幕截图缩小后文字退化为位图、发虚模糊的问题——文字始终以矢量字形按目标分辨率光栅化。

## 使用

直接用浏览器打开 `index.html`（无需服务器），拖入或点击选择文件，调整参数后点击"导出 PNG"。

## 开发

```bash
npm install          # 安装 exceljs(devDep 供单测) 与 @playwright/test
npm run fixtures     # 重新生成测试夹具
npm run test:unit    # Node 内置 test runner 单测
npm run test:e2e     # Playwright e2e（需 chromium）
```

## 说明

- xlsx 解析与样式还原使用 [ExcelJS](https://github.com/exceljs/exceljs)（MIT，已 vendored 到 `vendor/`）
- CSV 支持 UTF-8/GBK 自动识别、逗号/制表符分隔、引号转义
- 导出管线：DOM 表格 → SVG `foreignObject` → canvas 按倍率 `drawImage` → PNG
- 大表格（>8000 单元格）导出耗时较长属正常现象
```

- [ ] **Step 4: 全量测试**

```bash
npm run test:unit && npm run test:e2e
```
Expected: 单测 12 通过；e2e 7 通过。

- [ ] **Step 5: 手动视觉验证**

```bash
mkdir -p /tmp/opencode/hdshot && cd /tmp/opencode/hdshot
node -e "console.log('open index.html in browser for manual check')"
```

用 Playwright 截图对比：上传 test.xlsx 后对预览区截图，导出 2x PNG，目视确认中文文字清晰、合并单元格与填充色正确（可选步骤，交给用户验收）。

- [ ] **Step 6: Commit**

```bash
git add tests/e2e/hdshot.spec.js README.md
git commit -m "test: 错误路径e2e；docs: README"
```

---

## Self-Review 记录

- **Spec 覆盖**：上传xlsx/csv(Task 3,4,8)、多sheet选择(Task 8)、合并/样式还原(Task 4,5)、倍率/白边/背景/圆角(Task 7,8)、矢量导出(Task 7)、大文件警告(Task 8)、错误处理(Task 8,9)、验证(Task 2,9) ✓
- **占位符扫描**：无 TBD/TODO ✓
- **类型一致性**：`parseXlsx/parseCsv → {sheets}`、`renderSheet(sheet,{allBorders}) → string`、`exportPng(tableEl,{scale,margin,background,radius}) → {blob,width,height}` 在 parser/renderer/exporter/main/测试间一致；`cellCount`、`hidden`、`colWidths` 命名统一 ✓
