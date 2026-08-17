const test = require('node:test');
const assert = require('node:assert');
const fs = require('fs');
const path = require('path');
const parser = require('../../js/parser.js');
const ExcelJS = require('exceljs');

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

test('parseCsv: 字段中间的引号按字面量处理（不吞行）', () => {
  const buf = new TextEncoder().encode('a"b,c\nd,e\n');
  const { sheets } = parser.parseCsv(buf, 'q.csv');
  assert.strictEqual(sheets[0].rows.length, 2);
  assert.deepStrictEqual(sheets[0].rows[0].cells.map(c => c.text), ['a"b', 'c']);
  assert.deepStrictEqual(sheets[0].rows[1].cells.map(c => c.text), ['d', 'e']);
});

test('parseCsv: 仅一个空引号字段的文件产生1行1空单元格', () => {
  const buf = new TextEncoder().encode('""');
  const { sheets } = parser.parseCsv(buf, 'e.csv');
  assert.strictEqual(sheets[0].rows.length, 1);
  assert.strictEqual(sheets[0].rows[0].cells.length, 1);
  assert.strictEqual(sheets[0].rows[0].cells[0].text, '');
});

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
  assert.strictEqual(s.rows[0].height, 40);
  assert.strictEqual(s.colWidths[0], 182);
  assert.strictEqual(s.colWidths[3], 98);
});

test('parseXlsx: 列间空档不丢数据', async () => {
  const wb = new ExcelJS.Workbook();
  const ws = wb.addWorksheet('gap');
  ws.getCell('A1').value = 'left';
  ws.getCell('Z1').value = 'far right';
  const buf = await wb.xlsx.writeBuffer();
  const { sheets } = await parser.parseXlsx(buf);
  assert.strictEqual(sheets[0].rows[0].cells[0].text, 'left');
  assert.strictEqual(sheets[0].rows[0].cells[25].text, 'far right');
});

test('parseXlsx: 公式/富文本/超链接取值', async () => {
  const wb = new ExcelJS.Workbook();
  const ws = wb.addWorksheet('objects');
  ws.getCell('A1').value = { formula: 'B1+C1', result: 7 };
  ws.getCell('B1').value = 3;
  ws.getCell('C1').value = 4;
  ws.getCell('A2').value = { richText: [{ text: '红' }, { text: '蓝' }] };
  ws.getCell('A3').value = { text: '链接', hyperlink: 'https://example.com' };
  const buf = await wb.xlsx.writeBuffer();
  const { sheets } = await parser.parseXlsx(buf);
  assert.strictEqual(sheets[0].rows[0].cells[0].text, '7');
  assert.strictEqual(sheets[0].rows[1].cells[0].text, '红蓝');
  assert.strictEqual(sheets[0].rows[2].cells[0].text, '链接');
});

test('parseXlsx: 日期时间与12小时制（UTC无关）', async () => {
  const wb = new ExcelJS.Workbook();
  const ws = wb.addWorksheet('dt');
  ws.getCell('A1').value = new Date(Date.UTC(2026, 7, 1, 22, 30, 0));
  ws.getCell('A1').numFmt = 'yyyy-mm-dd hh:mm:ss';
  ws.getCell('A2').value = new Date(Date.UTC(2026, 7, 1, 22, 0, 0));
  ws.getCell('A2').numFmt = 'h:mm AM/PM';
  ws.getCell('A3').value = new Date(Date.UTC(2026, 7, 1, 0, 30, 45));
  ws.getCell('A3').numFmt = 'mm:ss';
  const buf = await wb.xlsx.writeBuffer();
  const { sheets } = await parser.parseXlsx(buf);
  assert.strictEqual(sheets[0].rows[0].cells[0].text, '2026-08-01 22:30:00');
  assert.strictEqual(sheets[0].rows[1].cells[0].text, '10:00 PM');
  assert.strictEqual(sheets[0].rows[2].cells[0].text, '30:45');
});

test('parseXlsx: 合并覆盖空单元格为hidden对象，非合并空格为null', async () => {
  const wb = new ExcelJS.Workbook();
  const ws = wb.addWorksheet('t');
  ws.mergeCells('A1:B1');
  ws.getCell('A1').value = 'M';
  ws.getCell('C1').value = 'X';
  ws.getCell('A2').value = 'a';
  ws.getCell('C2').value = 'c';
  const buf = await wb.xlsx.writeBuffer();
  const { sheets } = await parser.parseXlsx(buf);
  const r1 = sheets[0].rows[0].cells;
  assert.strictEqual(r1[0].text, 'M');
  assert.strictEqual(r1[0].colspan, 2);
  assert.strictEqual(r1[1].hidden, true);
  assert.strictEqual(r1[2].text, 'X');
  const r2 = sheets[0].rows[1].cells;
  assert.strictEqual(r2[0].text, 'a');
  assert.strictEqual(r2[1], null);
  assert.strictEqual(r2[2].text, 'c');
});

test('parseXlsx: 空白合并主单元格保留colspan', async () => {
  const wb = new ExcelJS.Workbook();
  const ws = wb.addWorksheet('t');
  ws.mergeCells('A1:B1');
  ws.getCell('C1').value = 'X';
  ws.getCell('A2').value = 'a';
  const buf = await wb.xlsx.writeBuffer();
  const { sheets } = await parser.parseXlsx(buf);
  const r1 = sheets[0].rows[0].cells;
  assert.strictEqual(r1[0].colspan, 2);
  assert.strictEqual(r1[0].hidden, false);
  assert.strictEqual(r1[0].text, '');
  assert.strictEqual(r1[1].hidden, true);
  assert.strictEqual(r1[2].text, 'X');
});
