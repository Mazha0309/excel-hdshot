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
