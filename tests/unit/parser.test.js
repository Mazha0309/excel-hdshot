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
