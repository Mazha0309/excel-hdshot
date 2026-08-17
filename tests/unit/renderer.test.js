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