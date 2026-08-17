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

test('空列占位保留列对齐（null→空td，hidden→跳过）', () => {
  const sparseSheet = {
    name: 's',
    rows: [{
      height: null,
      cells: [
        { text: 'A', rowspan: 1, colspan: 1, hidden: false, font: null, fill: null, align: null, border: null },
        null,
        { text: 'B', rowspan: 1, colspan: 1, hidden: false, font: null, fill: null, align: null, border: null }
      ]
    }],
    colWidths: [100, 100, 100],
    cellCount: 2
  };
  const html = renderer.renderSheet(sparseSheet);
  const rowHtml = html.slice(html.indexOf('<tr'), html.indexOf('</tr>') + 5);
  assert.strictEqual((rowHtml.match(/<td/g) || []).length, 3);
  assert.ok(rowHtml.includes('>A</td>'));
  assert.ok(rowHtml.includes('></td>'));
  assert.ok(rowHtml.includes('>B</td>'));
});

test('rowspan 属性输出', () => {
  const rsSheet = {
    name: 'r',
    rows: [{ height: null, cells: [{ text: 'v', rowspan: 2, colspan: 1, hidden: false, font: null, fill: null, align: null, border: null }] }],
    colWidths: [],
    cellCount: 1
  };
  const html = renderer.renderSheet(rsSheet);
  assert.ok(html.includes('rowspan="2"'));
});

test('溢出裁剪与换行保留换行符', () => {
  const wrapSheet = {
    name: 'w',
    rows: [{ height: null, cells: [{ text: 'a\nb', rowspan: 1, colspan: 1, hidden: false, font: null, fill: null, align: { h: null, v: null, wrap: true }, border: null }] }],
    colWidths: [100],
    cellCount: 1
  };
  const html = renderer.renderSheet(wrapSheet);
  assert.ok(html.includes('overflow:hidden'));
  assert.ok(html.includes('white-space:pre-wrap'));
});

test('字体名引号消毒', () => {
  const fSheet = {
    name: 'f',
    rows: [{ height: null, cells: [{ text: 'x', rowspan: 1, colspan: 1, hidden: false, font: { name: 'Bad\'"Name', size: 11, bold: false, italic: false, underline: false, color: null }, fill: null, align: null, border: null }] }],
    colWidths: [],
    cellCount: 1
  };
  const html = renderer.renderSheet(fSheet);
  assert.ok(html.includes("font-family:'BadName',sans-serif"));
});

test('单元格输出 data-r/data-c 定位属性', () => {
  const html = renderer.renderSheet(sheet);
  assert.ok(html.includes('data-r="0" data-c="0"'));
});

test('空列占位格输出 data-c', () => {
  const sparseSheet = {
    name: 's',
    rows: [{ height: null, cells: [
      { text: 'A', rowspan: 1, colspan: 1, hidden: false, font: null, fill: null, align: null, border: null },
      null,
      { text: 'B', rowspan: 1, colspan: 1, hidden: false, font: null, fill: null, align: null, border: null }
    ] }],
    colWidths: [100, 100, 100],
    cellCount: 2
  };
  const html = renderer.renderSheet(sparseSheet);
  assert.ok(html.includes('<td data-c="1"></td>'));
});