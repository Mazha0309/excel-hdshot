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
