const { test, expect } = require('@playwright/test');
const path = require('path');

const pageUrl = 'file://' + path.join(__dirname, '..', '..', 'index.html');
const fs = require('fs');
const xlsx = path.join(__dirname, '..', 'fixtures', 'test.xlsx');
const csv = path.join(__dirname, '..', 'fixtures', 'test.csv');

function pngSize(buf) {
  if (buf.length < 24 || buf.readUInt32BE(0) !== 0x89504e47) throw new Error('not a png');
  return { w: buf.readUInt32BE(16), h: buf.readUInt32BE(20) };
}

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

test('exportPng 内容真实渲染（非空白且铺满画布）', async ({ page }) => {
  await page.goto(pageUrl);
  const stats = await page.evaluate(async () => {
    document.body.insertAdjacentHTML('beforeend',
      '<table id="t3" style="border-collapse:collapse;width:120px;table-layout:fixed;height:30px;background:#fff;border:1px solid #000">' +
      '<tbody><tr><td style="background:#000;padding:4px;font-size:11pt">X</td>' +
      '<td style="padding:4px;font-size:11pt">Y</td></tr></tbody></table>');
    const r = await Exporter.exportPng(document.getElementById('t3'),
      { scale: 3, margin: 0, background: '#ffffff', radius: 0 });
    const canvas = document.createElement('canvas');
    canvas.width = r.width; canvas.height = r.height;
    const ctx = canvas.getContext('2d');
    const bmp = await createImageBitmap(r.blob);
    ctx.drawImage(bmp, 0, 0);
    const d = ctx.getImageData(0, 0, canvas.width, canvas.height).data;
    let minX = canvas.width, minY = canvas.height, maxX = -1, maxY = -1, content = 0, white = 0;
    for (let y = 0; y < canvas.height; y++) {
      for (let x = 0; x < canvas.width; x++) {
        const i = (y * canvas.width + x) * 4;
        if (d[i] < 250 || d[i+1] < 250 || d[i+2] < 250) { content++; if (x < minX) minX = x; if (x > maxX) maxX = x; if (y < minY) minY = y; if (y > maxY) maxY = y; }
        else if (d[i+3] === 255) white++;
      }
    }
    return { content, white, total: canvas.width * canvas.height,
      fillW: (maxX - minX + 1) / canvas.width, fillH: (maxY - minY + 1) / canvas.height };
  });
  expect(stats.content).toBeGreaterThan(0);
  expect(stats.white).toBeGreaterThan(0);
  expect(stats.fillW).toBeGreaterThan(0.9);
  expect(stats.fillH).toBeGreaterThan(0.9);
});

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
  await expect(page.locator('#preview td').nth(5)).toHaveText('含,逗号');
});

test('点击上传区打开文件选择器并上传', async ({ page }) => {
  await page.goto(pageUrl);
  const [chooser] = await Promise.all([
    page.waitForEvent('filechooser'),
    page.click('#drop-zone')
  ]);
  await chooser.setFiles(xlsx);
  await expect(page.locator('#preview td').first()).toHaveText('销售数据汇总');
});

test('非法文件类型提示且不显示工作区', async ({ page }) => {
  await page.goto(pageUrl);
  await page.setInputFiles('#file-input', path.join(__dirname, '..', 'fixtures', 'bad.txt'));
  await expect(page.locator('#toast')).toContainText('请选择 .xlsx 或 .csv 文件');
  await expect(page.locator('#workspace')).toBeHidden();
});

test('损坏xlsx解析失败提示', async ({ page }) => {
  const badPath = path.join(__dirname, '..', 'fixtures', 'corrupt.xlsx');
  fs.writeFileSync(badPath, 'this is not a valid xlsx file');
  try {
    await page.goto(pageUrl);
    await page.setInputFiles('#file-input', badPath);
    await expect(page.locator('#toast')).toContainText('解析失败');
  } finally {
    fs.unlinkSync(badPath);
  }
});

test('空表提示', async ({ page }) => {
  await page.goto(pageUrl);
  await page.setInputFiles('#file-input', xlsx);
  await page.selectOption('#sheet-select', '1');
  await expect(page.locator('#toast')).toContainText('该工作表无数据');
  await expect(page.locator('#dims-info')).toHaveText('');
});

test('单sheet文件禁用工作表选择', async ({ page }) => {
  await page.goto(pageUrl);
  await page.setInputFiles('#file-input', csv);
  await expect(page.locator('#sheet-select')).toBeDisabled();
});

test('无表格时导出按钮提示', async ({ page }) => {
  await page.goto(pageUrl);
  await page.setInputFiles('#file-input', xlsx);
  await page.selectOption('#sheet-select', '1');
  await page.click('#export-btn');
  await expect(page.locator('#toast')).toContainText('请先上传并选择有数据的表格');
});

test('真实拖拽上传触发预览', async ({ page }) => {
  await page.goto(pageUrl);
  const dataTransfer = await page.evaluateHandle(() => new DataTransfer());
  const bytes = fs.readFileSync(xlsx);
  const filePayload = await page.evaluateHandle(
    ({ b, name }) => new File([new Uint8Array(b)], name, {
      type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
    }),
    { b: Array.from(bytes), name: 'test.xlsx' }
  );
  await dataTransfer.evaluate((dt, file) => dt.items.add(file), filePayload);
  await page.dispatchEvent('#drop-zone', 'drop', { dataTransfer });
  await expect(page.locator('#preview td').first()).toHaveText('销售数据汇总');
});

test('主题色标题背景还原（theme+tint）', async ({ page }) => {
  await page.goto(pageUrl);
  await page.setInputFiles('#file-input', path.join(__dirname, '..', 'fixtures', 'theme.xlsx'));
  await expect(page.locator('#preview td').first()).toHaveText('主题色标题');
  const style = await page.locator('#preview td').first().getAttribute('style');
  expect(style).toContain('background:#8FAADC');
});

test('预览展示白边/背景/圆角（所见即所得）', async ({ page }) => {
  await page.goto(pageUrl);
  await page.setInputFiles('#file-input', xlsx);
  await page.fill('#margin-input', '24');
  await page.locator('#margin-input').dispatchEvent('change');
  await page.selectOption('#bg-select', '#ffffff');
  await page.fill('#radius-input', '12');
  await page.locator('#radius-input').dispatchEvent('change');
  const style = await page.locator('#preview .hs-wrap').getAttribute('style');
  expect(style).toContain('padding:24px');
  expect(style).toContain('border-radius:12px');
  const dims = await page.locator('#dims-info').textContent();
  expect(dims).not.toBe('');
});

test('点击单元格编辑列宽', async ({ page }) => {
  await page.goto(pageUrl);
  await page.setInputFiles('#file-input', xlsx);
  await page.locator('#preview td[data-r="2"][data-c="1"]').click();
  await expect(page.locator('#cell-edit')).toBeVisible();
  await expect(page.locator('#col-label')).toHaveText('B');
  await expect(page.locator('#row-label')).toHaveText('3');
  await page.fill('#col-width-input', '200');
  await page.locator('#col-width-input').dispatchEvent('change');
  const colStyle = await page.locator('#preview col').nth(1).getAttribute('style');
  expect(colStyle).toContain('width:200px');
  const selStillThere = await page.locator('#preview td.hs-selected').count();
  expect(selStillThere).toBe(1);
});

test('点击单元格编辑行高', async ({ page }) => {
  await page.goto(pageUrl);
  await page.setInputFiles('#file-input', xlsx);
  await page.locator('#preview td[data-r="0"][data-c="0"]').click();
  await page.fill('#row-height-input', '60');
  await page.locator('#row-height-input').dispatchEvent('change');
  const trStyle = await page.locator('#preview tr').first().getAttribute('style');
  expect(trStyle).toContain('height:60px');
});

test('CSV上传后编辑列宽自动回填默认', async ({ page }) => {
  await page.goto(pageUrl);
  await page.setInputFiles('#file-input', csv);
  await page.locator('#preview td[data-c="0"]').first().click();
  await page.fill('#col-width-input', '150');
  await page.locator('#col-width-input').dispatchEvent('change');
  const colStyle = await page.locator('#preview col').first().getAttribute('style');
  expect(colStyle).toContain('width:150px');
});
