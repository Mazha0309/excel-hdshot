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