(function (root, factory) {
  if (typeof module === 'object' && module.exports) {
    module.exports = factory();
  } else {
    root.Exporter = factory();
  }
})(typeof self !== 'undefined' ? self : this, function () {
  function buildSvgXml(el, w, h) {
    const xmlns = 'http://www.w3.org/2000/svg';
    const xhtml = 'http://www.w3.org/1999/xhtml';
    const svg = document.createElementNS(xmlns, 'svg');
    svg.setAttribute('xmlns', xmlns);
    svg.setAttribute('width', w);
    svg.setAttribute('height', h);
    const fo = document.createElementNS(xmlns, 'foreignObject');
    fo.setAttribute('width', '100%');
    fo.setAttribute('height', '100%');
    const body = document.createElementNS(xhtml, 'div');
    body.setAttribute('xmlns', xhtml);
    body.appendChild(el.cloneNode(true));
    fo.appendChild(body);
    svg.appendChild(fo);
    return new XMLSerializer().serializeToString(svg);
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
      if (!w || !h) throw new Error('表格无内容');
      const xml = buildSvgXml(wrap, w, h);
      const url = 'data:image/svg+xml;charset=utf-8,' + encodeURIComponent(xml);
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