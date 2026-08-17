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