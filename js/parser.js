(function (root, factory) {
  if (typeof module === 'object' && module.exports) {
    module.exports = factory(require('exceljs'));
  } else {
    root.XlsxParser = factory(root.ExcelJS);
  }
})(typeof self !== 'undefined' ? self : this, function (ExcelJS) {
  const DEFAULT_FONT = { name: 'Calibri', size: 11, bold: false, italic: false, underline: false, color: null };

  function decodeBuffer(buf) {
    try {
      return new TextDecoder('utf-8', { fatal: true }).decode(buf);
    } catch (e) {
      return new TextDecoder('gbk').decode(buf);
    }
  }

  function detectDelimiter(text) {
    const nl = text.indexOf('\n');
    const firstLine = text.slice(0, nl === -1 ? text.length : nl);
    const tabs = (firstLine.match(/\t/g) || []).length;
    const commas = (firstLine.match(/,/g) || []).length;
    return tabs > commas ? '\t' : ',';
  }

  function parseCsvText(text) {
    const delim = detectDelimiter(text);
    const rows = [];
    let row = [], field = '', inQuotes = false;
    for (let i = 0; i < text.length; i++) {
      const ch = text[i];
      if (inQuotes) {
        if (ch === '"') {
          if (text[i + 1] === '"') { field += '"'; i++; } else { inQuotes = false; }
        } else {
          field += ch;
        }
      } else if (ch === '"') {
        inQuotes = true;
      } else if (ch === delim) {
        row.push(field); field = '';
      } else if (ch === '\n') {
        row.push(field); rows.push(row); row = []; field = '';
      } else if (ch !== '\r') {
        field += ch;
      }
    }
    if (field !== '' || row.length) { row.push(field); rows.push(row); }
    return rows;
  }

  function makeCell(text) {
    return {
      text: String(text), rowspan: 1, colspan: 1, hidden: false,
      font: Object.assign({}, DEFAULT_FONT), fill: null, align: null, border: null
    };
  }

  function parseCsv(buf, fileName) {
    const text = decodeBuffer(buf).replace(/^\ufeff/, '');
    const rows = parseCsvText(text).map(r => ({ height: null, cells: r.map(makeCell) }));
    const cellCount = rows.reduce((n, r) => n + r.cells.length, 0);
    return {
      sheets: [{ name: fileName.replace(/\.[^.]+$/, ''), rows, colWidths: [], cellCount }]
    };
  }

  function parseXlsx() {
    throw new Error('parseXlsx 未实现');
  }

  return { parseXlsx, parseCsv };
});
