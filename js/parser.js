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
    let row = [], field = '', inQuotes = false, content = false;
    for (let i = 0; i < text.length; i++) {
      const ch = text[i];
      if (inQuotes) {
        if (ch === '"') {
          if (text[i + 1] === '"') { field += '"'; i++; } else { inQuotes = false; }
        } else {
          field += ch; content = true;
        }
      } else if (ch === '"' && field === '') {
        inQuotes = true; content = true;
      } else if (ch === delim) {
        row.push(field); field = '';
      } else if (ch === '\n') {
        row.push(field); rows.push(row); row = []; field = ''; content = false;
      } else if (ch !== '\r') {
        field += ch; content = true;
      }
    }
    if (content || field !== '' || row.length) { row.push(field); rows.push(row); }
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

  function colToIndex(ref) {
    let n = 0;
    for (const ch of ref) n = n * 26 + ch.charCodeAt(0) - 64;
    return n - 1;
  }

  function argbToCss(argb) {
    if (!argb || argb === 'FF000000') return null;
    if (argb.length === 8 && argb.startsWith('FF')) return '#' + argb.slice(2).toUpperCase();
    return null;
  }

  function normalizeFont(f) {
    if (!f) return Object.assign({}, DEFAULT_FONT);
    return {
      name: f.name || 'Calibri',
      size: f.size || 11,
      bold: !!f.bold,
      italic: !!f.italic,
      underline: !!f.underline,
      color: argbToCss(f.color && f.color.argb)
    };
  }

  function normalizeFill(f) {
    if (!f || f.type !== 'pattern') return null;
    const c = argbToCss(f.fgColor && f.fgColor.argb);
    return c ? { color: c } : null;
  }

  function side(s) {
    if (!s || !s.style) return null;
    return { style: s.style, color: argbToCss(s.color && s.color.argb) };
  }

  function normalizeBorder(b) {
    if (!b) return null;
    const out = { top: side(b.top), bottom: side(b.bottom), left: side(b.left), right: side(b.right) };
    return (out.top || out.bottom || out.left || out.right) ? out : null;
  }

  function normalizeAlign(a) {
    if (!a) return null;
    const h = { left: 'left', center: 'center', right: 'right', justify: 'justify' }[a.horizontal] || null;
    const v = { top: 'top', middle: 'middle', bottom: 'bottom' }[a.vertical] || null;
    return (h || v || a.wrapText) ? { h, v, wrap: !!a.wrapText } : null;
  }

  function hasContent(cell, st) {
    if (cell.value !== null && cell.value !== undefined) return true;
    if (st.fill) return true;
    if (st.border && (st.border.top || st.border.bottom || st.border.left || st.border.right)) return true;
    if (st.alignment) return true;
    if (st.font) {
      const f = st.font;
      if (f.bold || f.italic || f.underline || f.name || f.size || f.color) return true;
    }
    return false;
  }

  const MONTHS = ['January', 'February', 'March', 'April', 'May', 'June', 'July', 'August', 'September', 'October', 'November', 'December'];
  const DAYS = ['Sunday', 'Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday', 'Saturday'];

  function pad2(n) { return String(n).padStart(2, '0'); }

  function formatDateValue(d, fmt) {
    const hasTime = /(h+|s+)/i.test(fmt);
    const tokens = [
      ['yyyy', () => String(d.getFullYear())],
      ['yy', () => String(d.getFullYear()).slice(-2)],
      ['mmmm', () => MONTHS[d.getMonth()]],
      ['mmm', () => MONTHS[d.getMonth()].slice(0, 3)],
      ['mm', () => pad2(hasTime ? d.getMinutes() : d.getMonth() + 1)],
      ['m', () => hasTime ? d.getMinutes() : d.getMonth() + 1],
      ['dddd', () => DAYS[d.getDay()]],
      ['ddd', () => DAYS[d.getDay()].slice(0, 3)],
      ['dd', () => pad2(d.getDate())],
      ['d', () => d.getDate()],
      ['hh', () => pad2(d.getHours())],
      ['h', () => d.getHours()],
      ['ss', () => pad2(d.getSeconds())],
      ['s', () => d.getSeconds()],
      ['AM/PM', () => d.getHours() < 12 ? 'AM' : 'PM'],
      ['A/P', () => d.getHours() < 12 ? 'A' : 'P']
    ];
    let out = '';
    let i = 0;
    while (i < fmt.length) {
      const ch = fmt[i];
      if (ch === '"') {
        const end = fmt.indexOf('"', i + 1);
        if (end > i) { out += fmt.slice(i + 1, end); i = end + 1; continue; }
        out += ch; i++; continue;
      }
      if (ch === '\\') {
        if (i + 1 < fmt.length) out += fmt[i + 1];
        i += 2; continue;
      }
      let matched = false;
      for (const [tok, fn] of tokens) {
        if (fmt.startsWith(tok, i)) { out += fn(); i += tok.length; matched = true; break; }
      }
      if (!matched) { out += ch; i++; }
    }
    return out;
  }

  function formatNumberValue(v, fmt) {
    if (fmt == null || fmt === '' || fmt === 'General') return String(v);
    const sections = String(fmt).split(';');
    const isNeg = v < 0;
    const section = isNeg ? (sections[1] || sections[0]) : sections[0];
    const abs = Math.abs(v);
    const pct = section.includes('%');
    const pattern = section.replace(/%/g, '');
    const dot = pattern.indexOf('.');
    const intPat = dot >= 0 ? pattern.slice(0, dot) : pattern;
    const decPat = dot >= 0 ? pattern.slice(dot + 1) : '';
    const grouping = intPat.includes(',');
    const decPlaces = (decPat.match(/0/g) || []).length;
    const num = pct ? abs * 100 : abs;
    let str = num.toFixed(decPlaces);
    if (grouping) {
      const parts = str.split('.');
      str = parts[0].replace(/\B(?=(\d{3})+(?!\d))/g, ',') + (parts[1] !== undefined ? '.' + parts[1] : '');
    }
    if (pct) str += '%';
    return (isNeg ? '-' : '') + str;
  }

  function formatCellValue(v, numFmt) {
    if (v === null || v === undefined) return '';
    if (v instanceof Date) return formatDateValue(v, numFmt || 'yyyy-mm-dd');
    if (typeof v === 'number') return formatNumberValue(v, numFmt);
    if (Array.isArray(v)) return v.map(r => (r && r.text) || '').join('');
    if (v && typeof v === 'object' && v.error) return v.error;
    return String(v);
  }

  function textOf(cell) {
    return formatCellValue(cell.value, cell.numFmt);
  }

  async function parseXlsx(buf) {
    const wb = new ExcelJS.Workbook();
    await wb.xlsx.load(buf);
    const sheets = wb.worksheets.map(ws => {
      const merges = [];
      for (const spec of (ws.model.merges || [])) {
        const m = /^([A-Z]+)(\d+):([A-Z]+)(\d+)$/.exec(spec);
        if (!m) continue;
        merges.push({ r1: +m[2] - 1, c1: colToIndex(m[1]), r2: +m[4] - 1, c2: colToIndex(m[3]) });
      }
      const rowCount = ws.actualRowCount;
      const colCount = ws.actualColumnCount;
      const rows = [];
      let cellCount = 0;
      for (let r = 1; r <= rowCount; r++) {
        const xrow = ws.getRow(r);
        const cells = [];
        for (let c = 1; c <= colCount; c++) {
          const xc = xrow.getCell(c);
          const st = xc.style || {};
          if (!hasContent(xc, st)) { cells.push(null); continue; }
          const ri = r - 1, ci = c - 1;
          const master = merges.find(mm => mm.r1 === ri && mm.c1 === ci);
          const covered = merges.some(mm => mm.r1 <= ri && mm.c1 <= ci && mm.r2 >= ri && mm.c2 >= ci && !(mm.r1 === ri && mm.c1 === ci));
          cells.push({
            text: textOf(xc),
            rowspan: master ? master.r2 - master.r1 + 1 : 1,
            colspan: master ? master.c2 - master.c1 + 1 : 1,
            hidden: covered,
            font: normalizeFont(st.font),
            fill: normalizeFill(st.fill),
            align: normalizeAlign(st.alignment),
            border: normalizeBorder(st.border)
          });
          cellCount++;
        }
        const h = xrow.height;
        rows.push({ height: h ? Math.round(h * 1.3333) : null, cells });
      }
      const colWidths = [];
      for (let c = 1; c <= colCount; c++) {
        const w = ws.getColumn(c).width;
        colWidths.push(w ? Math.round(w * 7) : null);
      }
      return { name: ws.name, rows, colWidths, cellCount };
    });
    return { sheets };
  }

  return { parseXlsx, parseCsv };
});
