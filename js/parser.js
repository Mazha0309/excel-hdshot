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

  const STANDARD_THEME = [
    '000000', 'FFFFFF', '44546A', 'E7E6E6',
    '4472C4', 'ED7D31', 'A5A5A5', 'FFC000',
    '5B9BD5', '70AD47', '0563C1', '954F72'
  ];
  const INDEXED = [
    '000000','FFFFFF','FF0000','00FF00','0000FF','FFFF00','FF00FF','00FFFF',
    '000000','FFFFFF','FF0000','00FF00','0000FF','FFFF00','FF00FF','00FFFF',
    '800000','008000','000080','808000','800080','008080','C0C0C0','808080',
    '9999FF','993366','FFFFCC','CCFFFF','660066','FF8080','0066CC','CCCCFF',
    '000080','FF00FF','FFFF00','00FFFF','800080','800000','008080','0000FF',
    '00CCFF','CCFFFF','CCFFCC','FFFF99','99CCFF','FF99CC','CC99FF','FFCC99',
    '3366FF','33CCCC','99CC00','FFCC00','FF9900','FF6600','666699','969696',
    '003366','339966','003300','333300','993300','993366','333399','333333'
  ];

  function parseThemeXml(xml) {
    if (!xml) return null;
    const names = ['dk1', 'lt1', 'dk2', 'lt2', 'accent1', 'accent2', 'accent3', 'accent4', 'accent5', 'accent6', 'hlink', 'folHlink'];
    const scheme = [];
    for (const name of names) {
      const re = new RegExp('<a:' + name + '><a:(?:srgbClr val="([0-9A-Fa-f]{6})"|sysClr[^>]*?lastClr="([0-9A-Fa-f]{6})")', 'i');
      const m = re.exec(xml);
      scheme.push(m ? (m[1] || m[2] || null) : null);
    }
    return scheme;
  }

  function applyTint(hex, tint) {
    if (!tint) return hex;
    const chan = (v, f) => tint < 0 ? Math.round(v * f) : Math.round(v + (255 - v) * tint);
    const f = 1 + tint;
    const r = chan(parseInt(hex.slice(0, 2), 16), f);
    const g = chan(parseInt(hex.slice(2, 4), 16), f);
    const b = chan(parseInt(hex.slice(4, 6), 16), f);
    return [r, g, b].map(n => n.toString(16).padStart(2, '0').toUpperCase()).join('');
  }

  function resolveColor(c, theme) {
    if (!c) return null;
    if (c.argb) return argbToCss(c.argb);
    if (typeof c.theme === 'number' && c.theme >= 0 && c.theme <= 11) {
      const base = (theme && theme[c.theme]) || STANDARD_THEME[c.theme] || null;
      return base ? '#' + applyTint(base, c.tint || 0) : null;
    }
    if (typeof c.indexed === 'number' && c.indexed >= 0 && c.indexed <= 63) {
      return '#' + INDEXED[c.indexed];
    }
    return null;
  }

  function normalizeFont(f, theme) {
    if (!f) return Object.assign({}, DEFAULT_FONT);
    return {
      name: f.name || 'Calibri',
      size: f.size || 11,
      bold: !!f.bold,
      italic: !!f.italic,
      underline: !!f.underline,
      color: resolveColor(f.color, theme)
    };
  }

  function normalizeFill(f, theme) {
    if (!f || f.type !== 'pattern') return null;
    const c = resolveColor(f.fgColor, theme);
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

  function isMinute(fmt, i) {
    let j = i - 1;
    while (j >= 0 && fmt[j] === ':') j--;
    if (j >= 0 && fmt[j] === 'h') return true;
    let k = i + 1;
    while (k < fmt.length && fmt[k] === ':') k++;
    if (k < fmt.length && fmt[k] === 's') return true;
    return !/[yd]/i.test(fmt);
  }

  function formatDateValue(d, fmt) {
    const is12h = fmt.includes('AM/PM') || fmt.includes('A/P');
    const hr = d.getUTCHours();
    const hour = () => is12h ? (hr % 12 || 12) : hr;
    const tokens = [
      ['yyyy', () => String(d.getUTCFullYear())],
      ['yy', () => String(d.getUTCFullYear()).slice(-2)],
      ['mmmm', () => MONTHS[d.getUTCMonth()]],
      ['mmm', () => MONTHS[d.getUTCMonth()].slice(0, 3)],
      ['mm', (i) => pad2(isMinute(fmt, i) ? d.getUTCMinutes() : d.getUTCMonth() + 1)],
      ['m', (i) => isMinute(fmt, i) ? d.getUTCMinutes() : d.getUTCMonth() + 1],
      ['dddd', () => DAYS[d.getUTCDay()]],
      ['ddd', () => DAYS[d.getUTCDay()].slice(0, 3)],
      ['dd', () => pad2(d.getUTCDate())],
      ['d', () => d.getUTCDate()],
      ['hh', () => pad2(hour())],
      ['h', () => hour()],
      ['ss', () => pad2(d.getUTCSeconds())],
      ['s', () => d.getUTCSeconds()],
      ['AM/PM', () => hr < 12 ? 'AM' : 'PM'],
      ['A/P', () => hr < 12 ? 'A' : 'P']
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
        if (fmt.startsWith(tok, i)) { out += fn(i); i += tok.length; matched = true; break; }
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
    const decPlaces = (decPat.match(/[0#]/g) || []).length;
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
    if (v && typeof v === 'object') {
      if (v.error) return v.error;
      if (v.result !== undefined && v.result !== null) return formatCellValue(v.result, numFmt);
      if (Array.isArray(v.richText)) return v.richText.map(r => (r && r.text) || '').join('');
      if (v.text !== undefined) return String(v.text);
      return String(v);
    }
    return String(v);
  }

  function textOf(cell) {
    return formatCellValue(cell.value, cell.numFmt);
  }

  async function parseXlsx(buf) {
    const wb = new ExcelJS.Workbook();
    await wb.xlsx.load(buf);
    const themeKeys = wb.model.themes ? Object.keys(wb.model.themes) : [];
    const theme = parseThemeXml(themeKeys.length ? wb.model.themes[themeKeys[0]] : '');
    const sheets = wb.worksheets.map(ws => {
      const merges = [];
      for (const spec of (ws.model.merges || [])) {
        const m = /^([A-Z]+)(\d+):([A-Z]+)(\d+)$/.exec(spec);
        if (!m) continue;
        merges.push({ r1: +m[2] - 1, c1: colToIndex(m[1]), r2: +m[4] - 1, c2: colToIndex(m[3]) });
      }
      const rowCount = ws.actualRowCount;
      const colCount = ws.columnCount;
      const rows = [];
      let cellCount = 0;
      for (let r = 1; r <= rowCount; r++) {
        const xrow = ws.getRow(r);
        const cells = [];
        for (let c = 1; c <= colCount; c++) {
          const xc = xrow.getCell(c);
          const st = xc.style || {};
          const ri = r - 1, ci = c - 1;
          const master = merges.find(mm => mm.r1 === ri && mm.c1 === ci);
          const covered = merges.some(mm => mm.r1 <= ri && mm.c1 <= ci && mm.r2 >= ri && mm.c2 >= ci && !(mm.r1 === ri && mm.c1 === ci));
          if (!hasContent(xc, st) && !covered && !master) { cells.push(null); continue; }
          cells.push({
            text: textOf(xc),
            rowspan: master ? master.r2 - master.r1 + 1 : 1,
            colspan: master ? master.c2 - master.c1 + 1 : 1,
            hidden: covered,
            font: normalizeFont(st.font, theme),
            fill: normalizeFill(st.fill, theme),
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
        colWidths.push(w ? Math.round(w * 7) : 59);
      }
      return { name: ws.name, rows, colWidths, cellCount };
    });
    return { sheets };
  }

  return { parseXlsx, parseCsv };
});
