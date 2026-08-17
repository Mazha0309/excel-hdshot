(function () {
  var $ = function (id) { return document.getElementById(id); };
  var fileInput = $('file-input'), dropZone = $('drop-zone'), workspace = $('workspace');
  var sheetSelect = $('sheet-select'), scaleSelect = $('scale-select'), marginInput = $('margin-input');
  var bgSelect = $('bg-select'), radiusInput = $('radius-input'), exportBtn = $('export-btn');
  var preview = $('preview'), dimsInfo = $('dims-info'), toastEl = $('toast');
  var colLabel = $('col-label'), rowLabel = $('row-label');
  var colWidthInput = $('col-width-input'), rowHeightInput = $('row-height-input');
  var cellEdit = $('cell-edit');

  var state = { sheets: null, tableEl: null, fileName: 'sheet', isCsv: false, sel: null };

  function columnName(i) {
    var s = '';
    i++;
    while (i > 0) {
      s = String.fromCharCode(65 + ((i - 1) % 26)) + s;
      i = Math.floor((i - 1) / 26);
    }
    return s;
  }

  function esc(s) {
    return String(s).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');
  }

  function toast(msg, isErr) {
    toastEl.textContent = msg;
    toastEl.className = 'show' + (isErr ? ' err' : '');
    clearTimeout(toast._t);
    toast._t = setTimeout(function () { toastEl.className = ''; }, 3200);
  }

  function currentOpts() {
    return {
      scale: +scaleSelect.value,
      margin: +marginInput.value,
      background: bgSelect.value,
      radius: +radiusInput.value
    };
  }

  function updateDims() {
    if (!state.tableEl) return;
    var o = currentOpts();
    var w = Math.round((state.tableEl.offsetWidth + o.margin * 2) * o.scale);
    var h = Math.round((state.tableEl.offsetHeight + o.margin * 2) * o.scale);
    dimsInfo.textContent = w + ' × ' + h + ' px';
  }

  function previewHtml(sheet) {
    var o = currentOpts();
    return '<div class="hs-wrap" style="display:inline-block;padding:' + o.margin + 'px;background:' +
      o.background + ';border-radius:' + o.radius + 'px;overflow:hidden">' +
      TableRenderer.renderSheet(sheet, { allBorders: state.isCsv }) + '</div>';
  }

  function renderCurrentSheet() {
    var sheet = state.sheets[+sheetSelect.value];
    if (!sheet.rows.length) {
      toast('该工作表无数据', true);
      preview.innerHTML = '';
      state.tableEl = null;
      dimsInfo.textContent = '';
      clearSelection();
      return;
    }
    preview.innerHTML = previewHtml(sheet);
    state.tableEl = preview.querySelector('table');
    applyHighlight();
    updateDims();
  }

  function setSelection(sel) {
    state.sel = sel;
    var sheet = state.sheets[+sheetSelect.value];
    colLabel.textContent = columnName(sel.c);
    rowLabel.textContent = sel.r + 1;
    colWidthInput.value = sheet.colWidths[sel.c] || '';
    rowHeightInput.value = sheet.rows[sel.r].height || '';
    cellEdit.hidden = false;
    applyHighlight();
  }

  function applyHighlight() {
    if (!state.sel || !state.tableEl) return;
    var prev = state.tableEl.querySelector('.hs-selected');
    if (prev) prev.classList.remove('hs-selected');
    var td = state.tableEl.querySelector('td[data-r="' + state.sel.r + '"][data-c="' + state.sel.c + '"]');
    if (td) td.classList.add('hs-selected');
  }

  function clearSelection() {
    state.sel = null;
    cellEdit.hidden = true;
    if (state.tableEl) {
      var prev = state.tableEl.querySelector('.hs-selected');
      if (prev) prev.classList.remove('hs-selected');
    }
  }

  async function handleFile(file) {
    if (!file) return;
    if (!/\.(xlsx|csv)$/i.test(file.name)) {
      toast('请选择 .xlsx 或 .csv 文件', true);
      return;
    }
    try {
      var buf = await file.arrayBuffer();
      var isXlsx = /\.xlsx$/i.test(file.name);
      var parsed = isXlsx ? await XlsxParser.parseXlsx(buf) : XlsxParser.parseCsv(buf, file.name);
      state.sheets = parsed.sheets;
      state.fileName = file.name.replace(/\.[^.]+$/, '');
      state.isCsv = !isXlsx;
      var totalCells = state.sheets.reduce(function (n, s) { return n + s.cellCount; }, 0);
      if (totalCells > 8000) toast('表格较大，导出可能耗时较长');
      sheetSelect.innerHTML = state.sheets.map(function (s, i) {
        return '<option value="' + i + '">' + esc(s.name) + '</option>';
      }).join('');
      sheetSelect.disabled = state.sheets.length < 2;
      workspace.hidden = false;
      renderCurrentSheet();
    } catch (e) {
      console.error('解析失败', e);
      toast('解析失败：文件可能损坏或为加密文件', true);
    }
  }

  exportBtn.addEventListener('click', async function () {
    if (!state.tableEl) {
      toast('请先上传并选择有数据的表格', true);
      return;
    }
    try {
      exportBtn.disabled = true;
      var o = currentOpts();
      var selTd = state.tableEl.querySelector('.hs-selected');
      if (selTd) selTd.classList.remove('hs-selected');
      var res = await Exporter.exportPng(state.tableEl, o);
      var a = document.createElement('a');
      a.href = URL.createObjectURL(res.blob);
      a.download = state.fileName + '@' + o.scale + 'x.png';
      a.click();
      setTimeout(function () { URL.revokeObjectURL(a.href); }, 5000);
      toast('已导出 ' + res.width + '×' + res.height + ' PNG');
    } catch (e) {
      toast(e.message, true);
    } finally {
      exportBtn.disabled = false;
      applyHighlight();
    }
  });

  fileInput.addEventListener('change', function (e) { handleFile(e.target.files[0]); });
  dropZone.addEventListener('click', function () { fileInput.click(); });
  dropZone.addEventListener('dragover', function (e) { e.preventDefault(); dropZone.classList.add('over'); });
  dropZone.addEventListener('dragleave', function () { dropZone.classList.remove('over'); });
  dropZone.addEventListener('drop', function (e) {
    e.preventDefault();
    dropZone.classList.remove('over');
    handleFile(e.dataTransfer.files[0]);
  });
  sheetSelect.addEventListener('change', function () { clearSelection(); renderCurrentSheet(); });
  scaleSelect.addEventListener('change', updateDims);
  [marginInput, bgSelect, radiusInput].forEach(function (el) {
    el.addEventListener('change', function () { if (state.tableEl) renderCurrentSheet(); });
  });

  preview.addEventListener('click', function (e) {
    var td = e.target.closest('td');
    if (!td) { clearSelection(); return; }
    var r = td.dataset.r, c = td.dataset.c;
    if (r === undefined) {
      var tr = td.parentElement;
      r = Array.prototype.indexOf.call(tr.parentElement.children, tr);
    }
    setSelection({ r: +r, c: +c });
  });

  colWidthInput.addEventListener('change', function () {
    if (!state.sel) return;
    var v = +this.value;
    if (!(v > 0)) return;
    var sheet = state.sheets[+sheetSelect.value];
    var maxCol = 0;
    sheet.rows.forEach(function (r) { if (r.cells.length > maxCol) maxCol = r.cells.length; });
    var size = Math.max(maxCol, state.sel.c + 1);
    if (!sheet.colWidths.length) {
      sheet.colWidths = new Array(size).fill(59);
    } else if (sheet.colWidths.length < size) {
      var old = sheet.colWidths;
      sheet.colWidths = new Array(size).fill(59);
      for (var i = 0; i < old.length; i++) if (old[i]) sheet.colWidths[i] = old[i];
    }
    sheet.colWidths[state.sel.c] = v;
    renderCurrentSheet();
  });

  rowHeightInput.addEventListener('change', function () {
    if (!state.sel) return;
    var v = +this.value;
    if (!(v > 0)) return;
    state.sheets[+sheetSelect.value].rows[state.sel.r].height = v;
    renderCurrentSheet();
  });
})();