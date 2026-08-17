(function () {
  var $ = function (id) { return document.getElementById(id); };
  var fileInput = $('file-input'), dropZone = $('drop-zone'), workspace = $('workspace');
  var sheetSelect = $('sheet-select'), scaleSelect = $('scale-select'), marginInput = $('margin-input');
  var bgSelect = $('bg-select'), radiusInput = $('radius-input'), exportBtn = $('export-btn');
  var preview = $('preview'), dimsInfo = $('dims-info'), toastEl = $('toast');

  var state = { sheets: null, tableEl: null, fileName: 'sheet', isCsv: false };

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

  function renderCurrentSheet() {
    var sheet = state.sheets[+sheetSelect.value];
    if (!sheet.rows.length) {
      toast('该工作表无数据', true);
      preview.innerHTML = '';
      state.tableEl = null;
      dimsInfo.textContent = '';
      return;
    }
    preview.innerHTML = TableRenderer.renderSheet(sheet, { allBorders: state.isCsv });
    state.tableEl = preview.firstElementChild;
    updateDims();
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
  sheetSelect.addEventListener('change', renderCurrentSheet);
  [scaleSelect, marginInput, bgSelect, radiusInput].forEach(function (el) {
    el.addEventListener('change', updateDims);
  });
})();