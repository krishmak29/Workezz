/* ═══════════════════════════════════════════════════
   ui.js — Global state, DOM interaction, file handlers
   Dependencies: utils.js must be loaded first
═══════════════════════════════════════════════════ */

/* ── GLOBAL STATE ── */
var bomFile = null;
var siemFile = null;
var purFile = null;
var purRawData = null;
var siemRawData = null;
var bomSelectedSheet = null;
var allMakes = [];
var makeCheckedState = {};
var analysisResult = null;
var activeFilter = 'all';

/* ── NAVIGATION ── */
function goTo(n) {
  document.querySelectorAll('.screen').forEach(function (s) { s.classList.remove('active'); });
  document.getElementById('screen' + n).classList.add('active');
  [1, 2, 3].forEach(function (i) {
    var el = document.getElementById('step' + i);
    el.classList.remove('active', 'done');
    if (i === n) el.classList.add('active');
    else if (i < n) el.classList.add('done');
  });
  if (n === 3) populateExportScreen();
}

/* ── PANEL NAVIGATION ── */
function activatePanel(key) {
  document.querySelectorAll('.s1-nav-item').forEach(function (el) { el.classList.remove('active'); });
  document.querySelectorAll('.s1-panel').forEach(function (el) { el.classList.remove('active'); });
  var navEl = document.getElementById('nav-' + key);
  var panelEl = document.getElementById('panel-' + key);
  if (navEl) navEl.classList.add('active');
  if (panelEl) panelEl.classList.add('active');
}

/* ── COLUMN MAP TOGGLE ── */
function toggleColMap(panel) {
  var bodyId = panel === 'bom' ? 'bom-cmbody' : panel + '-cmBody';
  var body = document.getElementById(bodyId);
  var arrow = document.getElementById(panel + '-cmarrow');
  if (!body) return;
  var open = body.classList.toggle('open');
  if (arrow) arrow.style.transform = open ? 'rotate(90deg)' : '';
}

/* ── ADVANCED TOGGLE ── */
function toggleAdv(panel) {
  var btn = document.getElementById(panel + 'AdvBtn');
  var body = document.getElementById(panel + 'AdvBody');
  var open = body.classList.toggle('open');
  btn.classList.toggle('open', open);
}

/* ── DRAG & DROP ── */
function onDragOver(e, zoneId) {
  e.preventDefault();
  document.getElementById(zoneId).classList.add('drag-over');
}
function onDragLeave(zoneId) {
  document.getElementById(zoneId).classList.remove('drag-over');
}
function onDrop(e, type) {
  e.preventDefault();
  var files = e.dataTransfer.files;
  if (!files.length) return;
  if (type === 'bom') { var i = document.getElementById('bomInput'); i.files = files; handleBOM(i, files[0]); }
  else if (type === 'siem') { var i2 = document.getElementById('siemInput'); handleSiemens(i2, files[0]); }
  else if (type === 'pur') { var i3 = document.getElementById('purInput'); handlePurchase(i3, files[0]); }
  ['bomEmptyZone', 'siemEmptyZone', 'purEmptyZone'].forEach(function (id) {
    var el = document.getElementById(id); if (el) el.classList.remove('drag-over');
  });
}

/* ── SHOW / HIDE LOADED STATE ── */
function showLoaded(prefix, file, previewText) {
  document.getElementById(prefix + 'EmptyZone').style.display = 'none';
  document.getElementById(prefix + 'Loaded').style.display = '';
  document.getElementById(prefix + 'Ext').textContent = file.name.split('.').pop().toUpperCase();
  document.getElementById(prefix + 'Name').textContent = file.name;
  document.getElementById(prefix + 'Size').textContent = fmtSz(file.size);
  document.getElementById(prefix + 'Preview').innerHTML = '&#10003; ' + previewText;
}
function hideLoaded(prefix) {
  document.getElementById(prefix + 'EmptyZone').style.display = '';
  document.getElementById(prefix + 'Loaded').style.display = 'none';
}

/* ── BOM FILE HANDLER ── */
function handleBOM(input, fileOverride) {
  var file = fileOverride || input.files[0];
  if (!file) return;
  bomFile = file;
  bomSelectedSheet = null;
  if (input && !fileOverride) input.value = '';
  hideLoaded('bom');
  document.getElementById('bomSheetSelector').style.display = 'none';
  document.getElementById('bomSheetPills').innerHTML = '';
  updateChecklist();
  var fr = new FileReader();
  fr.onload = function (e) {
    try {
      var wb = XLSX.read(e.target.result, { type: 'array' });
      var names = wb.SheetNames;
      if (names.length === 1) {
        bomSelectedSheet = names[0];
        readSheetByName(bomFile, bomSelectedSheet).then(function (data) {
          var sr = parseInt(document.getElementById('bomStartRow').value) || 2;
          var cp = colIdx(document.getElementById('bomColPart').value || 'C');
          var c = 0;
          for (var i = sr - 1; i < data.length; i++) { if (String(data[i][cp] || '').trim()) c++; }
          showLoaded('bom', file, '<strong>' + c + ' part rows</strong> detected');
          renderBomPreview(data);
          updateChecklist();
        });
      } else {
        // Multiple sheets — show pills, wait for user pick
        document.getElementById('bomEmptyZone').style.display = 'none';
        document.getElementById('bomLoaded').style.display = '';
        document.getElementById('bomExt').textContent = file.name.split('.').pop().toUpperCase();
        document.getElementById('bomName').textContent = file.name;
        document.getElementById('bomSize').textContent = fmtSz(file.size);
        document.getElementById('bomPreview').innerHTML = '&#128196; Pick a sheet below';
        var html = '';
        names.forEach(function (n) {
          html += '<button class="bom-sheet-pill" data-sheet="' + escH(n) + '"'
            + ' onclick="selectBomSheet(\'' + escH(n).replace(/'/g, '&#39;') + '\')"'
            + ' style="padding:3px 12px;border-radius:12px;border:1.5px solid var(--border);'
            + 'background:var(--surface2);color:var(--text2);font-size:.7rem;font-weight:600;'
            + 'cursor:pointer;font-family:\'IBM Plex Sans\',sans-serif;transition:background .15s,color .15s;">'
            + escH(n) + '</button>';
        });
        document.getElementById('bomSheetPills').innerHTML = html;
        document.getElementById('bomSheetSelector').style.display = '';
        updateChecklist();
      }
    } catch (ex) { alert('Could not read file.'); }
  };
  fr.readAsArrayBuffer(file);
}

function selectBomSheet(name) {
  bomSelectedSheet = name;
  document.querySelectorAll('.bom-sheet-pill').forEach(function (p) {
    var sel = p.getAttribute('data-sheet') === name;
    p.style.background = sel ? 'var(--bom)' : 'var(--surface2)';
    p.style.color = sel ? '#fff' : 'var(--text2)';
    p.style.borderColor = sel ? 'var(--bom)' : 'var(--border)';
  });
  readSheetByName(bomFile, name).then(function (data) {
    var sr = parseInt(document.getElementById('bomStartRow').value) || 2;
    var cp = colIdx(document.getElementById('bomColPart').value || 'C');
    var c = 0;
    for (var i = sr - 1; i < data.length; i++) { if (String(data[i][cp] || '').trim()) c++; }
    document.getElementById('bomPreview').innerHTML = '&#10003; <strong>' + c + ' part rows</strong> in "' + escH(name) + '"';
    renderBomPreview(data);
    updateChecklist();
  });
}

function clearBOM() {
  bomFile = null;
  bomSelectedSheet = null;
  hideLoaded('bom');
  document.getElementById('bomInput').value = '';
  document.getElementById('bomSheetSelector').style.display = 'none';
  document.getElementById('bomSheetPills').innerHTML = '';
  document.getElementById('bomPreviewWrap').style.display = 'none';
  document.getElementById('bomPreviewTable').innerHTML = '';
  updateChecklist();
  updateEmptyState();
}

/* ── SIEMENS FILE HANDLER ── */
function handleSiemens(input, fileOverride) {
  var file = fileOverride || input.files[0];
  if (!file) return;
  siemFile = file;
  if (input && !fileOverride) input.value = '';
  readFirstSheetCached(siemFile).then(function (data) {
    siemRawData = data;
    var sr = parseInt(document.getElementById('siemStartRow').value) || 3;
    var cp = colIdx(document.getElementById('siemColPart').value || 'F');
    var cs = colIdx(document.getElementById('siemColStatus').value || 'D');
    var c = 0;
    for (var i = sr - 1; i < data.length; i++) {
      if (String(data[i][cs] || '').trim().toUpperCase() !== 'OPEN') continue;
      if (String(data[i][cp] || '').trim()) c++;
    }
    showLoaded('siem', siemFile, '<strong>' + c + ' OPEN</strong> parts detected');
    refreshSiemPreview();
    updateChecklist();
    updateEmptyState();
  });
  updateChecklist();
}

function clearSiemens() {
  siemFile = null;
  siemRawData = null;
  hideLoaded('siem');
  document.getElementById('siemInput').value = '';
  document.getElementById('siemPreviewWrap').style.display = 'none';
  document.getElementById('siemPreviewTable').innerHTML = '';
  updateChecklist();
  updateEmptyState();
}

/* ── PURCHASE FILE HANDLER ── */
function handlePurchase(input, fileOverride) {
  var file = fileOverride || input.files[0];
  if (!file) return;
  purFile = file;
  if (input && !fileOverride) input.value = '';
  readFirstSheetCached(purFile).then(function (data) {
    purRawData = data;
    var sr = parseInt(document.getElementById('purStartRow').value) || 2;
    var cp = colIdx(document.getElementById('purColPart').value || 'C');
    var c = 0;
    for (var i = sr - 1; i < data.length; i++) { if (String(data[i][cp] || '').trim()) c++; }
    showLoaded('pur', purFile, '<strong>' + c + ' rows</strong> in Purchase file');
    buildMakeList(data);
    refreshPurPreview();
    updateChecklist();
    updateEmptyState();
  });
  updateChecklist();
}

function clearPurchase() {
  purFile = null;
  purRawData = null;
  allMakes = [];
  hideLoaded('pur');
  document.getElementById('purInput').value = '';
  document.getElementById('makeFilterSection').style.display = 'none';
  document.getElementById('makeList').innerHTML = '';
  document.getElementById('makeSearch').value = '';
  document.getElementById('purPreviewWrap').style.display = 'none';
  document.getElementById('purPreviewTable').innerHTML = '';
  updateChecklist();
  updateEmptyState();
}

/* ── MAKE FILTER ── */
function buildMakeList(data) {
  var sr = parseInt(document.getElementById('purStartRow').value) || 2;
  var cm = colIdx(document.getElementById('purColMake').value || 'B');
  var makes = new Set();
  for (var i = sr - 1; i < data.length; i++) {
    var m = String(data[i][cm] || '').trim();
    if (m) makes.add(m);
  }
  allMakes = Array.from(makes).sort();
  document.getElementById('makeCount').textContent = allMakes.length + ' make' + (allMakes.length !== 1 ? 's' : '');
  document.getElementById('makeFilterSection').style.display = '';
  renderMakeList(allMakes);
  updateMakeSelectedCount();
}
function rebuildMakeList() { if (purRawData) buildMakeList(purRawData); }

function renderMakeList(list) {
  var html = '';
  if (!list.length) {
    html = '<div class="make-no-results">No makes found</div>';
  } else {
    list.forEach(function (m) {
      var safe = escH(m);
      var isChecked = makeCheckedState[m] !== false;
      html += '<label class="make-item"><input type="checkbox" value="' + safe + '"'
        + (isChecked ? ' checked' : '')
        + ' onchange="makeCheckedState[this.value]=this.checked;updateMakeSelectedCount()"> '
        + safe + '</label>';
    });
  }
  document.getElementById('makeList').innerHTML = html;
}

function filterMakeSearch() {
  var q = document.getElementById('makeSearch').value.toLowerCase().trim();
  var filtered = q ? allMakes.filter(function (m) { return m.toLowerCase().indexOf(q) >= 0; }) : allMakes;
  renderMakeList(filtered);
  updateMakeSelectedCount();
}
function updateMakeSelectedCount() {
  var total = allMakes.length;
  var sel = allMakes.filter(function (m) { return makeCheckedState[m] !== false; }).length;
  document.getElementById('makeSelectedCount').textContent = sel === total ? 'All selected' : sel + ' / ' + total + ' selected';
}
function selectAllMakes() { allMakes.forEach(function (m) { makeCheckedState[m] = true; }); filterMakeSearch(); }
function clearAllMakes() { allMakes.forEach(function (m) { makeCheckedState[m] = false; }); filterMakeSearch(); }
function getSelectedMakes() {
  var sel = new Set();
  allMakes.forEach(function (m) { if (makeCheckedState[m] !== false) sel.add(m); });
  return sel;
}

/* ── PREVIEW HELPERS ── */
function updateEmptyState() {
  var anyLoaded = !!(bomFile || siemFile || purFile);
  var es = document.getElementById('s1EmptyState');
  if (es) es.style.display = anyLoaded ? 'none' : '';
}

function renderStockPreview(data, tableId, wrapId, labelId, startRowInputId, partColInputId, accentColor) {
  var sr = parseInt(document.getElementById(startRowInputId).value) || 2;
  var cp = colIdx(document.getElementById(partColInputId).value || 'A');
  var hdrIdx = sr - 2;
  var startIdx = sr - 1;
  var maxC = 0;
  for (var i = Math.max(0, hdrIdx); i < Math.min(data.length, startIdx + 5); i++) {
    if (data[i] && data[i].length > maxC) maxC = data[i].length;
  }
  maxC = Math.min(maxC, 12);
  var html = '';
  if (hdrIdx >= 0 && data[hdrIdx]) {
    html += '<tr>';
    for (var ci = 0; ci < maxC; ci++) {
      var v = String(data[hdrIdx][ci] !== undefined ? data[hdrIdx][ci] : '');
      var isPN = (ci === cp);
      html += '<th style="' + (isPN ? 'background:' + accentColor + ';color:#fff;' : '') + '">' + (v ? escH(v) : ('Col ' + (ci + 1))) + '</th>';
    }
    html += '</tr>';
  }
  var shown = 0;
  for (var ri = startIdx; ri < data.length && shown < 5; ri++) {
    var row = data[ri];
    if (!row || !row.some(function (v) { return v !== ''; })) continue;
    html += '<tr>';
    for (var ci2 = 0; ci2 < maxC; ci2++) {
      var v2 = String(row[ci2] !== undefined ? row[ci2] : '');
      var isPN2 = (ci2 === cp);
      html += '<td style="font-weight:' + (isPN2 ? '700' : '400') + ';color:' + (isPN2 ? accentColor : '') + ';">' + escH(v2) + '</td>';
    }
    html += '</tr>';
    shown++;
  }
  if (!html) html = '<tr><td style="padding:10px;color:var(--muted);">No data at this start row</td></tr>';
  document.getElementById(tableId).innerHTML = html;
  if (labelId) document.getElementById(labelId).textContent = shown + ' rows shown';
  document.getElementById(wrapId).style.display = '';
}

function refreshSiemPreview() {
  if (!siemRawData) return;
  renderStockPreview(siemRawData, 'siemPreviewTable', 'siemPreviewWrap', 'siemPreviewLabel', 'siemStartRow', 'siemColPart', '#38bdf8');
}
function refreshPurPreview() {
  if (!purRawData) return;
  renderStockPreview(purRawData, 'purPreviewTable', 'purPreviewWrap', 'purPreviewLabel', 'purStartRow', 'purColPart', '#f59e0b');
}

function renderBomPreview(data) {
  var sr = parseInt(document.getElementById('bomStartRow').value) || 2;
  var cp = colIdx(document.getElementById('bomColPart').value || 'C');
  var hdrIdx = sr - 2;
  var startIdx = sr - 1;
  var html = '';
  var maxC = 0;
  for (var i = Math.max(0, hdrIdx); i < Math.min(data.length, startIdx + 5); i++) {
    if (data[i] && data[i].length > maxC) maxC = data[i].length;
  }
  maxC = Math.min(maxC, 12);
  if (hdrIdx >= 0 && data[hdrIdx]) {
    html += '<tr>';
    for (var ci = 0; ci < maxC; ci++) {
      var v = String(data[hdrIdx][ci] !== undefined ? data[hdrIdx][ci] : '');
      var isPN = (ci === cp);
      html += '<th style="' + (isPN ? 'background:#10b981;color:#fff;' : '') + '">' + (v ? escH(v) : ('Col ' + (ci + 1))) + '</th>';
    }
    html += '</tr>';
  }
  var shown = 0;
  for (var ri = startIdx; ri < data.length && shown < 5; ri++) {
    var row = data[ri];
    if (!row || !row.some(function (v) { return v !== ''; })) continue;
    html += '<tr>';
    for (var ci2 = 0; ci2 < maxC; ci2++) {
      var v2 = String(row[ci2] !== undefined ? row[ci2] : '');
      var isPN2 = (ci2 === cp);
      html += '<td style="font-weight:' + (isPN2 ? '700' : '400') + ';color:' + (isPN2 ? '#10b981' : '') + ';">' + escH(v2) + '</td>';
    }
    html += '</tr>';
    shown++;
  }
  if (!html) html = '<tr><td style="padding:10px;color:var(--muted);">No data at this start row</td></tr>';
  document.getElementById('bomPreviewTable').innerHTML = html;
  if (document.getElementById('bomPreviewLabel')) document.getElementById('bomPreviewLabel').textContent = shown + ' rows shown';
  document.getElementById('bomPreviewWrap').style.display = '';
  updateEmptyState();
}

function refreshBomPreview() {
  if (!bomFile || !bomSelectedSheet) return;
  readSheetByName(bomFile, bomSelectedSheet).then(function (data) {
    var sr = parseInt(document.getElementById('bomStartRow').value) || 2;
    var cp = colIdx(document.getElementById('bomColPart').value || 'C');
    var c = 0;
    for (var i = sr - 1; i < data.length; i++) { if (String(data[i][cp] || '').trim()) c++; }
    var name = bomSelectedSheet;
    document.getElementById('bomPreview').innerHTML = '&#10003; <strong>' + c + ' part rows</strong>' + (name ? ' in "' + escH(name) + '"' : '');
    renderBomPreview(data);
  });
}

/* ── CHECKLIST ── */
function updateChecklist() {
  var hasBom = !!(bomFile && bomSelectedSheet);
  var hasStock = !!(siemFile || purFile);

  var ckBom = document.getElementById('ck-bom');
  var ckStock = document.getElementById('ck-stock');
  if (ckBom) ckBom.className = 's1-run-ck' + (hasBom ? ' done' : '');
  if (ckStock) ckStock.className = 's1-run-ck' + (hasStock ? ' done' : '');

  var sbBom = document.getElementById('sb-ck-bom');
  var sbStock = document.getElementById('sb-ck-stock');
  if (sbBom) sbBom.className = 's1-foot-item' + (hasBom ? ' done' : '');
  if (sbStock) sbStock.className = 's1-foot-item' + (hasStock ? ' done' : (hasBom ? ' partial' : ''));

  document.getElementById('runBtn').disabled = !(hasBom && hasStock);

  var bomStatus = document.getElementById('nav-bom-status');
  var siemStatus = document.getElementById('nav-siem-status');
  var purStatus = document.getElementById('nav-pur-status');
  if (bomStatus) { bomStatus.className = 's1-nav-status ' + (hasBom ? 's1-ns-done' : 's1-ns-req'); bomStatus.textContent = hasBom ? 'Loaded' : 'Required'; }
  if (siemStatus) { siemStatus.className = 's1-nav-status ' + (siemFile ? 's1-ns-done' : 's1-ns-opt'); siemStatus.textContent = siemFile ? 'Loaded' : 'Optional'; }
  if (purStatus) { purStatus.className = 's1-nav-status ' + (purFile ? 's1-ns-done' : 's1-ns-opt'); purStatus.textContent = purFile ? 'Loaded' : 'Optional'; }

  var bomSub = document.getElementById('nav-bom-sub');
  if (bomSub && hasBom) { var bn = document.getElementById('bomName'); bomSub.textContent = bn && bn.textContent !== '—' ? bn.textContent : 'BOM loaded'; }
  else if (bomSub) bomSub.textContent = 'Bill of Materials';

  var siemSub = document.getElementById('nav-siem-sub');
  if (siemSub && siemFile) { var sn = document.getElementById('siemName'); siemSub.textContent = sn && sn.textContent !== '—' ? sn.textContent : 'Siemens loaded'; }
  else if (siemSub) siemSub.textContent = 'Open status only';

  var purSub = document.getElementById('nav-pur-sub');
  if (purSub && purFile) { var pn2 = document.getElementById('purName'); purSub.textContent = pn2 && pn2.textContent !== '—' ? pn2.textContent : 'Purchase loaded'; }
  else if (purSub) purSub.textContent = 'Filter by Make';
}

/* ── RESULTS TABLE FILTER ── */
function setFilter(f, btn) {
  activeFilter = f;
  document.querySelectorAll('.fpill').forEach(function (p) { p.classList.remove('on'); });
  btn.classList.add('on');
  renderTable();
}

/* ── DARK MODE ── */
function toggleDark() {
  var d = document.body.classList.toggle('dark');
  document.getElementById('darkIcon').textContent = d ? '\u2600\uFE0F' : '\uD83C\uDF19';
  document.getElementById('darkLabel').textContent = d ? 'Light Mode' : 'Dark Mode';
  try { localStorage.setItem('wz-theme', d ? 'dark' : 'light'); } catch (e) { }
}

/* ── INIT ── */
(function () {
  try {
    if (localStorage.getItem('wz-theme') === 'dark') {
      document.body.classList.add('dark');
      document.getElementById('darkIcon').textContent = '\u2600\uFE0F';
      document.getElementById('darkLabel').textContent = 'Light Mode';
    }
  } catch (e) { }
  updateChecklist();
})();
