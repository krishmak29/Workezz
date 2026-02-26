// ══ PROGRESS BAR ══
function showProgress(current, total, fileName) {
  const wrap = document.getElementById('progressWrap');
  const bar  = document.getElementById('progressBar');
  const txt  = document.getElementById('progressText');
  if (!wrap) return;
  wrap.style.display = '';
  const pct = total > 0 ? Math.round((current / total) * 100) : 0;
  bar.style.width = pct + '%';
  txt.textContent = total > 0
    ? `Processing file ${current + 1} of ${total} — ${fileName}`
    : 'Starting…';
  document.getElementById('mergeBtn').disabled = true;
}

function hideProgress() {
  const wrap = document.getElementById('progressWrap');
  if (wrap) wrap.style.display = 'none';
  const btn = document.getElementById('mergeBtn');
  if (btn) btn.disabled = false;
}

// ══ NAVIGATION ══
function goTo(n) {
  document.querySelectorAll('.screen').forEach(s => s.classList.remove('active'));
  document.querySelectorAll('.step').forEach((s, i) => {
    s.classList.remove('active', 'done');
    if (i + 1 < n) s.classList.add('done');
    if (i + 1 === n) s.classList.add('active');
  });
  document.getElementById('screen' + n).classList.add('active');
  if (n === 4) setExportNames();
}   

// ══ DRAG & DROP ══
const dz = document.getElementById('dropZone');
dz.addEventListener('dragover', e => { e.preventDefault(); dz.classList.add('dragover'); });
dz.addEventListener('dragleave', () => dz.classList.remove('dragover'));
dz.addEventListener('drop', e => {
  e.preventDefault(); dz.classList.remove('dragover');
  document.getElementById('fileInput').files = e.dataTransfer.files;
  addFiles();
});

// ══ ADD PANEL FILES ══
async function addFiles() {
  const input = document.getElementById('fileInput');
  if (!input.files.length) { alert('Please select files first.'); return; }
  for (const file of input.files) {
    if (state.files.find(f => f.name === file.name)) continue;
    const validation = await validateExcelFile(file);
    state.files.push({ file, name: file.name, size: file.size, validation });
  }
  input.value = '';
  renderFileTable();
}

function removeFile(name) {
  state.files = state.files.filter(f => f.name !== name);
  renderFileTable();
}

// ── DRAG AND DROP REORDER ──
let dragSrcIndex = null;

function renderFileTable() {
  const tbody = document.getElementById('fileTable');
  const count = state.files.length;
  document.getElementById('s1FileCount').textContent = count + ' file' + (count !== 1 ? 's' : '');
  document.getElementById('nextBtn1').disabled = count === 0;
  if (!count) {
    tbody.innerHTML = `<tr id="emptyRow1"><td colspan="4"><div class="empty-table">📋<p>No files added yet.<br>Upload .xls or .xlsx panel sheets above.</p></div></td></tr>`;
    return;
  }
  tbody.innerHTML = state.files.map((f, i) => {
    const ext = f.name.split('.').pop().toUpperCase();
    const v = f.validation || { status: 'ok' };
    let indicator = '';
    if (v.status === 'corrupt') {
      indicator = `<span class="file-status-indicator status-error">ⓘ<span class="tip">Could not read file — it may be corrupt or in an unsupported format</span></span>`;
    } else if (v.status === 'password') {
      indicator = `<span class="file-status-indicator status-error">ⓘ<span class="tip">File is password protected — cannot be read</span></span>`;
    } else if (v.status === 'multi') {
      indicator = `<span class="file-status-indicator status-warn">ⓘ<span class="tip">Multiple sheets detected (${v.sheetCount}) — using the first sheet</span></span>`;
    }
    return `<tr draggable="true" data-index="${i}" 
        ondragstart="bomDragStart(event,${i})" 
        ondragover="bomDragOver(event)" 
        ondragenter="bomDragEnter(event)"
        ondragleave="bomDragLeave(event)"
        ondrop="bomDrop(event,${i})" 
        ondragend="bomDragEnd(event)"
        style="cursor:grab;">
      <td class="sr-no" style="cursor:grab;">
        <span style="display:inline-flex;flex-direction:column;gap:1px;opacity:0.4;pointer-events:none;user-select:none;">
          <span style="display:block;width:14px;height:2px;background:currentColor;border-radius:1px;"></span>
          <span style="display:block;width:14px;height:2px;background:currentColor;border-radius:1px;"></span>
          <span style="display:block;width:14px;height:2px;background:currentColor;border-radius:1px;"></span>
        </span>
      </td>
      <td><div class="file-chip">
        <div class="file-chip-icon">${ext}</div>
        <div>
          <div class="file-chip-name">${f.name} ${indicator}</div>
          <div class="file-chip-size">${String(i + 1).padStart(2, '0')} · ${formatSize(f.size)}</div>
        </div>
      </div></td>
      <td style="font-size:.76rem;color:var(--muted);">${formatSize(f.size)}</td>
      <td style="text-align:right;"><button class="btn btn-danger-outline btn-sm" onclick="removeFile('${f.name}')">✕ Remove</button></td>
    </tr>`;
  }).join('');
}

function bomDragStart(e, i) {
  dragSrcIndex = i;
  e.dataTransfer.effectAllowed = 'move';
  e.currentTarget.style.opacity = '0.45';
}
function bomDragOver(e) {
  e.preventDefault();
  e.dataTransfer.dropEffect = 'move';
}
function bomDragEnter(e) {
  e.currentTarget.classList.add('drag-over');
}
function bomDragLeave(e) {
  e.currentTarget.classList.remove('drag-over');
}
function bomDrop(e, toIndex) {
  e.preventDefault();
  e.currentTarget.classList.remove('drag-over');
  if (dragSrcIndex === null || dragSrcIndex === toIndex) return;
  const moved = state.files.splice(dragSrcIndex, 1)[0];
  state.files.splice(toIndex, 0, moved);
  dragSrcIndex = null;
  renderFileTable();
}
function bomDragEnd(e) {
  e.currentTarget.style.opacity = '';
  document.querySelectorAll('#fileTable tr').forEach(r => r.classList.remove('drag-over'));
}



// ══ MRS FILE ══
document.getElementById('mrsInput').addEventListener('change', function() {
  if (!this.files.length) return;
  const file = this.files[0];
  state.mrsFile = file;
  parseMRS(file);
  document.getElementById('mrsFileName').innerHTML = `<span class="mrs-file-name">✓ ${file.name}</span>`;
  document.getElementById('mrsClearBtn').style.display = '';
  this.value = '';
});

function clearMRS() {
  state.mrsFile = null;
  state.mrsData = null;
  document.getElementById('mrsFileName').textContent = 'No file selected — MRS columns will be hidden';
  document.getElementById('mrsClearBtn').style.display = 'none';
}

async function parseMRS(file) {
  const data = await readExcelFile(file, 8);
  // MRS: data from row 8 (index 7)
  // Col B (idx 1) = Make, Col C (idx 2) = Part Number, Col P (idx 15) = Qty
  const mrsMap = new Map();
  for (let i = 7; i < data.length; i++) {
    const row = data[i];
    const rawPart = String(row[2] ?? '').trim();
    // Use raw numeric qty from _raw if available
    const rawQty = (row._raw && row._raw[15] !== undefined) ? row._raw[15] : row[15];
    if (!rawPart) continue;
    // Store both the raw part number AND the cleaned version as keys
    // so lookup always works regardless of whether prefixes are applied
    let qty = typeof rawQty === 'number' ? rawQty : parseFloat(String(rawQty ?? '').replace(/,/g, ''));
    if (isNaN(qty)) qty = 0;
    // Raw key
    const rawKey = rawPart.toLowerCase();
    mrsMap.set(rawKey, (mrsMap.get(rawKey) || 0) + qty);
    // Also store cleaned key (prefix/suffix stripped) so it matches panel cleaned parts
    let cleaned = rawPart;
    for (const p of state.prefixes) {
      const pt = p.trim();
      if (pt && cleaned.startsWith(pt)) { cleaned = cleaned.slice(pt.length).trim(); break; }
    }
    for (const s of state.suffixes) {
      const st = s.trim();
      if (st && cleaned.endsWith(st)) { cleaned = cleaned.slice(0, cleaned.length - st.length).trim(); break; }
    }
    const cleanedKey = cleaned.toLowerCase();
    if (cleanedKey !== rawKey) {
      mrsMap.set(cleanedKey, (mrsMap.get(cleanedKey) || 0) + qty);
    }
  }
  state.mrsData = mrsMap;
}

// ══ PREFIX/SUFFIX IMPORT ══
function importPrefixSuffix(input) {
  if (!input.files.length) return;
  const file = input.files[0];
  const reader = new FileReader();
  reader.onload = e => {
    try {
      const wb = XLSX.read(e.target.result, { type: 'array' });
      const ws = wb.Sheets[wb.SheetNames[0]];
      const data = XLSX.utils.sheet_to_json(ws, { header: 1, defval: '', raw: true });
      let addedP = 0, addedS = 0;
      // Row 1 = headers, data from row 2 (index 1)
      for (let i = 1; i < data.length; i++) {
        const row = data[i];
        const prefix = String(row[0] ?? '').trim();   // Col A
        const suffix = String(row[2] ?? '').trim();   // Col C
        if (prefix && !state.prefixes.includes(prefix)) { state.prefixes.push(prefix); addedP++; }
        if (suffix && !state.suffixes.includes(suffix)) { state.suffixes.push(suffix); addedS++; }
      }
      renderTags();
      savePrefixSuffix();
      alert(`✅ Imported ${addedP} prefix(es) and ${addedS} suffix(es) successfully!`);
    } catch (err) {
      alert('Error reading file: ' + err.message);
    }
  };
  reader.readAsArrayBuffer(file);
  input.value = '';
}

// ══ TAGS ══
function addTag(type, e) {
  if (e.key !== 'Enter') return;
  const input = document.getElementById(type + 'Input');
  const val = input.value.trim().replace(/[\u200B-\u200D\uFEFF]/g, ''); // strip zero-width chars
  if (!val) return;
  if (type === 'prefix') { if (!state.prefixes.includes(val)) state.prefixes.push(val); }
  else { if (!state.suffixes.includes(val)) state.suffixes.push(val); }
  input.value = '';
  renderTags();
  savePrefixSuffix();
}

function removeTag(type, val) {
  if (type === 'prefix') state.prefixes = state.prefixes.filter(v => v !== val);
  else state.suffixes = state.suffixes.filter(v => v !== val);
  renderTags();
  savePrefixSuffix();
}

function renderTags() {
  ['prefix', 'suffix'].forEach(type => {
    const wrap = document.getElementById(type + 'Wrap');
    const input = wrap.querySelector('input');
    const arr = type === 'prefix' ? state.prefixes : state.suffixes;
    wrap.querySelectorAll('.tag').forEach(t => t.remove());
    arr.forEach(v => {
      const tag = document.createElement('span');
      tag.className = 'tag';
      tag.innerHTML = `${v} <span class="tag-x" onclick="removeTag('${type}','${v}')">×</span>`;
      wrap.insertBefore(tag, input);
    });
  });
}

function savePrefixSuffix() {
  try {
    localStorage.setItem('wz-prefixes', JSON.stringify(state.prefixes));
    localStorage.setItem('wz-suffixes', JSON.stringify(state.suffixes));
  } catch(e) {}
}

function loadPrefixSuffix() {
  try {
    const p = localStorage.getItem('wz-prefixes');
    const s = localStorage.getItem('wz-suffixes');
    if (p) state.prefixes = JSON.parse(p);
    if (s) state.suffixes = JSON.parse(s);
    renderTags();
  } catch(e) {}
}

function addFRRow() {
  const id = Date.now();
  state.findReplace.push({ id, find: '', replace: '' });
  renderFRRows();
}

function renderFRRows() {
  const c = document.getElementById('frRows');
  c.innerHTML = state.findReplace.map(fr => `
    <div class="find-replace-row" style="margin-bottom:7px;">
      <input class="field-input" placeholder="Find…" value="${fr.find}" oninput="updateFR(${fr.id},'find',this.value)">
      <span>→</span>
      <div style="display:flex;gap:5px;">
        <input class="field-input" placeholder="Replace with…" value="${fr.replace}" oninput="updateFR(${fr.id},'replace',this.value)" style="flex:1;">
        <button class="btn btn-sm" style="background:var(--danger-bg);color:var(--danger);border:1px solid var(--danger-border);" onclick="removeFR(${fr.id})">✕</button>
      </div>
    </div>`).join('');
}

function updateFR(id, key, val) { const fr = state.findReplace.find(f => f.id === id); if (fr) fr[key] = val; }
function removeFR(id) { state.findReplace = state.findReplace.filter(f => f.id !== id); renderFRRows(); }



// ══ FILTER & PREVIEW ══
function setFilter(filter) {
  state.currentFilter = filter;
  document.querySelectorAll('.filter-btn').forEach(b => b.classList.remove('active'));
  const btn = document.querySelector(`.filter-btn.filter-${filter}`);
  if (btn) btn.classList.add('active');
  renderPreviewTable();
}

function renderPreview() {
  const counts = { new: 0, updated: 0, mismatch: 0, nomrs: 0 };
  state.merged.forEach(r => {
    if (r.status === 'new') counts.new++;
    else if (r.status === 'updated') counts.updated++;
    else if (r.status === 'mismatch') counts.mismatch++;
    // Count no MRS match
    if (state.mrsData) {
      const mrsQty = state.mrsData.get(r.partNum.toLowerCase()) || 0;
      if (!mrsQty) counts.nomrs++;
    }
  });

  document.getElementById('countAll').textContent      = state.merged.length;
  document.getElementById('countNew').textContent      = counts.new;
  document.getElementById('countUpdated').textContent  = counts.updated;
  document.getElementById('countMismatch').textContent = counts.mismatch;
  document.getElementById('countSkipped').textContent  = state.errors.length;
  // Show/hide No MRS Match button
  const noMrsBtn = document.getElementById('btnNoMrs');
  if (noMrsBtn) {
    noMrsBtn.style.display = state.mrsData ? '' : 'none';
    document.getElementById('countNoMrs').textContent = counts.nomrs;
  }

  const wp = document.getElementById('warnPanel');
  if (state.warnings.length) {
    wp.style.display = '';
    document.getElementById('warnList').innerHTML = state.warnings.map(w => `<li>⚠ ${w}</li>`).join('');
  } else wp.style.display = 'none';

  renderPreviewTable();
}

function renderPreviewTable() {
  const tbody = document.getElementById('previewTable');
  const filter = state.currentFilter;

  let rows;
  if (filter === 'skipped') {
    // Show skipped/error rows
    if (!state.errors.length) {
      tbody.innerHTML = `<tr><td colspan="6"><div class="empty-table">No skipped rows.</div></td></tr>`;
      return;
    }
    tbody.innerHTML = state.errors.map((r, i) => `
      <tr>
        <td class="sr-no">${String(i + 1).padStart(2, '0')}</td>
        <td style="font-size:.78rem;">${r.make}</td>
        <td style="font-family:'IBM Plex Mono',monospace;font-size:.75rem;">${r.partNum || '<span style="color:var(--muted)">—</span>'}</td>
        <td style="font-size:.78rem;">${r.partName}</td>
        <td style="font-size:.78rem;">${r.qty}</td>
        <td><span class="status-pill pill-skipped">⚠️ ${r.reason}</span></td>
      </tr>`).join('');
    return;
  }

  if (filter === 'nomrs') {
    rows = state.merged.filter(r => {
      const mrsQty = state.mrsData ? (state.mrsData.get(r.partNum.toLowerCase()) || 0) : 0;
      return !mrsQty;
    });
  } else {
    rows = filter === 'all' ? state.merged : state.merged.filter(r => r.status === filter);
  }

  if (!rows.length) {
    tbody.innerHTML = `<tr><td colspan="6"><div class="empty-table">No rows for this filter.</div></td></tr>`;
    return;
  }

  const pills = {
    new:      `<span class="status-pill pill-new">🟢 New</span>`,
    updated:  `<span class="status-pill pill-updated">🟡 Updated</span>`,
    mismatch: `<span class="status-pill pill-mismatch">🔴 Mismatch</span>`
  };

  tbody.innerHTML = rows.map((r, i) => `
    <tr class="${r.mismatch ? 'row-mismatch' : ''}">
      <td class="sr-no">${String(i + 1).padStart(2, '0')}</td>
      <td style="font-size:.78rem;">${r.make}</td>
      <td style="font-family:'IBM Plex Mono',monospace;font-size:.75rem;">${r.partNum}</td>
      <td style="font-size:.78rem;">${r.resolvedName}</td>
      <td style="font-family:'IBM Plex Mono',monospace;font-size:.78rem;font-weight:600;">${r.totalQty}</td>
      <td>${pills[r.status] || ''}</td>
    </tr>`).join('');
}

// ══ EXPORT ══
function todayStr() { return new Date().toISOString().slice(0, 10); }

function setExportNames() {
  const d = todayStr();
  document.getElementById('errorFileName').textContent  = `Workezz_Error_Report_${d}.xlsx`;
  document.getElementById('errorBtn').disabled = state.errors.length === 0;
  // Set default sheet name input
  const nameInput = document.getElementById('sheetNameInput');
  if (nameInput && !nameInput.dataset.touched) nameInput.value = `Master Sheet – ${d}`;
  // Build column checkboxes dynamically
  renderColSelector();
}

function renderColSelector() {
  const grid = document.getElementById('colSelectorGrid');
  if (!grid) return;
  const fileNames = state.files.map(f => f.name.replace(/\.[^/.]+$/, ''));
  const cols = [
    { id: 'col_sr',   label: 'SR',               badge: 'A',    always: true },
    { id: 'col_make', label: 'Make',              badge: 'B' },
    { id: 'col_part', label: 'Part Number',       badge: 'C' },
    { id: 'col_desc', label: 'Part Description',  badge: 'D' },
    ...fileNames.map((fn, i) => ({ id: `col_file_${i}`, label: fn, badge: 'File', fileName: fn })),
    { id: 'col_total', label: 'Total Qty',        badge: 'Qty' },
    ...(state.mrsData ? [
      { id: 'col_mrsqty',   label: 'MRS Qty',      badge: 'MRS', mrs: true },
      { id: 'col_mrsorder', label: 'Qty to Order', badge: 'MRS', mrs: true }
    ] : [])
  ];
  grid.innerHTML = cols.map(c => `
    <div class="col-item ${c.mrs ? 'mrs' : ''} checked" id="${c.id}" onclick="toggleColItem(this)">
      <div class="col-checkbox">
        <svg width="10" height="8" viewBox="0 0 10 8" fill="none">
          <path d="M1 4L3.5 6.5L9 1" stroke="white" stroke-width="1.8" stroke-linecap="round" stroke-linejoin="round"/>
        </svg>
      </div>
      <span class="col-name">${c.label}</span>
      <span class="col-badge">${c.badge}</span>
    </div>`).join('');
  updateSelectAllBtn();
}

function toggleColItem(el) {
  el.classList.toggle('checked');
  updateSelectAllBtn();
}

function updateSelectAllBtn() {
  const items = document.querySelectorAll('.col-item');
  if (!items.length) return;
  const allChecked = [...items].every(i => i.classList.contains('checked'));
  const btn = document.getElementById('selectAllBtn');
  if (btn) btn.textContent = allChecked ? 'Deselect All' : 'Select All';
}

function toggleAllCols() {
  const items = document.querySelectorAll('.col-item');
  const allChecked = [...items].every(i => i.classList.contains('checked'));
  items.forEach(i => allChecked ? i.classList.remove('checked') : i.classList.add('checked'));
  updateSelectAllBtn();
}

async function exportMerged() {
  const fileNames = state.files.map(f => f.name.replace(/\.[^/.]+$/, ''));

  const nameInput = document.getElementById('sheetNameInput');
  const sheetName = (nameInput && nameInput.value.trim()) || `Master Sheet – ${todayStr()}`;

  const isChecked = id => { const el = document.getElementById(id); return el ? el.classList.contains('checked') : true; };
  const allHeaders = ['SR', 'Make', 'Part Number', 'Part Description', ...fileNames, 'Total Qty'];
  const allHeaderIds = ['col_sr','col_make','col_part','col_desc',...fileNames.map((_,i)=>`col_file_${i}`),'col_total'];
  if (state.mrsData) { allHeaders.push('MRS Qty','Qty to Order'); allHeaderIds.push('col_mrsqty','col_mrsorder'); }

  const selectedIdx = allHeaderIds.map((id,i)=>({id,i})).filter(({id})=>isChecked(id)).map(({i})=>i);
  const header = selectedIdx.map(i => allHeaders[i]);

  const mrsColPos   = state.mrsData ? selectedIdx.indexOf(4 + fileNames.length + 1) : -1;
  const orderColPos = state.mrsData ? selectedIdx.indexOf(4 + fileNames.length + 2) : -1;
  const allWidths   = [5,20,28,42,...fileNames.map(()=>14),12,...(state.mrsData?[12,14]:[])];

  const dataRows = [];
  state.merged.forEach((r, ri) => {
    const fullRow = [ri+1, r.make, r.partNum, r.resolvedName];
    fileNames.forEach(fn => fullRow.push(r.fileQtys[fn] || 0));
    fullRow.push(r.totalQty);
    if (state.mrsData) {
      const mrsQty = state.mrsData.get(r.partNum.toLowerCase()) || 0;
      fullRow.push(mrsQty);
      fullRow.push(r.totalQty - mrsQty);
    }
    dataRows.push(selectedIdx.map(i => fullRow[i]));
  });

  const wb = new ExcelJS.Workbook();
  const ws = wb.addWorksheet(sheetName, { views:[{ state:'frozen', ySplit:1 }] });
  ws.columns = selectedIdx.map(i => ({ width: allWidths[i] || 14 }));

  // Header
  const hRow = ws.addRow(header);
  hRow.height = 28;
  hRow.eachCell(cell => {
    cell.fill = { type:'pattern', pattern:'solid', fgColor:{ argb:'FF1E3A5F' } };
    cell.font = { bold:true, color:{ argb:'FFFFFFFF' }, name:'Calibri', size:10 };
    cell.alignment = { horizontal:'center', vertical:'middle', wrapText:true };
    cell.border = { bottom:{ style:'thin', color:{ argb:'FF2563EB' } } };
  });

  // Data rows
  dataRows.forEach((row, ri) => {
    const exRow = ws.addRow(row);
    exRow.height = 18;
    const isAlt = ri % 2 === 0;
    exRow.eachCell({ includeEmpty:true }, (cell, colNum) => {
      const ci = colNum - 1;
      cell.font = { name:'Calibri', size:10 };
      cell.alignment = { vertical:'middle' };
      cell.border = {
        top:{ style:'thin', color:{ argb:'FFDEE2E6' } }, bottom:{ style:'thin', color:{ argb:'FFDEE2E6' } },
        left:{ style:'thin', color:{ argb:'FFDEE2E6' } }, right:{ style:'thin', color:{ argb:'FFDEE2E6' } }
      };
      if (mrsColPos >= 0 && ci === mrsColPos) {
        cell.fill = { type:'pattern', pattern:'solid', fgColor:{ argb:'FFDBEAFE' } };
        cell.font = { bold:true, color:{ argb:'FF1E40AF' }, name:'Calibri', size:10 };
        cell.alignment = { horizontal:'center', vertical:'middle' };
      } else if (orderColPos >= 0 && ci === orderColPos) {
        const val = cell.value;
        if (typeof val === 'number' && val <= 0) {
          cell.fill = { type:'pattern', pattern:'solid', fgColor:{ argb:'FFDCFCE7' } };
          cell.font = { bold:true, color:{ argb:'FF15803D' }, name:'Calibri', size:10 };
        } else {
          cell.fill = { type:'pattern', pattern:'solid', fgColor:{ argb:'FFFEE2E2' } };
          cell.font = { bold:true, color:{ argb:'FFB91C1C' }, name:'Calibri', size:10 };
        }
        cell.alignment = { horizontal:'center', vertical:'middle' };
      } else if (isAlt) {
        cell.fill = { type:'pattern', pattern:'solid', fgColor:{ argb:'FFF7F8FB' } };
      }
    });
  });

  ws.autoFilter = { from:{ row:1, column:1 }, to:{ row:1, column:header.length } };

  const buffer = await wb.xlsx.writeBuffer();
  const blob = new Blob([buffer], { type:'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
  const exportFileName = (nameInput && nameInput.value.trim()
    ? nameInput.value.trim().replace(/[^a-zA-Z0-9_\-. ]/g, '_')
    : 'Workezz_Merged') + `_${todayStr()}.xlsx`;
  saveAs(blob, exportFileName);
}

async function exportErrors() {
  const wb = new ExcelJS.Workbook();
  const ws = wb.addWorksheet('Error Report');
  ws.columns = [{ width:25 },{ width:12 },{ width:20 },{ width:25 },{ width:40 },{ width:10 },{ width:30 }];
  const hRow = ws.addRow(['File Name','Row Number','Make','Part Number','Part Description','Quantity','Error Reason']);
  hRow.height = 24;
  hRow.eachCell(cell => {
    cell.fill = { type:'pattern', pattern:'solid', fgColor:{ argb:'FF1E3A5F' } };
    cell.font = { bold:true, color:{ argb:'FFFFFFFF' }, name:'Calibri', size:10 };
    cell.alignment = { horizontal:'center', vertical:'middle' };
  });
  state.errors.forEach(e => ws.addRow([e.file, e.rowNum, e.make, e.partNum, e.partName, e.qty, e.reason]));
  const buffer = await wb.xlsx.writeBuffer();
  const blob = new Blob([buffer], { type:'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
  saveAs(blob, `Workezz_Error_Report_${todayStr()}.xlsx`);
}

function resetApp() {
  state.files = []; state.prefixes = []; state.suffixes = [];
  state.findReplace = []; state.merged = []; state.errors = []; state.warnings = [];
  state.mrsFile = null; state.mrsData = null; state.currentFilter = 'all';
  renderFileTable(); renderTags(); renderFRRows();
  clearMRS();
  goTo(1);
}

// ── DARK MODE TOGGLE ──
function toggleDark() {
  const isDark = document.body.classList.toggle('dark');
  document.getElementById('darkIcon').textContent = isDark ? '☀️' : '🌙';
  document.getElementById('darkLabel').textContent = isDark ? 'Light Mode' : 'Dark Mode';
  localStorage.setItem('wz-dark', isDark ? '1' : '0');
}
// Restore preference
(function(){
  try { if(localStorage.getItem('wz-dark')==='1'){document.body.classList.add('dark');document.getElementById('darkIcon').textContent='☀️';document.getElementById('darkLabel').textContent='Light Mode';} } catch(e){}
})();

// Init
loadPrefixSuffix();
renderFRRows();