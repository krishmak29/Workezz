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
function addFiles() {
  const input = document.getElementById('fileInput');
  if (!input.files.length) { alert('Please select files first.'); return; }
  for (const file of input.files) {
    if (state.files.find(f => f.name === file.name)) continue;
    state.files.push({ file, name: file.name, size: file.size });
  }
  input.value = '';
  renderFileTable();
}

function removeFile(name) {
  state.files = state.files.filter(f => f.name !== name);
  renderFileTable();
}

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
    return `<tr>
      <td class="sr-no">${String(i + 1).padStart(2, '0')}</td>
      <td><div class="file-chip">
        <div class="file-chip-icon">${ext}</div>
        <div><div class="file-chip-name">${f.name}</div><div class="file-chip-size">${formatSize(f.size)}</div></div>
      </div></td>
      <td style="font-size:.76rem;color:var(--muted);">${formatSize(f.size)}</td>
      <td style="text-align:right;"><button class="btn btn-danger-outline btn-sm" onclick="removeFile('${f.name}')">✕ Remove</button></td>
    </tr>`;
  }).join('');
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
  // Col B (idx 1) = Make, Col C (idx 2) = Part Number, Col O (idx 14) = Qty
  const mrsMap = new Map();
  for (let i = 7; i < data.length; i++) {
    const row = data[i];
    const make = String(row[1] ?? '').trim();
    const rawPart = String(row[2] ?? '').trim();
    const rawQty = row[15]; // Column P = index 15
    if (!rawPart) continue;
    // Strip prefixes from MRS part number so it matches cleaned panel part numbers
    let part = rawPart;
    for (const p of state.prefixes) {
      const pt = p.trim();
      if (pt && part.startsWith(pt)) { part = part.slice(pt.length).trim(); break; }
    }
    let qty = typeof rawQty === 'number' ? rawQty : parseFloat(String(rawQty ?? '').replace(/,/g, ''));
    if (isNaN(qty)) qty = 0;
    const key = part.toLowerCase();
    mrsMap.set(key, (mrsMap.get(key) || 0) + qty);
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
}

function removeTag(type, val) {
  if (type === 'prefix') state.prefixes = state.prefixes.filter(v => v !== val);
  else state.suffixes = state.suffixes.filter(v => v !== val);
  renderTags();
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
  document.querySelector(`.filter-btn.filter-${filter}`).classList.add('active');
  renderPreviewTable();
}

function renderPreview() {
  const counts = { new: 0, updated: 0, mismatch: 0 };
  state.merged.forEach(r => {
    if (r.status === 'new') counts.new++;
    else if (r.status === 'updated') counts.updated++;
    else if (r.status === 'mismatch') counts.mismatch++;
  });

  document.getElementById('countAll').textContent     = state.merged.length;
  document.getElementById('countNew').textContent     = counts.new;
  document.getElementById('countUpdated').textContent = counts.updated;
  document.getElementById('countMismatch').textContent= counts.mismatch;
  document.getElementById('countSkipped').textContent = state.errors.length;

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

  rows = filter === 'all' ? state.merged : state.merged.filter(r => r.status === filter);

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
  document.getElementById('exportFileName').textContent = `Workezz_Merged_${d}.xlsx`;
  document.getElementById('errorFileName').textContent  = `Workezz_Error_Report_${d}.xlsx`;
  document.getElementById('errorBtn').disabled = state.errors.length === 0;
  // Show MRS cols info
  const mrsEl = document.getElementById('exportMrsCols');
  if (state.mrsData) mrsEl.textContent = ' · MRS Qty · Qty to Order';
  else mrsEl.textContent = '';
}

function exportMerged() {
  const fileNames = state.files.map(f => f.name.replace(/\.[^/.]+$/, ''));

  // Build header
  const header = ['SR', 'Make', 'Part Number', 'Part Description', ...fileNames, 'Total Qty'];
  if (state.mrsData) { header.push('MRS Qty'); header.push('Qty to Order'); }

  const data = [header];

  state.merged.forEach((r, i) => {
    const row = [i + 1, r.make, r.partNum, r.resolvedName];
    fileNames.forEach(fn => row.push(r.fileQtys[fn] || 0));
    row.push(r.totalQty);

    if (state.mrsData) {
      let mrsQty = state.mrsData.get(r.partNum.toLowerCase()) || 0;
      if (!mrsQty) {
        for (const [mrsKey, mrsVal] of state.mrsData) {
          let stripped = mrsKey;
          for (const p of state.prefixes) {
            const pt = p.trim().toLowerCase();
            if (pt && stripped.startsWith(pt)) { stripped = stripped.slice(pt.length).trim(); break; }
          }
          if (stripped === r.partNum.toLowerCase()) { mrsQty = mrsVal; break; }
        }
      }
      const toOrder = r.totalQty - mrsQty;
      row.push(mrsQty);
      row.push(toOrder);
    }

    data.push(row);
  });

  const ws = XLSX.utils.aoa_to_sheet(data);

  // Column widths
  const cols = [
    { wch: 5 }, { wch: 20 }, { wch: 28 }, { wch: 42 },
    ...fileNames.map(() => ({ wch: 14 })),
    { wch: 12 }
  ];
  if (state.mrsData) { cols.push({ wch: 12 }); cols.push({ wch: 14 }); }
  ws['!cols'] = cols;

  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, 'Master Sheet');
  XLSX.writeFile(wb, `Workezz_Merged_${todayStr()}.xlsx`);
}

function exportErrors() {
  const data = [['File Name', 'Row Number', 'Make', 'Part Number', 'Part Description', 'Quantity', 'Error Reason']];
  state.errors.forEach(e => data.push([e.file, e.rowNum, e.make, e.partNum, e.partName, e.qty, e.reason]));
  const ws = XLSX.utils.aoa_to_sheet(data);
  ws['!cols'] = [{ wch: 25 }, { wch: 12 }, { wch: 20 }, { wch: 25 }, { wch: 40 }, { wch: 10 }, { wch: 30 }];
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, 'Error Report');
  XLSX.writeFile(wb, `Workezz_Error_Report_${todayStr()}.xlsx`);
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
renderFRRows();