// ══ MERGE ENGINE ══
async function runMerge() {
  const totalFiles = state.files.length;
  showProgress(0, totalFiles, '');

  const colMake  = colLetter(document.getElementById('colMake').value  || 'B');
  const colPart  = colLetter(document.getElementById('colPart').value  || 'C');
  const colName  = colLetter(document.getElementById('colName').value  || 'D');
  const colQty   = colLetter(document.getElementById('colQty').value   || 'P');
  const startRow = Math.max(1, parseInt(document.getElementById('startRow').value) || 9);

  state.merged = []; state.errors = []; state.warnings = [];
  // mergeMap: key -> { make, partNum, names, totalQty, fileQtys:{filename:qty}, status, mismatch }
  const mergeMap = new Map();

  for (let fi = 0; fi < state.files.length; fi++) {
    const fileObj = state.files[fi];
    showProgress(fi, totalFiles, fileObj.name);
    const rows = await readExcelFile(fileObj.file, startRow);
    const cleanFileName = fileObj.name.replace(/\.[^/.]+$/, ''); // strip extension

    for (let i = startRow - 1; i < rows.length; i++) {
      const row = rows[i];
      const rawPart  = String(row[colPart]  ?? '').trim();
      const make     = String(row[colMake]  ?? '').trim();
      const partName = String(row[colName]  ?? '').trim();
      const rawQty   = row[colQty];

      // Only skip if Part Number is truly blank
      if (!rawPart) {
        if (make || partName || (rawQty !== '' && rawQty !== null && rawQty !== undefined))
          state.errors.push({ file: fileObj.name, rowNum: i + 1, make, partNum: '', partName, qty: rawQty, reason: 'Blank Part Number' });
        continue;
      }

      // Parse qty — never skip for non-numeric
      let qty = 0;
      if (typeof rawQty === 'number') {
        qty = rawQty;
      } else if (rawQty !== '' && rawQty !== null && rawQty !== undefined) {
        const parsed = parseFloat(String(rawQty).replace(/,/g, ''));
        qty = isNaN(parsed) ? 0 : parsed;
      }
      if (qty < 0) {
        state.warnings.push(`Row ${i + 1} in ${fileObj.name}: negative qty (${qty}) for part ${rawPart} — included as 0`);
        qty = 0;
      }

      const cleaned = cleanPartNumber(rawPart);
      const key = (make + '||' + cleaned).toLowerCase();

      if (mergeMap.has(key)) {
        const ex = mergeMap.get(key);
        ex.totalQty += qty;
        ex.fileQtys[cleanFileName] = (ex.fileQtys[cleanFileName] || 0) + qty;
        ex.names[partName] = (ex.names[partName] || 0) + 1;
        ex.status = 'updated';
      } else {
        const names = {}; names[partName] = 1;
        const fileQtys = {}; fileQtys[cleanFileName] = qty;
        mergeMap.set(key, { make, partNum: cleaned, names, totalQty: qty, fileQtys, status: 'new', mismatch: false });
      }
    }
  }

  // Resolve names + mismatch
  for (const [, entry] of mergeMap) {
    const ne = Object.entries(entry.names);
    ne.sort((a, b) => b[1] - a[1]);
    entry.resolvedName = ne[0][0];
    if (ne.length > 1) { entry.mismatch = true; entry.status = 'mismatch'; }
    state.merged.push(entry);
  }

  // Warnings
  const bc = state.errors.filter(e => e.reason === 'Blank Part Number').length;
  if (bc) state.warnings.push(`${bc} row(s) skipped — blank Part Number`);

  hideProgress();
  state.currentFilter = 'all';
  renderPreview();
  goTo(3);
}
