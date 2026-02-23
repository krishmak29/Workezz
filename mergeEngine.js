// ══ MERGE ENGINE ══
async function runMerge() {
  const totalFiles = state.files.length;
  showProgress(0, totalFiles, '');

  const colMake  = colLetter(document.getElementById('colMake').value  || 'A');
  const colPart  = colLetter(document.getElementById('colPart').value  || 'B');
  const colName  = colLetter(document.getElementById('colName').value  || 'C');
  const colQty   = colLetter(document.getElementById('colQty').value   || 'O');
  const startRow = Math.max(1, parseInt(document.getElementById('startRow').value) || 9);

  state.merged = []; state.errors = []; state.warnings = [];
  // mergeMap: key = cleaned Part Number (lowercase)
  // value -> { makes:{makeName:count}, partNum, names:{}, totalQty, fileQtys:{}, status, mismatch }
  const mergeMap = new Map();

  for (let fi = 0; fi < state.files.length; fi++) {
    const fileObj = state.files[fi];
    showProgress(fi, totalFiles, fileObj.name);
    const rows = await readExcelFile(fileObj.file, startRow);
    const cleanFileName = fileObj.name.replace(/\.[^/.]+$/, ''); // strip extension

    for (let i = startRow - 1; i < rows.length; i++) {
      const row = rows[i];
      // Use formatted string value for Part Number to preserve leading zeros etc.
      const rawPart  = String(row[colPart]  ?? '').trim();
      const make     = String(row[colMake]  ?? '').trim();
      const partName = String(row[colName]  ?? '').trim();
      // Use raw numeric value for qty (attached as row._raw by fileReader)
      const rawQty   = (row._raw && row._raw[colQty] !== undefined) ? row._raw[colQty] : row[colQty];

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
      // KEY = Part Number only (not Make + Part Number)
      const key = cleaned.toLowerCase();

      if (mergeMap.has(key)) {
        const ex = mergeMap.get(key);
        ex.totalQty += qty;
        ex.fileQtys[cleanFileName] = (ex.fileQtys[cleanFileName] || 0) + qty;
        ex.names[partName] = (ex.names[partName] || 0) + 1;
        // Track makes — most frequent wins
        if (make) ex.makes[make] = (ex.makes[make] || 0) + 1;
        ex.status = 'updated';
      } else {
        const names = {}; names[partName] = 1;
        const fileQtys = {}; fileQtys[cleanFileName] = qty;
        const makes = {}; if (make) makes[make] = 1;
        mergeMap.set(key, { makes, partNum: cleaned, names, totalQty: qty, fileQtys, status: 'new', mismatch: false });
      }
    }
  }

  // Resolve names, makes + mismatch
  for (const [, entry] of mergeMap) {
    // Resolve Part Name — most frequent wins
    const ne = Object.entries(entry.names);
    ne.sort((a, b) => b[1] - a[1]);
    entry.resolvedName = ne[0][0];
    if (ne.length > 1) { entry.mismatch = true; entry.status = 'mismatch'; }

    // Resolve Make — most frequent wins, fallback to blank
    const me = Object.entries(entry.makes || {});
    me.sort((a, b) => b[1] - a[1]);
    entry.make = me.length ? me[0][0] : '';

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
