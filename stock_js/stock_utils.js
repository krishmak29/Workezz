/* ═══════════════════════════════════════════════════
   utils.js — Pure helper functions (no DOM, no state)
   Dependencies: xlsx library (global XLSX)
═══════════════════════════════════════════════════ */

/* ── STRING / FORMAT HELPERS ── */
function colIdx(l) {
  return l.toUpperCase().charCodeAt(0) - 65;
}

function fmtSz(b) {
  if (b < 1024) return b + ' B';
  if (b < 1048576) return (b / 1024).toFixed(1) + ' KB';
  return (b / 1048576).toFixed(1) + ' MB';
}

function escH(s) {
  return String(s || '').replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');
}

/* ── PART NUMBER NORMALIZATION ── */
function normKey(str) {
  return String(str || '').replace(/[^a-zA-Z0-9]/g, '').replace(/^0+(?=[a-zA-Z0-9])/, '').toLowerCase();
}

/* ── STOCK LOOKUP (exact then normalized) ── */
function stockLookup(bomKey, map, normMap) {
  if (map.has(bomKey)) return { entry: map.get(bomKey), isNormalized: false };
  var nk = normKey(bomKey);
  if (nk && normMap.has(nk)) return { entry: normMap.get(nk), isNormalized: true };
  return null;
}

/* ── EXCEL FILE READERS ── */

// Read a specific sheet by name, returns array-of-arrays (column-index safe)
function readSheetByName(file, sheetName) {
  return new Promise(function (resolve) {
    var r = new FileReader();
    r.onload = function (e) {
      try {
        var wb = XLSX.read(e.target.result, { type: 'array' });
        var sn = (sheetName && wb.SheetNames.indexOf(sheetName) >= 0) ? sheetName : wb.SheetNames[0];
        var ws = wb.Sheets[sn];
        var ref = ws['!ref'];
        if (!ref) { resolve([]); return; }
        var range = XLSX.utils.decode_range(ref);
        var rows = XLSX.utils.sheet_to_json(ws, { header: 'A', defval: '', raw: true });
        var result = rows.map(function (row) {
          var arr = [];
          for (var ci = 0; ci <= range.e.c; ci++) {
            var cl = XLSX.utils.encode_col(ci);
            arr.push(row[cl] !== undefined ? row[cl] : '');
          }
          return arr;
        });
        resolve(result);
      } catch (ex) { resolve([]); }
    };
    r.readAsArrayBuffer(file);
  });
}

// Read first sheet only (used for Siemens & Purchase — single-sheet files)
function readFirstSheetCached(file) {
  return new Promise(function (resolve) {
    var r = new FileReader();
    r.onload = function (e) {
      try {
        var wb = XLSX.read(e.target.result, { type: 'array' });
        resolve(XLSX.utils.sheet_to_json(wb.Sheets[wb.SheetNames[0]], { header: 1, defval: '', raw: true }));
      } catch (ex) { resolve([]); }
    };
    r.readAsArrayBuffer(file);
  });
}

// Read the sheet with the most data rows (fallback helper)
function readBestSheet(file) {
  return new Promise(function (resolve) {
    var r = new FileReader();
    r.onload = function (e) {
      try {
        var wb = XLSX.read(e.target.result, { type: 'array' });
        var best = [], bestN = 0;
        for (var i = 0; i < wb.SheetNames.length; i++) {
          var d = XLSX.utils.sheet_to_json(wb.Sheets[wb.SheetNames[i]], { header: 1, defval: '', raw: true });
          var n = d.filter(function (row) { return row.some(function (v) { return v !== ''; }); }).length;
          if (n > bestN) { bestN = n; best = d; }
        }
        resolve(best);
      } catch (ex) { resolve([]); }
    };
    r.readAsArrayBuffer(file);
  });
}
