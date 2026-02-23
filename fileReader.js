// ══ CORE FILE READER ══
function readExcelFile(file, startRowHint) {
  return new Promise(resolve => {
    const reader = new FileReader();
    reader.onload = e => {
      try {
        const wb = XLSX.read(e.target.result, { type: 'array' });
        // Smart sheet selection: pick sheet with most data from startRow onwards
        const sr = (startRowHint || 9) - 1;
        let bestData = [];
        let bestCount = 0;
        for (const name of wb.SheetNames) {
          const ws = wb.Sheets[name];
          // Read twice: raw:false for text/formatted values (preserves leading zeros in part numbers)
          // raw:true for a parallel read to get actual numeric qty values
          const dataFormatted = XLSX.utils.sheet_to_json(ws, { header: 1, defval: '', raw: false });
          const dataRaw       = XLSX.utils.sheet_to_json(ws, { header: 1, defval: '', raw: true });
          // Merge: use formatted string for all cells EXCEPT qty col — we patch qty in mergeEngine
          // Store both reads together as { fmt, raw } pair
          let count = 0;
          for (let i = sr; i < dataFormatted.length; i++) {
            const row = dataFormatted[i];
            if (row[1] || row[2] || row[3]) count++;
          }
          if (count > bestCount) {
            bestCount = count;
            // Merge: formatted data but attach raw numeric values for qty lookup
            bestData = dataFormatted.map((row, ri) => {
              const rawRow = dataRaw[ri] || [];
              // Attach raw values as a hidden property for qty parsing
              row._raw = rawRow;
              return row;
            });
          }
        }
        resolve(bestData);
      } catch (err) { console.error(err); resolve([]); }
    };
    reader.readAsArrayBuffer(file);
  });
}

// Validate a file — returns { status: 'ok'|'corrupt'|'password'|'multi', sheetCount }
function validateExcelFile(file) {
  return new Promise(resolve => {
    const reader = new FileReader();
    reader.onload = e => {
      try {
        const wb = XLSX.read(e.target.result, { type: 'array' });
        const sheetCount = wb.SheetNames.length;
        if (sheetCount > 1) {
          resolve({ status: 'multi', sheetCount });
        } else {
          resolve({ status: 'ok', sheetCount });
        }
      } catch (err) {
        const msg = err.message || '';
        if (msg.toLowerCase().includes('password') || msg.toLowerCase().includes('encrypted')) {
          resolve({ status: 'password' });
        } else {
          resolve({ status: 'corrupt' });
        }
      }
    };
    reader.onerror = () => resolve({ status: 'corrupt' });
    reader.readAsArrayBuffer(file);
  });
}

function formatSize(b) {
  if (b < 1024) return b + ' B';
  if (b < 1048576) return (b / 1024).toFixed(1) + ' KB';
  return (b / 1048576).toFixed(1) + ' MB';
}

// ══ COLUMN LETTER TO INDEX ══
function colLetter(l) { return l.toUpperCase().charCodeAt(0) - 65; }
