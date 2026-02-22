// ══ CORE FILE READER ══
function readExcelFile(file, startRowHint) {
  return new Promise(resolve => {
    const reader = new FileReader();
    reader.onload = e => {
      try {
        const wb = XLSX.read(e.target.result, { type: 'array' });
        // Always use first sheet
        const ws = wb.Sheets[wb.SheetNames[0]];
        const data = XLSX.utils.sheet_to_json(ws, { header: 1, defval: '', raw: true });
        resolve(data);
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
