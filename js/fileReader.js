// ══ CORE FILE READER ══
function readExcelFile(file, startRowHint) {
  return new Promise(resolve => {
    const reader = new FileReader();
    reader.onload = e => {
      try {
        const wb = XLSX.read(e.target.result, { type: 'array' });
        // Try all sheets, pick the one with most data rows from startRow onwards
        const sr = (startRowHint || 9) - 1;
        let bestData = [];
        let bestCount = 0;
        for (const name of wb.SheetNames) {
          const ws = wb.Sheets[name];
          const data = XLSX.utils.sheet_to_json(ws, { header: 1, defval: '', raw: true });
          // Count rows from startRow that have content in cols B,C,D (idx 1,2,3)
          let count = 0;
          for (let i = sr; i < data.length; i++) {
            const row = data[i];
            if (row[1] || row[2] || row[3]) count++;
          }
          if (count > bestCount) { bestCount = count; bestData = data; }
        }
        resolve(bestData);
      } catch (err) { console.error(err); resolve([]); }
    };
    reader.readAsArrayBuffer(file);
  });
}

function formatSize(b) {
  if (b < 1024) return b + ' B';
  if (b < 1048576) return (b / 1024).toFixed(1) + ' KB';
  return (b / 1048576).toFixed(1) + ' MB';
}

// ══ CLEAN PART NUMBER ══
function colLetter(l) { return l.toUpperCase().charCodeAt(0) - 65; }
