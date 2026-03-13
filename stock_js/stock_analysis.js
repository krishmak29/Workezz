/* ═══════════════════════════════════════════════════
   analysis.js — Core logic, analysis engine & Excel export
   Dependencies: utils.js, ui.js must be loaded first
═══════════════════════════════════════════════════ */

/* ── BUILD SIEMENS MAP ── */
function buildSiemMap(data) {
  var sr = parseInt(document.getElementById('siemStartRow').value) || 3;
  var cp = colIdx(document.getElementById('siemColPart').value || 'F');
  var cs = colIdx(document.getElementById('siemColStatus').value || 'D');
  var cq = colIdx(document.getElementById('siemColQty').value || 'H');
  var cAllocStatus = 10; // col K (0-indexed)
  var cAllocStart = 11;  // col L onwards — project reservation columns
  var map = new Map();
  var normMap = new Map();
  var maxCols = 0;

  for (var i = sr - 1; i < data.length; i++) {
    var row = data[i];
    if (String(row[cs] || '').trim().toUpperCase() !== 'OPEN') continue;
    var pn = String(row[cp] || '').trim();
    if (!pn) continue;
    var k = pn.toLowerCase();
    var nk = normKey(pn);

    // Raw qty from col H
    var rq = row[cq];
    var rawQty = typeof rq === 'number' ? rq : (parseFloat(String(rq || '').replace(/,/g, '')) || 0);

    // Subtract project reservations from col L onwards if col K is marked
    var usedQty = 0;
    var statusMark = String(row[cAllocStatus] || '').trim();
    if (statusMark !== '') {
      for (var ac = cAllocStart; ac < row.length; ac++) {
        var av = row[ac];
        usedQty += typeof av === 'number' ? av : (parseFloat(String(av || '').replace(/,/g, '')) || 0);
      }
    }
    var q = Math.max(0, rawQty - usedQty); // net available qty

    if (!map.has(k)) map.set(k, { displayPart: pn, qty: 0, rows: [] });
    map.get(k).qty += q;
    map.get(k).rows.push(row);

    if (nk) {
      if (!normMap.has(nk)) normMap.set(nk, { displayPart: pn, qty: 0, rows: [] });
      normMap.get(nk).qty += q;
      normMap.get(nk).rows.push(row);
    }
    if (row.length > maxCols) maxCols = row.length;
  }
  return { map: map, normMap: normMap, maxCols: Math.min(Math.max(maxCols, 8), 12) };
}

/* ── BUILD PURCHASE MAP ── */
function buildPurMap(data, selectedMakes) {
  var sr = parseInt(document.getElementById('purStartRow').value) || 2;
  var cp = colIdx(document.getElementById('purColPart').value || 'C');
  var cm = colIdx(document.getElementById('purColMake').value || 'B');
  var cq = colIdx(document.getElementById('purColQty').value || 'F');
  var useMakeFilter = selectedMakes && selectedMakes.size > 0 && selectedMakes.size < allMakes.length;
  var map = new Map();
  var normMap = new Map();
  var maxCols = 0;

  for (var i = sr - 1; i < data.length; i++) {
    var row = data[i];
    var make = String(row[cm] || '').trim();
    if (useMakeFilter && !selectedMakes.has(make)) continue;
    var pn = String(row[cp] || '').trim();
    if (!pn) continue;
    var k = pn.toLowerCase();
    var nk = normKey(pn);
    var rq = row[cq];
    var q = typeof rq === 'number' ? rq : (parseFloat(String(rq || '').replace(/,/g, '')) || 0);

    if (!map.has(k)) map.set(k, { displayPart: pn, qty: 0, rows: [] });
    map.get(k).qty += q;
    map.get(k).rows.push(row);

    if (nk) {
      if (!normMap.has(nk)) normMap.set(nk, { displayPart: pn, qty: 0, rows: [] });
      normMap.get(nk).qty += q;
      normMap.get(nk).rows.push(row);
    }
    if (row.length > maxCols) maxCols = row.length;
  }
  return { map: map, normMap: normMap, maxCols: Math.min(Math.max(maxCols, 8), 20) };
}

/* ── RUN ANALYSIS ── */
async function runAnalysis() {
  var btn = document.getElementById('runBtn');
  btn.textContent = 'Analysing\u2026';
  btn.disabled = true;
  await new Promise(function (r) { setTimeout(r, 50); });
  try {
    var bomSR = parseInt(document.getElementById('bomStartRow').value) || 2;
    var bomCP = colIdx(document.getElementById('bomColPart').value || 'C');
    var toRead = [readSheetByName(bomFile, bomSelectedSheet)];
    if (siemFile) toRead.push(readFirstSheetCached(siemFile));
    var res = await Promise.all(toRead);
    var bomData = res[0] || [];
    var siemData = siemFile ? res[1] : null;
    var purData = purRawData;

    // Auto-detect "total qty" column
    var bomCQ = -1;
    var scanRows = Math.min(5, bomData.length);
    outer: for (var si = 0; si < scanRows; si++) {
      var scanRow = bomData[si] || [];
      for (var hi = 0; hi < scanRow.length; hi++) {
        var hv = String(scanRow[hi] || '').toLowerCase().trim();
        if (hv.indexOf('total') >= 0 && hv.indexOf('qty') >= 0) { bomCQ = hi; break outer; }
      }
    }
    if (bomCQ === -1) {
      outer2: for (var si2 = 0; si2 < scanRows; si2++) {
        var scanRow2 = bomData[si2] || [];
        for (var hi2 = 0; hi2 < scanRow2.length; hi2++) {
          var hv2 = String(scanRow2[hi2] || '').toLowerCase().trim();
          if (hv2.indexOf('qty') >= 0 && hi2 !== bomCP) { bomCQ = hi2; break outer2; }
        }
      }
    }

    var siemRes = siemData ? buildSiemMap(siemData) : { map: new Map(), normMap: new Map(), maxCols: 8 };
    var siemMap = siemRes.map, siemNormMap = siemRes.normMap, siemMaxCols = siemRes.maxCols;
    var selMakes = getSelectedMakes();
    var purRes = purData ? buildPurMap(purData, selMakes) : { map: new Map(), normMap: new Map(), maxCols: 8 };
    var purMap = purRes.map, purNormMap = purRes.normMap, purMaxCols = purRes.maxCols;
    // Detect real header row for BOM:
    // - Candidate is bomSR - 2 (row just before data, 0-indexed)
    // - If all non-empty cells in that row share the same value, it's a filler row
    //   → step back one more to bomSR - 3
    // - bomMaxCols = last non-empty column index in real header + 1 (no trailing blanks)
    var bomHdrIdx = bomSR - 2; // default: row just above data
    if (bomHdrIdx >= 0) {
      var candRow = bomData[bomHdrIdx] || [];
      var nonEmpty = candRow.filter(function(v) { return String(v || '').trim() !== ''; });
      if (nonEmpty.length > 0) {
        var firstVal = String(nonEmpty[0]).trim();
        var allSame = nonEmpty.every(function(v) { return String(v).trim() === firstVal; });
        if (allSame && bomHdrIdx - 1 >= 0) {
          bomHdrIdx = bomHdrIdx - 1; // filler row — use the row above instead
        }
      }
    }
    var bomHdrRaw = (bomHdrIdx >= 0 ? bomData[bomHdrIdx] : null) || [];
    // Build bomColMap: indices of columns that have a non-empty header value
    // This strips leading, trailing AND middle empty columns from the export
    var bomColMap = [];
    for (var hci = 0; hci < bomHdrRaw.length; hci++) {
      if (String(bomHdrRaw[hci] || '').trim() !== '') bomColMap.push(hci);
    }
    if (bomColMap.length < 8) {
      var maxIdx = bomColMap.length ? bomColMap[bomColMap.length - 1] : 7;
      for (var fi = bomColMap.length; fi < 8; fi++) bomColMap.push(maxIdx + 1 + (fi - bomColMap.length));
    }
    // bomHdr = only the header values for mapped columns (packed)
    var bomHdr = bomColMap.map(function(ci) { return bomHdrRaw[ci] !== undefined ? bomHdrRaw[ci] : ''; });
    var bomMaxCols = bomColMap.length;
    // Remap bomCP to its packed position in bomColMap
    var bomCPMapped = bomColMap.indexOf(bomCP);
    if (bomCPMapped === -1) {
      bomCPMapped = 0;
      for (var mi = 0; mi < bomColMap.length; mi++) {
        if (bomColMap[mi] <= bomCP) bomCPMapped = mi;
      }
    }
    var bomRows = [];

    for (var bi = bomSR - 1; bi < bomData.length; bi++) {
      var brow = bomData[bi];
      if (!brow.some(function (v) { return v !== ''; })) continue;
      var pn = String(brow[bomCP] || '').trim();
      if (!pn) continue; // skip rows with no part number (totals, blanks, etc.)
      var k = pn.toLowerCase();
      var siemLk = stockLookup(k, siemMap, siemNormMap);
      var purLk = stockLookup(k, purMap, purNormMap);
      var sm = !!(siemLk && siemLk.entry.qty >= 1);
      var pm = !!(purLk && purLk.entry.qty >= 1);
      var rqRaw = bomCQ >= 0 ? brow[bomCQ] : undefined;
      var reqQty = rqRaw != null && rqRaw !== '' ? (typeof rqRaw === 'number' ? rqRaw : (parseFloat(String(rqRaw).replace(/,/g, '')) || null)) : null;
      bomRows.push({
        raw: brow, pn: pn, siemMatch: sm, purMatch: pm, matched: sm || pm,
        siemQty: sm ? siemLk.entry.qty : 0, purQty: pm ? purLk.entry.qty : 0,
        reqQty: reqQty, siemLk: siemLk, purLk: purLk
      });
    }

    analysisResult = {
      bomRows: bomRows, bomHdr: bomHdr, bomMaxCols: bomMaxCols, bomCP: bomCP, bomColMap: bomColMap, bomCPMapped: bomCPMapped,
      siemMap: siemMap, siemNormMap: siemNormMap, purMap: purMap, purNormMap: purNormMap,
      siemMaxCols: siemMaxCols, purMaxCols: purMaxCols,
      siemData: siemData, purData: purData, selMakes: selMakes
    };

    renderScreen2();
    goTo(2);
  } catch (err) {
    console.error(err);
    alert('Analysis failed: ' + err.message);
  }
  btn.textContent = '\u26a1 Run Stock Check \u203a';
  btn.disabled = false;
  updateChecklist();
}

/* ── SCREEN 2 RENDER ── */
function renderScreen2() {
  var d = analysisResult;
  var nS = 0, nP = 0, nB = 0, nM = 0;
  d.bomRows.forEach(function (r) {
    if (r.siemMatch && r.purMatch) nB++;
    else if (r.siemMatch) nS++;
    else if (r.purMatch) nP++;
    else nM++;
  });
  var total = d.bomRows.length, matched = nS + nP + nB;
  var pct = total > 0 ? Math.round(matched / total * 100) : 0;

  ['stTotal', 'stSiem', 'stPur', 'stBoth', 'stMiss'].forEach(function (id, i) {
    document.getElementById(id).textContent = [total, nS, nP, nB, nM][i];
  });
  ['fc-all', 'fc-siem', 'fc-pur', 'fc-both', 'fc-miss'].forEach(function (id, i) {
    document.getElementById(id).textContent = [total, nS, nP, nB, nM][i];
  });
  ['lgBoth', 'lgSiem', 'lgPur', 'lgMiss'].forEach(function (id, i) {
    document.getElementById(id).textContent = [nB, nS, nP, nM][i];
  });

  document.getElementById('ringPct').textContent = pct + '%';
  var circ = 2 * Math.PI * 56;
  function seg(id, val, off) {
    var el = document.getElementById(id);
    el.setAttribute('stroke-dasharray', (total > 0 ? (val / total) * circ : 0) + ' ' + circ);
    el.setAttribute('stroke-dashoffset', -off);
  }
  seg('rBoth', nB, 0);
  seg('rSiem', nS, (nB / total || 0) * circ);
  seg('rPur', nP, ((nB + nS) / total || 0) * circ);
  seg('rMiss', nM, ((nB + nS + nP) / total || 0) * circ);

  activeFilter = 'all';
  document.querySelectorAll('.fpill').forEach(function (p) { p.classList.remove('on'); });
  document.querySelector('.fpill.fp-all').classList.add('on');
  renderTable();
}

/* ── RESULTS TABLE ── */
function renderTable() {
  if (!analysisResult) return;
  var f = activeFilter;
  var filtered = analysisResult.bomRows.filter(function (r) {
    if (f === 'all') return true;
    if (f === 'siem') return r.siemMatch && !r.purMatch;
    if (f === 'pur') return r.purMatch && !r.siemMatch;
    if (f === 'both') return r.siemMatch && r.purMatch;
    if (f === 'miss') return !r.matched;
  });

  var headCols;
  if (f === 'all') {
    headCols = '<th>#</th><th>Part Number</th><th>Source</th><th class="ctr">Required Qty</th><th class="ctr">Total Available</th>';
  } else if (f === 'siem') {
    headCols = '<th>#</th><th>Part Number</th><th class="ctr">Required Qty</th><th class="ctr">Siemens Available</th>';
  } else if (f === 'pur') {
    headCols = '<th>#</th><th>Part Number</th><th class="ctr">Required Qty</th><th class="ctr">Purchase Available</th>';
  } else if (f === 'both') {
    headCols = '<th>#</th><th>Part Number</th><th class="ctr">Required Qty</th><th class="ctr">Siemens Qty</th><th class="ctr">Purchase Qty</th><th class="ctr">Total Available</th>';
  } else {
    headCols = '<th>#</th><th>Part Number</th><th class="ctr">Required Qty</th>';
  }
  document.getElementById('resultsHead').innerHTML = '<tr>' + headCols + '</tr>';

  var colspan = headCols.split('<th').length - 1;
  var tbody = document.getElementById('resultsBody');
  if (!filtered.length) {
    tbody.innerHTML = '<tr><td colspan="' + colspan + '" style="text-align:center;padding:36px;color:var(--muted);font-size:.8rem;">No parts in this category</td></tr>';
    return;
  }

  var pnStyle = 'font-family:\'IBM Plex Mono\',monospace;font-size:.76rem;font-weight:600;';
  var idxStyle = 'color:var(--muted);font-family:\'IBM Plex Mono\',monospace;font-size:.72rem;';

  function qCell(val) {
    return val > 0 ? '<span class="qnum y">' + val + '</span>' : '<span class="qnum n">&mdash;</span>';
  }
  function reqCell(r) {
    return r.reqQty != null
      ? '<span class="qnum" style="color:var(--text2);text-align:center;display:block;">' + r.reqQty + '</span>'
      : '<span class="qnum n">&mdash;</span>';
  }

  var html = '';
  filtered.forEach(function (r, i) {
    var idx = '<td style="' + idxStyle + '">' + (i + 1) + '</td>';
    var pn = '<td style="' + pnStyle + '">' + escH(r.pn) + '</td>';
    var req = '<td>' + reqCell(r) + '</td>';
    var sq = qCell(r.siemQty);
    var pq = qCell(r.purQty);
    var tot = r.siemQty + r.purQty;
    var tq = tot > 0 ? '<span class="qnum y" style="font-size:.82rem;">' + tot + '</span>' : '<span class="qnum n">&mdash;</span>';
    var src;
    if (r.siemMatch && r.purMatch) src = '<td><span class="src-tag tag-both">&#9989; Both</span></td>';
    else if (r.siemMatch)          src = '<td><span class="src-tag tag-siem">&#127981; Siemens</span></td>';
    else if (r.purMatch)           src = '<td><span class="src-tag tag-pur">&#128722; Purchase</span></td>';
    else                           src = '<td><span class="src-tag tag-miss">&#10060; Not Found</span></td>';

    var cells;
    if (f === 'all') {
      cells = idx + pn + src + req + '<td>' + tq + '</td>';
    } else if (f === 'siem') {
      cells = idx + pn + req + '<td>' + sq + '</td>';
    } else if (f === 'pur') {
      cells = idx + pn + req + '<td>' + pq + '</td>';
    } else if (f === 'both') {
      cells = idx + pn + req + '<td>' + sq + '</td><td>' + pq + '</td><td>' + tq + '</td>';
    } else {
      cells = idx + pn + req;
    }
    html += '<tr>' + cells + '</tr>';
  });
  tbody.innerHTML = html;
}

/* ── SCREEN 3 — EXPORT SUMMARY ── */
function populateExportScreen() {
  if (!analysisResult) return;
  var d = analysisResult, total = d.bomRows.length;
  var matched = d.bomRows.filter(function (r) { return r.matched; }).length;
  var nS = d.bomRows.filter(function (r) { return r.siemMatch; }).length;
  var nP = d.bomRows.filter(function (r) { return r.purMatch; }).length;
  var pct = total > 0 ? Math.round(matched / total * 100) : 0;
  document.getElementById('ex-summary').innerHTML =
    '<strong>' + matched + ' of ' + total + '</strong> parts matched (' + pct + '%) &nbsp;·&nbsp; '
    + '<span style="color:var(--stk);">' + nS + ' Siemens</span> &nbsp;·&nbsp; '
    + '<span style="color:var(--pur);">' + nP + ' Purchase</span>';
  document.getElementById('exportOk').style.display = 'none';
}

/* ── EXPORT TO EXCEL ── */
async function runExport() {
  var btn = document.getElementById('exportBtn');
  btn.textContent = 'Generating\u2026';
  btn.disabled = true;
  document.getElementById('exportOk').style.display = 'none';
  await new Promise(function (r) { setTimeout(r, 50); });
  try {
    var d = analysisResult;
    var bomRows = d.bomRows, bomHdr = d.bomHdr, bomMaxCols = d.bomMaxCols, bomCP = d.bomCP, bomColMap = d.bomColMap, bomCPMapped = d.bomCPMapped;
    var siemMaxCols = d.siemMaxCols, purMaxCols = d.purMaxCols;
    var siemData = d.siemData, purData = d.purData, selMakes = d.selMakes;
    var matchCount = bomRows.filter(function (r) { return r.matched; }).length;
    var siemU = bomRows.filter(function (r) { return r.siemMatch; }).length;
    var purU = bomRows.filter(function (r) { return r.purMatch; }).length;

    var wb = new ExcelJS.Workbook();
    // Colour palette
    var W = 'FFFFFFFF', T = 'FF1E2235', M = 'FF8890AA';
    var BH = 'FF065F46', GR1 = 'FFF7F8FB', GR2 = 'FFECEEf4';
    var SH = 'FF075985', SL = 'FFB9E6FB', SF = 'FF0369A1';
    var PH = 'FF92400E', PL = 'FFFDE68A', PF = 'FFD97706', GF = 'FF15803D';

    function cs(cell, bg, fg, bold, mono, ah, bc) {
      cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: bg } };
      cell.font = { bold: bold || false, size: 9, color: { argb: fg || T }, name: mono ? 'Courier New' : 'Calibri' };
      cell.alignment = { vertical: 'middle', horizontal: ah || 'center' };
      cell.border = { bottom: { style: 'thin', color: { argb: bc || 'FFE5E7EB' } }, right: { style: 'thin', color: { argb: bc || 'FFE5E7EB' } } };
    }

    /* ── Sheet 1: BOM Stock Check ── */
    var ws1 = wb.addWorksheet('BOM - Stock Check');
    ws1.views = [{ state: 'frozen', xSplit: 0, ySplit: 2 }];
    for (var ci = 0; ci < bomMaxCols; ci++) ws1.getColumn(ci + 1).width = ci === bomCPMapped ? 28 : (ci < 4 ? 18 : 12);
    ws1.getColumn(bomMaxCols + 1).width = 14;
    var lr = ws1.addRow([]); lr.height = 20;
    ws1.mergeCells(1, 1, 1, bomMaxCols + 1);
    var lc = lr.getCell(1);
    lc.value = matchCount + ' of ' + bomRows.length + ' BOM parts matched  |  BLUE = Siemens  |  AMBER = Purchase';
    lc.font = { italic: true, size: 8, color: { argb: GF }, name: 'Calibri' };
    lc.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFD1FAE5' } };
    lc.alignment = { vertical: 'middle', horizontal: 'left' };
    var h1 = ws1.addRow([]); h1.height = 22;
    for (var hi = 0; hi < bomMaxCols; hi++) {
      var hc = h1.getCell(hi + 1);
      hc.value = (bomHdr[hi] !== undefined && bomHdr[hi] !== '') ? bomHdr[hi] : '';
      cs(hc, BH, W, true, false, 'center', 'FFD1D5DB');
    }
    var eh = h1.getCell(bomMaxCols + 1);
    eh.value = 'In-Stock Qty';
    cs(eh, 'FF059669', W, true, false, 'center', 'FF6EE7B7');

    for (var ri = 0; ri < bomRows.length; ri++) {
      if (ri % 100 === 0 && ri > 0) await new Promise(function (r) { setTimeout(r, 0); });
      var en = bomRows[ri]; var dr = ws1.addRow([]); dr.height = 17;
      var bg, nc;
      if (en.siemMatch) { bg = SL; nc = SF; }
      else if (en.purMatch) { bg = PL; nc = PF; }
      else { bg = ri % 2 === 0 ? GR1 : GR2; nc = 'FF1E3A5F'; }
      for (var ci2 = 0; ci2 < bomMaxCols; ci2++) {
        var srcIdx = bomColMap[ci2]; // source column index in original raw row
        var val = srcIdx !== undefined && en.raw[srcIdx] !== undefined && en.raw[srcIdx] !== '' ? en.raw[srcIdx] : null;
        var isPN = (ci2 === bomCPMapped); var c2 = dr.getCell(ci2 + 1);
        c2.value = val;
        cs(c2, bg, isPN ? nc : T, isPN, isPN);
      }
      var qc = dr.getCell(bomMaxCols + 1);
      qc.value = en.matched ? (en.siemQty + en.purQty) : null;
      cs(qc, bg, en.siemMatch ? SF : (en.purMatch ? PF : M), en.matched, false, 'center');
    }

    await new Promise(function (r) { setTimeout(r, 0); });

    /* ── Sheet 2: Siemens Stock ── */
    var ws2 = wb.addWorksheet('Siemens Stock - Required');
    ws2.views = [{ state: 'frozen', xSplit: 0, ySplit: 2 }];
    for (var si2 = 0; si2 < siemMaxCols; si2++) ws2.getColumn(si2 + 1).width = [24, 6, 16, 9, 14, 22, 24, 13][si2] || 14;
    var lr2 = ws2.addRow([]); lr2.height = 20;
    ws2.mergeCells(1, 1, 1, siemMaxCols);
    var lc2 = lr2.getCell(1);
    lc2.value = 'Siemens Open Stock - BOM-matched OPEN parts  |  ' + siemU + ' BOM parts matched';
    lc2.font = { italic: true, size: 8, color: { argb: SF }, name: 'Calibri' };
    lc2.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFE0F2FE' } };
    lc2.alignment = { vertical: 'middle', horizontal: 'left' };
    var siemHdrR = siemData && parseInt(document.getElementById('siemStartRow').value) >= 2 ? (siemData[parseInt(document.getElementById('siemStartRow').value) - 2] || []) : [];
    var sfb = ['PO No.', 'RM', 'Rack No.', 'Office', 'Remark', 'Type No. (Part)', 'OR No.', 'In Stock Qty'];
    var h2 = ws2.addRow([]); h2.height = 22;
    for (var hi2 = 0; hi2 < siemMaxCols; hi2++) {
      var hv2 = (siemHdrR[hi2] !== undefined && siemHdrR[hi2] !== '') ? siemHdrR[hi2] : (sfb[hi2] || '');
      var hc2 = h2.getCell(hi2 + 1);
      hc2.value = hv2;
      cs(hc2, SH, W, true, false, 'center', 'FF7DD3FC');
    }
    if (!siemData) {
      ws2.addRow(['No Siemens stock file uploaded.']).getCell(1).font = { italic: true, size: 9, color: { argb: M } };
    } else {
      var s2i = 0;
      var scP = colIdx(document.getElementById('siemColPart').value || 'F');
      var scQ = colIdx(document.getElementById('siemColQty').value || 'H');
      var cAllocStatus = 10;
      var cAllocStart = 11;
      for (var s2ri = 0; s2ri < bomRows.length; s2ri++) {
        var s2e = bomRows[s2ri];
        if (!s2e.siemMatch || !s2e.siemLk) continue;
        s2e.siemLk.entry.rows.forEach(function (raw) {
          var rawQty = typeof raw[scQ] === 'number' ? raw[scQ] : (parseFloat(String(raw[scQ] || '').replace(/,/g, '')) || 0);
          var statusMark = String(raw[cAllocStatus] || '').trim();
          var usedQty = 0;
          if (statusMark !== '') {
            for (var ac = cAllocStart; ac < raw.length; ac++) {
              var av = raw[ac];
              usedQty += typeof av === 'number' ? av : (parseFloat(String(av || '').replace(/,/g, '')) || 0);
            }
          }
          var netQty = Math.max(0, rawQty - usedQty);
          var rowBg;
          if (statusMark !== '' && netQty === 0) {
            rowBg = 'FFFCA5A5'; // red — fully allocated
          } else if (statusMark !== '' && netQty > 0) {
            rowBg = 'FFFDE68A'; // amber — partially allocated
          } else {
            rowBg = 'FFB9E6FB'; // blue — untouched
          }
          var dr2 = ws2.addRow([]); dr2.height = 17;
          for (var ci3 = 0; ci3 < siemMaxCols; ci3++) {
            var v = raw[ci3] !== undefined && raw[ci3] !== '' ? raw[ci3] : null;
            var c3 = dr2.getCell(ci3 + 1);
            c3.value = v;
            cs(c3, rowBg, (ci3 === scP) ? SF : T, (ci3 === scP || ci3 === scQ), (ci3 === scP), (ci3 === scQ) ? 'center' : 'left', rowBg);
          }
          s2i++;
        });
      }
      if (s2i === 0) ws2.addRow(['No Siemens rows matched this BOM.']).getCell(1).font = { italic: true, size: 9, color: { argb: M } };
    }

    await new Promise(function (r) { setTimeout(r, 0); });

    /* ── Sheet 3: Purchase Stock ── */
    var ws3 = wb.addWorksheet('Purchase Stock - Required');
    ws3.views = [{ state: 'frozen', xSplit: 0, ySplit: 2 }];
    for (var pi = 0; pi < purMaxCols; pi++) ws3.getColumn(pi + 1).width = pi < 4 ? 20 : 15;
    var makeLabel = selMakes.size > 0 ? Array.from(selMakes).join(', ') : 'All makes';
    var lr3 = ws3.addRow([]); lr3.height = 20;
    ws3.mergeCells(1, 1, 1, purMaxCols);
    var lc3 = lr3.getCell(1);
    lc3.value = 'Purchase Stock - BOM-matched parts  |  ' + purU + ' BOM parts matched  |  Makes: ' + makeLabel;
    lc3.font = { italic: true, size: 8, color: { argb: PF }, name: 'Calibri' };
    lc3.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFFFFBEB' } };
    lc3.alignment = { vertical: 'middle', horizontal: 'left' };
    var purSR2 = parseInt(document.getElementById('purStartRow').value) || 2;
    var purHdrR = purData && purSR2 >= 2 ? (purData[purSR2 - 2] || []) : [];
    var h3 = ws3.addRow([]); h3.height = 22;
    for (var hi3 = 0; hi3 < purMaxCols; hi3++) {
      var hv3 = (purHdrR[hi3] !== undefined && purHdrR[hi3] !== '') ? purHdrR[hi3] : ('Col ' + (hi3 + 1));
      var hc3 = h3.getCell(hi3 + 1);
      hc3.value = hv3;
      cs(hc3, PH, W, true, false, 'center', 'FFFDE68A');
    }
    if (!purData) {
      ws3.addRow(['No Purchase stock file uploaded.']).getCell(1).font = { italic: true, size: 9, color: { argb: M } };
    } else {
      var s3i = 0;
      var pcP = colIdx(document.getElementById('purColPart').value || 'C');
      var pcQ = colIdx(document.getElementById('purColQty').value || 'F');
      for (var s3ri = 0; s3ri < bomRows.length; s3ri++) {
        var s3e = bomRows[s3ri];
        if (!s3e.purMatch || !s3e.purLk) continue;
        s3e.purLk.entry.rows.forEach(function (raw) {
          var dr3 = ws3.addRow([]); dr3.height = 17;
          for (var ci4 = 0; ci4 < purMaxCols; ci4++) {
            var v2 = raw[ci4] !== undefined && raw[ci4] !== '' ? raw[ci4] : null;
            var c4 = dr3.getCell(ci4 + 1);
            c4.value = v2;
            cs(c4, PL, (ci4 === pcP) ? PF : T, (ci4 === pcP || ci4 === pcQ), (ci4 === pcP), (ci4 === pcQ) ? 'center' : 'left', 'FFFDE68A');
          }
          s3i++;
        });
      }
      if (s3i === 0) ws3.addRow(['No Purchase rows matched this BOM with selected makes.']).getCell(1).font = { italic: true, size: 9, color: { argb: M } };
    }

    var buffer = await wb.xlsx.writeBuffer();
    var blob = new Blob([buffer], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
    var fname = (document.getElementById('exportFileName').value || 'Stock_Check').trim().replace(/[^a-zA-Z0-9_\-. ]/g, '_');
    saveAs(blob, fname + '.xlsx');
    document.getElementById('exportOk').style.display = '';
    btn.textContent = '\u2193 Download Excel';
    btn.disabled = false;
  } catch (err) {
    console.error(err);
    alert('Export failed: ' + err.message);
    btn.textContent = '\u2193 Download Excel';
    btn.disabled = false;
  }
}
