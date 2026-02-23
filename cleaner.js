
function cleanPartNumber(val) {
  let s = String(val ?? '').trim();
  // Remove matching prefix — trim each stored prefix to remove invisible chars
  for (const p of state.prefixes) {
    const pt = p.trim();
    if (pt && s.startsWith(pt)) { s = s.slice(pt.length).trim(); break; }
  }
  // Remove matching suffix
  for (const sf of state.suffixes) {
    const st = sf.trim();
    if (st && s.endsWith(st)) { s = s.slice(0, s.length - st.length).trim(); break; }
  }
  for (const fr of state.findReplace) if (fr.find) s = s.split(fr.find).join(fr.replace);
  if (document.getElementById('togTrim').checked) s = s.trim();
  if (document.getElementById('togSpecial').checked) s = s.replace(/[^a-zA-Z0-9\s\-\.]/g, '');
  if (document.getElementById('togUpper').checked) s = s.toUpperCase();
  else if (document.getElementById('togProper').checked) s = s.replace(/\w\S*/g, t => t.charAt(0).toUpperCase() + t.slice(1).toLowerCase());
  return s.trim();
}