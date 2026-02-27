window.state = {
  files: [],
  prefixes: [],
  suffixes: [],
  findReplace: [],
  merged: [],
  errors: [],
  warnings: [],
  mrsFile: null,
  mrsData: null,
  shortageFile: null,
  shortageData: null,   // Map<partNum_lower, qty>  — shortage qty per part
  currentFilter: 'all'
};
