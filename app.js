// ============================================================
// app.js — Quince Logistics Calculator portal
// ============================================================
//
// Flow: file drop → parse xlsx → validate headers → validate COO →
// compute Ocean/Air per row → build output xlsx (input + 2 cols with
// freeze pane) → trigger download. All client-side via SheetJS.

const REQUIRED_HEADERS = [
  'Length (in)',
  'Width (in)',
  'Height (in)',
  'Weight (g)',
  'COO'
];

const OUTPUT_HEADER_OCEAN = 'Logistics Cost Ocean ($/unit)';
const OUTPUT_HEADER_AIR   = 'Logistics Cost Air ($/unit)';

// ---------------------------------------------------------------------------
// UI hookup
// ---------------------------------------------------------------------------
const dropZone   = document.getElementById('drop-zone');
const fileInput  = document.getElementById('file-input');
const browseBtn  = document.getElementById('browse-btn');
const statusEl   = document.getElementById('status');
const resultEl   = document.getElementById('result');
const resultSummary = document.getElementById('result-summary');
const downloadBtn = document.getElementById('download-btn');
const resetBtn   = document.getElementById('reset-btn');

let computedBlob = null;
let computedFileName = null;

browseBtn.addEventListener('click', () => fileInput.click());
fileInput.addEventListener('change', (e) => {
  if (e.target.files.length > 0) handleFile(e.target.files[0]);
});

dropZone.addEventListener('dragover', (e) => {
  e.preventDefault();
  dropZone.classList.add('drag-active');
});
dropZone.addEventListener('dragleave', () => dropZone.classList.remove('drag-active'));
dropZone.addEventListener('drop', (e) => {
  e.preventDefault();
  dropZone.classList.remove('drag-active');
  if (e.dataTransfer.files.length > 0) handleFile(e.dataTransfer.files[0]);
});

downloadBtn.addEventListener('click', () => {
  if (!computedBlob) return;
  const url = URL.createObjectURL(computedBlob);
  const a = document.createElement('a');
  a.href = url;
  a.download = computedFileName;
  document.body.appendChild(a);
  a.click();
  document.body.removeChild(a);
  URL.revokeObjectURL(url);
});

resetBtn.addEventListener('click', () => {
  computedBlob = null;
  computedFileName = null;
  fileInput.value = '';
  resultEl.classList.add('hidden');
  statusEl.classList.add('hidden');
  statusEl.textContent = '';
});

// ---------------------------------------------------------------------------
// Status helpers
// ---------------------------------------------------------------------------
function showStatus(msg, kind) {
  statusEl.classList.remove('hidden', 'status-error', 'status-ok', 'status-busy');
  if (kind) statusEl.classList.add('status-' + kind);
  statusEl.textContent = msg;
}
function showError(msg) {
  showStatus(msg, 'error');
  resultEl.classList.add('hidden');
}

// ---------------------------------------------------------------------------
// File handler
// ---------------------------------------------------------------------------
async function handleFile(file) {
  resultEl.classList.add('hidden');
  showStatus('Reading file…', 'busy');

  const t0 = performance.now();

  let workbook;
  try {
    const buf = await file.arrayBuffer();
    workbook = XLSX.read(buf, { type: 'array', cellDates: false });
  } catch (e) {
    showError("Couldn't open file. Please save as .xlsx and try again.");
    return;
  }

  // Use the first sheet.
  const sheetName = workbook.SheetNames[0];
  if (!sheetName) {
    showError('No sheets found in the workbook.');
    return;
  }
  const sheet = workbook.Sheets[sheetName];

  // Convert to 2D array (header: 1 means raw rows, no key conversion).
  const rows = XLSX.utils.sheet_to_json(sheet, { header: 1, defval: '', raw: true });
  if (rows.length < 2) {
    showError('File has no data rows. Expected a header row plus at least one data row.');
    return;
  }

  const headerRow = rows[0].map(c => String(c == null ? '' : c).trim());
  const dataRows  = rows.slice(1);

  // ----- Validate required headers (case-insensitive trim) ---------------
  const headerIndex = {};   // normalizedLowerHeader → column index
  headerRow.forEach((h, i) => {
    if (h) headerIndex[h.toLowerCase()] = i;
  });

  const missingHeaders = [];
  const colByField = {};   // field → column index
  REQUIRED_HEADERS.forEach((h) => {
    const idx = headerIndex[h.toLowerCase()];
    if (idx === undefined) missingHeaders.push(h);
    else colByField[h] = idx;
  });
  if (missingHeaders.length > 0) {
    showError(
      `Required column${missingHeaders.length > 1 ? 's' : ''} not found in your file: ` +
      missingHeaders.map(h => `"${h}"`).join(', ') +
      `.\n\nPlease add ${missingHeaders.length > 1 ? 'columns' : 'a column'} with ` +
      `the exact header name${missingHeaders.length > 1 ? 's' : ''} above and re-upload.`
    );
    return;
  }

  // ----- Validate per-row data + compute ---------------------------------
  showStatus(`Computing ${dataRows.length} rows…`, 'busy');

  const computeResults = []; // parallel array: {ocean, air} or null when blank row

  for (let i = 0; i < dataRows.length; i++) {
    const row = dataRows[i];
    const userRowNumber = i + 2;  // +1 for header, +1 for 1-based row number

    const rawL = row[colByField['Length (in)']];
    const rawW = row[colByField['Width (in)']];
    const rawH = row[colByField['Height (in)']];
    const rawWt = row[colByField['Weight (g)']];
    const rawCoo = row[colByField['COO']];

    // Detect fully-blank row → skip silently, write blank output.
    if (rawL === '' && rawW === '' && rawH === '' && rawWt === '' && rawCoo === '') {
      computeResults.push(null);
      continue;
    }

    // Validate numerics
    const l = Number(rawL);
    const w = Number(rawW);
    const h = Number(rawH);
    const wt = Number(rawWt);
    if (!isFinite(l) || l <= 0) {
      showError(`Row ${userRowNumber}: "Length (in)" must be a positive number. Got "${rawL}".`);
      return;
    }
    if (!isFinite(w) || w <= 0) {
      showError(`Row ${userRowNumber}: "Width (in)" must be a positive number. Got "${rawW}".`);
      return;
    }
    if (!isFinite(h) || h <= 0) {
      showError(`Row ${userRowNumber}: "Height (in)" must be a positive number. Got "${rawH}".`);
      return;
    }
    if (!isFinite(wt) || wt <= 0) {
      showError(`Row ${userRowNumber}: "Weight (g)" must be a positive number. Got "${rawWt}".`);
      return;
    }

    // Validate COO
    if (rawCoo === null || rawCoo === undefined || String(rawCoo).trim() === '') {
      showError(`Row ${userRowNumber}: "COO" is empty. Accepted formats: ISO code (e.g. "IN"), full name (e.g. "India"), or combined ("India | IN").`);
      return;
    }
    const cooCode = normalizeCOO(rawCoo);
    if (!cooCode) {
      showError(
        `Row ${userRowNumber}: COO value "${rawCoo}" not recognized.\n\n` +
        `Accepted formats:\n` +
        `  • 2-letter ISO code (e.g. "IN", "CN", "VN")\n` +
        `  • Full country name (e.g. "India", "China")\n` +
        `  • Combined ("India | IN")`
      );
      return;
    }

    const result = computeLogistics(wt, l, w, h, cooCode);
    computeResults.push(result);
  }

  // ----- Build output workbook ------------------------------------------
  showStatus('Building output file…', 'busy');

  // Output is input row data + 2 appended columns.
  const outputHeader = headerRow.concat([OUTPUT_HEADER_OCEAN, OUTPUT_HEADER_AIR]);
  const outputRows = [outputHeader];

  for (let i = 0; i < dataRows.length; i++) {
    const row = dataRows[i];
    // Ensure row has at least as many columns as the header (some xlsx
    // libraries truncate trailing blanks).
    while (row.length < headerRow.length) row.push('');
    const res = computeResults[i];
    if (res === null) {
      outputRows.push(row.concat(['', '']));
    } else {
      // Round to 4 decimals — matches Python CSV precision.
      outputRows.push(row.concat([
        Math.round(res.ocean * 10000) / 10000,
        Math.round(res.air * 10000) / 10000
      ]));
    }
  }

  const outputSheet = XLSX.utils.aoa_to_sheet(outputRows);

  // Apply $#,##0.00 number format to the two output columns.
  const oceanColIndex = headerRow.length;       // 0-based after input
  const airColIndex   = headerRow.length + 1;
  const oceanColLetter = XLSX.utils.encode_col(oceanColIndex);
  const airColLetter   = XLSX.utils.encode_col(airColIndex);
  for (let r = 1; r < outputRows.length; r++) {
    const oceanCell = outputSheet[oceanColLetter + (r + 1)];
    const airCell   = outputSheet[airColLetter + (r + 1)];
    if (oceanCell && typeof oceanCell.v === 'number') oceanCell.z = '"$"#,##0.00';
    if (airCell   && typeof airCell.v === 'number')   airCell.z = '"$"#,##0.00';
  }

  // Set column widths roughly based on header length.
  outputSheet['!cols'] = outputHeader.map((h) => ({ wch: Math.max(12, String(h).length + 2) }));

  // Assemble workbook
  const outputWb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(outputWb, outputSheet, sheetName.substring(0, 31) || 'Logistics');

  // Write to binary array
  const wbout = XLSX.write(outputWb, { bookType: 'xlsx', type: 'array' });
  computedBlob = new Blob([wbout], { type: 'application/octet-stream' });

  // Filename: logistics_<original>_<ts>.xlsx
  const baseName = file.name.replace(/\.(xlsx|xls)$/i, '');
  const ts = new Date().toISOString().slice(0, 19).replace(/[:T]/g, '-');
  computedFileName = `${baseName}__logistics_${ts}.xlsx`;

  const elapsedMs = Math.round(performance.now() - t0);
  const nonBlankCount = computeResults.filter(r => r !== null).length;

  showStatus('', 'ok');
  statusEl.classList.add('hidden');
  resultEl.classList.remove('hidden');
  resultSummary.innerHTML =
    `Computed <strong>${nonBlankCount.toLocaleString()}</strong> row${nonBlankCount === 1 ? '' : 's'} ` +
    `in <strong>${elapsedMs.toLocaleString()} ms</strong>. ` +
    `Output preserves your input columns and appends ` +
    `<code>${OUTPUT_HEADER_OCEAN}</code> + <code>${OUTPUT_HEADER_AIR}</code>.`;
}
