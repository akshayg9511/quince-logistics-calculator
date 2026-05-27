// ============================================================
// app.js — Quince Logistics Calculator portal
// ============================================================
//
// Flow: file drop → parse xlsx → detect header row (1-5) → loop rows →
// compute Ocean/Air or write a Notes message → build output xlsx
// (header + data + 3 appended columns) → trigger download.
// All client-side via SheetJS.

const REQUIRED_HEADERS = [
  'Length (in)',
  'Width (in)',
  'Height (in)',
  'Weight (g)',
  'COO'
];

const MAX_HEADER_SEARCH_ROW = 5;

const OUTPUT_HEADER_OCEAN = 'Logistics Cost Ocean ($/unit)';
const OUTPUT_HEADER_AIR   = 'Logistics Cost Air ($/unit)';
const OUTPUT_HEADER_NOTES = 'Notes';

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
// Header detection — scan first MAX_HEADER_SEARCH_ROW rows for one that
// contains every required header name.
// ---------------------------------------------------------------------------
function findHeaderRow(rows) {
  const requiredLower = REQUIRED_HEADERS.map(h => h.toLowerCase());
  const searchLimit = Math.min(MAX_HEADER_SEARCH_ROW, rows.length);
  for (let r = 0; r < searchLimit; r++) {
    const cells = (rows[r] || []).map(c => String(c == null ? '' : c).trim().toLowerCase());
    const cellSet = new Set(cells);
    if (requiredLower.every(h => cellSet.has(h))) return r;
  }
  return -1;
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
    showError('File is empty or has no data rows. Expected a header row plus at least one data row.');
    return;
  }

  // ----- Find header row in rows 1..MAX_HEADER_SEARCH_ROW ---------------
  const headerRowIdx = findHeaderRow(rows);
  if (headerRowIdx === -1) {
    showError(
      `Could not find the required column headers ` +
      REQUIRED_HEADERS.map(h => `"${h}"`).join(', ') +
      ` in rows 1–${MAX_HEADER_SEARCH_ROW} of your file. ` +
      `Please make sure these exact column names appear in one of the first ${MAX_HEADER_SEARCH_ROW} rows, then re-upload.`
    );
    return;
  }

  const headerRow = rows[headerRowIdx].map(c => String(c == null ? '' : c).trim());
  const dataRows  = rows.slice(headerRowIdx + 1);

  // Build the column-index lookup for required fields.
  const headerIndex = {};   // lowercase-trimmed header → column index
  headerRow.forEach((h, i) => {
    if (h) headerIndex[h.toLowerCase()] = i;
  });
  const colByField = {};
  REQUIRED_HEADERS.forEach((h) => { colByField[h] = headerIndex[h.toLowerCase()]; });

  // ----- Per-row compute -----------------------------------------------
  showStatus(`Computing ${dataRows.length.toLocaleString()} rows…`, 'busy');

  // computeResults[i] = { ocean, air, note }  (ocean/air may be null on bad row)
  const computeResults = [];
  let okCount = 0;
  let issueCount = 0;
  let blankCount = 0;

  for (let i = 0; i < dataRows.length; i++) {
    const row = dataRows[i];

    const rawL  = row[colByField['Length (in)']];
    const rawW  = row[colByField['Width (in)']];
    const rawH  = row[colByField['Height (in)']];
    const rawWt = row[colByField['Weight (g)']];
    const rawCoo = row[colByField['COO']];

    // Fully-blank row → preserve in output, no Note (silent skip).
    if ((rawL === '' || rawL == null) &&
        (rawW === '' || rawW == null) &&
        (rawH === '' || rawH == null) &&
        (rawWt === '' || rawWt == null) &&
        (rawCoo === '' || rawCoo == null)) {
      computeResults.push({ ocean: null, air: null, note: '' });
      blankCount++;
      continue;
    }

    // Validate each numeric input. Collect problems instead of aborting.
    const problems = [];
    const numericChecks = [
      ['Length (in)', rawL],
      ['Width (in)',  rawW],
      ['Height (in)', rawH],
      ['Weight (g)',  rawWt]
    ];
    const numericValues = {};
    for (const [field, raw] of numericChecks) {
      if (raw === '' || raw == null) {
        problems.push(`${field} is empty`);
        numericValues[field] = NaN;
        continue;
      }
      const n = Number(raw);
      if (!isFinite(n)) {
        problems.push(`${field} is not a number (got "${raw}")`);
        numericValues[field] = NaN;
      } else if (n <= 0) {
        problems.push(`${field} is ${n}`);
        numericValues[field] = n;
      } else {
        numericValues[field] = n;
      }
    }

    // Validate COO.
    let cooCode = null;
    if (rawCoo === null || rawCoo === undefined || String(rawCoo).trim() === '') {
      problems.push('COO is empty');
    } else {
      cooCode = normalizeCOO(rawCoo);
      if (!cooCode) {
        problems.push(`COO "${rawCoo}" not recognized`);
      }
    }

    if (problems.length > 0) {
      computeResults.push({ ocean: null, air: null, note: problems.join('; ') });
      issueCount++;
      continue;
    }

    // Compute.
    const r = computeLogistics(
      numericValues['Weight (g)'],
      numericValues['Length (in)'],
      numericValues['Width (in)'],
      numericValues['Height (in)'],
      cooCode
    );
    computeResults.push({ ocean: r.ocean, air: r.air, note: '' });
    okCount++;
  }

  // ----- Build output workbook ------------------------------------------
  showStatus('Building output file…', 'busy');

  const outputHeader = headerRow.concat([OUTPUT_HEADER_OCEAN, OUTPUT_HEADER_AIR, OUTPUT_HEADER_NOTES]);
  const outputRows = [outputHeader];

  for (let i = 0; i < dataRows.length; i++) {
    const row = dataRows[i];
    while (row.length < headerRow.length) row.push('');
    const res = computeResults[i];
    if (res.ocean === null && res.air === null) {
      // Blank or invalid row.
      outputRows.push(row.concat(['', '', res.note || '']));
    } else {
      outputRows.push(row.concat([
        Math.round(res.ocean * 10000) / 10000,
        Math.round(res.air * 10000) / 10000,
        ''   // empty Notes for clean rows
      ]));
    }
  }

  const outputSheet = XLSX.utils.aoa_to_sheet(outputRows);

  // Column indices for the three output columns we appended.
  const oceanColIndex = headerRow.length;
  const airColIndex   = headerRow.length + 1;
  const notesColIndex = headerRow.length + 2;
  const oceanColLetter = XLSX.utils.encode_col(oceanColIndex);
  const airColLetter   = XLSX.utils.encode_col(airColIndex);
  const notesColLetter = XLSX.utils.encode_col(notesColIndex);

  // Cell styles: darker green for the output header row, light green for
  // every data cell in the three appended output columns. xlsx-js-style
  // honors these on write; viewers (Excel, Google Sheets, Numbers) render
  // the fill correctly.
  const HEADER_FILL = { patternType: 'solid', fgColor: { rgb: 'A9D08E' } };
  const DATA_FILL   = { patternType: 'solid', fgColor: { rgb: 'E2EFDA' } };
  const HEADER_FONT = { bold: true };

  // Style the header cells (row 1 = output row index 0).
  [oceanColLetter, airColLetter, notesColLetter].forEach((letter) => {
    const ref = letter + '1';
    if (outputSheet[ref]) {
      outputSheet[ref].s = { fill: HEADER_FILL, font: HEADER_FONT };
    }
  });

  // Style every data cell in the three output columns. Apply $#,##0.00
  // number format to the two numeric columns at the same time.
  for (let r = 1; r < outputRows.length; r++) {
    const xlsxRowNum = r + 1;
    [oceanColLetter, airColLetter, notesColLetter].forEach((letter) => {
      const ref = letter + xlsxRowNum;
      // Ensure the cell exists even if blank (so the fill renders).
      if (!outputSheet[ref]) {
        outputSheet[ref] = { t: 's', v: '' };
      }
      outputSheet[ref].s = { fill: DATA_FILL };
    });
    const oceanCell = outputSheet[oceanColLetter + xlsxRowNum];
    const airCell   = outputSheet[airColLetter + xlsxRowNum];
    if (oceanCell && typeof oceanCell.v === 'number') oceanCell.z = '"$"#,##0.00';
    if (airCell   && typeof airCell.v === 'number')   airCell.z = '"$"#,##0.00';
  }

  // Column widths roughly based on header length; Notes a bit wider.
  outputSheet['!cols'] = outputHeader.map((h) => ({
    wch: h === OUTPUT_HEADER_NOTES ? 40 : Math.max(12, String(h).length + 2)
  }));

  // Assemble workbook + write.
  const outputWb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(outputWb, outputSheet, sheetName.substring(0, 31) || 'Logistics');
  const wbout = XLSX.write(outputWb, { bookType: 'xlsx', type: 'array' });
  computedBlob = new Blob([wbout], { type: 'application/octet-stream' });

  // Filename.
  const baseName = file.name.replace(/\.(xlsx|xls)$/i, '');
  const ts = new Date().toISOString().slice(0, 19).replace(/[:T]/g, '-');
  computedFileName = `${baseName}__logistics_${ts}.xlsx`;

  const elapsedMs = Math.round(performance.now() - t0);
  const totalRows = okCount + issueCount + blankCount;

  showStatus('', 'ok');
  statusEl.classList.add('hidden');
  resultEl.classList.remove('hidden');

  const issueSentence = issueCount > 0
    ? ` <strong>${issueCount.toLocaleString()}</strong> row${issueCount === 1 ? '' : 's'} had issues — see the <code>Notes</code> column for details.`
    : '';
  const blankNote = blankCount > 0
    ? ` (${blankCount.toLocaleString()} blank row${blankCount === 1 ? '' : 's'} preserved.)`
    : '';

  resultSummary.innerHTML =
    `Computed <strong>${okCount.toLocaleString()} of ${totalRows.toLocaleString()}</strong> rows ` +
    `in <strong>${elapsedMs.toLocaleString()} ms</strong>.` +
    issueSentence +
    blankNote +
    ` Output contains your input columns plus <code>${OUTPUT_HEADER_OCEAN}</code>, ` +
    `<code>${OUTPUT_HEADER_AIR}</code>, and <code>${OUTPUT_HEADER_NOTES}</code>.`;
}
