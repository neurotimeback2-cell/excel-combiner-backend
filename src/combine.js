import * as XLSX from 'xlsx';

export const SOURCE_COLUMN = 'Source file';
export const COMPANY_COLUMN = 'Company';

const EXCEL_MAX_ROWS = 1048576;
const EXCEL_MAX_COLS = 16384;
// Serial date difference between the 1904 and 1900 date systems.
const DATE_1904_OFFSET = 1462;

const READ_OPTS = { type: 'buffer', dense: true, cellNF: true, cellDates: false, cellFormula: false, cellHTML: false };

/** Error caused by the uploaded files (shown to the user as-is). */
export class UserError extends Error {}

const normKey = (s) => String(s).trim().replace(/\s+/g, ' ').toLowerCase();
const normNewlines = (s) => s.replace(/\r\n?/g, '\n');

/** First of `base`, `base (2)`, `base (3)`, … whose key is not in `taken`. */
function uniqueName(base, taken) {
  let name = base;
  for (let n = 2; taken.has(normKey(name)); n++) name = `${base} (${n})`;
  return name;
}

/** "Capital Bank .xlsx" → "Capital Bank" */
export function companyName(fileName) {
  return fileName.replace(/\.[^.]+$/, '').trim() || fileName;
}

/** Columns placed before the data columns, filled from the file each row came from. */
function extraColumns({ addCompanyColumn = false, addSourceColumn = false }) {
  return [
    ...(addCompanyColumn ? [{ name: COMPANY_COLUMN, value: (file) => companyName(file.name) }] : []),
    ...(addSourceColumn ? [{ name: SOURCE_COLUMN, value: (file) => file.name }] : []),
  ];
}

function isEmpty(cell) {
  return !cell || cell.t === 'z' || cell.v == null || (cell.t === 's' && cell.v === '');
}

function cellText(cell) {
  return isEmpty(cell) ? '' : String(cell.w ?? cell.v).trim();
}

/** Keep only the value, number format and hyperlink; formulas are replaced by their computed values. */
function cleanCell(cell, date1904) {
  const out = { t: cell.t, v: cell.v };
  if (cell.z && cell.z !== 'General') out.z = cell.z;
  if (cell.l) out.l = cell.l;
  if (date1904 && out.t === 'n' && out.z && XLSX.SSF.is_date(out.z)) out.v += DATE_1904_OFFSET;
  return out;
}

/**
 * The first non-empty row is the header. Every following non-empty row is data.
 * Columns without a header but with data get a generated name, so nothing is dropped.
 */
function parseSheet(ws, name, date1904) {
  const sheet = { name, key: normKey(name), columns: [], rows: [], blankRows: 0 };
  const data = ws['!data'] ?? [];
  const isBlankRow = (row) => !row || !row.some((cell) => !isEmpty(cell));

  let headerIdx = 0;
  while (headerIdx < data.length && isBlankRow(data[headerIdx])) headerIdx++;
  if (headerIdx >= data.length) return sheet;
  const headerRow = data[headerIdx];

  const dataRows = [];
  const colHasData = [];
  for (let r = headerIdx + 1; r < data.length; r++) {
    const row = data[r];
    if (isBlankRow(row)) {
      // Blank rows after the last data row are just the sheet's formatted area, not skipped data.
      dataRows.push(null);
      continue;
    }
    row.forEach((cell, c) => { if (!isEmpty(cell)) colHasData[c] = true; });
    dataRows.push(row);
  }
  while (dataRows.length && dataRows[dataRows.length - 1] === null) dataRows.pop();

  const width = Math.max(headerRow.length, colHasData.length);
  const usedKeys = new Set();
  const colIndexes = [];
  for (let c = 0; c < width; c++) {
    const text = cellText(headerRow[c]);
    if (!text && !colHasData[c]) continue;
    const colName = uniqueName(text || `Column ${XLSX.utils.encode_col(c)}`, usedKeys);
    usedKeys.add(normKey(colName));
    sheet.columns.push({ name: colName, key: normKey(colName) });
    colIndexes.push(c);
  }

  for (const row of dataRows) {
    if (!row) {
      sheet.blankRows++;
      continue;
    }
    sheet.rows.push(colIndexes.map((c) => (isEmpty(row[c]) ? null : cleanCell(row[c], date1904))));
  }
  return sheet;
}

export function parseWorkbook(buffer, fileName) {
  let wb;
  try {
    wb = XLSX.read(buffer, READ_OPTS);
  } catch (err) {
    throw new UserError(`Cannot read "${fileName}": ${err.message}`);
  }
  const date1904 = Boolean(wb.Workbook?.WBProps?.date1904);
  const sheets = wb.SheetNames
    .filter((name) => !wb.Sheets[name]['!type'] || wb.Sheets[name]['!type'] === 'sheet') // skip chart/macro sheets
    .map((name) => parseSheet(wb.Sheets[name], name, date1904));
  return { name: fileName, sheets };
}

/** Merge parsed files sheet-by-sheet (matched by sheet name) and column-by-column (matched by header). */
export function combine(files, opts = {}) {
  const extras = extraColumns(opts);
  const outSheets = new Map();

  files.forEach((file, fileIdx) => {
    for (const sheet of file.sheets) {
      let out = outSheets.get(sheet.key);
      if (!out) {
        out = { name: sheet.name, columns: [], colIndex: new Map(), rows: [], perFile: new Map(), blankRows: 0 };
        outSheets.set(sheet.key, out);
      }
      const mapping = sheet.columns.map((col) => {
        if (!out.colIndex.has(col.key)) {
          out.colIndex.set(col.key, out.columns.length);
          out.columns.push(col.name);
        }
        return out.colIndex.get(col.key);
      });
      for (const row of sheet.rows) {
        const outRow = [];
        row.forEach((cell, j) => { if (cell) outRow[mapping[j]] = cell; });
        out.rows.push({ fileIdx, cells: outRow });
      }
      const stats = out.perFile.get(fileIdx) ?? { rows: 0, keys: new Set() };
      stats.rows += sheet.rows.length;
      sheet.columns.forEach((col) => stats.keys.add(col.key));
      out.perFile.set(fileIdx, stats);
      out.blankRows += sheet.blankRows;
    }
  });

  const workbook = XLSX.utils.book_new();
  const offset = extras.length;
  const reportSheets = [];

  for (const out of outSheets.values()) {
    // Renamed to "Company (2)" etc. if the uploaded sheets already have a column with that name.
    const header = [...extras.map((col) => uniqueName(col.name, out.colIndex)), ...out.columns];
    if (out.rows.length + 1 > EXCEL_MAX_ROWS) {
      throw new UserError(`Sheet "${out.name}" would have ${out.rows.length} rows, more than Excel's limit of ${EXCEL_MAX_ROWS - 1}.`);
    }
    if (header.length > EXCEL_MAX_COLS) {
      throw new UserError(`Sheet "${out.name}" would have ${header.length} columns, more than Excel's limit of ${EXCEL_MAX_COLS}.`);
    }

    const data = [header.map((v) => ({ t: 's', v }))];
    for (const { fileIdx, cells } of out.rows) {
      const row = extras.map((col) => ({ t: 's', v: col.value(files[fileIdx]) }));
      cells.forEach((cell, j) => { row[offset + j] = cell; });
      data.push(row);
    }

    const ws = { '!data': data, '!ref': 'A1' };
    if (header.length) {
      const range = { s: { r: 0, c: 0 }, e: { r: data.length - 1, c: header.length - 1 } };
      ws['!ref'] = XLSX.utils.encode_range(range);
      ws['!autofilter'] = { ref: ws['!ref'] };
      ws['!cols'] = columnWidths(data, header.length);
    }
    XLSX.utils.book_append_sheet(workbook, ws, out.name);

    const allKeys = [...out.colIndex.keys()];
    reportSheets.push({
      name: out.name,
      columns: out.columns.length,
      totalRows: out.rows.length,
      blankRowsSkipped: out.blankRows,
      perFile: files.map((_, i) => out.perFile.get(i)?.rows ?? null),
      missingColumns: [...out.perFile.entries()]
        .map(([i, stats]) => ({
          file: files[i].name,
          columns: allKeys.filter((k) => !stats.keys.has(k)).map((k) => out.columns[out.colIndex.get(k)]),
        }))
        .filter((m) => m.columns.length),
    });
  }

  if (!reportSheets.length) throw new UserError('The uploaded files contain no worksheets.');

  const report = {
    files: files.map((f) => f.name),
    sheets: reportSheets,
    totalRows: reportSheets.reduce((sum, s) => sum + s.totalRows, 0),
  };
  return { workbook, report };
}

function columnWidths(data, width) {
  const widths = new Array(width).fill(10);
  for (const row of data.slice(0, 200)) {
    row.forEach((cell, c) => {
      const len = cell ? String(cell.v).split('\n')[0].length + 2 : 0;
      if (len > widths[c]) widths[c] = Math.min(len, 60);
    });
  }
  return widths.map((wch) => ({ wch }));
}

function sameValue(expected, actual) {
  if (isEmpty(actual) || expected.t !== actual.t) return false;
  if (expected.t === 's') return normNewlines(expected.v) === normNewlines(actual.v);
  return expected.v === actual.v;
}

/**
 * Read the generated file back and check that every non-empty input cell is present,
 * in the right sheet, row and column, and that the output contains nothing extra.
 */
export function verify(files, buffer, opts = {}) {
  const fail = (msg) => { throw new Error(`Verification failed: ${msg}. The combined file was not produced.`); };
  const wb = XLSX.read(buffer, READ_OPTS);
  const extras = extraColumns(opts);
  const offset = extras.length;

  const outByKey = new Map();
  for (const name of wb.SheetNames) {
    const data = wb.Sheets[name]['!data'] ?? [];
    const header = data[0] ?? [];
    const colByKey = new Map();
    for (let c = offset; c < header.length; c++) {
      if (!isEmpty(header[c])) colByKey.set(normKey(cellText(header[c])), c);
    }
    let lastRow = data.length - 1;
    while (lastRow > 0 && !data[lastRow]) lastRow--;
    let cellCount = 0;
    for (let r = 1; r <= lastRow; r++) {
      (data[r] ?? []).forEach((cell, c) => { if (c >= offset && !isEmpty(cell)) cellCount++; });
    }
    outByKey.set(normKey(name), { name, data, colByKey, dataRows: Math.max(lastRow, 0), cellCount, cursor: 1, expectedCells: 0 });
  }

  for (const file of files) {
    for (const sheet of file.sheets) {
      const out = outByKey.get(sheet.key);
      if (!out) fail(`sheet "${sheet.name}" from "${file.name}" is missing`);
      const cols = sheet.columns.map((col) => {
        const c = out.colByKey.get(col.key);
        if (c === undefined) fail(`column "${col.name}" of sheet "${sheet.name}" from "${file.name}" is missing`);
        return c;
      });
      for (const row of sheet.rows) {
        const outRow = out.data[out.cursor] ?? [];
        const where = `sheet "${out.name}", row ${out.cursor + 1} (from "${file.name}")`;
        extras.forEach((col, i) => {
          if (outRow[i]?.v !== col.value(file)) fail(`wrong "${col.name}" value in ${where}`);
        });
        row.forEach((cell, j) => {
          if (!cell) return;
          if (!sameValue(cell, outRow[cols[j]])) fail(`value in ${where}, column "${sheet.columns[j].name}" does not match the input`);
          out.expectedCells++;
        });
        out.cursor++;
      }
    }
  }

  for (const out of outByKey.values()) {
    if (out.cursor - 1 !== out.dataRows) fail(`sheet "${out.name}" has ${out.dataRows} data rows, expected ${out.cursor - 1}`);
    if (out.cellCount !== out.expectedCells) fail(`sheet "${out.name}" has ${out.cellCount} filled cells, expected ${out.expectedCells}`);
  }
}

/** @param {{name: string, buffer: Buffer}[]} inputs */
export function combineFiles(inputs, opts = {}) {
  const files = inputs.map((f) => parseWorkbook(f.buffer, f.name));
  const { workbook, report } = combine(files, opts);
  // Inline strings are decoded twice by SheetJS on read-back (e.g. literal
  // "&quot;" becomes a quote), causing a false verification failure. Shared
  // strings preserve those values through the write/read cycle.
  // SheetJS does not escape literal Excel _xHHHH_ sequences when writing XML.
  // Prefix their underscore with _x005F_ so Excel and SheetJS read them as text.
  // Restore the in-memory workbook after writing.
  const escaped = [];
  let buffer;
  try {
    for (const name of workbook.SheetNames) {
      for (const row of workbook.Sheets[name]['!data'] ?? []) {
        if (!row) continue;
        for (const cell of row) {
          if (cell?.t !== 's' || typeof cell.v !== 'string' || !/_x[\da-f]{4}_/i.test(cell.v)) continue;
          escaped.push([cell, cell.v]);
          cell.v = cell.v.replace(/_x[\da-f]{4}_/gi, (match) => `_x005F${match}`);
        }
      }
    }
    buffer = XLSX.write(workbook, { type: 'buffer', bookType: 'xlsx', compression: true, bookSST: true });
  } finally {
    for (const [cell, value] of escaped) cell.v = value;
  }
  verify(files, buffer, opts);
  report.verified = true;
  return { buffer, report };
}
