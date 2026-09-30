import { test } from 'node:test';
import assert from 'node:assert/strict';
import * as XLSX from 'xlsx';
import { combineFiles, combine, parseWorkbook, verify, SOURCE_COLUMN } from '../src/combine.js';

/** Build an .xlsx buffer from { sheetName: arrayOfRows }. */
function makeXlsx(sheets, { date1904 = false, bookType = 'xlsx' } = {}) {
  const wb = XLSX.utils.book_new();
  for (const [name, aoa] of Object.entries(sheets)) {
    XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet(aoa), name);
  }
  if (date1904) wb.Workbook = { WBProps: { date1904: true } };
  return XLSX.write(wb, { type: 'buffer', bookType });
}

function readBack(buffer) {
  const wb = XLSX.read(buffer, { type: 'buffer' });
  return Object.fromEntries(wb.SheetNames.map((n) => [n, XLSX.utils.sheet_to_json(wb.Sheets[n], { header: 1, defval: null })]));
}

test('combines rows of the same sheet across files, in file order', () => {
  const a = makeXlsx({ Facebook: [['Date', 'Clicks'], ['d1', 1], ['d2', 2]], Instagram: [['Date', 'Likes'], ['d1', 10]] });
  const b = makeXlsx({ Facebook: [['Date', 'Clicks'], ['d3', 3]], Instagram: [['Date', 'Likes'], ['d2', 20], ['d3', 30]] });
  const { buffer, report } = combineFiles([{ name: 'a.xlsx', buffer: a }, { name: 'b.xlsx', buffer: b }]);

  const out = readBack(buffer);
  assert.deepEqual(Object.keys(out), ['Facebook', 'Instagram']);
  assert.deepEqual(out.Facebook, [['Date', 'Clicks'], ['d1', 1], ['d2', 2], ['d3', 3]]);
  assert.deepEqual(out.Instagram, [['Date', 'Likes'], ['d1', 10], ['d2', 20], ['d3', 30]]);
  assert.equal(report.totalRows, 6);
  assert.equal(report.verified, true);
  assert.deepEqual(report.sheets.map((s) => s.perFile), [[2, 1], [1, 2]]);
});

test('matches columns by header even when order differs, and keeps extra columns', () => {
  const a = makeXlsx({ LinkedIn: [['Date', 'Clicks', 'Reach'], ['d1', 1, 100]] });
  const b = makeXlsx({ LinkedIn: [['Reach', 'date ', 'Saves'], [200, 'd2', 5]] });
  const { buffer, report } = combineFiles([{ name: 'a.xlsx', buffer: a }, { name: 'b.xlsx', buffer: b }]);

  assert.deepEqual(readBack(buffer).LinkedIn, [
    ['Date', 'Clicks', 'Reach', 'Saves'],
    ['d1', 1, 100, null],
    ['d2', null, 200, 5],
  ]);
  assert.deepEqual(report.sheets[0].missingColumns, [
    { file: 'a.xlsx', columns: ['Saves'] },
    { file: 'b.xlsx', columns: ['Clicks'] },
  ]);
});

test('sheets present only in some files are still included; sheet names match case-insensitively', () => {
  const a = makeXlsx({ Facebook: [['A'], [1]] });
  const b = makeXlsx({ facebook: [['A'], [2]], TikTok: [['B'], ['x']] });
  const { buffer, report } = combineFiles([{ name: 'a.xlsx', buffer: a }, { name: 'b.xlsx', buffer: b }]);

  assert.deepEqual(readBack(buffer), { Facebook: [['A'], [1], [2]], TikTok: [['B'], ['x']] });
  assert.deepEqual(report.sheets[1].perFile, [null, 1]);
});

test('keeps data in columns without a header and in duplicate-named columns', () => {
  const a = makeXlsx({ S: [['Name', null, 'Name'], ['n1', 'orphan', 'n1-dup']] });
  const { buffer } = combineFiles([{ name: 'a.xlsx', buffer: a }]);
  assert.deepEqual(readBack(buffer).S, [['Name', 'Column B', 'Name (2)'], ['n1', 'orphan', 'n1-dup']]);
});

test('skips fully blank rows but keeps everything else, and reports the skipped count', () => {
  const a = makeXlsx({ S: [[], ['H1', 'H2'], ['x', null], [], [null, 'y'], [], []] });
  const { buffer, report } = combineFiles([{ name: 'a.xlsx', buffer: a }]);
  assert.deepEqual(readBack(buffer).S, [['H1', 'H2'], ['x', null], [null, 'y']]);
  assert.equal(report.sheets[0].blankRowsSkipped, 1);
});

test('adds a source file column when requested', () => {
  const a = makeXlsx({ S: [['A'], [1]] });
  const b = makeXlsx({ S: [['A'], [2]] });
  const { buffer } = combineFiles([{ name: 'a.xlsx', buffer: a }, { name: 'b.xlsx', buffer: b }], { addSourceColumn: true });
  assert.deepEqual(readBack(buffer).S, [[SOURCE_COLUMN, 'A'], ['a.xlsx', 1], ['b.xlsx', 2]]);
});

test('preserves value types, number formats, and odd strings', () => {
  const wb = XLSX.utils.book_new();
  const ws = XLSX.utils.aoa_to_sheet([
    ['Date', 'Rate', 'Text', 'Flag', 'Big'],
    [45292, 0.125, '  leading/trailing  ', true, 1234567890.123456],
    [45293, 0.5, 'line1\nline2 <&> "q" 🙂', false, -0.000001],
    [45294, 1, '00123', null, 1e21],
  ]);
  for (const r of [2, 3, 4]) {
    ws[`A${r}`].z = 'yyyy-mm-dd';
    ws[`B${r}`].z = '0.0%';
  }
  XLSX.utils.book_append_sheet(wb, ws, 'S');
  const input = XLSX.write(wb, { type: 'buffer', bookType: 'xlsx' });

  const { buffer } = combineFiles([{ name: 'a.xlsx', buffer: input }]);
  const out = XLSX.read(buffer, { type: 'buffer', cellNF: true }).Sheets.S;
  assert.equal(out.A2.v, 45292);
  assert.equal(out.A2.w, '2024-01-01');
  assert.equal(out.B2.w, '12.5%');
  assert.equal(out.C2.v, '  leading/trailing  ');
  assert.equal(out.C3.v, 'line1\nline2 <&> "q" 🙂');
  assert.equal(out.C4.v, '00123');
  assert.equal(out.C4.t, 's');
  assert.equal(out.D2.v, true);
  assert.equal(out.E2.v, 1234567890.123456);
  assert.equal(out.E3.v, -0.000001);
  assert.equal(out.E4.v, 1e21);
});

test('converts dates from 1904-date-system workbooks', () => {
  const wb = XLSX.utils.book_new();
  const ws = XLSX.utils.aoa_to_sheet([['Date', 'N'], [43830, 43830]]); // 2024-01-01 in the 1904 system
  ws.A2.z = 'yyyy-mm-dd';
  XLSX.utils.book_append_sheet(wb, ws, 'S');
  wb.Workbook = { WBProps: { date1904: true } };
  const input = XLSX.write(wb, { type: 'buffer', bookType: 'xlsx' });

  const out = XLSX.read(combineFiles([{ name: 'a.xlsx', buffer: input }]).buffer, { type: 'buffer', cellNF: true }).Sheets.S;
  assert.equal(out.A2.w, '2024-01-01');
  assert.equal(out.B2.v, 43830); // plain numbers are not dates and must not change
});

test('reads legacy .xls files', () => {
  const a = makeXlsx({ S: [['A'], [1]] }, { bookType: 'biff8' });
  const b = makeXlsx({ S: [['A'], [2]] });
  const { buffer } = combineFiles([{ name: 'a.xls', buffer: a }, { name: 'b.xlsx', buffer: b }]);
  assert.deepEqual(readBack(buffer).S, [['A'], [1], [2]]);
});

test('rejects unreadable files instead of silently skipping them', () => {
  assert.throws(
    () => combineFiles([{ name: 'broken.xlsx', buffer: Buffer.from('PK\x03\x04garbage') }]),
    /Cannot read "broken.xlsx"/,
  );
});

test('verification catches a lost value', () => {
  const files = [parseWorkbook(makeXlsx({ S: [['A', 'B'], [1, 2], [3, 4]] }), 'a.xlsx')];
  const { workbook } = combine(files);
  delete workbook.Sheets.S['!data'][2][1]; // simulate a dropped cell
  const buffer = XLSX.write(workbook, { type: 'buffer', bookType: 'xlsx' });
  assert.throws(() => verify(files, buffer), /Verification failed/);
});

test('verification catches an extra or missing row', () => {
  const files = [parseWorkbook(makeXlsx({ S: [['A'], [1], [2]] }), 'a.xlsx')];
  const { workbook } = combine(files);
  workbook.Sheets.S['!data'].pop();
  workbook.Sheets.S['!ref'] = 'A1:A2';
  assert.throws(() => verify(files, XLSX.write(workbook, { type: 'buffer', bookType: 'xlsx' })), /Verification failed/);
});

test('handles many files with many rows', () => {
  const inputs = [];
  let expected = 0;
  for (let f = 0; f < 20; f++) {
    const rows = 500 + f * 37;
    expected += rows;
    const aoa = [['Id', 'File', 'Value']];
    for (let i = 0; i < rows; i++) aoa.push([i, f, Math.random()]);
    inputs.push({ name: `f${f}.xlsx`, buffer: makeXlsx({ Facebook: aoa, Instagram: aoa }) });
  }
  const { report } = combineFiles(inputs);
  assert.equal(report.sheets[0].totalRows, expected);
  assert.equal(report.sheets[1].totalRows, expected);
  assert.equal(report.verified, true);
});
