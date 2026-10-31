import { test } from 'node:test';
import assert from 'node:assert/strict';
import * as XLSX from 'xlsx';
import { addAllSheet, describeSheet, parseDateText, verifyAll } from '../src/all-sheet.js';

const POST_HEADER = ['ID', 'Source', 'Author', 'Content', 'Date', 'Day', 'Sentiment', 'Category', 'Keywords', 'URL', 'Comment Count', 'Like Count', 'Share Count', 'View Count'];
const COMMENT_HEADER = ['post_url', 'post_content', 'author', 'comment_url', 'content', 'category', 'sentiment', 'comment_date'];

/** Build an .xlsx buffer from { sheetName: arrayOfRows }; `edit(wb)` can adjust the workbook before writing. */
function makeXlsx(sheets, edit = () => {}) {
  const wb = XLSX.utils.book_new();
  for (const [name, aoa] of Object.entries(sheets)) {
    XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet(aoa), name);
  }
  edit(wb);
  return XLSX.write(wb, { type: 'buffer', bookType: 'xlsx' });
}

// Date serials compared to the second, avoiding floating-point noise.
const secs = (v) => (typeof v === 'number' ? Math.round(v * 86400) : v);

function readBack(buffer) {
  const wb = XLSX.read(buffer, { type: 'buffer', cellNF: true });
  // Raw cell values: sheet_to_json would turn date-formatted numbers into JS dates.
  const rows = (name) => {
    const range = XLSX.utils.decode_range(wb.Sheets[name]['!ref']);
    return Array.from({ length: range.e.r + 1 }, (_, r) => Array.from({ length: range.e.c + 1 }, (_, c) => wb.Sheets[name][XLSX.utils.encode_cell({ r, c })]?.v ?? null));
  };
  // "All" as objects keyed by header, without empty cells.
  const [header, ...data] = rows('All');
  const all = data.map((r) => Object.fromEntries(header.map((h, i) => [h, r[i]]).filter(([, v]) => v !== null)));
  return { wb, rows, header, all };
}

test('merges post and comment sheets of every platform into a first "All" sheet', () => {
  const input = makeXlsx({
    Facebook: [POST_HEADER, [1, 'Facebook', 'Ann', 'fb post', 45950.5, '2025-10-20', 'Müsbət', 'digər', 'm10', 'https://fb/1', 2, 3, 0, null]],
    Telegram: [POST_HEADER, [2, 'Telegram', 'Chan', 'tg post', 45951.25, '2025-10-21', 'Neytral', 'digər', 'm10', 'https://t.me/2', '', '', '', 900]],
    'Facebook comments': [COMMENT_HEADER, ['https://fb/1', 'fb post', 'Bob', 'https://fb/1/c1', 'nice', 'ictimai rəy', 1, '2025-10-20 13:30:00']],
  });
  const { buffer, report } = addAllSheet(input, 'export.xlsx');
  const { wb, rows, header, all } = readBack(buffer);

  assert.deepEqual(wb.SheetNames, ['All', 'Facebook', 'Telegram', 'Facebook comments']);
  assert.deepEqual(header, ['Platform', 'Type', 'ID', 'Author', 'Content', 'Date', 'Day', 'Sentiment', 'Category', 'Keywords', 'URL', 'Comment Count', 'Like Count', 'Share Count', 'View Count', 'Post URL', 'Post content']);
  assert.deepEqual(all.map((r) => [r.Platform, r.Type, r.Content]), [['Facebook', 'Post', 'fb post'], ['Telegram', 'Post', 'tg post'], ['Facebook', 'Comment', 'nice']]);
  assert.deepEqual(all[2], {
    Platform: 'Facebook', Type: 'Comment', Author: 'Bob', Content: 'nice', Date: 45950 + 13.5 / 24, Day: 45950,
    Sentiment: 'Müsbət', Category: 'İctimai rəy', URL: 'https://fb/1/c1', 'Post URL': 'https://fb/1', 'Post content': 'fb post',
  });
  // Telegram's empty counters stay empty.
  assert.deepEqual(all[1], {
    Platform: 'Telegram', Type: 'Post', ID: 2, Author: 'Chan', Content: 'tg post', Date: 45951.25, Day: 45951,
    Sentiment: 'Neytral', Category: 'Digər', Keywords: 'm10', URL: 'https://t.me/2', 'View Count': 900,
  });
  // The original sheets are untouched.
  assert.deepEqual(rows('Facebook comments')[1], ['https://fb/1', 'fb post', 'Bob', 'https://fb/1/c1', 'nice', 'ictimai rəy', 1, '2025-10-20 13:30:00']);
  assert.equal(report.totalRows, 3);
  assert.equal(report.verified, true);
  assert.deepEqual(report.sheets.map((s) => [s.name, s.platform, s.type, s.rows]), [
    ['Facebook', 'Facebook', 'Post', 1], ['Telegram', 'Telegram', 'Post', 1], ['Facebook comments', 'Facebook', 'Comment', 1],
  ]);
});

test('leaves out YouTube\'s Title and a Source column that only repeats the platform', () => {
  const input = makeXlsx({
    YouTube: [['ID', 'Source', 'Author', 'Title', 'Content'], [1, 'YouTube', 'Baku TV', 'Video title', 'description']],
    TikTok: [['Source', 'Content'], ['Tiktok', 'clip']],
    News: [['Source', 'Content'], ['report.az', 'article']],
  });
  const { buffer, report } = addAllSheet(input, 'export.xlsx');
  const { header, all, rows } = readBack(buffer);

  assert.deepEqual(header, ['Platform', 'Type', 'ID', 'Author', 'Content', 'Source']);
  assert.deepEqual(all, [
    { Platform: 'YouTube', Type: 'Post', ID: 1, Author: 'Baku TV', Content: 'description' },
    { Platform: 'TikTok', Type: 'Post', Content: 'clip' },
    { Platform: 'News', Type: 'Post', Content: 'article', Source: 'report.az' },
  ]);
  assert.deepEqual(report.sheets.map((s) => s.leftOut), [['Source', 'Title'], ['Source'], []]);
  assert.deepEqual(rows('YouTube')[1], [1, 'YouTube', 'Baku TV', 'Video title', 'description']);
});

test('matches columns by name regardless of order, case and underscores', () => {
  const input = makeXlsx({
    Threads: [['Category', 'content', 'URL', 'Like_Count'], ['digər', 'thread', 'https://th/1', 4]],
    'YouTube comments': [['like_count', 'Comment URL', 'Content'], [7, 'https://yt/c', 'cool']],
  });
  const { all } = readBack(addAllSheet(input, 'x.xlsx').buffer);
  assert.deepEqual(all, [
    { Platform: 'Threads', Type: 'Post', Content: 'thread', Category: 'Digər', URL: 'https://th/1', 'Like Count': 4 },
    { Platform: 'YouTube', Type: 'Comment', Content: 'cool', URL: 'https://yt/c', 'Like Count': 7 },
  ]);
});

test('unifies sentiment and category spellings', () => {
  const input = makeXlsx({
    Posts: [['Sentiment', 'Category'], ['NEYTRAL', 'ictimai_rəy'], ['Neqativ', 'digər'], ['müsbət', 'kadr_siyasəti'], ['unknown', 'fraud']],
    'Posts comments': [['sentiment', 'category'], [1, 'İctimai rəy'], [0, 'Digər'], [-1, 'Fraud'], [0.5, 'Faiz dərəcəsi/Kredit']],
  });
  const { buffer, report } = addAllSheet(input, 'x.xlsx');
  const { all } = readBack(buffer);
  assert.deepEqual(all.map((r) => [r.Sentiment, r.Category]), [
    ['Neytral', 'İctimai rəy'], ['Mənfi', 'Digər'], ['Müsbət', 'Kadr siyasəti'], ['unknown', 'Fraud'],
    ['Müsbət', 'İctimai rəy'], ['Neytral', 'Digər'], ['Mənfi', 'Fraud'], [0.5, 'Faiz dərəcəsi/Kredit'],
  ]);
  assert.equal(report.changes.find((c) => c.column === 'Sentiment').count, 6);
});

test('turns date texts into real dates, fills Day from Date, and keeps unreadable dates as text', () => {
  const input = makeXlsx({
    Facebook: [['Date', 'Day'], [45950.75, '2025-10-20'], [null, '2025-10-21']],
    'Facebook comments': [['comment_date_text'], ['17:03, 23 Dekabr 2025, Çərşənbə axşamı'], ['5 İyun 2025'], ['vaxt yoxdur']],
  });
  const { buffer, report } = addAllSheet(input, 'x.xlsx');
  const { all, wb } = readBack(buffer);
  assert.deepEqual(all.map((r) => [secs(r.Date), secs(r.Day)]), [
    [secs(45950.75), secs(45950)],
    [undefined, secs(45951)],
    [secs(46014 + (17 * 60 + 3) / 1440), secs(46014)],
    [secs(45813), secs(45813)],
    ['vaxt yoxdur', undefined],
  ]);
  assert.equal(wb.Sheets.All.C4.w, '2025-12-23 17:03:00');
  assert.equal(wb.Sheets.All.D4.w, '2025-12-23');
  assert.equal(report.unreadDates, 1);
  assert.equal(report.daysFilled, 2);
});

test('parses the date formats found in exports', () => {
  const parsed = (text) => secs(parseDateText(text));
  assert.equal(parsed('2025-09-24 10:03:37'), secs(45924) + 10 * 3600 + 3 * 60 + 37);
  assert.equal(parsed('2025-09-24'), secs(45924));
  assert.equal(parsed('24.09.2025 10:03'), secs(45924) + 10 * 3600 + 3 * 60);
  assert.equal(parsed('09:07, 12 Dekabr 2025, Cümə'), secs(46003) + 9 * 3600 + 7 * 60);
  assert.equal(parsed('12 December 2025'), secs(46003));
  assert.equal(parseDateText('31 Fevral 2025'), null);
  assert.equal(parseDateText('2025-13-01'), null);
  assert.equal(parseDateText('soon'), null);
});

test('keeps unknown columns at the end and leaves out columns no sheet has', () => {
  const input = makeXlsx({
    Instagram: [['Content', 'Reach', 'Saves'], ['ig', 100, '5']],
    VK: [['Content', 'reach', 'Reposts'], ['vk', 7, 2]],
  });
  const { header, all } = readBack(addAllSheet(input, 'x.xlsx').buffer);
  assert.deepEqual(header, ['Platform', 'Type', 'Content', 'Reach', 'Saves', 'Reposts']);
  assert.deepEqual(all, [
    { Platform: 'Instagram', Type: 'Post', Content: 'ig', Reach: 100, Saves: '5' },
    { Platform: 'VK', Type: 'Post', Content: 'vk', Reach: 7, Reposts: 2 },
  ]);
});

test('puts Company and Source file columns from combined files first', () => {
  const input = makeXlsx({
    Facebook: [['Company', 'Content'], ['Pasha Bank', 'a']],
    'Facebook comments': [['Company', 'content'], ['Kapital Bank', 'b']],
  });
  const { header, all } = readBack(addAllSheet(input, 'combined.xlsx').buffer);
  assert.deepEqual(header, ['Company', 'Platform', 'Type', 'Content']);
  assert.deepEqual(all.map((r) => r.Company), ['Pasha Bank', 'Kapital Bank']);
});

test('rebuilds an existing "All" sheet instead of merging it into itself', () => {
  const first = addAllSheet(makeXlsx({ Facebook: [['Content'], ['a'], ['b']] }), 'x.xlsx').buffer;
  const { buffer, report } = addAllSheet(first, 'x.xlsx');
  const { wb, all } = readBack(buffer);
  assert.deepEqual(wb.SheetNames, ['All', 'Facebook']);
  assert.equal(all.length, 2);
  assert.equal(report.replacedExisting, true);
});

test('keeps sheet-scoped names such as autofilters on their own sheets', () => {
  const input = makeXlsx({ Facebook: [['Content'], ['a']], Instagram: [['Content'], ['b']] }, (wb) => {
    wb.Workbook = { Names: [{ Name: '_xlnm._FilterDatabase', Sheet: 1, Hidden: true, Ref: 'Instagram!$A$1:$A$2' }] };
  });
  const { wb } = readBack(addAllSheet(input, 'x.xlsx').buffer);
  const filter = wb.Workbook.Names.find((n) => n.Ref.startsWith('Instagram'));
  assert.equal(wb.SheetNames[filter.Sheet], 'Instagram');
});

test('handles sheets whose declared range is larger than their data', () => {
  const input = makeXlsx({ Facebook: [['Content'], ['a']] }, (wb) => { wb.Sheets.Facebook['!ref'] = 'A1:N1000'; });
  const { all } = readBack(addAllSheet(input, 'x.xlsx').buffer);
  assert.deepEqual(all, [{ Platform: 'Facebook', Type: 'Post', Content: 'a' }]);
});

test('keeps dates right in 1904-date-system workbooks', () => {
  const input = makeXlsx({ Facebook: [['Day'], ['2025-10-20']] }, (wb) => { wb.Workbook = { WBProps: { date1904: true } }; });
  const { wb } = readBack(addAllSheet(input, 'x.xlsx').buffer);
  assert.equal(wb.Sheets.All.C2.w, '2025-10-20');
});

test('verification catches a value lost from "All"', () => {
  const input = makeXlsx({ Facebook: [['Content', 'URL'], ['a', 'u1'], ['b', 'u2']] });
  const { buffer } = addAllSheet(input, 'x.xlsx');
  const wb = XLSX.read(buffer, { type: 'buffer', dense: true });
  delete wb.Sheets.All['!data'][2][3];
  const tampered = XLSX.write(wb, { type: 'buffer', bookType: 'xlsx' });
  assert.throws(() => verifyAll(input, tampered, 'x.xlsx'), /Verification failed/);
});

test('verification catches a changed original sheet', () => {
  const input = makeXlsx({ Facebook: [['Content'], ['a']] });
  const wb = XLSX.read(addAllSheet(input, 'x.xlsx').buffer, { type: 'buffer', dense: true });
  wb.Sheets.Facebook['!data'][1][0].v = 'changed';
  const tampered = XLSX.write(wb, { type: 'buffer', bookType: 'xlsx' });
  assert.throws(() => verifyAll(input, tampered, 'x.xlsx'), /Verification failed: sheet "Facebook" changed/);
});

test('describes sheets by name', () => {
  assert.deepEqual(describeSheet('Instagram comments'), { platform: 'Instagram', type: 'Comment' });
  assert.deepEqual(describeSheet('YouTube Comments '), { platform: 'YouTube', type: 'Comment' });
  assert.deepEqual(describeSheet('Telegram'), { platform: 'Telegram', type: 'Post' });
});

test('rejects files without data', () => {
  assert.throws(() => addAllSheet(makeXlsx({ Empty: [] }), 'empty.xlsx'), /no sheets with data/);
});

test('cuts texts longer than Excel allows in the original sheets and in "All"', () => {
  const long = 'x'.repeat(40000);
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet([['ID', 'Content'], [1, long]]), 'Facebook');
  // Excel can't hold such a cell, but files written by other tools can; ODS lets the test write one.
  const { buffer, report } = addAllSheet(XLSX.write(wb, { type: 'buffer', bookType: 'ods' }), 'posts.ods');

  const out = XLSX.read(buffer, { type: 'buffer' });
  assert.equal(out.Sheets.Facebook.B2.v, 'x'.repeat(32767));
  assert.deepEqual(report.truncatedTexts, [{ sheet: 'All', cells: ['D2'] }, { sheet: 'Facebook', cells: ['B2'] }]);
  assert.equal(report.verified, true);
});
