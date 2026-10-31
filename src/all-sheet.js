import * as XLSX from 'xlsx';
import {
  DATE_1904_OFFSET, EXCEL_MAX_ROWS, UserError,
  cellText, columnWidths, isEmpty, longTexts, normKey, parseSheets, readWorkbook, sameValue, uniqueName, writeWorkbook,
} from './combine.js';

export const ALL_SHEET = 'All';

const DATETIME_FORMAT = 'yyyy-mm-dd hh:mm:ss';
const DAY_FORMAT = 'yyyy-mm-dd';
const MAX_DATE_SERIAL = 2958465; // 9999-12-31

// Read everything SheetJS can write back, so the original sheets survive the round trip.
const KEEP_OPTS = { cellStyles: true, cellFormula: true };

/** Header comparison key: case, spaces, underscores and dashes are ignored ("comment_url" = "Comment URL"). */
const headerKey = (s) => String(s).toLowerCase().replace(/[\s_-]+/g, ' ').trim();

/*
 * Value normalizers. Each takes a source cell and returns the cell to put in "All"
 * (the same cell when nothing needs to change). They must be pure: verification calls them again.
 */

const MONTHS = new Map(
  [
    ['yanvar', 'fevral', 'mart', 'aprel', 'may', 'iyun', 'iyul', 'avqust', 'sentyabr', 'oktyabr', 'noyabr', 'dekabr'],
    ['january', 'february', 'march', 'april', 'may', 'june', 'july', 'august', 'september', 'october', 'november', 'december'],
  ].flatMap((names) => names.map((name, i) => [name, i + 1])),
);

/** Excel serial number (1900 date system), or null for a date that doesn't exist. */
function toSerial(y, mo, d, h = 0, mi = 0, s = 0) {
  if (h > 23 || mi > 59 || s > 59) return null;
  const ms = Date.UTC(y, mo - 1, d, h, mi, s);
  const check = new Date(ms);
  if (check.getUTCFullYear() !== y || check.getUTCMonth() !== mo - 1 || check.getUTCDate() !== d) return null;
  return ms / 86400000 + 25569;
}

const ISO_DATE = /^(\d{4})-(\d{1,2})-(\d{1,2})(?:[ T](\d{1,2}):(\d{2})(?::(\d{2}))?)?$/;
const DOTTED_DATE = /^(\d{1,2})\.(\d{1,2})\.(\d{4})(?:,?\s+(\d{1,2}):(\d{2})(?::(\d{2}))?)?$/;
// "17:03, 23 Dekabr 2025, Çərşənbə axşamı" (Facebook comments) or "23 Dekabr 2025"
const WORDY_DATE = /^(?:(\d{1,2}):(\d{2})(?::(\d{2}))?,?\s+)?(\d{1,2})\s+(\p{L}+)\s+(\d{4})(?:,.*)?$/u;

export function parseDateText(text) {
  const s = text.trim();
  const n = (v) => Number(v ?? 0);
  let m;
  if ((m = ISO_DATE.exec(s))) return toSerial(n(m[1]), n(m[2]), n(m[3]), n(m[4]), n(m[5]), n(m[6]));
  if ((m = DOTTED_DATE.exec(s))) return toSerial(n(m[3]), n(m[2]), n(m[1]), n(m[4]), n(m[5]), n(m[6]));
  if ((m = WORDY_DATE.exec(s))) {
    const month = MONTHS.get(m[5].toLocaleLowerCase('az')) ?? MONTHS.get(m[5].toLowerCase());
    return month ? toSerial(n(m[6]), month, n(m[4]), n(m[1]), n(m[2]), n(m[3])) : null;
  }
  return null;
}

/** Numbers stay as they are; recognised date texts become real dates. Anything else is kept unchanged. */
function toDate(cell, ctx, format) {
  let serial = null;
  if (cell.t === 'n') serial = cell.v;
  else if (cell.t === 's') {
    serial = parseDateText(cell.v);
    if (serial !== null && ctx.date1904) serial -= DATE_1904_OFFSET;
  }
  if (serial === null || !(serial > 0 && serial <= MAX_DATE_SERIAL)) return cell;
  return { t: 'n', v: serial, z: format };
}

const SENTIMENTS = new Map([
  ['1', 'Müsbət'], ['müsbət', 'Müsbət'], ['pozitiv', 'Müsbət'], ['positive', 'Müsbət'],
  ['0', 'Neytral'], ['neytral', 'Neytral'], ['neutral', 'Neytral'],
  ['-1', 'Mənfi'], ['mənfi', 'Mənfi'], ['neqativ', 'Mənfi'], ['negativ', 'Mənfi'], ['negative', 'Mənfi'],
]);

function toSentiment(cell) {
  const text = String(cell.v).trim();
  const label = SENTIMENTS.get(text.toLocaleLowerCase('az')) ?? SENTIMENTS.get(text.toLowerCase());
  return label ? { t: 's', v: label } : cell;
}

// "ictimai_rəy", "ictimai rəy" and "İctimai rəy" are the same category.
const categoryKey = (text) => text.trim().replace(/[\s_]+/g, ' ').toLocaleLowerCase('az');
function categoryLabel(text) {
  const s = text.trim().replace(/[\s_]+/g, ' ');
  return s.charAt(0).toLocaleUpperCase('az') + s.slice(1);
}

function toCategory(cell, ctx) {
  return cell.t === 's' && cell.v.trim() ? { t: 's', v: ctx.categories.get(categoryKey(cell.v)) } : cell;
}

function trimText(cell) {
  return cell.t === 's' && cell.v.trim() && cell.v.trim() !== cell.v ? { ...cell, v: cell.v.trim() } : cell;
}

function toCount(cell) {
  return cell.t === 's' && /^\s*\d+\s*$/.test(cell.v) ? { t: 'n', v: Number(cell.v) } : cell;
}

/*
 * Columns of the "All" sheet, in this order. A source column goes to the column whose `from`
 * list has its header. Platform and Type come from the sheet name. Columns that no sheet uses
 * are left out; source columns that match nothing are added after these, so no data is dropped.
 */
const COLUMNS = [
  { name: 'Company', from: ['company'] },
  { name: 'Source file', from: ['source file'] },
  { name: 'Platform' },
  { name: 'Type' },
  { name: 'ID', from: ['id'] },
  { name: 'Author', from: ['author'], normalize: trimText, note: 'trimmed of spaces around the name' },
  { name: 'Content', from: ['content'] },
  { name: 'Date', from: ['date', 'comment date', 'comment date text'], normalize: (c, ctx) => toDate(c, ctx, DATETIME_FORMAT), note: 'turned from text into real Excel dates' },
  { name: 'Day', from: ['day'], normalize: (c, ctx) => toDate(c, ctx, DAY_FORMAT), note: 'turned from text into real Excel dates' },
  { name: 'Sentiment', from: ['sentiment'], normalize: toSentiment, note: 'written as Müsbət / Neytral / Mənfi (comments use 1 / 0 / -1; some posts NEYTRAL or Neqativ)' },
  { name: 'Category', from: ['category'], normalize: toCategory, note: 'written in one spelling (e.g. “ictimai_rəy” → “İctimai rəy”, “digər” → “Digər”)' },
  { name: 'Keywords', from: ['keywords'] },
  { name: 'URL', from: ['url', 'comment url'] },
  { name: 'Comment Count', from: ['comment count'], normalize: toCount, note: 'turned from text into numbers' },
  { name: 'Like Count', from: ['like count'], normalize: toCount, note: 'turned from text into numbers' },
  { name: 'Share Count', from: ['share count'], normalize: toCount, note: 'turned from text into numbers' },
  { name: 'View Count', from: ['view count'], normalize: toCount, note: 'turned from text into numbers' },
  { name: 'Post URL', from: ['post url'] },
  { name: 'Post content', from: ['post content'] },
];
const COLUMN_BY_NAME = new Map(COLUMNS.map((c) => [c.name, c]));

// Left out of "All": YouTube's video title (its Content is used instead).
const LEFT_OUT = new Set(['title']);

/** "Facebook comments" → Facebook / Comment, "Facebook" → Facebook / Post. */
export function describeSheet(name) {
  const trimmed = name.trim();
  const isComments = /comments?$/i.test(trimmed);
  const platform = trimmed.replace(/[\s_-]*comments?$/i, '').trim() || trimmed;
  return { platform, type: isComments ? 'Comment' : 'Post' };
}

/** A "Source" column is redundant when it only repeats the platform ("Tiktok" on the TikTok sheet). */
function repeatsPlatform(sheet, j, platform) {
  return sheet.rows.every((row) => !row[j] || normKey(cellText(row[j])) === normKey(platform));
}

/** Decide where every source column goes, and how category spellings are unified. */
function planAll(sheets, date1904) {
  const extras = new Map(); // header key → display name
  const used = new Set(['Platform', 'Type']);
  const sources = [];

  for (const sheet of sheets) {
    const { platform, type } = describeSheet(sheet.name);
    const taken = new Set();
    const leftOut = [];
    const targets = sheet.columns.map((col, j) => {
      const key = headerKey(col.name);
      if (LEFT_OUT.has(key) || (key === 'source' && repeatsPlatform(sheet, j, platform))) {
        leftOut.push(col.name);
        return null;
      }
      const known = COLUMNS.find((c) => c.from?.includes(key) && !taken.has(c.name));
      if (known) {
        taken.add(known.name);
        used.add(known.name);
        return { known: known.name };
      }
      if (!extras.has(key)) extras.set(key, col.name);
      return { extra: key };
    });
    sources.push({ sheet, platform, type, targets, leftOut });
  }
  if (used.has('Date')) used.add('Day'); // Day is filled from Date where it's missing

  const header = COLUMNS.filter((c) => used.has(c.name)).map((c) => c.name);
  const taken = new Set(header.map(normKey));
  const extraNames = new Map();
  for (const [key, name] of extras) {
    const unique = uniqueName(name, taken);
    taken.add(normKey(unique));
    extraNames.set(key, unique);
    header.push(unique);
  }
  const index = new Map(header.map((name, i) => [name, i]));
  for (const src of sources) {
    src.columns = src.targets.map((t) => {
      if (!t) return null;
      const name = t.known ?? extraNames.get(t.extra);
      return { name, index: index.get(name), normalize: COLUMN_BY_NAME.get(name)?.normalize };
    });
  }

  // The first spelling seen (capitalised) becomes the label for each category.
  const categories = new Map();
  for (const { sheet, columns } of sources) {
    const j = columns.findIndex((c) => c?.name === 'Category');
    if (j < 0) continue;
    for (const row of sheet.rows) {
      const cell = row[j];
      if (cell?.t === 's' && cell.v.trim() && !categories.has(categoryKey(cell.v))) {
        categories.set(categoryKey(cell.v), categoryLabel(cell.v));
      }
    }
  }

  return { header, index, sources, ctx: { date1904, categories } };
}

/** Plan "All" from every worksheet except an existing "All". */
function planFor(wb) {
  const date1904 = Boolean(wb.Workbook?.WBProps?.date1904);
  const sheets = parseSheets(wb, { keepDateSystem: true }).filter((s) => normKey(s.name) !== normKey(ALL_SHEET));
  return planAll(sheets, date1904);
}

/** The cells one source row produces in "All", keyed by column index. */
function allRow(plan, src, row) {
  const { index, ctx } = plan;
  const out = [];
  out[index.get('Platform')] = { t: 's', v: src.platform };
  out[index.get('Type')] = { t: 's', v: src.type };
  row.forEach((cell, j) => {
    const col = src.columns[j];
    if (cell && col) out[col.index] = col.normalize ? col.normalize(cell, ctx) : cell;
  });
  const date = out[index.get('Date')];
  if (index.has('Day') && !out[index.get('Day')] && date?.t === 'n') {
    out[index.get('Day')] = { t: 'n', v: Math.floor(date.v), z: DAY_FORMAT };
  }
  return out;
}

const changed = (a, b) => a.t !== b.t || a.v !== b.v;

function buildAllSheet(plan) {
  const { header, index, sources } = plan;
  const data = [header.map((v) => ({ t: 's', v }))];
  const changes = new Map();
  let daysFilled = 0;
  let unreadDates = 0;

  for (const src of sources) {
    const dayCol = src.columns.findIndex((c) => c?.name === 'Day');
    for (const row of src.sheet.rows) {
      const out = allRow(plan, src, row);
      row.forEach((cell, j) => {
        const col = src.columns[j];
        if (!cell || !col) return;
        const value = out[col.index];
        if (changed(cell, value)) changes.set(col.name, (changes.get(col.name) ?? 0) + 1);
        if ((col.name === 'Date' || col.name === 'Day') && value.t !== 'n') unreadDates++;
      });
      if (index.has('Day') && out[index.get('Day')] && !row[dayCol]) daysFilled++;
      data.push(out);
    }
  }
  if (data.length > EXCEL_MAX_ROWS) {
    throw new UserError(`The "${ALL_SHEET}" sheet would have ${data.length - 1} rows, more than Excel's limit of ${EXCEL_MAX_ROWS - 1}.`);
  }

  const range = { s: { r: 0, c: 0 }, e: { r: data.length - 1, c: header.length - 1 } };
  const ws = { '!data': data, '!ref': XLSX.utils.encode_range(range) };
  ws['!autofilter'] = { ref: ws['!ref'] };
  ws['!cols'] = columnWidths(data, header.length);
  return { ws, changes, daysFilled, unreadDates };
}

/** Put the sheet first, keeping sheet-scoped names (e.g. autofilters) pointing at their own sheets. */
function insertFirst(wb, name, ws) {
  wb.SheetNames.unshift(name);
  wb.Sheets[name] = ws;
  wb.Workbook?.Sheets?.unshift({ name, Hidden: 0 });
  for (const n of wb.Workbook?.Names ?? []) if (n.Sheet != null) n.Sheet += 1;
  for (const view of wb.Workbook?.WBView ?? []) view.activeTab = 0;
}

function removeSheet(wb, name) {
  const idx = wb.SheetNames.indexOf(name);
  wb.SheetNames.splice(idx, 1);
  delete wb.Sheets[name];
  wb.Workbook?.Sheets?.splice(idx, 1);
  if (wb.Workbook?.Names) {
    wb.Workbook.Names = wb.Workbook.Names
      .filter((n) => n.Sheet !== idx)
      .map((n) => (n.Sheet != null && n.Sheet > idx ? { ...n, Sheet: n.Sheet - 1 } : n));
  }
}

/**
 * Read the uploaded and the generated file again from scratch and check that the original sheets
 * are unchanged and that every source row is in "All" with every value in the right column.
 */
export function verifyAll(inputBuffer, buffer, fileName) {
  const fail = (msg) => { throw new Error(`Verification failed: ${msg}. The file was not produced.`); };
  const before = readWorkbook(inputBuffer, fileName);
  const after = readWorkbook(buffer, fileName);
  const plan = planFor(before);
  const originals = before.SheetNames.filter((n) => normKey(n) !== normKey(ALL_SHEET));

  if (after.SheetNames.join('\n') !== [ALL_SHEET, ...originals].join('\n')) fail('the sheet list is not as expected');

  for (const name of originals) {
    const a = before.Sheets[name]['!data'] ?? [];
    const b = after.Sheets[name]['!data'] ?? [];
    let expected = 0;
    a.forEach((row, r) => (row ?? []).forEach((cell, c) => {
      if (isEmpty(cell)) return;
      expected++;
      if (!sameValue(cell, b[r]?.[c])) fail(`sheet "${name}" changed at ${XLSX.utils.encode_cell({ r, c })}`);
    }));
    const actual = b.reduce((sum, row) => sum + (row ?? []).filter((cell) => !isEmpty(cell)).length, 0);
    if (actual !== expected) fail(`sheet "${name}" has ${actual} filled cells, expected ${expected}`);
  }

  const data = after.Sheets[ALL_SHEET]['!data'] ?? [];
  const colOf = new Map((data[0] ?? []).map((cell, c) => [cellText(cell), c]));
  if (plan.header.some((name) => !colOf.has(name))) fail(`a column of "${ALL_SHEET}" is missing`);

  let r = 1;
  let expectedCells = 0;
  for (const src of plan.sources) {
    for (const row of src.sheet.rows) {
      const outRow = data[r] ?? [];
      allRow(plan, src, row).forEach((cell, i) => {
        if (!cell) return;
        expectedCells++;
        if (!sameValue(cell, outRow[colOf.get(plan.header[i])])) {
          fail(`"${ALL_SHEET}" row ${r + 1} (from sheet "${src.sheet.name}"), column "${plan.header[i]}" does not match the source`);
        }
      });
      r++;
    }
  }
  const rows = data.slice(1).filter(Boolean);
  if (rows.length !== r - 1) fail(`"${ALL_SHEET}" has ${rows.length} rows, expected ${r - 1}`);
  const cells = rows.reduce((sum, row) => sum + row.filter((cell) => !isEmpty(cell)).length, 0);
  if (cells !== expectedCells) fail(`"${ALL_SHEET}" has ${cells} filled cells, expected ${expectedCells}`);
}

/** Add an "All" sheet that merges every sheet of the workbook (posts and comments of all platforms). */
export function addAllSheet(inputBuffer, fileName) {
  const wb = readWorkbook(inputBuffer, fileName, KEEP_OPTS);
  const existing = wb.SheetNames.find((n) => normKey(n) === normKey(ALL_SHEET));
  if (existing) removeSheet(wb, existing); // left over from an earlier run: rebuilt from scratch

  const plan = planFor(wb);
  if (!plan.sources.some((src) => src.sheet.columns.length)) throw new UserError(`"${fileName}" has no sheets with data.`);
  const { ws, changes, daysFilled, unreadDates } = buildAllSheet(plan);
  insertFirst(wb, ALL_SHEET, ws);

  const buffer = writeWorkbook(wb);
  verifyAll(inputBuffer, buffer, fileName);

  const report = {
    file: fileName,
    columns: plan.header,
    sheets: plan.sources.map((src) => ({
      name: src.sheet.name,
      platform: src.platform,
      type: src.type,
      rows: src.sheet.rows.length,
      leftOut: src.leftOut,
      extraColumns: src.columns.filter((c) => c && !COLUMN_BY_NAME.has(c.name)).map((c) => c.name),
    })),
    totalRows: plan.sources.reduce((sum, src) => sum + src.sheet.rows.length, 0),
    changes: [...changes]
      .sort(([a], [b]) => plan.index.get(a) - plan.index.get(b))
      .map(([column, count]) => ({ column, count, note: COLUMN_BY_NAME.get(column)?.note })),
    daysFilled,
    unreadDates,
    replacedExisting: Boolean(existing),
    truncatedTexts: longTexts(wb),
    verified: true,
  };
  return { buffer, report };
}
