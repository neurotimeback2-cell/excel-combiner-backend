# Excel Combiner — backend

Express API that merges uploaded Excel files sheet by sheet (all "Instagram" sheets into one "Instagram" sheet, and so on) and returns the combined `.xlsx`. The UI is in the separate frontend repo; on the server, nginx serves the UI and forwards `/api` to this API.

## Run locally

```bash
npm install
npm run dev     # http://127.0.0.1:3000, restarts on code changes
npm test
```

## Deploy on the VPS

Requires Node.js 22+. The steps assume the repos are cloned into `/opt/excel-combiner/`; if you use another path, change it in `deploy/excel-combiner-api.service`.

```bash
sudo mkdir -p /opt/excel-combiner && sudo chown $USER /opt/excel-combiner
cd /opt/excel-combiner
git clone <backend-repo-url> backend
cd backend
npm ci --omit=dev
cp .env.example .env        # optional, defaults are fine

sudo cp deploy/excel-combiner-api.service /etc/systemd/system/
sudo systemctl daemon-reload
sudo systemctl enable --now excel-combiner-api
curl http://127.0.0.1:3000/api/health     # {"ok":true}
```

Then set up the frontend and nginx (see the frontend repo README).

Update to a new version:

```bash
cd /opt/excel-combiner/backend && git pull && npm ci --omit=dev && sudo systemctl restart excel-combiner-api
```

Logs: `journalctl -u excel-combiner-api -f`

### Settings (`.env`)

| Name | Default | Meaning |
|---|---|---|
| `PORT` | `3000` | API port. If changed, update `proxy_pass` in the nginx config too. |
| `HOST` | `127.0.0.1` | Only nginx on the same server can reach the API. `0.0.0.0` exposes it to the network. |
| `MAX_FILE_MB` | `100` | Max size of one uploaded file. |
| `CORS_ORIGIN` | unset | Comma-separated frontend origins permitted to call the API directly, e.g. `https://app.example.com`. Not needed when nginx proxies `/api` on the same origin. |

### Memory

Files are held in RAM while a request is processed: plan for about 8× the total upload size (e.g. 30 files of 5 MB ≈ 150 MB upload ≈ 1.2 GB RAM at peak). Memory is freed after each request.

## API

- `POST /api/combine` — `multipart/form-data` with one or more `files` fields and optional `addCompanyColumn=true` and `addSourceColumn=true`. Returns `{ report, file }`, where `file` is the combined workbook as base64. Errors return `{ error }` with status 400 (bad input) or 500.
- `POST /api/all-sheet` — `multipart/form-data` with one `file` field. Returns `{ report, file }`, where `file` is the same workbook with an added "All" sheet, as base64. Errors as above.
- `GET /api/health` — returns `{ ok: true }`.

## How combining works

- Sheets are matched by name (case and surrounding spaces ignored). A sheet present in only some files is still included.
- The first non-empty row of each sheet is the header. Columns are matched by header name, so column order may differ. A column present only in some files becomes an extra column, empty for the other files.
- Rows are appended in upload order. Completely empty rows are skipped; everything else is kept.
- Columns with data but no header get a name like `Column F`; duplicate headers become `Name (2)`.
- With `addCompanyColumn`, every sheet gets a first column `Company` holding the file name without its extension (`Pasha Bank.xlsx` → `Pasha Bank`). With `addSourceColumn`, a `Source file` column holds the full file name. If the uploads already have a column with that name, the added one becomes `Company (2)`.
- Values and number formats (dates, percentages) are kept. Formulas are replaced by their calculated values. Cell styles are not copied.

## How the "All" sheet works

`POST /api/all-sheet` takes one workbook with a sheet per platform — posts (`Facebook`, `Instagram`, `Telegram`, …) and comments (`Facebook comments`, …) — and adds a first sheet `All` with every row of every sheet, one sheet after another. The original sheets are not changed. Code: `src/all-sheet.js`.

- **Platform** and **Type** come from the sheet name: `Facebook comments` → Facebook / Comment, `Facebook` → Facebook / Post. New platforms need no code change.
- Columns are matched by name, ignoring case, spaces and underscores (`comment_url` = `Comment URL`). The column list and the names each column accepts are in `COLUMNS` in `src/all-sheet.js`; comment columns map like this: `comment_url` → URL, `post_url` → Post URL, `post_content` → Post content, `comment_date` / `comment_date_text` → Date, `like_count` → Like Count.
- Left out: `Title` (YouTube; Content holds the text) and `Source` when it only repeats the platform. Columns that match nothing are added at the end of All, so no data is dropped. Columns no sheet has are not shown.
- Values are unified so the sheet can be filtered and pivoted:
  - Sentiment → `Müsbət` / `Neytral` / `Mənfi` (from `1` / `0` / `-1`, `NEYTRAL`, `Neqativ`, …).
  - Category → one spelling per category (`ictimai_rəy`, `ictimai rəy`, `İctimai rəy` → `İctimai rəy`).
  - Date and Day → real Excel dates, including texts like `2025-09-24 10:03:37` and `17:03, 23 Dekabr 2025, Çərşənbə axşamı`. Day is filled from Date where it's missing. Unreadable dates stay as text.
  - Counts stored as text become numbers; Author is trimmed.
- An existing `All` sheet is rebuilt, not merged into itself.

## No data loss

After building the output, the server reads the generated file back and checks that every non-empty cell of every upload is in the right sheet, row and column, with no extra rows or cells. If anything doesn't match, the request fails instead of returning a file. Unreadable files fail the whole request; they are never skipped.

For the "All" sheet, the server reads the uploaded and the generated file again from scratch and checks that every original sheet is unchanged cell by cell, and that every source row is in All with every value in the right column (after the unifying described above).

Nothing is stored: uploads and results exist only in memory during the request, and there is no database.
