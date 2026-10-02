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
- `GET /api/health` — returns `{ ok: true }`.

## How combining works

- Sheets are matched by name (case and surrounding spaces ignored). A sheet present in only some files is still included.
- The first non-empty row of each sheet is the header. Columns are matched by header name, so column order may differ. A column present only in some files becomes an extra column, empty for the other files.
- Rows are appended in upload order. Completely empty rows are skipped; everything else is kept.
- Columns with data but no header get a name like `Column F`; duplicate headers become `Name (2)`.
- With `addCompanyColumn`, every sheet gets a first column `Company` holding the file name without its extension (`Pasha Bank.xlsx` → `Pasha Bank`). With `addSourceColumn`, a `Source file` column holds the full file name. If the uploads already have a column with that name, the added one becomes `Company (2)`.
- Values and number formats (dates, percentages) are kept. Formulas are replaced by their calculated values. Cell styles are not copied.

## No data loss

After building the output, the server reads the generated file back and checks that every non-empty cell of every upload is in the right sheet, row and column, with no extra rows or cells. If anything doesn't match, the request fails instead of returning a file. Unreadable files fail the whole request; they are never skipped.

Nothing is stored: uploads and results exist only in memory during the request, and there is no database.
