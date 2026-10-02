import express from 'express';
import multer from 'multer';
import path from 'node:path';
import { combineFiles, UserError } from './combine.js';

// Optional .env in the working directory; real environment variables take precedence.
try {
  process.loadEnvFile();
} catch (err) {
  if (err.code !== 'ENOENT') throw err;
}

const PORT = Number(process.env.PORT) || 3000;
// 127.0.0.1 keeps the API reachable only through nginx on the same server.
const HOST = process.env.HOST || '127.0.0.1';
const MAX_FILE_MB = Number(process.env.MAX_FILE_MB) || 100;
const CORS_ORIGINS = new Set(
  (process.env.CORS_ORIGIN || '').split(',').map((origin) => origin.trim()).filter(Boolean),
);
const EXTENSIONS = ['.xlsx', '.xlsm', '.xlsb', '.xls', '.ods'];

// Files are kept in memory only and never written to disk.
const upload = multer({
  storage: multer.memoryStorage(),
  limits: { fileSize: MAX_FILE_MB * 1024 * 1024, files: 500 },
});

const app = express();

// Required only when the frontend calls this API from another origin.
app.use((req, res, next) => {
  const origin = req.get('origin');
  if (!origin || (!CORS_ORIGINS.has('*') && !CORS_ORIGINS.has(origin))) return next();

  res.set('Access-Control-Allow-Origin', CORS_ORIGINS.has('*') ? '*' : origin);
  if (!CORS_ORIGINS.has('*')) res.vary('Origin');
  res.set('Access-Control-Allow-Methods', 'GET, POST, OPTIONS');
  res.set('Access-Control-Allow-Headers', 'Content-Type');
  if (req.method === 'OPTIONS') return res.sendStatus(204);
  next();
});

app.get('/api/health', (req, res) => res.json({ ok: true }));

app.post('/api/combine', upload.array('files'), (req, res) => {
  const files = req.files ?? [];
  if (!files.length) return res.status(400).json({ error: 'No files uploaded.' });

  // Browsers send UTF-8 file names, but multer decodes them as latin1.
  const inputs = files.map((f) => ({ name: Buffer.from(f.originalname, 'latin1').toString('utf8'), buffer: f.buffer }));
  const unsupported = inputs.filter((f) => !EXTENSIONS.includes(path.extname(f.name).toLowerCase()));
  if (unsupported.length) {
    return res.status(400).json({ error: `Not an Excel file: ${unsupported.map((f) => f.name).join(', ')}` });
  }
  const started = Date.now();
  try {
    const { buffer, report } = combineFiles(inputs, {
      addCompanyColumn: req.body.addCompanyColumn === 'true',
      addSourceColumn: req.body.addSourceColumn === 'true',
    });
    console.log(`Combined ${inputs.length} files, ${report.totalRows} rows in ${Date.now() - started} ms`);
    res.json({ report, file: buffer.toString('base64') });
  } catch (err) {
    if (err instanceof UserError) return res.status(400).json({ error: err.message });
    console.error(err);
    res.status(500).json({ error: err.message || 'Unexpected error while combining files.' });
  }
});

app.use((err, req, res, next) => {
  if (err instanceof multer.MulterError) {
    const message = err.code === 'LIMIT_FILE_SIZE' ? `A file is larger than ${MAX_FILE_MB} MB.` : err.message;
    return res.status(400).json({ error: message });
  }
  console.error(err);
  res.status(500).json({ error: 'Unexpected server error.' });
});

app.listen(PORT, HOST, (err) => {
  if (err) throw err;
  console.log(`Excel Combiner API running on http://${HOST}:${PORT}`);
});
