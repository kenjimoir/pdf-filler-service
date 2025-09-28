// index.cjs — PDF fill → upload to Drive (Shared Drives OK)
// Requires deps: express, cors, pdf-lib, googleapis
// Render-compatible: uses process.env.PORT and os.tmpdir()

const fs = require('fs');
const os = require('os');
const path = require('path');
const express = require('express');
const cors = require('cors');
const { PDFDocument } = require('pdf-lib');
const { google } = require('googleapis');

const PORT = process.env.PORT || 8080;
const TMP  = path.join(os.tmpdir(), 'pdf-filler');
ensureDir(TMP);

// === Default fallback folder (used ONLY when no folderId is provided in request) ===
const OUTPUT_FOLDER_ID = process.env.OUTPUT_FOLDER_ID || '';

function log(...args) { console.log(new Date().toISOString(), '-', ...args); }
function ensureDir(p) { if (!fs.existsSync(p)) fs.mkdirSync(p, { recursive: true }); }

// ---- Google Drive client (supports Shared Drives) ----
function getDriveClient() {
  // Option A: GOOGLE_CREDENTIALS_JSON (paste JSON into Render env var)
  // Option B: GOOGLE_APPLICATION_CREDENTIALS (path to a mounted JSON file)
  let credentials = null;
  if (process.env.GOOGLE_CREDENTIALS_JSON) {
    try { credentials = JSON.parse(process.env.GOOGLE_CREDENTIALS_JSON); }
    catch (e) { log('ERROR parsing GOOGLE_CREDENTIALS_JSON:', e && e.message); }
  }
  const auth = new google.auth.GoogleAuth({
    scopes: ['https://www.googleapis.com/auth/drive'],
    ...(credentials ? { credentials } : {}) // fallback to ADC / GOOGLE_APPLICATION_CREDENTIALS
  });
  return google.drive({ version: 'v3', auth });
}

// ---- Drive helpers ----
async function downloadDriveFile(fileId, destPath) {
  const drive = getDriveClient();
  log('Downloading template', fileId, '→', destPath);
  const res = await drive.files.get(
    { fileId, alt: 'media', supportsAllDrives: true },
    { responseType: 'stream' }
  );
  await new Promise((resolve, reject) => {
    const out = fs.createWriteStream(destPath);
    res.data.on('error', reject).pipe(out).on('finish', resolve);
  });
  return destPath;
}

/**
 * Upload localPath to Drive.
 * - If parentId is provided (from request), it is used.
 * - Else if OUTPUT_FOLDER_ID env is set, it is used.
 * - Else it uploads to My Drive root of the service account.
 */
async function uploadToDrive(localPath, name, parentId) {
  const drive = getDriveClient();
  const parents =
    parentId ? [parentId] :
    (OUTPUT_FOLDER_ID ? [OUTPUT_FOLDER_ID] : undefined);

  const fileMetadata = { name, parents };
  const media = { mimeType: 'application/pdf', body: fs.createReadStream(localPath) };
  const res = await drive.files.create({
    requestBody: fileMetadata,
    media,
    fields: 'id, name, webViewLink, parents',
    supportsAllDrives: true
  });
  return res.data;
}

// ---- PDF fill (handles text / checkbox / radio / dropdown) ----
async function fillPdf(srcPath, outPath, fields = {}) {
  const bytes = fs.readFileSync(srcPath);
  const pdfDoc = await PDFDocument.load(bytes, { updateFieldAppearances: true });

  let filled = 0;
  let form = null;
  try { form = pdfDoc.getForm(); } catch (_) {}

  if (form) {
    for (const [key, rawVal] of Object.entries(fields)) {
      try {
        const f = form.getField(String(key));
        const typeName = (f && f.constructor && f.constructor.name) || '';

        const val = rawVal == null ? '' : String(rawVal);
        if (typeName.includes('Text')) {
          f.setText(val);
          filled++;
        } else if (typeName.includes('Check')) {
          const on = val.toLowerCase();
          if (on === 'true' || on === 'yes' || on === '1' || on === 'on') f.check();
          else f.uncheck();
          filled++;
        } else if (typeName.includes('Radio')) {
          try { f.select(val); filled++; } catch (_) {}
        } else if (typeName.includes('Dropdown')) {
          try { f.select(val); filled++; } catch (_) {}
        }
      } catch (_) {
        // field not found — ignore missing field
      }
    }
    try { form.updateFieldAppearances(); } catch (_) {}
  }

  const outBytes = await pdfDoc.save();
  fs.writeFileSync(outPath, outBytes);
  return { outPath, filled, size: outBytes.length };
}

// ---- HTTP server ----
const app = express();
app.use(cors());
app.use(express.json({ limit: '10mb' }));

app.get('/', (req, res) => {
  res.type('text/plain').send('PDF filler is up. Try GET /health');
});

app.get('/health', (req, res) => {
  res.json({
    ok: true,
    tmpDir: TMP,
    hasOUTPUT_FOLDER_ID: !!OUTPUT_FOLDER_ID,
    credMode: process.env.GOOGLE_CREDENTIALS_JSON
      ? 'env-json'
      : (process.env.GOOGLE_APPLICATION_CREDENTIALS ? 'file-path' : 'adc/unknown')
  });
});

// GET /fields?fileId=XXXXXXXX — lists PDF form field names
app.get('/fields', async (req, res) => {
  try {
    const fileId = (req.query.fileId || '').trim();
    if (!fileId) return res.status(400).json({ error: 'fileId is required' });

    const localPath = path.join(TMP, `template_${fileId}.pdf`);
    await downloadDriveFile(fileId, localPath);

    const bytes = fs.readFileSync(localPath);
    const pdfDoc = await PDFDocument.load(bytes);
    let names = [];
    try {
      const form = pdfDoc.getForm();
      const fields = form ? form.getFields() : [];
      names = fields.map(f => f.getName());
    } catch (_) {}
    res.json({ count: names.length, names });
  } catch (e) {
    log('List fields failed:', e && (e.stack || e));
    res.status(500).json({ error: 'List fields failed', detail: e && (e.message || String(e)) });
  }
});

// POST /fill
// body: { templateFileId: "<Drive file id>", fields: {...}, outputName?: "baseName", folderId?: "<Drive folder id>" }
app.post('/fill', async (req, res) => {
  try {
    const { templateFileId, fields, outputName, folderId } = req.body || {};
    if (!templateFileId) {
      return res.status(400).json({ error: 'templateFileId is required' });
    }

    // 1) Download template → 2) Fill → 3) Upload to Drive
    const tmpTemplate = path.join(TMP, `template_${templateFileId}.pdf`);
    await downloadDriveFile(templateFileId, tmpTemplate);

    const base = (outputName && String(outputName).trim()) || `filled_${Date.now()}`;
    const outName = base.toLowerCase().endsWith('.pdf') ? base : `${base}.pdf`;
    const outPath = path.join(TMP, outName);

    const result = await fillPdf(tmpTemplate, outPath, fields || {});
    log(`Filled PDF -> ${result.outPath} (${result.size} bytes, fields filled: ${result.filled})`);

    const uploaded = await uploadToDrive(result.outPath, outName, folderId);
    log('Uploaded to Drive:', uploaded);

    res.json({
      ok: true,
      filledCount: result.filled,
      driveFile: uploaded
    });
  } catch (err) {
    log('ERROR /fill:', err && (err.stack || err));
    res.status(500).json({ error: 'Fill failed', detail: err && (err.message || String(err)) });
  }
});

app.listen(PORT, () => {
  log(`Server listening on ${PORT}`);
  log(`OUTPUT_FOLDER_ID (fallback): ${OUTPUT_FOLDER_ID || '(none set)'}`);
  log(`Creds: ${process.env.GOOGLE_CREDENTIALS_JSON ? 'GOOGLE_CREDENTIALS_JSON' :
               (process.env.GOOGLE_APPLICATION_CREDENTIALS ? 'GOOGLE_APPLICATION_CREDENTIALS' : 'ADC/unknown')}`);
});
