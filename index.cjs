// index.cjs — PDF fill → upload to Drive (Shared Drives OK)
// Requires deps: express, cors, pdf-lib, googleapis, iconv-lite, @pdf-lib/fontkit
// Render-compatible: uses process.env.PORT and os.tmpdir()

const fs = require('fs');
const os = require('os');
const path = require('path');
const express = require('express');
const cors = require('cors');
const { PDFDocument, rgb, degrees } = require('pdf-lib');
const { google } = require('googleapis');
const fontkit = require('@pdf-lib/fontkit');

const PORT = process.env.PORT || 8080;
const TMP = path.join(os.tmpdir(), 'pdf-filler');
ensureDir(TMP);

const OUTPUT_FOLDER_ID = process.env.OUTPUT_FOLDER_ID || '';

function log(...args) {
  console.log(new Date().toISOString(), '-', ...args);
}

function ensureDir(p) {
  if (!fs.existsSync(p)) fs.mkdirSync(p, { recursive: true });
}

// ---- Google Drive client (supports Shared Drives) ----
function getDriveClient() {
  let credentials = null;
  if (process.env.GOOGLE_CREDENTIALS_JSON) {
    try {
      credentials = JSON.parse(process.env.GOOGLE_CREDENTIALS_JSON);
    } catch (e) {
      log('ERROR parsing GOOGLE_CREDENTIALS_JSON:', e && e.message);
    }
  }
  const auth = new google.auth.GoogleAuth({
    scopes: ['https://www.googleapis.com/auth/drive'],
    ...(credentials ? { credentials } : {}),
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
    supportsAllDrives: true,
  });
  return res.data;
}

// ---- PDF fill (handles text / checkbox / radio / dropdown) ----
async function fillPdf(srcPath, outPath, fields = {}, opts = {}) {
  const bytes = fs.readFileSync(srcPath);
  const pdfDoc = await PDFDocument.load(bytes, { updateFieldAppearances: true });

  // === fontkit 登録 ===
  try {
    pdfDoc.registerFontkit(fontkit);
  } catch (e) {
    log('fontkit register skipped (maybe already registered):', e.message);
  }

  // === 日本語フォント読み込み・埋め込み ===
  let customFont = null;
  const fontPath = process.env.FONT_TTF_PATH || 'fonts/NotoSansJP-VariableFont_wght.ttf';
  try {
    if (fs.existsSync(fontPath)) {
      const fontBytes = fs.readFileSync(fontPath);
      customFont = await pdfDoc.embedFont(fontBytes, { subset: true });
      log('Embedded custom font from:', fontPath);
    } else {
      log('Custom font not found, continue without. path=', fontPath);
    }
  } catch (e) {
    log('Font embed failed (continue without):', e && e.message);
  }

  let filled = 0;
  let form = null;
  try { form = pdfDoc.getForm(); } catch (_) {}

  if (form) {
    for (const [key, rawVal] of Object.entries(fields)) {
      try {
        const f = form.getField(String(key));
        const typeName = (f && f.constructor && f.constructor.name) || '';
        const val = rawVal == null ? '' : String(rawVal);

        // ---- Text ----
        if (typeName.includes('Text')) {
          if (customFont) {
            try { f.updateAppearances(customFont); } catch (_) {}
          }
          f.setText(val);
          filled++;

        // ---- Checkbox ----
        } else if (typeName.includes('Check')) {
          const on = val.toLowerCase();
          if (['true', 'yes', '1', 'on'].includes(on)) f.check();
          else f.uncheck();
          filled++;

        // ---- Radio ----
        } else if (typeName.includes('Radio')) {
          try { f.select(val); filled++; } catch (_) {}

        // ---- Dropdown ----
        } else if (typeName.includes('Dropdown')) {
          try { f.select(val); filled++; } catch (_) {}
        }
      } catch (_) {
        // ignore missing field
      }
    }
    try { form.updateFieldAppearances(customFont || undefined); } catch (_) {}
  }
  // === 確認用の文字（透かし）を入れる部分 ===
  const wmText = opts.watermarkText && String(opts.watermarkText).trim();
  if (wmText) {
    const pages = pdfDoc.getPages();
    for (const page of pages) {
      const { width, height } = page.getSize();
      page.drawText(wmText, {
        x: width / 2 - 200,    // 中央に配置
        y: height / 2 - 40,
        size: 80,              // 文字の大きさ
        opacity: 0.12,         // うっすら見える程度
        rotate: degrees(45),   // 斜めに表示
        color: rgb(0.85, 0.1, 0.1)  // 赤みのある色
      });
    }
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
      : (process.env.GOOGLE_APPLICATION_CREDENTIALS ? 'file-path' : 'adc/unknown'),
  });
});

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
    res.status(500).json({ error: 'List fields failed', detail: e.message });
  }
});

app.post('/fill', async (req, res) => {
  try {
    const { templateFileId, fields, outputName, folderId, mode, watermarkText } = req.body || {};
    if (!templateFileId) {
      return res.status(400).json({ error: 'templateFileId is required' });
    }

    const tmpTemplate = path.join(TMP, `template_${templateFileId}.pdf`);
    await downloadDriveFile(templateFileId, tmpTemplate);

    const base = (outputName && String(outputName).trim()) || `filled_${Date.now()}`;
    const outName = base.toLowerCase().endsWith('.pdf') ? base : `${base}.pdf`;
    const outPath = path.join(TMP, outName);

    const wm = watermarkText || (mode === 'review' ? '確認用 / DRAFT' : '');
    const result = await fillPdf(tmpTemplate, outPath, fields || {}, { watermarkText: wm });
    log(`Filled PDF -> ${result.outPath} (${result.size} bytes, fields filled: ${result.filled})`);

    const uploaded = await uploadToDrive(result.outPath, outName, folderId);
    log('Uploaded to Drive:', uploaded);

    res.json({
      ok: true,
      filledCount: result.filled,
      driveFile: uploaded,
    });
  } catch (err) {
    log('ERROR /fill:', err && (err.stack || err));
    res.status(500).json({ error: 'Fill failed', detail: err.message });
  }
});

app.listen(PORT, () => {
  log(`Server listening on ${PORT}`);
  log(`OUTPUT_FOLDER_ID (fallback): ${OUTPUT_FOLDER_ID || '(none set)'}`);
  log(`Creds: ${process.env.GOOGLE_CREDENTIALS_JSON ? 'GOOGLE_CREDENTIALS_JSON' :
    (process.env.GOOGLE_APPLICATION_CREDENTIALS ? 'GOOGLE_APPLICATION_CREDENTIALS' : 'ADC/unknown')}`);
  log(`Font path: ${process.env.FONT_TTF_PATH || 'fonts/NotoSansJP-VariableFont_wght.ttf'}`);
});
