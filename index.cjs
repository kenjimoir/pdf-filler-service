// index.cjs — optimized PDF fill → Drive upload
// Deps: express, cors, pdf-lib, googleapis, @pdf-lib/fontkit
// Uses envs: PORT, OUTPUT_FOLDER_ID, GOOGLE_APPLICATION_CREDENTIALS or GOOGLE_CREDENTIALS_JSON,
//            FONT_TTF_PATH  (e.g., fonts/KozMinPr6N-Subset.otf or fonts/NotoSansJP-Regular.ttf)

const fs = require('fs');
const os = require('os');
const path = require('path');
const express = require('express');
const cors = require('cors');
const { PDFDocument, rgb, degrees, PDFName, PDFBool, PDFDict, PDFString } = require('pdf-lib');
const { google } = require('googleapis');
const fontkit = require('@pdf-lib/fontkit');

const PORT = process.env.PORT || 8080;
const TMP = path.join(os.tmpdir(), 'pdf-filler');
ensureDir(TMP);
const ROOT = process.cwd();

const OUTPUT_FOLDER_ID = process.env.OUTPUT_FOLDER_ID || '';

function log(...args) {
  console.log(new Date().toISOString(), '-', ...args);
}
function ensureDir(p) {
  if (!fs.existsSync(p)) fs.mkdirSync(p, { recursive: true });
}

// ───────────────────────────────────────────────────────────────────────────────
// Google Drive helpers
// ───────────────────────────────────────────────────────────────────────────────
function getDriveClient() {
  let credentials = null;
  if (process.env.GOOGLE_CREDENTIALS_JSON) {
    try { credentials = JSON.parse(process.env.GOOGLE_CREDENTIALS_JSON); }
    catch (e) { log('ERROR parsing GOOGLE_CREDENTIALS_JSON:', e && e.message); }
  }
  const auth = new google.auth.GoogleAuth({
    scopes: ['https://www.googleapis.com/auth/drive'],
    ...(credentials ? { credentials } : {}),
  });
  return google.drive({ version: 'v3', auth });
}

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

// ───────────────────────────────────────────────────────────────────────────────
// PDF fill (optimized)
// ───────────────────────────────────────────────────────────────────────────────
async function fillPdf(srcPath, outPath, fields = {}, opts = {}) {
  const bytes = fs.readFileSync(srcPath);
  const pdfDoc = await PDFDocument.load(bytes, { updateFieldAppearances: true });

  // Register fontkit once
  try { pdfDoc.registerFontkit(fontkit); } catch (_) {}

  // ── Embed JP font (subset when safe) ─────────────────────────────────────────
  let customFont = null;
  global.LAST_EMBEDDED_FONT = '';
  const pref = (process.env.FONT_TTF_PATH || '').trim();
  const candidates = [
    pref,
    path.join(ROOT, 'fonts/NotoSansJP-Regular.ttf'),
    path.join(ROOT, 'fonts/NotoSansJP-Regular.otf'),
  ].filter(Boolean);

  for (const p of candidates) {
    try {
      const exists = p && fs.existsSync(p);
      log('Font candidate:', p || '(none)', exists ? '(exists)' : '(missing)');
      if (!exists) continue;

      const fontBytes = fs.readFileSync(p);
      const ext = path.extname(p).toLowerCase();

      // Best size when TTF (safe subsetting). OTF (CFF) can be touchy — try subset then fallback.
      const trySubset = ext === '.ttf' || ext === '.ttc' ? true : true;

      try {
        customFont = await pdfDoc.embedFont(fontBytes, { subset: trySubset });
        global.LAST_EMBEDDED_FONT = p;
        log('Embedded JP font:', p, '(subset:', trySubset, ')');
        break;
      } catch (e) {
        log('Embed failed (subset=', trySubset, ') → retry no-subset:', e.message);
        customFont = await pdfDoc.embedFont(fontBytes, { subset: false });
        global.LAST_EMBEDDED_FONT = p;
        log('Embedded JP font (no-subset):', p);
        break;
      }
    } catch (e) {
      log('Font try failed:', p, e.message);
    }
  }
  if (!customFont) log('WARNING: No JP font embedded; CJK text may appear blank.');

  // ── Set AcroForm defaults (use our font everywhere) ──────────────────────────
  try {
    const acroFormRef = pdfDoc.catalog.get(PDFName.of('AcroForm'));
    if (acroFormRef) {
      const acroForm = pdfDoc.context.lookup(acroFormRef, PDFDict);

      // DR.Font /F0 -> our font
      const dr = (acroForm.get(PDFName.of('DR')) || pdfDoc.context.obj({}));
      const drFont = (dr.get(PDFName.of('Font')) || pdfDoc.context.obj({}));
      if (customFont) drFont.set(PDFName.of('F0'), customFont.ref);
      dr.set(PDFName.of('Font'), drFont);
      acroForm.set(PDFName.of('DR'), dr);

      // DA -> "/F0 10 Tf 0 g" (10pt, black)
      acroForm.set(PDFName.of('DA'), PDFString.of('/F0 10 Tf 0 g'));

      // Ask viewers to regenerate if needed
      acroForm.set(PDFName.of('NeedAppearances'), PDFBool.True);
    }
  } catch (e) {
    log('AcroForm DR/DA setup failed:', e.message);
  }

  // ── Fill fields (force per-field DA on text fields) ──────────────────────────
  let filled = 0;
  let form = null;
  try { form = pdfDoc.getForm(); } catch (_) {}

  if (form) {
    for (const [key, rawVal] of Object.entries(fields)) {
      try {
        const field = form.getField(String(key));
        const typeName = (field && field.constructor && field.constructor.name) || '';
        const val = rawVal == null ? '' : String(rawVal);

        if (typeName.includes('Text')) {
          // Override field-level DA to our font
          try {
            const acro = field.acroField || field._acroField || field['acroField'];
            if (acro && acro.dict) acro.dict.set(PDFName.of('DA'), PDFString.of('/F0 10 Tf 0 g'));
          } catch (_) {}

          if (customFont) { try { field.updateAppearances(customFont); } catch (_) {} }
          field.setText(val);
          filled++;

        } else if (typeName.includes('Check')) {
          const on = val.toLowerCase();
          if (['true', 'yes', '1', 'on', '✓'].includes(on)) field.check();
          else field.uncheck();
          filled++;

        } else if (typeName.includes('Radio')) {
          try { field.select(val); filled++; } catch (_) {}

        } else if (typeName.includes('Dropdown')) {
          try { field.select(val); filled++; } catch (_) {}
        }
      } catch (_) {}
    }

    // Rebuild appearances with our font
    try { form.updateFieldAppearances(customFont || undefined); } catch (_) {}

    // Flatten (bakes visuals; enables stripping form later)
    try { form.flatten(); } catch (_) {}
  }

  // ── Optional watermark ───────────────────────────────────────────────────────
  const wmText = opts.watermarkText && String(opts.watermarkText).trim();
  if (wmText) {
    const pages = pdfDoc.getPages();
    for (const page of pages) {
      const { width, height } = page.getSize();
      page.drawText(wmText, {
        x: width / 2 - 200,
        y: height / 2 - 40,
        size: 80,
        opacity: 0.12,
        rotate: degrees(45),
        color: rgb(0.85, 0.1, 0.1),
      });
    }
  }

  // ── Strip AcroForm & XMP metadata (smaller, cleaner) ─────────────────────────
  try {
    const acroFormRef = pdfDoc.catalog.get(PDFName.of('AcroForm'));
    if (acroFormRef) pdfDoc.catalog.set(PDFName.of('AcroForm'), undefined);
  } catch (_) {}
  try {
    const metaRef = pdfDoc.catalog.get(PDFName.of('Metadata'));
    if (metaRef) pdfDoc.catalog.set(PDFName.of('Metadata'), undefined);
  } catch (_) {}

  // ── Save with object streams (smaller/faster) ────────────────────────────────
  let outBytes;
  try {
    outBytes = await pdfDoc.save({ useObjectStreams: true });
  } catch (e) {
    log('pdfDoc.save(useObjectStreams:true) failed → retry false:', e.message);
    outBytes = await pdfDoc.save({ useObjectStreams: false });
  }

  fs.writeFileSync(outPath, outBytes);
  return { outPath, filled, size: outBytes.length };
}

// ───────────────────────────────────────────────────────────────────────────────
// HTTP server
// ───────────────────────────────────────────────────────────────────────────────
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
    fontEnv: process.env.FONT_TTF_PATH || '(none)',
    lastEmbeddedFont: global.LAST_EMBEDDED_FONT || '(none yet)',
  });
});

// List fields (for mapping work)
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
      const flds = form ? form.getFields() : [];
      names = flds.map(f => f.getName());
    } catch (_) {}
    res.json({ count: names.length, names });
  } catch (e) {
    log('List fields failed:', e && (e.stack || e));
    res.status(500).json({ error: 'List fields failed', detail: e.message });
  }
});

// Main fill
app.post('/fill', async (req, res) => {
  try {
    const { templateFileId, fields, outputName, folderId, mode, watermarkText } = req.body || {};
    if (!templateFileId) return res.status(400).json({ error: 'templateFileId is required' });

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

    res.json({ ok: true, filledCount: result.filled, driveFile: uploaded, webViewLink: uploaded.webViewLink });
  } catch (err) {
    log('ERROR /fill:', err && (err.stack || err));
    res.status(500).json({ error: 'Fill failed', detail: err.message });
  }
});

// ── Debug helpers you already used ─────────────────────────────────────────────
app.get('/debug/passthrough', async (req, res) => {
  try {
    const fileId = (req.query.fileId || '').trim();
    if (!fileId) return res.status(400).json({ ok:false, error:'fileId required' });
    const local = path.join(TMP, `pt_${fileId}.pdf`);
    await downloadDriveFile(fileId, local);
    const uploaded = await uploadToDrive(local, `PASSTHROUGH_${Date.now()}.pdf`);
    res.json({ ok:true, link: uploaded.webViewLink });
  } catch (e) {
    res.status(500).json({ ok:false, error:e.message });
  }
});

app.get('/debug/roundtrip', async (req, res) => {
  try {
    const fileId = (req.query.fileId || '').trim();
    if (!fileId) return res.status(400).json({ ok:false, error:'fileId required' });
    const src = path.join(TMP, `rt_${fileId}.pdf`);
    await downloadDriveFile(fileId, src);
    const bytes = fs.readFileSync(src);
    const doc = await PDFDocument.load(bytes, { updateFieldAppearances: false });
    const outBytes = await doc.save({ useObjectStreams: true });
    const out = path.join(TMP, `ROUNDTRIP_${Date.now()}.pdf`);
    fs.writeFileSync(out, outBytes);
    const uploaded = await uploadToDrive(out, `ROUNDTRIP_${Date.now()}.pdf`);
    res.json({ ok:true, link: uploaded.webViewLink });
  } catch (e) {
    res.status(500).json({ ok:false, error:e.message });
  }
});

app.get('/debug/overlay', async (req, res) => {
  try {
    const fileId = (req.query.fileId || '').trim();
    if (!fileId) return res.status(400).json({ ok:false, error:'fileId required' });

    const src = path.join(TMP, `ov_${fileId}.pdf`);
    await downloadDriveFile(fileId, src);

    const bytes = fs.readFileSync(src);
    const doc = await PDFDocument.load(bytes, { updateFieldAppearances: false });
    try { doc.registerFontkit(fontkit); } catch (_) {}

    // Embed env font or fallback Noto
    let fontBytes = null;
    const envP = process.env.FONT_TTF_PATH || '';
    if (envP && fs.existsSync(envP)) fontBytes = fs.readFileSync(envP);
    else {
      const tries = [
        path.join(ROOT, 'fonts/NotoSansJP-Regular.ttf'),
        path.join(ROOT, 'fonts/NotoSerifCJKjp-Regular.otf'),
      ];
      for (const t of tries) if (!fontBytes && fs.existsSync(t)) fontBytes = fs.readFileSync(t);
    }
    const font = fontBytes ? await doc.embedFont(fontBytes, { subset: false }) : undefined;

    const page = doc.getPages()[0];
    page.drawText('VISIBLE OVERLAY TEST あいうえお 山田太郎', { x: 48, y: 720, size: 14, font, color: rgb(0,0,0) });

    const outBytes = await doc.save({ useObjectStreams: true });
    const out = path.join(TMP, `OVERLAY_${Date.now()}.pdf`);
    fs.writeFileSync(out, outBytes);
    const uploaded = await uploadToDrive(out, `OVERLAY_${Date.now()}.pdf`);
    res.json({ ok:true, link: uploaded.webViewLink });
  } catch (e) {
    res.status(500).json({ ok:false, error:e.message });
  }
});

// ───────────────────────────────────────────────────────────────────────────────
app.listen(PORT, () => {
  log(`Server listening on ${PORT}`);
  log(`OUTPUT_FOLDER_ID (fallback): ${OUTPUT_FOLDER_ID || '(none set)'}`);
  log(`Creds: ${process.env.GOOGLE_CREDENTIALS_JSON ? 'GOOGLE_CREDENTIALS_JSON' :
    (process.env.GOOGLE_APPLICATION_CREDENTIALS ? 'GOOGLE_APPLICATION_CREDENTIALS' : 'ADC/unknown')}`);
  log(`FONT_TTF_PATH: ${process.env.FONT_TTF_PATH || '(not set — will try Noto fallback)'}`);
});
