// index.cjs — PDF fill → upload to Drive (Shared Drives OK)
// deps: express, cors, pdf-lib, googleapis, @pdf-lib/fontkit
// runs on Render (uses process.env.PORT) and writes temp files in os.tmpdir()

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

// ---------- Google Drive client ----------
function getDriveClient() {
  let credentials = null;
  if (process.env.GOOGLE_CREDENTIALS_JSON) {
    try { credentials = JSON.parse(process.env.GOOGLE_CREDENTIALS_JSON); } catch (e) {
      log('ERROR parsing GOOGLE_CREDENTIALS_JSON:', e && e.message);
    }
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
  const parents = parentId ? [parentId] : (OUTPUT_FOLDER_ID ? [OUTPUT_FOLDER_ID] : undefined);
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

// ---------- helpers for flexible matching ----------
function normLoose(s) {
  // NFKC normalize, trim, collapse spaces, to lower, convert full-width alnum to half
  const t = String(s ?? '').normalize('NFKC').trim().replace(/\s+/g, '').toLowerCase();
  // Full-width alnum already normalized by NFKC in most cases; just return
  return t;
}
function normYesNo(s) {
  const n = normLoose(s);
  if (n === 'はい' || n === 'yes' || n === 'true' || n === 'y' || n === '1' ) return 'yes';
  if (n === 'いいえ' || n === 'no' || n === 'false' || n === 'n' || n === '0') return 'no';
  return '';
}

// ---------- PDF fill ----------
async function fillPdf(srcPath, outPath, fields = {}, opts = {}) {
  const bytes = fs.readFileSync(srcPath);
  const pdfDoc = await PDFDocument.load(bytes, {
    updateFieldAppearances: true, // we'll pass our font later
  });

  // 1) Register fontkit (needed for custom fonts)
  try { pdfDoc.registerFontkit(fontkit); } catch (_) {}

  // 2) Embed a Japanese font (KozMin or Noto fallbacks)
  let customFont = null;
  let chosenFontPath = null;

  const fontCandidates = [
    process.env.FONT_TTF_PATH, // e.g., fonts/KozMinPr6N-Regular.otf
    path.join(ROOT, 'fonts/KozMinPr6N-Regular.otf'),
    path.join(ROOT, 'fonts/NotoSerifCJKjp-Regular.otf'),
    path.join(ROOT, 'fonts/NotoSansJP-Regular.ttf'),
    path.join(ROOT, 'fonts/NotoSansJP-Regular.otf'),
  ].filter(Boolean);

  for (const p of fontCandidates) {
    try {
      if (p && fs.existsSync(p)) {
        const fontBytes = fs.readFileSync(p);
        // Avoid subset for OTF/CFF fonts to prevent rare CFF subsetting issues
        const ext = path.extname(p).toLowerCase();
        const allowSubset = ext !== '.otf';
        customFont = await pdfDoc.embedFont(fontBytes, { subset: allowSubset });
        chosenFontPath = p;
        log('Embedded JP font:', p, '(subset:', allowSubset, ')');
        break;
      }
    } catch (e) {
      log('Font embed failed for', p, e && e.message);
    }
  }

  if (!customFont) {
    throw new Error('CJK font not embedded (FONT_TTF_PATH missing/unreadable). Cannot safely render Japanese text.');
  }

  // 3) Make AcroForm defaults point to our JP font
  let acroFormRef = pdfDoc.catalog.get(PDFName.of('AcroForm'));
  let acroForm = acroFormRef ? pdfDoc.context.lookup(acroFormRef, PDFDict) : null;
  if (!acroForm) {
    acroForm = pdfDoc.context.obj({});
    pdfDoc.catalog.set(PDFName.of('AcroForm'), acroForm);
  }

  const dr = acroForm.get(PDFName.of('DR')) || pdfDoc.context.obj({});
  const drFont = dr.get(PDFName.of('Font')) || pdfDoc.context.obj({});
  drFont.set(PDFName.of('F0'), customFont.ref);
  dr.set(PDFName.of('Font'), drFont);
  acroForm.set(PDFName.of('DR'), dr);
  acroForm.set(PDFName.of('DA'), PDFString.of('/F0 10 Tf 0 g'));
  acroForm.set(PDFName.of('NeedAppearances'), PDFBool.True);

  // 4) Fill fields
  const form = pdfDoc.getForm();
  let filled = 0;

  for (const [name, rawVal] of Object.entries(fields || {})) {
    try {
      const fld = form.getField(String(name));
      const typeName = (fld && fld.constructor && fld.constructor.name) || '';
      const val = rawVal == null ? '' : String(rawVal);

      if (typeName.includes('Text')) {
        fld.setText(val);
        try { fld.updateAppearances(customFont); } catch (_) {}
        filled++;

      } else if (typeName.includes('Check')) {
        const on = normLoose(val);
        if (['true','yes','1','on','はい','チェック','有','あり'].includes(on)) fld.check();
        else fld.uncheck();
        filled++;

      } else if (typeName.includes('Radio')) {
        // ---- Robust radio selection ----
        let ok = false;

        // 1) try direct select with given value
        try { fld.select(val); ok = true; filled++; } catch (_) {}

        // 2) try normalize & match against options
        if (!ok) {
          try {
            const options = (fld.getOptions && Array.isArray(fld.getOptions())) ? fld.getOptions() : [];
            const target = normLoose(val);
            for (const opt of options) {
              const optNorm = normLoose(opt);
              if (optNorm === target) {
                fld.select(opt);
                ok = true; filled++;
                break;
              }
            }
          } catch (_) {}
        }

        // 3) last resort: treat as yes/no
        if (!ok) {
          const yn = normYesNo(val);
          if (yn === 'yes') {
            const candidates = ['はい','Yes','yes','true','1'];
            for (const c of candidates) { try { fld.select(c); ok = true; filled++; break; } catch (_) {} }
          } else if (yn === 'no') {
            const candidates = ['いいえ','No','no','false','0'];
            for (const c of candidates) { try { fld.select(c); ok = true; filled++; break; } catch (_) {} }
          }
        }

      } else if (typeName.includes('Dropdown')) {
        // dropdown select
        try {
          fld.select(val); filled++;
        } catch (_) {
          // try normalized match among options
          try {
            const options = (fld.getOptions && Array.isArray(fld.getOptions())) ? fld.getOptions() : [];
            const target = normLoose(val);
            for (const opt of options) {
              if (normLoose(opt) === target) { fld.select(opt); filled++; break; }
            }
          } catch (_) {}
        }
      }
    } catch {
      // ignore unknown field names
    }
  }

  // 5) Bulk rebuild appearances (ALWAYS pass the font!)
  try { form.updateFieldAppearances(customFont); } catch (_) {}

  // 6) Optional watermark
  const wmText = opts.watermarkText && String(opts.watermarkText).trim();
  if (wmText) {
    for (const page of pdfDoc.getPages()) {
      const { width, height } = page.getSize();
      page.drawText(wmText, {
        x: width / 2 - 200,
        y: height / 2 - 40,
        size: 80,
        opacity: 0.12,
        rotate: degrees(45),
        color: rgb(0.85, 0.1, 0.1)
      });
    }
  }

  // 7) Save with conservative options
  let outBytes;
  try {
    outBytes = await pdfDoc.save({
      useObjectStreams: false,
      addDefaultPage: false,
      updateFieldAppearances: false,
    });
  } catch (e) {
    log('pdfDoc.save() failed, retrying with minimal options:', e.message);
    outBytes = await pdfDoc.save();
  }

  fs.writeFileSync(outPath, outBytes);
  return { outPath, filled, size: outBytes.length, fontPath: chosenFontPath };
}

// ---------- HTTP server ----------
const app = express();
app.use(cors());
app.use(express.json({ limit: '10mb' }));

app.get('/', (_req, res) => {
  res.type('text/plain').send('PDF filler is up. Try GET /health');
});

app.get('/health', (_req, res) => {
  res.json({
    ok: true,
    tmpDir: TMP,
    hasOUTPUT_FOLDER_ID: !!OUTPUT_FOLDER_ID,
    credMode: process.env.GOOGLE_CREDENTIALS_JSON
      ? 'env-json'
      : (process.env.GOOGLE_APPLICATION_CREDENTIALS ? 'file-path' : 'adc/unknown'),
    fontEnv: process.env.FONT_TTF_PATH || '(none)',
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
      const flds = form ? form.getFields() : [];
      names = flds.map(f => f.getName());
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

// ---------- Debug endpoints ----------
app.get('/debug/passthrough', async (req, res) => {
  try {
    const fileId = (req.query.fileId || '').trim();
    if (!fileId) return res.status(400).json({ ok: false, error: 'fileId required' });

    const local = path.join(TMP, `pt_${fileId}.pdf`);
    await downloadDriveFile(fileId, local);
    const uploaded = await uploadToDrive(local, `PASSTHROUGH_${Date.now()}.pdf`);
    res.json({ ok: true, link: uploaded.webViewLink });
  } catch (e) {
    res.status(500).json({ ok: false, error: e.message });
  }
});

app.get('/debug/roundtrip', async (req, res) => {
  try {
    const fileId = (req.query.fileId || '').trim();
    if (!fileId) return res.status(400).json({ ok: false, error: 'fileId required' });

    const src = path.join(TMP, `rt_${fileId}.pdf`);
    await downloadDriveFile(fileId, src);

    const bytes = fs.readFileSync(src);
    const doc = await PDFDocument.load(bytes, { updateFieldAppearances: false });
    const outBytes = await doc.save({ useObjectStreams: false });
    const out = path.join(TMP, `ROUNDTRIP_${Date.now()}.pdf`);
    fs.writeFileSync(out, outBytes);

    const uploaded = await uploadToDrive(out, `ROUNDTRIP_${Date.now()}.pdf`);
    res.json({ ok: true, link: uploaded.webViewLink });
  } catch (e) {
    res.status(500).json({ ok: false, error: e.message });
  }
});

app.get('/debug/overlay', async (req, res) => {
  try {
    const fileId = (req.query.fileId || '').trim();
    if (!fileId) return res.status(400).json({ ok: false, error: 'fileId required' });

    const src = path.join(TMP, `ov_${fileId}.pdf`);
    await downloadDriveFile(fileId, src);

    const bytes = fs.readFileSync(src);
    const doc = await PDFDocument.load(bytes, { updateFieldAppearances: false });
    try { doc.registerFontkit(fontkit); } catch (_) {}

    // Use same font source as fillPdf
    let fontBytes = null;
    const pref = process.env.FONT_TTF_PATH;
    const candidates = [
      pref,
      path.join(ROOT, 'fonts/KozMinPr6N-Regular.otf'),
      path.join(ROOT, 'fonts/NotoSerifCJKjp-Regular.otf'),
      path.join(ROOT, 'fonts/NotoSansJP-Regular.ttf'),
      path.join(ROOT, 'fonts/NotoSansJP-Regular.otf'),
    ].filter(Boolean);
    for (const p of candidates) {
      if (p && fs.existsSync(p)) { fontBytes = fs.readFileSync(p); break; }
    }
    const font = fontBytes ? await doc.embedFont(fontBytes, { subset: false }) : undefined;

    const page = doc.getPages()[0];
    page.drawText('VISIBLE OVERLAY TEST あいうえお 山田太郎', {
      x: 48, y: 720, size: 14, font, color: rgb(0, 0, 0)
    });

    const outBytes = await doc.save({ useObjectStreams: false });
    const out = path.join(TMP, `OVERLAY_${Date.now()}.pdf`);
    fs.writeFileSync(out, outBytes);

    const uploaded = await uploadToDrive(out, `OVERLAY_${Date.now()}.pdf`);
    res.json({ ok: true, link: uploaded.webViewLink });
  } catch (e) {
    res.status(500).json({ ok: false, error: e.message });
  }
});

app.listen(PORT, () => {
  log(`Server listening on ${PORT}`);
  log(`OUTPUT_FOLDER_ID (fallback): ${OUTPUT_FOLDER_ID || '(none set)'}`);
  log(`Creds: ${process.env.GOOGLE_CREDENTIALS_JSON ? 'GOOGLE_CREDENTIALS_JSON'
    : (process.env.GOOGLE_APPLICATION_CREDENTIALS ? 'GOOGLE_APPLICATION_CREDENTIALS' : 'ADC/unknown')}`);
  log(`FONT_TTF_PATH: ${process.env.FONT_TTF_PATH || '(none)'}`);
});
