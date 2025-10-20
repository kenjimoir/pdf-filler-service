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

/* ===== Yes/No 正規化 & 補助 ===== */
function normJaEnYesNo(v) {
  const s = String(v ?? '').trim().toLowerCase().normalize('NFKC');
  if (['はい','yes','y','true','1','on'].includes(s)) return 'yes';
  if (['いいえ','no','n','false','0','off'].includes(s)) return 'no';
  return '';
}
const YES_SUFFIXES = ['_はい','_yes','-はい','-yes',' (はい)',' (yes)','_Yes','-Yes',' (Yes)'];
const NO_SUFFIXES  = ['_いいえ','_no','-いいえ','-no',' (いいえ)',' (no)','_No','-No',' (No)'];

function stripKnownSuffix(name) {
  const nk = name.normalize('NFKC').toLowerCase();
  for (const suf of YES_SUFFIXES) {
    const s = suf.normalize('NFKC').toLowerCase();
    if (nk.endsWith(s)) return { base: name.slice(0, name.length - suf.length), kind: 'yes', suffix: suf };
  }
  for (const suf of NO_SUFFIXES) {
    const s = suf.normalize('NFKC').toLowerCase();
    if (nk.endsWith(s)) return { base: name.slice(0, name.length - suf.length), kind: 'no', suffix: suf };
  }
  return { base: name, kind: '' };
}
function tryGetField(form, base, suffixes) {
  for (const suf of suffixes) {
    try { return form.getField(base + suf); } catch (_) {}
  }
  return null;
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

// ---------- PDF fill ----------
async function fillPdf(srcPath, outPath, fields = {}, opts = {}) {
  const bytes = fs.readFileSync(srcPath);
  const pdfDoc = await PDFDocument.load(bytes, { updateFieldAppearances: true });

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
        const ext = path.extname(p).toLowerCase();
        const allowSubset = ext !== '.otf'; // OTF(CFF) はsubsetting回避
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

  // 3) AcroForm defaults → our JP font
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

      // ---- Text ----
      if (typeName.includes('Text')) {
        fld.setText(val);
        try { fld.updateAppearances(customFont); } catch (_) {}
        filled++;

      // ---- Radio / Checkbox（はい・いいえ完全対応）----
      } else if (typeName.includes('Radio') || typeName.includes('Check')) {
        const yn = normJaEnYesNo(val);

        if (typeName.includes('Radio')) {
          // まずそのまま
          let ok = false;
          try { fld.select(val); ok = true; filled++; } catch (_) {}

          if (!ok) {
            const candidates = [val];
            if (yn === 'yes') candidates.push('はい','Yes','yes','TRUE','true','1');
            if (yn === 'no')  candidates.push('いいえ','No','no','FALSE','false','0');
            for (const c of candidates) {
              try { fld.select(String(c)); ok = true; filled++; break; } catch (_) {}
            }
          }
        } else {
          // チェックボックス
          // 1) 名前末尾に *_はい / *_いいえ / *_yes / *_no が付いているかを判定
          const info = stripKnownSuffix(name);
          let handled = false;

          if (info.kind) {
            // このフィールド自体が yes/no 片方を担っている
            if (info.kind === 'yes') {
              if (yn === 'yes') fld.check(); else fld.uncheck();
              handled = true; filled++;
            } else if (info.kind === 'no') {
              if (yn === 'no') fld.check(); else fld.uncheck();
              handled = true; filled++;
            }
          }

          // 2) 相方（*_はい / *_いいえ 等）を探して両方制御
          if (!handled) {
            const base = info.base;
            const yesField = tryGetField(form, base, YES_SUFFIXES);
            const noField  = tryGetField(form, base, NO_SUFFIXES);

            if (yesField || noField) {
              if (yn === 'yes') {
                try { yesField && yesField.check(); } catch(_) {}
                try { noField && noField.uncheck(); } catch(_) {}
                filled++;
              } else if (yn === 'no') {
                try { noField && noField.check(); } catch(_) {}
                try { yesField && yesField.uncheck(); } catch(_) {}
                filled++;
              } else {
                // 値が yes/no でない場合は何もしない（既定の状態を尊重）
              }
              handled = true;
            }
          }

          // 3) それでも相方が見つからない＝単独チェックボックス
          if (!handled) {
            if (yn) {
              if (yn === 'no') fld.check(); else fld.uncheck(); // No=チェック という運用に合わせる
              filled++;
            } else {
              // 汎用トグル（true/1/on/はい/yes ならチェック）
              const on = val.trim().toLowerCase().normalize('NFKC');
              if (['true','1','on','はい','yes','y'].includes(on)) fld.check(); else fld.uncheck();
              filled++;
            }
          }
        }

      // ---- Dropdown ----
      } else if (typeName.includes('Dropdown')) {
        try { fld.select(val); filled++; } catch (_) {}
      }

    } catch {
      // unknown field name → 無視
    }
  }

  // 5) Bulk rebuild appearances
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

  // 7) Save with conservative options (size & compatibility)
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

    // same font as fillPdf
    let fontBytes = null;
    const candidates = [
      process.env.FONT_TTF_PATH,
      path.join(ROOT, 'fonts/KozMinPr6N-Regular.otf'),
      path.join(ROOT, 'fonts/NotoSerifCJKjp-Regular.otf'),
      path.join(ROOT, 'fonts/NotoSansJP-Regular.ttf'),
      path.join(ROOT, 'fonts/NotoSansJP-Regular.otf'),
    ].filter(Boolean);
    for (const p of candidates) { if (p && fs.existsSync(p)) { fontBytes = fs.readFileSync(p); break; } }
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
    : (process.env.GOOGLE_APPLICATION_CREDENTIALS ? 'GOOGLE_APPLICATIONS_CREDENTIALS' : 'ADC/unknown')}`);
  log(`FONT_TTF_PATH: ${process.env.FONT_TTF_PATH || '(none)'}`);
});
