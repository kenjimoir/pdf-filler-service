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

// ---------- Helpers for form logic ----------
function norm(s) {
  return String(s == null ? '' : s).trim();
}
function yn(s) {
  const v = norm(s).toLowerCase();
  if (['true', 'yes', 'y', '1', 'on'].includes(v)) return 'yes';
  if (['false', 'no', 'n', '0', 'off'].includes(v)) return 'no';
  if (v === 'はい') return 'yes';
  if (v === 'いいえ') return 'no';
  return v; // それ以外はそのまま返す（場所など）
}

// base_yes/base_no の両方がテンプレにあるとき、回答（yes/no）で片方をcheck・片方をuncheck
function applyYesNoPair(form, base, answer) {
  const yesName = `${base}_yes`;
  const noName  = `${base}_no`;
  const ans = yn(answer);
  let touched = 0;
  try {
    const fYes = form.getFieldMaybe ? form.getFieldMaybe(yesName) : null;
    const fNo  = form.getFieldMaybe ? form.getFieldMaybe(noName)  : null;
    if (!fYes && !fNo) return 0;

    if (ans === 'yes') {
      if (fYes && fYes.check) { fYes.check(); touched++; }
      if (fNo  && fNo.uncheck) { fNo.uncheck(); }
    } else if (ans === 'no') {
      if (fNo && fNo.check) { fNo.check(); touched++; }
      if (fYes && fYes.uncheck) { fYes.uncheck(); }
    }
    return touched;
  } catch { return 0; }
}

// 単一選択だがテンプレ側が「Base_アジア」「Base_北米」…のように複数 CheckBox で表現されている場合
// → payload の { Base: 'アジア' } で Base_アジア に check、同じBase_前方一致の他は uncheck
function applySingleChoiceBySuffix(form, base, value) {
  const target = `${base}_${norm(value)}`;
  let touched = 0;
  try {
    // form.getFields() から該当 prefix を拾って制御
    const all = form.getFields();
    const rel = all.filter(f => {
      const name = String(f.getName ? f.getName() : '');
      return name === target || name.startsWith(`${base}_`);
    });
    if (!rel.length) return 0;
    for (const f of rel) {
      const name = String(f.getName ? f.getName() : '');
      if (name === target) { f.check && f.check(); touched++; }
      else { f.uncheck && f.uncheck(); }
    }
    return touched;
  } catch { return 0; }
}

// pdf-lib のフォームに安全にアクセスするユーティリティ
function augmentForm(form) {
  if (!form.getFieldMaybe) {
    form.getFieldMaybe = function(name) {
      try { return this.getField(String(name)); } catch { return null; }
    };
  }
  return form;
}

// ---------- PDF fill ----------
async function fillPdf(srcPath, outPath, fields = {}, opts = {}) {
  const bytes = fs.readFileSync(srcPath);
  const pdfDoc = await PDFDocument.load(bytes, {
    updateFieldAppearances: true,
  });

  // 1) Register fontkit
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
        const allowSubset = ext !== '.otf'; // OTF(CFF)は無サブセット安定運用
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
    throw new Error('CJK font not embedded (FONT_TTF_PATH missing/unreadable).');
  }

  // 3) AcroForm defaults → F0 に我々のフォント
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
  const form = augmentForm(pdfDoc.getForm());
  let filled = 0;

  // 4-1) まず Yes/No の基底回答（例: { TreatmentNow: 'yes' }）を見て base_yes/base_no を処理
  for (const [rawKey, rawVal] of Object.entries(fields || {})) {
    const key = String(rawKey);
    const val = norm(rawVal);
    const ynVal = yn(val);

    // Base 回答での yes/no ペア処理
    if (ynVal === 'yes' || ynVal === 'no') {
      filled += applyYesNoPair(form, key, ynVal);
    }

    // 単一選択チェック群（Destination のように Base_value で置いてあるケース）
    // フィールド名が Base だけで、値が「アジア」などのときに Base_アジア を check する
    if (!key.includes('_') && val) {
      filled += applySingleChoiceBySuffix(form, key, val);
    }
  }

  // 4-2) 通常の “フィールド名 → 値” 直指定（テキスト / ドロップダウン / ラジオ個別 / チェック個別）
  for (const [rawKey, rawVal] of Object.entries(fields || {})) {
    const name = String(rawKey);
    const value = norm(rawVal);

    try {
      const fld = form.getFieldMaybe(name);
      if (!fld) continue;
      const typeName = (fld && fld.constructor && fld.constructor.name) || '';

      if (typeName.includes('Text')) {
        fld.setText(value);
        try { fld.updateAppearances(customFont); } catch (_) {}
        filled++;
      } else if (typeName.includes('Dropdown')) {
        try { fld.select(value); filled++; } catch (_) {}
      } else if (typeName.includes('Radio')) {
        // ラジオは select(ExportValue) で選択。日本語/英語の yes/no も受ける
        const v = yn(value);
        try { fld.select(v || value); filled++; } catch (_) {}
      } else if (typeName.includes('Check')) {
        // CheckBox は on/off しかないので、yes/true系で check
        const v = yn(value);
        if (['yes', 'true', '1', 'on'].includes(v)) { fld.check(); filled++; }
        else if (['no', 'false', '0', 'off'].includes(v)) { fld.uncheck(); }
        else {
          // 値を直接一致させる型はないため、個別チェック名に対して “存在すれば check”
          // （例: Destination_アジア: 'アジア' と送る運用は上の applySingleChoiceBySuffix で賄う）
          fld.check && fld.check(); filled++;
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

  // 7) Save
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
