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

/* -----------------------------------------------------------
   日本語 → ASCII（英語）マッピング
   - ラジオ/ドロップダウンの「書き出し値」を ASCII に統一するため
   - テンプレート側の Export Value を ASCII に直した前提
   - 直していない場合でも、英語で失敗 → 日本語でも select() を試すフォールバックあり
----------------------------------------------------------- */
const MAP_YESNO = {
  'はい': 'yes', 'いいえ': 'no',
  'YES': 'yes', 'Yes': 'yes', 'yes': 'yes',
  'NO': 'no', 'No': 'no', 'no': 'no',
};

const MAP_REGION = {
  'アジア': 'asia',
  'ヨーロッパ': 'europe',
  'オセアニア': 'oceania',
  '北米': 'north_america',
  '中南米': 'latin_america',
  'アフリカ': 'africa',
  '中東': 'middle_east',
  'その他': 'other',
};

const JP_OF = {
  yes: 'はい', no: 'いいえ',
  asia: 'アジア',
  europe: 'ヨーロッパ',
  oceania: 'オセアニア',
  north_america: '北米',
  latin_america: '中南米',
  africa: 'アフリカ',
  middle_east: '中東',
  other: 'その他',
};

/** フィールド名でラジオ/ドロップダウンっぽいものをざっくり判定 */
function looksLikeRadioName(name) {
  if (!name) return false;
  // 例: Q1_TreatmentNow, Q6_Sanctioned, Q7_DangerousWork, Q8_JobDuty
  if (/^Q\d+_/i.test(name)) return true;
  // 例: Region/TravelRegion/MainRegion 等
  if (/region/i.test(name)) return true;
  // yes/no 系
  if (/YesNo$/i.test(name)) return true;
  return false;
}

/** 日本語→英語の正規化（該当しなければそのまま返す） */
function normalizeRadioValue(fieldName, raw) {
  const v = String(raw ?? '').trim();
  if (!v) return v;
  if (MAP_YESNO[v] != null) return MAP_YESNO[v];
  if (MAP_REGION[v] != null) return MAP_REGION[v];
  // すでに英語ならそのまま
  if (/^[A-Za-z0-9_\-]+$/.test(v)) return v;
  return v;
}

/* -----------------------------------------------------------
   Google Drive helpers
----------------------------------------------------------- */
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

/* -----------------------------------------------------------
   PDF fill
----------------------------------------------------------- */
async function fillPdf(srcPath, outPath, fields = {}, opts = {}) {
  const bytes = fs.readFileSync(srcPath);
  const pdfDoc = await PDFDocument.load(bytes, {
    updateFieldAppearances: true,
  });

  // fontkit
  try { pdfDoc.registerFontkit(fontkit); } catch (_) {}

  // フォント埋め込み（KozMin 推奨、Noto 系 fallback）
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
        const allowSubset = ext !== '.otf'; // OTF/CFF は非サブセット安定
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

  // AcroForm → F0 デフォルト設定
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

  // ---- 事前に 日本語→英語 へ正規化（ラジオ/ドロップダウン想定フィールドのみ）----
  const normalized = {};
  for (const [k, v] of Object.entries(fields || {})) {
    if (looksLikeRadioName(k)) {
      normalized[k] = normalizeRadioValue(k, v);
    } else {
      normalized[k] = v;
    }
  }

  // ---- 塗り ----
  const form = pdfDoc.getForm();
  let filled = 0;

  for (const [name, rawVal] of Object.entries(normalized)) {
    try {
      const fld = form.getField(String(name));
      const typeName = (fld && fld.constructor && fld.constructor.name) || '';
      const origVal = fields[name]; // 入力の生値（日本語の可能性あり）
      const val = rawVal == null ? '' : String(rawVal);

      // テキスト
      if (typeName.includes('Text')) {
        fld.setText(val);
        try { fld.updateAppearances(customFont); } catch (_) {}
        filled++;

      // チェックボックス（yes/no/true/1/はい → ON）
      } else if (typeName.includes('Check')) {
        const on = String(fields[name]).toLowerCase();
        if (
          ['true','1','on','yes','y','はい','有','あり','true/yes'].includes(on)
        ) fld.check(); else fld.uncheck();
        filled++;

      // ラジオ
      } else if (typeName.includes('Radio')) {
        let selected = false;

        // ① ASCII（英語）で試す
        try { fld.select(val); selected = true; } catch (_) {}

        // ② ダメなら日本語に戻して試す（テンプレ側が未更新でも動く）
        if (!selected) {
          const jp = JP_OF[val] || String(origVal || '');
          if (jp) {
            try { fld.select(jp); selected = true; } catch(_) {}
          }
        }

        // ③ ダメなら大文字小文字差分などいくつか試す
        if (!selected && val) {
          try { fld.select(val.toUpperCase()); selected = true; } catch(_) {}
          if (!selected) { try { fld.select(val.toLowerCase()); selected = true; } catch(_) {} }
        }

        if (!selected) {
          log(`WARN radio select failed for field="${name}" with val="${val}" (orig="${origVal}")`);
        } else {
          filled++;
        }

      // ドロップダウン
      } else if (typeName.includes('Dropdown')) {
        let selected = false;
        try { fld.select(val); selected = true; } catch(_) {}
        if (!selected) {
          const jp = JP_OF[val] || String(origVal || '');
          if (jp) { try { fld.select(jp); selected = true; } catch(_) {} }
        }
        if (!selected) {
          log(`WARN dropdown select failed for field="${name}" val="${val}" orig="${origVal}"`);
        } else {
          filled++;
        }
      }
    } catch (e) {
      // 不明フィールド名は無視
    }
  }

  // 全体の外観再生成（フォント指定）
  try { form.updateFieldAppearances(customFont); } catch (_) {}

  // 透かし（任意）
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

  // 保存（互換性寄り）
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

/* -----------------------------------------------------------
   HTTP server
----------------------------------------------------------- */
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

/* -----------------------------------------------------------
   Debug endpoints（任意）
----------------------------------------------------------- */
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

    // フォント解決
    let fontBytes = null;
    const candidates = [
      process.env.FONT_TTF_PATH,
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
