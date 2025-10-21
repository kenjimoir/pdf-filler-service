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

/* ---------------- helpers: normalization & alias ---------------- */

function stripWeird(s) {
  if (s == null) return '';
  let t = String(s);
  // remove ASCII quotes, Japanese quotes, zero-width etc, and trim
  t = t
    .replace(/[\u200B-\u200D\uFEFF]/g, '')        // zero-width
    .replace(/[“”„‟〝〞＂"]/g, '')                 // various double quotes
    .replace(/[‘’‚‛′＇']/g, '')                   // various single quotes
    .replace(/\s+/g, ' ')
    .trim();
  return t;
}
function normalizeYesNo(vRaw) {
  const v = stripWeird(vRaw).toLowerCase();
  if (!v) return '';
  if (['yes','true','1','on'].includes(v) || v === 'はい') return 'yes';
  if (['no','false','0','off'].includes(v) || v === 'いいえ') return 'no';
  // allow first char match (は / い) as last resort
  if (v.startsWith('は')) return 'yes';
  if (v.startsWith('い')) return 'no';
  return v; // return as-is; caller will compare further if needed
}

function normalizeRegion(vRaw) {
  const v = stripWeird(vRaw);
  const t = v.toLowerCase();
  // map English to Japanese labels you use on the PDF
  const map = {
    'asia': 'アジア',
    'europe': 'ヨーロッパ',
    'oceania': 'オセアニア',
    'north america': '北米',
    'south america': '中南米',
    'latin america': '中南米',
    'africa': 'アフリカ',
    'middle east': '中東',
    'other': 'その他',
  };
  return map[t] || v; // if already Japanese, just return
}

// map possible incoming keys → logical base keys used for _yes/_no pairs
function buildAliasView(fieldsIn) {
  const f = fieldsIn || {};
  const out = { ...f };

  const aliasPairs = [
    ['TreatmentNow', 'Q1_TreatmentNow'],
    ['SeriousHistory', 'Q2_SeriousHistory'],
    ['LuggageClaims5Plus', 'Q3_LuggageClaims5Plus'],
    ['DuplicateContracts', 'Q4_DuplicateContracts'],
    ['SanctionedCountries', 'Q6_SanctionedCountries'],
    ['WorkDuringTravel', 'Q7_JobDuringTravel'],
  ];
  for (const [base, alt] of aliasPairs) {
    if (out[base] == null && out[alt] != null) out[base] = out[alt];
  }

  // DestinationRegion: we keep both (Q5_* is your GAS key)
  if (out['DestinationRegion'] == null && out['Q5_DestinationRegion'] != null) {
    out['DestinationRegion'] = out['Q5_DestinationRegion'];
  }

  return out;
}

/* ---------------- PDF fill core ---------------- */

async function fillPdf(srcPath, outPath, fields = {}, opts = {}) {
  const bytes = fs.readFileSync(srcPath);
  const pdfDoc = await PDFDocument.load(bytes, {
    updateFieldAppearances: true,
  });

  // 1) font
  try { pdfDoc.registerFontkit(fontkit); } catch (_) {}
  let customFont = null;
  let chosenFontPath = null;
  const fontCandidates = [
    process.env.FONT_TTF_PATH,
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
        const allowSubset = true;
        customFont = await pdfDoc.embedFont(fontBytes, { subset: allowSubset });
        chosenFontPath = p;
        log('Embedded JP font:', p, '(subset:', allowSubset, ')');
        break;
      }
    } catch (e) {
      log('Font embed failed for', p, e && e.message);
    }
  }
  if (!customFont) throw new Error('CJK font not embedded. Set FONT_TTF_PATH.');

  // 2) AcroForm default appearance
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
  acroForm.set(PDFName.of('DA'), PDFString.of('/F0 0 Tf 0 g'));
  acroForm.set(PDFName.of('NeedAppearances'), PDFBool.True);

  // 3) fill (robust yes/no + alias)
  const form = pdfDoc.getForm();
  const allFields = form.getFields();
  const valueBy = buildAliasView(fields);
  let filled = 0;

  for (const f of allFields) {
    const name = f.getName ? f.getName() : '';
    const ctor = f.constructor && f.constructor.name || '';
    const valRaw = fields[name]; // exact-name match first

    // Text/Dropdown/Radio by exact key
    if (ctor.includes('Text')) {
      if (valRaw != null) {
        f.setText(String(valRaw));
        try { f.updateAppearances(customFont); } catch (_) {}
        filled++;
      }
      continue;
    }
    if (ctor.includes('Dropdown')) {
      if (valRaw != null) {
        try { f.select(String(valRaw)); filled++; } catch (_) {}
      }
      continue;
    }
    if (ctor.includes('Radio')) {
      if (valRaw != null) {
        try { f.select(String(valRaw)); filled++; } catch (_) {}
      }
      continue;
    }

    // Checkboxes: 3パターン
    if (ctor.includes('Check')) {
      const n = String(name);

      // 3-1) ペア方式: Foo_yes / Foo_no
      const m = n.match(/^(.*)_(yes|no)$/i);
      if (m) {
        const base = m[1];
        const isYesBox = m[2].toLowerCase() === 'yes';
        const baseVal = valueBy[base];
        if (baseVal != null) {
          const yn = normalizeYesNo(baseVal);
          if ((isYesBox && yn === 'yes') || (!isYesBox && yn === 'no')) f.check();
          else f.uncheck();
          filled++;
        } else {
          // no info -> uncheck to be safe
          f.uncheck();
        }
        continue;
      }

      // 3-2) DestinationRegion_アジア のような単項目一致  ← (reverted as requested)
      const rm = n.match(/^DestinationRegion_(.+)$/);
      if (rm) {
        const want = stripWeird(rm[1]);
        const given = normalizeRegion(valueBy['DestinationRegion'] || '');
        if (want && given && want === given) f.check(); else f.uncheck();
        filled++;
        continue;
      }

      // 3-3) 単独の yes/no キー（フィールド名そのものが Foo で、値が yes/no）
      if (valueBy[n] != null) {
        const yn = normalizeYesNo(valueBy[n]);
        if (yn === 'yes' || yn === 'on' || yn === '1' || yn === 'true') f.check();
        else f.uncheck();
        filled++;
        continue;
      }

      // それ以外は扱わない（名寄せできない）
    }
  }

  // 4) rebuild appearances
  try { form.updateFieldAppearances(customFont); } catch (_) {}

  // 5) watermark
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

  // 6) save
  let outBytes;
  try {
    outBytes = await pdfDoc.save({
      useObjectStreams: false,
      addDefaultPage: false,
      updateFieldAppearances: false,
    });
  } catch (e) {
    log('pdfDoc.save() failed, retry with defaults:', e.message);
    outBytes = await pdfDoc.save();
  }

  fs.writeFileSync(outPath, outBytes);
  return { outPath, filled, size: outBytes.length, fontPath: chosenFontPath };
}

/* ---------------- HTTP server ---------------- */

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

    // debug: log the exact incoming values after normalization
    const aliasView = buildAliasView(fields || {});
    const probe = {};
    for (const k of ['TreatmentNow','SeriousHistory','LuggageClaims5Plus','DuplicateContracts','SanctionedCountries','WorkDuringTravel','DestinationRegion']) {
      if (aliasView[k] != null) probe[k] = stripWeird(aliasView[k]);
    }
    log('INCOMING (probe):', JSON.stringify(probe));

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

/* ---------------- Debug endpoints (unchanged) ---------------- */

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
