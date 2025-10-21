// index.cjs — PDF fill → upload to Drive (Shared Drives OK)
// deps: express, cors, pdf-lib, googleapis, @pdf-lib/fontkit

const fs = require('fs');
const os = require('os');
const path = require('path');
const express = require('express');
const cors = require('cors');
const {
  PDFDocument, rgb, degrees, PDFName, PDFBool, PDFDict, PDFString
} = require('pdf-lib');
const { google } = require('googleapis');
const fontkit = require('@pdf-lib/fontkit');

const PORT = process.env.PORT || 8080;
const TMP  = path.join(os.tmpdir(), 'pdf-filler');
ensureDir(TMP);
const ROOT = process.cwd();

const OUTPUT_FOLDER_ID = process.env.OUTPUT_FOLDER_ID || '';

function log(...a){ console.log(new Date().toISOString(), '-', ...a); }
function ensureDir(p){ if(!fs.existsSync(p)) fs.mkdirSync(p, {recursive:true}); }

// ---------- Google Drive ----------
function getDriveClient() {
  let credentials = null;
  if (process.env.GOOGLE_CREDENTIALS_JSON) {
    try { credentials = JSON.parse(process.env.GOOGLE_CREDENTIALS_JSON); }
    catch(e){ log('ERROR parsing GOOGLE_CREDENTIALS_JSON:', e && e.message); }
  }
  const auth = new google.auth.GoogleAuth({
    scopes: ['https://www.googleapis.com/auth/drive'],
    ...(credentials ? { credentials } : {}),
  });
  return google.drive({ version: 'v3', auth });
}

async function downloadDriveFile(fileId, dest) {
  const drive = getDriveClient();
  log('Downloading template', fileId, '→', dest);
  const res = await drive.files.get(
    { fileId, alt:'media', supportsAllDrives:true },
    { responseType:'stream' }
  );
  await new Promise((resolve, reject) => {
    const out = fs.createWriteStream(dest);
    res.data.on('error', reject).pipe(out).on('finish', resolve);
  });
  return dest;
}
async function uploadToDrive(localPath, name, parentId) {
  const drive = getDriveClient();
  const parents = parentId ? [parentId] : (OUTPUT_FOLDER_ID ? [OUTPUT_FOLDER_ID] : undefined);
  const fileMetadata = { name, parents };
  const media = { mimeType:'application/pdf', body: fs.createReadStream(localPath) };
  const res = await drive.files.create({
    requestBody: fileMetadata,
    media,
    fields:'id,name,webViewLink,parents',
    supportsAllDrives:true,
  });
  return res.data;
}

// ---------- Helpers for value mapping ----------
const YES_WORDS  = ['true','1','on','yes'];
const NO_WORDS   = ['false','0','off','no'];
const JA_YES     = 'はい';
const JA_NO      = 'いいえ';

// 英→日マップ（旅行先など）
const REGION_MAP_EN2JA = {
  'asia':'アジア', 'europe':'ヨーロッパ', 'oceania':'オセアニア', 'north america':'北米',
  'latin america':'中南米', 'central & south america':'中南米', 'central and south america':'中南米',
  'africa':'アフリカ', 'middle east':'中東', 'other':'その他', 'others':'その他'
};

// 特定フィールド向け変換（名前は PDF 側のフィールド名で判定）
function normalizeForField(fieldName, raw) {
  if (raw == null) return '';
  const v = String(raw).trim();

  // ラジオ/ドロップダウンで「はい/いいえ」系を日本語に（PDF が日本語 Export Value の場合）
  if (/_?(YesNo|Yes|No|YN)$/i.test(fieldName) || /Q\d+_/.test(fieldName)) {
    const lower = v.toLowerCase();
    if (YES_WORDS.includes(lower)) return JA_YES;
    if (NO_WORDS.includes(lower))  return JA_NO;
  }

  // 旅行先（例: Destination / MainDestination / TravelRegion）
  if (/(destination|region|travelRegion|mainDestination)/i.test(fieldName)) {
    const lower = v.toLowerCase();
    if (REGION_MAP_EN2JA[lower]) return REGION_MAP_EN2JA[lower];
  }

  return v;
}

// 可能ならチェックボックスの On 値（Export Value）を取り出す
function getCheckBoxOnValueSafe(pdfField) {
  try {
    // pdf-lib の内部: PDFAcroCheckBox.getOnValue() -> PDFName('/Yes' 等)
    const acro = pdfField?.acroField;
    const onVal = acro?.getOnValue?.();
    if (onVal) {
      const s = String(onVal); // '/Yes' など
      return s.replace(/^\//, '');
    }
  } catch(_) {}
  return ''; // 取れないなら空
}

// name が *_yes / *_no の二個組チェックボックス命名か？
function splitYesNoPair(name) {
  const m = String(name).match(/^(.*)_(yes|no)$/i);
  if (!m) return null;
  return { base: m[1], which: m[2].toLowerCase() }; // {base:'Q1_TreatmentNow', which:'yes'|'no'}
}

function truthyWord(s) { return YES_WORDS.includes(String(s).trim().toLowerCase()); }
function falsyWord(s)  { return NO_WORDS.includes(String(s).trim().toLowerCase()); }

// ---------- Fill PDF ----------
async function fillPdf(srcPath, outPath, fields = {}, opts = {}) {
  const bytes = fs.readFileSync(srcPath);
  const pdfDoc = await PDFDocument.load(bytes, { updateFieldAppearances:true });
  try { pdfDoc.registerFontkit(fontkit); } catch(_) {}

  // 日本語フォント埋め込み
  let customFont = null;
  let chosenFont = null;
  const candidates = [
    process.env.FONT_TTF_PATH,
    path.join(ROOT, 'fonts/KozMinPr6N-Regular.otf'),
    path.join(ROOT, 'fonts/NotoSerifCJKjp-Regular.otf'),
    path.join(ROOT, 'fonts/NotoSansJP-Regular.ttf'),
    path.join(ROOT, 'fonts/NotoSansJP-Regular.otf'),
  ].filter(Boolean);

  for (const p of candidates) {
    try {
      if (p && fs.existsSync(p)) {
        const fb = fs.readFileSync(p);
        const ext = path.extname(p).toLowerCase();
        const allowSubset = ext !== '.otf'; // OTF(CFF)はノーサブセット安定
        customFont = await pdfDoc.embedFont(fb, { subset: allowSubset });
        chosenFont = p;
        log('Embedded JP font:', p, '(subset:', allowSubset, ')');
        break;
      }
    } catch(e){ log('Font embed failed:', p, e && e.message); }
  }
  if (!customFont) throw new Error('CJK font not embedded');

  // AcroForm Default Appearance (DA) / Default Resources (DR)
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

  const form = pdfDoc.getForm();
  let filled = 0;

  // 先にフィールド辞書を構築（相方のチェックを外す用途）
  const fieldByName = {};
  for (const f of form.getFields()) fieldByName[f.getName()] = f;

  // 入力ループ
  for (const [rawName, rawVal] of Object.entries(fields || {})) {
    const name = String(rawName);
    const value = normalizeForField(name, rawVal); // 英→日や はい/いいえ 変換
    const vLower = String(value).trim().toLowerCase();

    let fld;
    try { fld = form.getField(name); } catch { fld = null; }
    if (!fld) {
      // yes/no の二個組対応（*_yes / *_no）
      const pair = splitYesNoPair(name);
      if (pair) {
        const yesName = `${pair.base}_yes`;
        const noName  = `${pair.base}_no`;
        const yesFld = fieldByName[yesName];
        const noFld  = fieldByName[noName];

        const wantYes = truthyWord(vLower) || vLower === JA_YES.toLowerCase();
        const wantNo  = falsyWord(vLower)  || vLower === JA_NO.toLowerCase();

        if (yesFld) {
          try {
            const onVal = getCheckBoxOnValueSafe(yesFld);
            if (wantYes || (onVal && vLower === onVal.toLowerCase())) yesFld.check(); else yesFld.uncheck();
            filled++;
          } catch(_) {}
        }
        if (noFld) {
          try {
            const onVal = getCheckBoxOnValueSafe(noFld);
            if (wantNo || (onVal && vLower === onVal.toLowerCase())) noFld.check(); else noFld.uncheck();
            filled++;
          } catch(_) {}
        }
        continue;
      }
      // 見つからない名前はスキップ
      continue;
    }

    const typeName = (fld && fld.constructor && fld.constructor.name) || '';

    try {
      if (typeName.includes('Text')) {
        fld.setText(String(value));
        try { fld.updateAppearances(customFont); } catch(_) {}
        filled++;
      }
      else if (typeName.includes('Check')) {
        // チェックボックス：Export Value と比較
        const onVal = getCheckBoxOnValueSafe(fld); // 例: 'no', 'Yes', 'On'
        let shouldCheck = false;

        if (onVal) {
          // Export Value と完全一致か、一般的 truthy/ falsy の規則で判定
          if (vLower === onVal.toLowerCase()) shouldCheck = true;
          else if (truthyWord(vLower) && ['yes','on','true','1'].includes(onVal.toLowerCase())) shouldCheck = true;
          else if (falsyWord(vLower) && ['no','off','false','0'].includes(onVal.toLowerCase())) shouldCheck = true;
        } else {
          // On 値が取れない場合のフォールバック
          shouldCheck = truthyWord(vLower) || vLower === JA_YES.toLowerCase();
        }

        if (shouldCheck) fld.check(); else fld.uncheck();
        filled++;

        // もし *_yes/*_no の相方が居れば片方外す
        const pair = splitYesNoPair(name);
        if (pair) {
          const otherName = pair.which === 'yes' ? `${pair.base}_no` : `${pair.base}_yes`;
          const otherFld  = fieldByName[otherName];
          if (otherFld) {
            try {
              const otherOn = getCheckBoxOnValueSafe(otherFld);
              // このフィールドをチェックした場合は相方を外す
              if (shouldCheck) otherFld.uncheck();
              // このフィールドを外した場合でも、値が完全に yes/no の時は相方側を合わせておく
              else {
                if (pair.which === 'yes' && (falsyWord(vLower) || vLower === JA_NO.toLowerCase() || (otherOn && vLower === otherOn.toLowerCase()))) {
                  otherFld.check();
                }
                if (pair.which === 'no' && (truthyWord(vLower) || vLower === JA_YES.toLowerCase() || (otherOn && vLower === otherOn.toLowerCase()))) {
                  otherFld.check();
                }
              }
            } catch(_) {}
          }
        }
      }
      else if (typeName.includes('Radio')) {
        // ラジオは Export Value と完全一致が必要。normalizeForField 済みの値で試す
        try { fld.select(String(value)); filled++; }
        catch {
          // それでも NG の場合、はい/いいえ英語→日本語の再マップを試す
          const alt = (truthyWord(vLower) ? JA_YES : (falsyWord(vLower) ? JA_NO : value));
          try { fld.select(String(alt)); filled++; } catch(_) {}
        }
      }
      else if (typeName.includes('Dropdown')) {
        // ドロップダウンも同様
        try { fld.select(String(value)); filled++; }
        catch {
          const alt = (truthyWord(vLower) ? JA_YES : (falsyWord(vLower) ? JA_NO : value));
          try { fld.select(String(alt)); filled++; } catch(_) {}
        }
      }
    } catch(_) {}
  }

  // 全体の外観再構築 & flatten
  try { form.updateFieldAppearances(customFont); } catch(_) {}
  try { form.flatten(); } catch(_) {}

  // 透かし
  const wm = (opts.watermarkText && String(opts.watermarkText).trim()) || '';
  if (wm) {
    for (const page of pdfDoc.getPages()) {
      const { width, height } = page.getSize();
      page.drawText(wm, {
        x: width/2 - 200, y: height/2 - 40, size: 80, opacity: 0.12,
        rotate: degrees(45), color: rgb(0.85, 0.1, 0.1)
      });
    }
  }

  // 保存（互換寄り）
  let outBytes;
  try {
    outBytes = await pdfDoc.save({
      useObjectStreams:false,
      addDefaultPage:false,
      updateFieldAppearances:false,
    });
  } catch(e){
    log('pdfDoc.save() failed, retry minimal:', e.message);
    outBytes = await pdfDoc.save();
  }
  fs.writeFileSync(outPath, outBytes);
  return { outPath, filled, size: outBytes.length, fontPath: chosenFont };
}

// ---------- HTTP server ----------
const app = express();
app.use(cors());
app.use(express.json({ limit:'10mb' }));

app.get('/', (_req,res)=> res.type('text/plain').send('PDF filler is up. Try GET /health'));

app.get('/health', (_req,res)=>{
  res.json({
    ok:true,
    tmpDir: TMP,
    hasOUTPUT_FOLDER_ID: !!OUTPUT_FOLDER_ID,
    credMode: process.env.GOOGLE_CREDENTIALS_JSON ? 'env-json'
      : (process.env.GOOGLE_APPLICATION_CREDENTIALS ? 'file-path' : 'adc/unknown'),
    fontEnv: process.env.FONT_TTF_PATH || '(none)',
  });
});

app.get('/fields', async (req,res)=>{
  try {
    const fileId = (req.query.fileId || '').trim();
    if (!fileId) return res.status(400).json({ error:'fileId is required' });

    const local = path.join(TMP, `template_${fileId}.pdf`);
    await downloadDriveFile(fileId, local);

    const bytes = fs.readFileSync(local);
    const pdfDoc = await PDFDocument.load(bytes);
    let names = [];
    try {
      const form = pdfDoc.getForm();
      names = (form ? form.getFields() : []).map(f => f.getName());
    } catch(_) {}
    res.json({ count:names.length, names });
  } catch(e){
    log('List fields failed:', e && (e.stack || e));
    res.status(500).json({ error:'List fields failed', detail:e.message });
  }
});

app.post('/fill', async (req,res)=>{
  try {
    const { templateFileId, fields, outputName, folderId, mode, watermarkText } = req.body || {};
    if (!templateFileId) return res.status(400).json({ error:'templateFileId is required' });

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

    res.json({ ok:true, filledCount: result.filled, driveFile: uploaded, webViewLink: uploaded.webViewLink });
  } catch(err){
    log('ERROR /fill:', err && (err.stack || err));
    res.status(500).json({ error:'Fill failed', detail: err.message });
  }
});

// ---- Debug helpers (unchanged) ----
app.get('/debug/passthrough', async (req,res)=>{
  try {
    const fileId = (req.query.fileId || '').trim();
    if (!fileId) return res.status(400).json({ ok:false, error:'fileId required' });
    const local = path.join(TMP, `pt_${fileId}.pdf`);
    await downloadDriveFile(fileId, local);
    const up = await uploadToDrive(local, `PASSTHROUGH_${Date.now()}.pdf`);
    res.json({ ok:true, link: up.webViewLink });
  } catch(e){ res.status(500).json({ ok:false, error:e.message }); }
});

app.get('/debug/roundtrip', async (req,res)=>{
  try {
    const fileId = (req.query.fileId || '').trim();
    if (!fileId) return res.status(400).json({ ok:false, error:'fileId required' });
    const src = path.join(TMP, `rt_${fileId}.pdf`);
    await downloadDriveFile(fileId, src);
    const bytes = fs.readFileSync(src);
    const doc = await PDFDocument.load(bytes, { updateFieldAppearances:false });
    const outBytes = await doc.save({ useObjectStreams:false });
    const out = path.join(TMP, `ROUNDTRIP_${Date.now()}.pdf`);
    fs.writeFileSync(out, outBytes);
    const up = await uploadToDrive(out, `ROUNDTRIP_${Date.now()}.pdf`);
    res.json({ ok:true, link: up.webViewLink });
  } catch(e){ res.status(500).json({ ok:false, error:e.message }); }
});

app.get('/debug/overlay', async (req,res)=>{
  try {
    const fileId = (req.query.fileId || '').trim();
    if (!fileId) return res.status(400).json({ ok:false, error:'fileId required' });
    const src = path.join(TMP, `ov_${fileId}.pdf`);
    await downloadDriveFile(fileId, src);

    const bytes = fs.readFileSync(src);
    const doc = await PDFDocument.load(bytes, { updateFieldAppearances:false });
    try { doc.registerFontkit(fontkit); } catch(_) {}

    let fontBytes = null;
    const cands = [
      process.env.FONT_TTF_PATH,
      path.join(ROOT, 'fonts/KozMinPr6N-Regular.otf'),
      path.join(ROOT, 'fonts/NotoSerifCJKjp-Regular.otf'),
      path.join(ROOT, 'fonts/NotoSansJP-Regular.ttf'),
      path.join(ROOT, 'fonts/NotoSansJP-Regular.otf'),
    ].filter(Boolean);
    for (const p of cands) { if (p && fs.existsSync(p)) { fontBytes = fs.readFileSync(p); break; } }
    const font = fontBytes ? await doc.embedFont(fontBytes, { subset:false }) : undefined;

    const page = doc.getPages()[0];
    page.drawText('VISIBLE OVERLAY TEST あいうえお 山田太郎', {
      x:48, y:720, size:14, font, color:rgb(0,0,0)
    });

    const outBytes = await doc.save({ useObjectStreams:false });
    const out = path.join(TMP, `OVERLAY_${Date.now()}.pdf`);
    fs.writeFileSync(out, outBytes);
    const up = await uploadToDrive(out, `OVERLAY_${Date.now()}.pdf`);
    res.json({ ok:true, link: up.webViewLink });
  } catch(e){ res.status(500).json({ ok:false, error:e.message }); }
});

app.listen(PORT, ()=>{
  log(`Server listening on ${PORT}`);
  log(`OUTPUT_FOLDER_ID (fallback): ${OUTPUT_FOLDER_ID || '(none set)'}`);
  log(`Creds: ${
    process.env.GOOGLE_CREDENTIALS_JSON ? 'GOOGLE_CREDENTIALS_JSON'
      : (process.env.GOOGLE_APPLICATION_CREDENTIALS ? 'GOOGLE_APPLICATION_CREDENTIALS' : 'ADC/unknown')
  }`);
  log(`FONT_TTF_PATH: ${process.env.FONT_TTF_PATH || '(none)'}`);
});
