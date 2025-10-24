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

// === new flags ===
const RESPECT_TEMPLATE_APPEARANCE =
  String(process.env.RESPECT_TEMPLATE_APPEARANCE || 'true').toLowerCase() === 'true';
const FORCE_BURN_IN =
  String(process.env.FORCE_BURN_IN || 'false').toLowerCase() === 'true';

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
  // remove zero-width, quotes (JP/ASCII), collapse spaces
  t = t
    .replace(/[\u200B-\u200D\uFEFF]/g, '')
    .replace(/[“”„‟〝〞＂"]/g, '')
    .replace(/[‘’‚‛′＇']/g, '')
    .replace(/\s+/g, ' ')
    .trim();
  return t;
}
function normalizeYesNo(vRaw) {
  const v = stripWeird(vRaw).toLowerCase();
  if (!v) return '';
  if (['yes','true','1','on'].includes(v) || v === 'はい') return 'yes';
  if (['no','false','0','off'].includes(v) || v === 'いいえ') return 'no';
  if (v.startsWith('は')) return 'yes';
  if (v.startsWith('い')) return 'no';
  return v;
}

function normalizeRegion(vRaw) {
  const v = stripWeird(vRaw);
  const t = v.toLowerCase();
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
  return map[t] || v;
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

  if (out['DestinationRegion'] == null && out['Q5_DestinationRegion'] != null) {
    out['DestinationRegion'] = out['Q5_DestinationRegion'];
  }

  return out;
}

/** Resolve a value for a PDF field name - simplified direct lookup. */
function resolveValue(name, fields, aliasView) {
  // 1. Direct field match (most common case)
  if (fields[name] != null && fields[name] !== '') return fields[name];
  
  // 2. Check alias view (for backward compatibility with existing aliases)
  if (aliasView && aliasView[name] != null && aliasView[name] !== '') return aliasView[name];
  
  // 3. No match found
  return '';
}

/* ---------------- PDF fill core ---------------- */

async function fillPdf(srcPath, outPath, fields = {}, opts = {}) {
  const bytes = fs.readFileSync(srcPath);
  const pdfDoc = await PDFDocument.load(bytes, { updateFieldAppearances: false });

  // 1) font
  try { pdfDoc.registerFontkit(fontkit); } catch (_) {}
  let customFont = null;
  let chosenFontPath = null;
  const fontCandidates = [
    process.env.FONT_TTF_PATH,
    path.join(ROOT, 'fonts/NotoSansJP-Regular.ttf'),
    path.join(ROOT, 'fonts/NotoSansJP-Regular.otf'),
    path.join(ROOT, 'fonts/NotoSerifCJKjp-Regular.otf'),
    path.join(ROOT, 'fonts/KozMinPr6N-Regular.otf'),
  ].filter(Boolean);
  for (const p of fontCandidates) {
    try {
      if (p && fs.existsSync(p)) {
        const fontBytes = fs.readFileSync(p);
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
  if (!customFont) throw new Error('CJK font not embedded. Set FONT_TTF_PATH to a .ttf shipped with the repo.');

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

  if (!RESPECT_TEMPLATE_APPEARANCE) {
    acroForm.set(PDFName.of('DA'), PDFString.of('/F0 12 Tf 0 g'));
    acroForm.set(PDFName.of('NeedAppearances'), PDFBool.True);
  }

  // 3) fill
  const form = pdfDoc.getForm();
  const allFields = form.getFields();
  const valueBy = buildAliasView(fields);
  let filled = 0;

  for (const f of allFields) {
    const name = f.getName ? f.getName() : '';
    const ctor = f.constructor && f.constructor.name || '';
    const valRaw = resolveValue(name, fields, valueBy);

    if (ctor.includes('Text')) {
      if (valRaw != null && valRaw !== '') {
        f.setText(String(valRaw));
        if (!RESPECT_TEMPLATE_APPEARANCE) {
          try { f.updateAppearances(customFont); } catch (_) {}
        }
        filled++;
      }
      continue;
    }
    if (ctor.includes('Dropdown')) {
      if (valRaw != null && valRaw !== '') {
        try { f.select(String(valRaw)); filled++; } catch (_) {}
      }
      continue;
    }
    if (ctor.includes('Radio')) {
      if (valRaw != null && valRaw !== '') {
        try { f.select(String(valRaw)); filled++; } catch (_) {}
      }
      continue;
    }

    if (ctor.includes('Check')) {
      const n = String(name);
      const m = n.match(/^(.*)_(yes|no)$/i);
      if (m) {
        const base = m[1];
        const isYesBox = m[2].toLowerCase() === 'yes';
        const baseVal = valueBy[base] != null ? valueBy[base] : fields[base];
        if (baseVal != null) {
          const yn = normalizeYesNo(baseVal);
          if ((isYesBox && yn === 'yes') || (!isYesBox && yn === 'no')) f.check();
          else f.uncheck();
          filled++;
        } else f.uncheck();
        continue;
      }
      const rm = n.match(/^DestinationRegion_(.+)$/);
      if (rm) {
        const want = stripWeird(rm[1]);
        const given = normalizeRegion(valueBy['DestinationRegion'] || fields['DestinationRegion'] || '');
        if (want && given && want === given) f.check(); else f.uncheck();
        filled++;
        continue;
      }
      const single = resolveValue(n, fields, valueBy);
      if (single !== '') {
        // Check if this is a numeric value (like 19, 20) that should be checked
        const numericValue = String(single).trim();
        if (/^\d+$/.test(numericValue)) {
          // For numeric values, check based on field name pattern
          try {
            log(`Checkbox ${n}: value="${numericValue}"`);
            
            // Check if this is an era field with valid values
            const isEraField = n.includes('Era19or20');
            const shouldCheck = isEraField && (numericValue === '19' || numericValue === '20');
            
            if (shouldCheck) {
              f.check();
              if (!RESPECT_TEMPLATE_APPEARANCE) {
                try { f.updateAppearances(customFont); } catch (_) {}
              }
              filled++;
              log(`✅ Checked checkbox ${n} (era field with value ${numericValue})`);
            } else {
              f.uncheck();
              if (!RESPECT_TEMPLATE_APPEARANCE) {
                try { f.updateAppearances(customFont); } catch (_) {}
              }
              log(`❌ Unchecked checkbox ${n} (not era field or wrong value)`);
            }
          } catch (e) {
            f.uncheck();
            log(`❌ Error with checkbox ${n}:`, e.message);
          }
        } else {
          // Check for PhoneType and TravelerSex checkboxes (Japanese text values)
          const isPhoneTypeField = n.includes('PhoneType') || n.includes('電話');
          const isTravelerSexField = n.includes('TravelerSex') || n.includes('Sex') || n.includes('性別');
          const phoneTypeValue = String(single).trim();
          const travelerSexValue = String(single).trim();
          
          if (isPhoneTypeField) {
            log(`PhoneType checkbox ${n}: value="${phoneTypeValue}"`);
            
            // Check if this field should be checked based on the value
            const shouldCheck = (
              // Handle explicit field names (Option 1 approach)
              (n === 'PhoneType_自宅' && phoneTypeValue === 'yes') ||
              (n === 'PhoneType_勤務先' && phoneTypeValue === 'yes') ||
              (n === 'PhoneType_携帯' && phoneTypeValue === 'yes') ||
              // Handle generic 'PhoneType' field name matching its value (fallback)
              (n === 'PhoneType' && phoneTypeValue === '自宅') ||
              (n === 'PhoneType' && phoneTypeValue === '勤務先') ||
              (n === 'PhoneType' && phoneTypeValue === '携帯') ||
              // Handle specific field names containing the type
              (n.includes('自宅') && phoneTypeValue === '自宅') ||
              (n.includes('勤務先') && phoneTypeValue === '勤務先') ||
              (n.includes('携帯') && phoneTypeValue === '携帯') ||
              (n.includes('Home') && phoneTypeValue === '自宅') ||
              (n.includes('Work') && phoneTypeValue === '勤務先') ||
              (n.includes('Mobile') && phoneTypeValue === '携帯')
            );
            
            if (shouldCheck) {
              // Simple checkbox logic for explicit field names
              f.check();
              if (!RESPECT_TEMPLATE_APPEARANCE) {
                try { f.updateAppearances(customFont); } catch (_) {}
              }
              filled++;
              log(`✅ Checked PhoneType ${n} (value: ${phoneTypeValue})`);
            } else {
              f.uncheck();
              if (!RESPECT_TEMPLATE_APPEARANCE) {
                try { f.updateAppearances(customFont); } catch (_) {}
              }
              log(`❌ Unchecked PhoneType checkbox ${n} (no match for value: ${phoneTypeValue})`);
            }
          } else if (isTravelerSexField) {
            log(`TravelerSex checkbox ${n}: value="${travelerSexValue}"`);
            
            // Check if this field should be checked based on the value
            const shouldCheck = (
              // Handle explicit field names (if you rename them)
              (n === 'TravelerSex_男性' && travelerSexValue === 'yes') ||
              (n === 'TravelerSex_女性' && travelerSexValue === 'yes') ||
              // Handle generic 'TravelerSex' field name matching its value
              (n === 'TravelerSex' && travelerSexValue === '男性') ||
              (n === 'TravelerSex' && travelerSexValue === '女性') ||
              (n === 'TravelerSex' && travelerSexValue === 'Male') ||
              (n === 'TravelerSex' && travelerSexValue === 'Female') ||
              // Handle specific field names containing the type
              (n.includes('男性') && travelerSexValue === '男性') ||
              (n.includes('女性') && travelerSexValue === '女性') ||
              (n.includes('Male') && travelerSexValue === '男性') ||
              (n.includes('Female') && travelerSexValue === '女性')
            );
            
            if (shouldCheck) {
              // Simple checkbox logic for explicit field names
              f.check();
              if (!RESPECT_TEMPLATE_APPEARANCE) {
                try { f.updateAppearances(customFont); } catch (_) {}
              }
              filled++;
              log(`✅ Checked TravelerSex ${n} (value: ${travelerSexValue})`);
            } else {
              f.uncheck();
              if (!RESPECT_TEMPLATE_APPEARANCE) {
                try { f.updateAppearances(customFont); } catch (_) {}
              }
              log(`❌ Unchecked TravelerSex checkbox ${n} (no match for value: ${travelerSexValue})`);
            }
          } else {
            // For other non-numeric values, use the existing yes/no logic
            const yn = normalizeYesNo(single);
            if (yn === 'yes' || yn === 'on' || yn === '1' || yn === 'true') f.check();
            else f.uncheck();
            filled++;
          }
        }
        continue;
      }
    }
  }

  // 4) Burn-in (optional)
  if (FORCE_BURN_IN && !RESPECT_TEMPLATE_APPEARANCE) {
    try {
      let burned = 0;
      for (const f of allFields) {
        const name = f.getName ? f.getName() : '';
        const ctor = f.constructor && f.constructor.name || '';
        let raw = resolveValue(name, fields, valueBy);
        if (!raw && ctor.includes('Dropdown')) {
          try { raw = f.getSelected() || ''; } catch (_) {}
        }
        if ((!ctor.includes('Text') && !ctor.includes('Dropdown')) || !raw) continue;

        const widgets = (f.acroField && f.acroField.getWidgets) ? f.acroField.getWidgets() : [];
        for (const w of widgets) {
          const page = w.getPage && w.getPage();
          if (!page) continue;
          const rect = w.getRectangle && w.getRectangle();
          if (!rect) continue;
          const text = String(raw);
          const padding = 2;
          let size = Math.min(12, rect.height - 2 * padding);
          const maxWidth = rect.width - 2 * padding;
          while (size > 5.5 && customFont.widthOfTextAtSize(text, size) > maxWidth) size -= 0.5;

          page.drawText(text, {
            x: rect.x + padding,
            y: rect.y + (rect.height - size) / 2,
            size,
            font: customFont,
            color: rgb(0, 0, 0)
          });
          burned++;
        }
      }
      log('burn-in count:', burned);
    } catch (e) {
      log('Burn-in fallback failed:', e && e.message);
    }
  }

  // 5) watermark (optional)
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

app.get('/', (_req, res) => res.type('text/plain').send('PDF filler is up. Try GET /health'));

app.get('/health', (_req, res) => {
  res.json({
    ok: true,
    tmpDir: TMP,
    hasOUTPUT_FOLDER_ID: !!OUTPUT_FOLDER_ID,
    credMode: process.env.GOOGLE_CREDENTIALS_JSON
      ? 'env-json'
      : (process.env.GOOGLE_APPLICATION_CREDENTIALS ? 'file-path' : 'adc/unknown'),
    fontEnv: process.env.FONT_TTF_PATH || '(none)',
    RESPECT_TEMPLATE_APPEARANCE,
    FORCE_BURN_IN,
  });
});

app.post('/fill', async (req, res) => {
  try {
    const { templateFileId, fields, outputName, folderId, mode, watermarkText } = req.body || {};
    if (!templateFileId) return res.status(400).json({ error: 'templateFileId is required' });

    const tmpTemplate = path.join(TMP, `template_${templateFileId}.pdf`);
    await downloadDriveFile(templateFileId, tmpTemplate);

    log('===== FULL FIELD MAP RECEIVED FROM GAS =====');
    if (fields && typeof fields === 'object') {
      Object.entries(fields).forEach(([k, v]) => log(`${k}: ${JSON.stringify(v)}`));
      log('===== END FIELD MAP =====');
    } else log('⚠️ No fields object received:', typeof fields);

    const base = (outputName && String(outputName).trim()) || `filled_${Date.now()}`;
    const outName = base.toLowerCase().endsWith('.pdf') ? base : `${base}.pdf`;
    const outPath = path.join(TMP, outName);

    const aliasView = buildAliasView(fields || {});
    const probe = {};
    for (const k of [
      'TreatmentNow','SeriousHistory','LuggageClaims5Plus','DuplicateContracts',
      'SanctionedCountries','WorkDuringTravel','DestinationRegion'
    ]) {
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

app.listen(PORT, () => {
  log(`Server listening on ${PORT}`);
  log(`OUTPUT_FOLDER_ID: ${OUTPUT_FOLDER_ID || '(none)'}`);
  log(`FONT_TTF_PATH: ${process.env.FONT_TTF_PATH || '(none)'}`);
  log(`RESPECT_TEMPLATE_APPEARANCE: ${RESPECT_TEMPLATE_APPEARANCE}`);
  log(`FORCE_BURN_IN: ${FORCE_BURN_IN}`);
  log(
    `Creds: ${
      process.env.GOOGLE_CREDENTIALS_JSON
        ? 'GOOGLE_CREDENTIALS_JSON'
        : process.env.GOOGLE_APPLICATION_CREDENTIALS
        ? 'GOOGLE_APPLICATION_CREDENTIALS'
        : 'ADC/unknown'
    }`
  );
});
