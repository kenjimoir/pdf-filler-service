// index.cjs — PDF fill → upload to Drive (Shared Drives OK)
// deps: express, cors, pdf-lib, googleapis, @pdf-lib/fontkit

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

function log(...a){ console.log(new Date().toISOString(), '-', ...a); }
function ensureDir(p){ if(!fs.existsSync(p)) fs.mkdirSync(p,{recursive:true}); }

// ---------- Google Drive ----------
function getDriveClient(){
  let credentials=null;
  if(process.env.GOOGLE_CREDENTIALS_JSON){
    try{ credentials=JSON.parse(process.env.GOOGLE_CREDENTIALS_JSON); }catch(e){ log('ERROR parsing GOOGLE_CREDENTIALS_JSON:', e?.message); }
  }
  const auth=new google.auth.GoogleAuth({ scopes:['https://www.googleapis.com/auth/drive'], ...(credentials?{credentials}:{}) });
  return google.drive({version:'v3', auth});
}
async function downloadDriveFile(fileId, dest){
  const drive=getDriveClient();
  log('Downloading template', fileId, '→', dest);
  const res=await drive.files.get({fileId, alt:'media', supportsAllDrives:true}, {responseType:'stream'});
  await new Promise((resolve,reject)=>{ const out=fs.createWriteStream(dest); res.data.on('error',reject).pipe(out).on('finish',resolve); });
  return dest;
}
async function uploadToDrive(local, name, parentId){
  const drive=getDriveClient();
  const parents = parentId ? [parentId] : (OUTPUT_FOLDER_ID ? [OUTPUT_FOLDER_ID] : undefined);
  const res=await drive.files.create({
    requestBody:{name, parents},
    media:{mimeType:'application/pdf', body:fs.createReadStream(local)},
    fields:'id,name,webViewLink,parents',
    supportsAllDrives:true,
  });
  return res.data;
}

// ---------- normalization helpers ----------
const norm = s => String(s||'').trim().toLowerCase().replace(/[ \t\u3000]/g,'');
function mapYesNo(v){
  const n=norm(v);
  if(['はい','yes','true','1','on','可','有','あり','する'].some(k=>n.includes(k))) return 'はい';
  if(['いいえ','no','false','0','off','不可','無','なし','しない'].some(k=>n.includes(k))) return 'いいえ';
  return v;
}
function mapRegion(v){
  const n=norm(v);
  if(['asia','ａｓｉａ','アジア'].some(k=>n.includes(k))) return 'アジア';
  if(['europe','ヨーロッパ','欧州'].some(k=>n.includes(k))) return 'ヨーロッパ';
  if(['oceania','オセアニア'].some(k=>n.includes(k))) return 'オセアニア';
  if(['northamerica','北米'].some(k=>n.includes(k))) return '北米';
  if(['latin','south','中南米','南米'].some(k=>n.includes(k))) return '中南米';
  if(['africa','アフリカ'].some(k=>n.includes(k))) return 'アフリカ';
  if(['middleeast','中東'].some(k=>n.includes(k))) return '中東';
  if(['other','その他','そのた'].some(k=>n.includes(k))) return 'その他';
  return v;
}

function setCheckBoxByAnswer(cb, answer){
  const acro = cb.acroField || cb._acroField || cb['acroField'];
  let onName=null;
  try{ const n=acro?.getOnValue?.(); onName = n ? String(n) : null; }catch{}
  const yn=mapYesNo(answer);
  if(onName){
    if(/^はい$/i.test(onName)) { yn==='はい' ? cb.check() : cb.uncheck(); return; }
    if(/^いいえ$/i.test(onName)) { yn==='いいえ' ? cb.check() : cb.uncheck(); return; }
  }
  if(['はい','yes','true','1','on'].includes(norm(answer))) cb.check(); else cb.uncheck();
}
function selectRadioSafe(radio, wanted){
  const trySelect = v=>{ try{ radio.select(v); return true; }catch{ return false; } };
  let opts=[];
  try{ opts = radio.getOptions ? radio.getOptions().map(String) : []; }catch{}
  const yn=mapYesNo(wanted), reg=mapRegion(wanted);

  if(!opts.length) return trySelect(wanted)||trySelect(yn)||trySelect(reg);

  if(opts.includes(wanted)) return trySelect(wanted);
  if(opts.includes(yn))     return trySelect(yn);
  if(opts.includes(reg))    return trySelect(reg);

  const yes = opts.find(o=>/^(はい|yes)$/i.test(o));
  const no  = opts.find(o=>/^(いいえ|no)$/i.test(o));
  if(yn==='はい' && yes) return trySelect(yes);
  if(yn==='いいえ' && no) return trySelect(no);

  const hit=opts.find(o=>norm(o)===norm(yn)||norm(o)===norm(reg));
  return hit ? trySelect(hit) : false;
}

/** 渡航先が Radio ではなく CheckBox 群のときにまとめて判定 */
function applyRegionCheckboxPass(form, regionValue){
  if(!regionValue) return 0;
  const wanted = mapRegion(regionValue);
  let hits = 0;
  try{
    for(const f of form.getFields()){
      const name = f.getName?.() || '';
      const ctor = f.constructor?.name || '';
      if(!ctor.includes('Check')) continue;
      const acro = f.acroField || f._acroField || f['acroField'];
      let onName=null;
      try{ const n = acro?.getOnValue?.(); onName = n ? String(n) : null; }catch{}
      if(onName && norm(onName) === norm(wanted)){
        try{ f.check(); hits++; }catch{}
      }
    }
  }catch{}
  return hits;
}

// ---------- PDF fill ----------
async function fillPdf(srcPath, outPath, fields = {}, opts = {}) {
  const bytes = fs.readFileSync(srcPath);
  const pdfDoc = await PDFDocument.load(bytes, { updateFieldAppearances: true });
  try{ pdfDoc.registerFontkit(fontkit); }catch{}

  // Embed JP font
  let customFont=null, chosenFontPath=null;
  const fontCandidates = [
    process.env.FONT_TTF_PATH,
    path.join(ROOT,'fonts/KozMinPr6N-Regular.otf'),
    path.join(ROOT,'fonts/NotoSerifCJKjp-Regular.otf'),
    path.join(ROOT,'fonts/NotoSansJP-Regular.ttf'),
    path.join(ROOT,'fonts/NotoSansJP-Regular.otf'),
  ].filter(Boolean);
  for(const p of fontCandidates){
    try{
      if(p && fs.existsSync(p)){
        const fontBytes=fs.readFileSync(p);
        const allowSubset = path.extname(p).toLowerCase() !== '.otf';
        customFont = await pdfDoc.embedFont(fontBytes,{subset:allowSubset});
        chosenFontPath=p; log('Embedded JP font:', p, '(subset:', allowSubset, ')'); break;
      }
    }catch(e){ log('Font embed failed for', p, e?.message); }
  }
  if(!customFont) throw new Error('CJK font not embedded (FONT_TTF_PATH).');

  // AcroForm DR/DA
  let acroFormRef = pdfDoc.catalog.get(PDFName.of('AcroForm'));
  let acroForm = acroFormRef ? pdfDoc.context.lookup(acroFormRef, PDFDict) : null;
  if(!acroForm){ acroForm = pdfDoc.context.obj({}); pdfDoc.catalog.set(PDFName.of('AcroForm'), acroForm); }
  const dr = acroForm.get(PDFName.of('DR')) || pdfDoc.context.obj({});
  const drFont = dr.get(PDFName.of('Font')) || pdfDoc.context.obj({});
  drFont.set(PDFName.of('F0'), customFont.ref);
  dr.set(PDFName.of('Font'), drFont);
  acroForm.set(PDFName.of('DR'), dr);
  acroForm.set(PDFName.of('DA'), PDFString.of('/F0 10 Tf 0 g'));
  acroForm.set(PDFName.of('NeedAppearances'), PDFBool.True);

  const form = pdfDoc.getForm();

  // --- 1st pass: direct fill by name ---
  let filled=0;
  for(const [name, rawVal] of Object.entries(fields||{})){
    try{
      const fld = form.getField(String(name));
      const type = fld?.constructor?.name || '';
      const val  = rawVal==null ? '' : String(rawVal);

      if(type.includes('Text')){
        fld.setText(val);
        try{ fld.updateAppearances(customFont); }catch{}
        filled++;
      } else if(type.includes('Check')){
        setCheckBoxByAnswer(fld, val);
        filled++;
      } else if(type.includes('Radio')){
        if(selectRadioSafe(fld, val)) filled++;
      } else if(type.includes('Dropdown')){
        try{ fld.select(val); filled++; }catch{}
      }
    }catch{}
  }

  // --- 2nd pass (region as checkboxes) ---
  const regionKey = ['TravelRegion','Region','Destination'].find(k=>fields && fields[k]!=null);
  if(regionKey){
    const add = applyRegionCheckboxPass(form, fields[regionKey]);
    if(add>0) log(`Region checkbox pass matched: ${regionKey}=${fields[regionKey]} → ${add} check(s)`);
  }

  // Appearance rebuild
  try{ form.updateFieldAppearances(customFont); }catch{}

  // Optional watermark
  const wmText = (opts.watermarkText||'').trim();
  if(wmText){
    for(const page of pdfDoc.getPages()){
      const {width,height} = page.getSize();
      page.drawText(wmText,{
        x: width/2-200, y: height/2-40, size: 80, opacity: .12, rotate: degrees(45), color: rgb(.85,.1,.1)
      });
    }
  }

  // --- flatten for final (preview安定化) ---
  if(String(opts.mode||'').toLowerCase() !== 'review'){
    try{ form.flatten(); }catch{}
  }

  // Save
  let outBytes;
  try{
    outBytes = await pdfDoc.save({ useObjectStreams:false, addDefaultPage:false, updateFieldAppearances:false });
  }catch(e){
    log('pdfDoc.save() failed, retrying:', e.message);
    outBytes = await pdfDoc.save();
  }
  fs.writeFileSync(outPath, outBytes);
  return { outPath, filled, size: outBytes.length, fontPath: chosenFontPath };
}

// ---------- HTTP ----------
const app = express();
app.use(cors());
app.use(express.json({limit:'10mb'}));

app.get('/', (_req,res)=>{ res.type('text/plain').send('PDF filler is up. Try GET /health'); });
app.get('/health', (_req,res)=>{
  res.json({
    ok:true, tmpDir:TMP,
    hasOUTPUT_FOLDER_ID: !!OUTPUT_FOLDER_ID,
    credMode: process.env.GOOGLE_CREDENTIALS_JSON ? 'env-json' :
      (process.env.GOOGLE_APPLICATION_CREDENTIALS ? 'file-path' : 'adc/unknown'),
    fontEnv: process.env.FONT_TTF_PATH || '(none)',
  });
});
app.get('/fields', async (req,res)=>{
  try{
    const fileId=(req.query.fileId||'').trim();
    if(!fileId) return res.status(400).json({error:'fileId is required'});
    const local=path.join(TMP,`template_${fileId}.pdf`);
    await downloadDriveFile(fileId, local);
    const bytes=fs.readFileSync(local);
    const pdfDoc=await PDFDocument.load(bytes);
    let names=[];
    try{ const form=pdfDoc.getForm(); names=(form?form.getFields():[]).map(f=>f.getName()); }catch{}
    res.json({count:names.length, names});
  }catch(e){
    log('List fields failed:', e?.stack||e);
    res.status(500).json({error:'List fields failed', detail:e.message});
  }
});

app.post('/fill', async (req,res)=>{
  try{
    const { templateFileId, fields, outputName, folderId, mode, watermarkText } = req.body||{};
    if(!templateFileId) return res.status(400).json({error:'templateFileId is required'});

    const tmp=path.join(TMP,`template_${templateFileId}.pdf`);
    await downloadDriveFile(templateFileId, tmp);

    const base = (outputName && String(outputName).trim()) || `filled_${Date.now()}`;
    const outName = base.toLowerCase().endsWith('.pdf') ? base : `${base}.pdf`;
    const outPath = path.join(TMP, outName);

    const result = await fillPdf(tmp, outPath, fields||{}, { watermarkText, mode });
    log(`Filled PDF -> ${result.outPath} (${result.size} bytes, fields filled: ${result.filled})`);

    const uploaded = await uploadToDrive(result.outPath, outName, folderId);
    log('Uploaded to Drive:', uploaded);
    res.json({ ok:true, filledCount: result.filled, driveFile: uploaded, webViewLink: uploaded.webViewLink });
  }catch(err){
    log('ERROR /fill:', err?.stack||err);
    res.status(500).json({ error:'Fill failed', detail: err.message });
  }
});

// Debug endpoints (unchanged)
app.get('/debug/passthrough', async (req,res)=>{ try{
  const fileId=(req.query.fileId||'').trim(); if(!fileId) return res.status(400).json({ok:false,error:'fileId required'});
  const local=path.join(TMP,`pt_${fileId}.pdf`); await downloadDriveFile(fileId, local);
  const up=await uploadToDrive(local,`PASSTHROUGH_${Date.now()}.pdf`); res.json({ok:true,link:up.webViewLink});
}catch(e){ res.status(500).json({ok:false,error:e.message}); }});
app.get('/debug/roundtrip', async (req,res)=>{ try{
  const fileId=(req.query.fileId||'').trim(); if(!fileId) return res.status(400).json({ok:false,error:'fileId required'});
  const src=path.join(TMP,`rt_${fileId}.pdf`); await downloadDriveFile(fileId, src);
  const bytes=fs.readFileSync(src); const doc=await PDFDocument.load(bytes,{updateFieldAppearances:false});
  const outBytes=await doc.save({useObjectStreams:false}); const out=path.join(TMP,`ROUNDTRIP_${Date.now()}.pdf`);
  fs.writeFileSync(out,outBytes); const up=await uploadToDrive(out,`ROUNDTRIP_${Date.now()}.pdf`);
  res.json({ok:true,link:up.webViewLink});
}catch(e){ res.status(500).json({ok:false,error:e.message}); }});
app.get('/debug/overlay', async (req,res)=>{ try{
  const fileId=(req.query.fileId||'').trim(); if(!fileId) return res.status(400).json({ok:false,error:'fileId required'});
  const src=path.join(TMP,`ov_${fileId}.pdf`); await downloadDriveFile(fileId, src);
  const bytes=fs.readFileSync(src); const doc=await PDFDocument.load(bytes,{updateFieldAppearances:false});
  try{ doc.registerFontkit(fontkit); }catch{}
  let fontBytes=null; const pref=process.env.FONT_TTF_PATH;
  const cands=[pref, path.join(ROOT,'fonts/KozMinPr6N-Regular.otf'), path.join(ROOT,'fonts/NotoSerifCJKjp-Regular.otf'),
               path.join(ROOT,'fonts/NotoSansJP-Regular.ttf'), path.join(ROOT,'fonts/NotoSansJP-Regular.otf')].filter(Boolean);
  for(const p of cands){ if(p&&fs.existsSync(p)){ fontBytes=fs.readFileSync(p); break; } }
  const font=fontBytes?await doc.embedFont(fontBytes,{subset:false}):undefined;
  const page=doc.getPages()[0];
  page.drawText('VISIBLE OVERLAY TEST あいうえお 山田太郎',{x:48,y:720,size:14,font,color:rgb(0,0,0)});
  const outBytes=await doc.save({useObjectStreams:false}); const out=path.join(TMP,`OVERLAY_${Date.now()}.pdf`);
  fs.writeFileSync(out,outBytes); const up=await uploadToDrive(out,`OVERLAY_${Date.now()}.pdf`);
  res.json({ok:true,link:up.webViewLink});
}catch(e){ res.status(500).json({ok:false,error:e.message}); }});

app.listen(PORT, ()=>{
  log(`Server listening on ${PORT}`);
  log(`OUTPUT_FOLDER_ID (fallback): ${OUTPUT_FOLDER_ID || '(none set)'}`);
  log(`Creds: ${process.env.GOOGLE_CREDENTIALS_JSON ? 'GOOGLE_CREDENTIALS_JSON' :
    (process.env.GOOGLE_APPLICATION_CREDENTIALS ? 'GOOGLE_APPLICATION_CREDENTIALS' : 'ADC/unknown')}`);
  log(`FONT_TTF_PATH: ${process.env.FONT_TTF_PATH || '(none)'}`);
});
