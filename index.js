// PDF Filler using PDFtk - Reliable Japanese text support
// This service uses PDFtk (PDF Toolkit) instead of pdf-lib for better Unicode support

const fs = require('fs');
const os = require('os');
const path = require('path');
const express = require('express');
const cors = require('cors');
const { exec } = require('child_process');
const { promisify } = require('util');
const { google } = require('googleapis');

const execAsync = promisify(exec);

const PORT = process.env.PORT || 8080;
const TMP = path.join(os.tmpdir(), 'pdf-filler-pdftk');
const OUTPUT_FOLDER_ID = process.env.OUTPUT_FOLDER_ID;

// Ensure temp directory exists
if (!fs.existsSync(TMP)) {
  fs.mkdirSync(TMP, { recursive: true });
}

// Initialize Google Drive client
function getDriveClient() {
  // Option 1: JSON string in environment variable
  if (process.env.GOOGLE_CREDENTIALS_JSON) {
    try {
      const creds = JSON.parse(process.env.GOOGLE_CREDENTIALS_JSON);
      const auth = new google.auth.GoogleAuth({
        credentials: creds,
        scopes: ['https://www.googleapis.com/auth/drive'],
      });
      return google.drive({ version: 'v3', auth });
    } catch (parseError) {
      throw new Error(`Invalid GOOGLE_CREDENTIALS_JSON: ${parseError.message}`);
    }
  }
  
  // Option 2: File path in environment variable (Render file-based secrets)
  if (process.env.GOOGLE_APPLICATION_CREDENTIALS) {
    const auth = new google.auth.GoogleAuth({
      keyFile: process.env.GOOGLE_APPLICATION_CREDENTIALS,
      scopes: ['https://www.googleapis.com/auth/drive'],
    });
    return google.drive({ version: 'v3', auth });
  }
  
  throw new Error('Either GOOGLE_CREDENTIALS_JSON or GOOGLE_APPLICATION_CREDENTIALS must be set');
}

// Generate FDF (Form Data Format) content from fields
// Text is encoded as UTF-16BE hex so Japanese renders correctly
// Checkbox/radio values use name objects (e.g. /Yes, /On, /1) and also set /AS
// onNameMap: optional map { fieldName: onExportName }
function generateFDF(fields, onNameMap) {
  const toUtf16Hex = (s) => {
    const str = String(s);
    // Build UTF-16BE buffer with BOM FE FF
    const be = Buffer.alloc(2 + str.length * 2);
    be[0] = 0xFE; be[1] = 0xFF;
    for (let i = 0; i < str.length; i++) {
      const code = str.charCodeAt(i);
      be[2 + i * 2] = (code >> 8) & 0xFF;
      be[3 + i * 2] = code & 0xFF;
    }
    return '<' + be.toString('hex').toUpperCase() + '>';
  };

  let fdf = '%FDF-1.2\n';
  fdf += '1 0 obj\n';
  fdf += '<<\n';
  fdf += '/FDF\n';
  fdf += '<<\n';
  fdf += '/Fields [\n';

  for (const [rawName, rawVal] of Object.entries(fields)) {
    if (rawVal == null || rawVal === '') continue;
    const fieldName = String(rawName).replace(/[()\\]/g, '\\$&');
    const val = String(rawVal);
    const isOn = /^(on|yes|true|1)$/i.test(val);
    const isOff = /^(off|no|false)$/i.test(val);

    fdf += '<<\n';
    fdf += `/T (${fieldName})\n`;
    if (isOn || isOff) {
      const detectedOn = onNameMap && onNameMap[fieldName] ? onNameMap[fieldName] : 'Yes';
      const name = isOn ? detectedOn : 'Off';
      fdf += `/V /${name}\n`;
      fdf += `/AS /${name}\n`;
    } else {
      fdf += `/V ${toUtf16Hex(val)}\n`;
    }
    fdf += '>>\n';
  }

  fdf += ']\n';
  fdf += '>>\n';
  fdf += '>>\n';
  fdf += 'endobj\n';
  fdf += 'trailer\n';
  fdf += '<<\n';
  fdf += '/Root 1 0 R\n';
  fdf += '>>\n';
  fdf += '%%EOF\n';

  return fdf;
}

// Detect checkbox/radio on-state names from template via pdftk dump
async function buildOnNameMap(templatePath) {
  const map = {};
  try {
    const cmd = `pdftk "${templatePath}" dump_data_fields_utf8`;
    const { stdout } = await execAsync(cmd);
    const lines = stdout.split(/\r?\n/);
    let current = {};
    const flush = () => {
      if (current.FieldType === 'Button' && current.FieldName) {
        const options = current.FieldStateOption || [];
        const onOpt = options.find((o) => o && o.toLowerCase() !== 'off');
        if (onOpt) {
          map[current.FieldName] = onOpt;
        }
      }
      current = {};
    };
    for (const line of lines) {
      if (line.trim() === '---') {
        flush();
        continue;
      }
      const idx = line.indexOf(':');
      if (idx > -1) {
        const key = line.slice(0, idx).trim();
        const value = line.slice(idx + 1).trim();
        if (key === 'FieldStateOption') {
          if (!current.FieldStateOption) current.FieldStateOption = [];
          current.FieldStateOption.push(value);
        } else {
          current[key] = value;
        }
      }
    }
    // flush last
    flush();
  } catch (e) {
    console.warn(`⚠️ Failed to dump field data for on-name detection: ${e.message}`);
  }
  return map;
}

// Fill PDF using PDFtk, then flatten with Ghostscript to avoid checkbox/font blob issues
async function fillPdfWithPDFtk(templatePath, outputPath, fields, opts) {
  const options = Object.assign({ flatten: true, flattenMethod: 'gs' }, opts);
  if (String(options.flattenMethod || '').toLowerCase() === 'none') {
    options.flatten = false;
  }
  console.log(`📝 Filling PDF with PDFtk...`);
  console.log(`   Template: ${templatePath}`);
  console.log(`   Output: ${outputPath}`);
  console.log(`   Fields: ${Object.keys(fields).length}`);
  
  // Generate FDF content
  const fdfContent = generateFDF(fields, options.onNameMap || null);
  const fdfPath = path.join(TMP, `data_${Date.now()}.fdf`);
  fs.writeFileSync(fdfPath, fdfContent, 'utf8');
  
  try {
    if (!options.flatten || options.flattenMethod === 'pdftk') {
      // Single-step pdftk (optionally flatten). If not flattening, also set NeedAppearances.
      if (!options.flatten) {
        const filledPath = path.join(TMP, `filled_${Date.now()}.pdf`);
        const pdftkFill = `pdftk "${templatePath}" fill_form "${fdfPath}" output "${filledPath}"`;
        console.log(`🔧 Running: ${pdftkFill}`);
        const { stderr: pdftkErr1 } = await execAsync(pdftkFill);
        if (pdftkErr1) console.warn(`⚠️ PDFtk stderr: ${pdftkErr1}`);
        // Step 1: write NeedAppearances flag
        const pdftkNeedApp = `pdftk "${filledPath}" output "${outputPath}" need_appearances`;
        console.log(`🔧 Running: ${pdftkNeedApp}`);
        const { stderr: pdftkErr2 } = await execAsync(pdftkNeedApp);
        if (pdftkErr2) console.warn(`⚠️ PDFtk stderr: ${pdftkErr2}`);
        // SKIP: No qpdf appearance stripping (this preserves checkboxes)
        try { if (fs.existsSync(filledPath)) fs.unlinkSync(filledPath); } catch (_) {}
      } else {
        const pdftkCmd = `pdftk "${templatePath}" fill_form "${fdfPath}" output "${outputPath}" flatten drop_xfa`;
        console.log(`🔧 Running: ${pdftkCmd}`);
        const { stderr: pdftkErr } = await execAsync(pdftkCmd);
        if (pdftkErr) console.warn(`⚠️ PDFtk stderr: ${pdftkErr}`);
      }
    } else {
      // Two-step: pdftk fill (no flatten) → ensure appearances → drop XFA → Ghostscript flatten
      const filledPath = path.join(TMP, `filled_${Date.now()}.pdf`);
      const filledAppearPath = path.join(TMP, `filled_appear_${Date.now()}.pdf`);
      const pdftkCmd = `pdftk "${templatePath}" fill_form "${fdfPath}" output "${filledPath}"`; 
      console.log(`🔧 Running: ${pdftkCmd}`);
      const { stderr: pdftkErr } = await execAsync(pdftkCmd);
      if (pdftkErr) console.warn(`⚠️ PDFtk stderr: ${pdftkErr}`);

      // Force NeedAppearances and drop XFA so widgets have visual appearances for GS to bake
      const pdftkAppear = `pdftk "${filledPath}" output "${filledAppearPath}" need_appearances drop_xfa`;
      console.log(`🔧 Running: ${pdftkAppear}`);
      const { stderr: pdftkErr2 } = await execAsync(pdftkAppear);
      if (pdftkErr2) console.warn(`⚠️ PDFtk stderr: ${pdftkErr2}`);

      const gsCmd = `gs -dBATCH -dNOPAUSE -dSAFER -sDEVICE=pdfwrite -dPDFSETTINGS=/prepress -dDetectDuplicateImages=true -dCompressFonts=true -sOutputFile="${outputPath}" "${filledAppearPath}"`;
      console.log(`🔧 Running: ${gsCmd}`);
      const { stderr: gsErr } = await execAsync(gsCmd);
      if (gsErr) console.warn(`⚠️ Ghostscript stderr: ${gsErr}`);
      try { if (fs.existsSync(filledPath)) fs.unlinkSync(filledPath); } catch (_) {}
      try { if (fs.existsSync(filledAppearPath)) fs.unlinkSync(filledAppearPath); } catch (_) {}
    }

    console.log(`✅ PDF filled successfully (${options.flatten ? 'flattened via ' + options.flattenMethod : 'not flattened'})`);
    console.log(`   Output size: ${fs.statSync(outputPath).size} bytes`);
    
    // Clean up FDF file
    fs.unlinkSync(fdfPath);
    // no intermediate file when flattening directly
    
    return { success: true, size: fs.statSync(outputPath).size };
    
  } catch (error) {
    console.error(`❌ PDFtk error: ${error.message}`);
    
    // Clean up FDF file on error
    if (fs.existsSync(fdfPath)) {
      fs.unlinkSync(fdfPath);
    }
    
    throw new Error(`PDFtk failed: ${error.message}`);
  }
}

// Upload to Google Drive
async function uploadToDrive(drive, filePath, fileName, folderId) {
  // If folderId is provided, verify it exists and is accessible
  const finalFolderId = folderId || OUTPUT_FOLDER_ID;
  if (finalFolderId) {
    try {
      await drive.files.get({
        fileId: finalFolderId,
        fields: 'id, name, mimeType',
        supportsAllDrives: true,
      });
      console.log(`✅ Verified folder exists: ${finalFolderId}`);
    } catch (error) {
      if (error.code === 404) {
        throw new Error(`Folder not found: ${finalFolderId}. Please check:\n1. Folder ID is correct\n2. Folder is shared with service account\n3. If in Shared Drive, service account has access`);
      } else if (error.code === 403) {
        throw new Error(`Access denied to folder: ${finalFolderId}. Please share the folder with your service account email.`);
      }
      throw error;
    }
  }
  
  const parents = finalFolderId ? [finalFolderId] : [];
  
  const fileMetadata = {
    name: fileName,
    parents: parents.length > 0 ? parents : undefined,
  };
  
  const media = {
    mimeType: 'application/pdf',
    body: fs.createReadStream(filePath),
  };
  
  const file = await drive.files.create({
    requestBody: fileMetadata,
    media: media,
    fields: 'id, name, webViewLink, webContentLink',
    supportsAllDrives: true,
  });
  
  return file.data;
}

// Check if PDFtk is available
async function checkPDFtk() {
  try {
    const { stdout } = await execAsync('pdftk --version');
    console.log(`✅ PDFtk found: ${stdout.trim()}`);
    return true;
  } catch (error) {
    console.error(`❌ PDFtk not found: ${error.message}`);
    console.error(`   Please install PDFtk on your system`);
    console.error(`   Ubuntu/Debian: sudo apt-get install pdftk`);
    console.error(`   macOS: brew install pdftk-java`);
    console.error(`   Or use: apt-get install pdftk-java`);
    return false;
  }
}

// HTTP Server
const app = express();
app.use(cors());
app.use(express.json({ limit: '10mb' }));

app.get('/', (_req, res) => {
  res.json({ service: 'PDF Filler PDFtk', status: 'running' });
});

app.get('/health', async (_req, res) => {
  const pdftkAvailable = await checkPDFtk();
  res.json({
    ok: true,
    pdftkAvailable,
    outputFolder: OUTPUT_FOLDER_ID || 'not set',
    tempDir: TMP,
  });
});

app.post('/fill', async (req, res) => {
  const { templateFileId, fields, outputName, folderId, mode, flattenMethod } = req.body;
  
  if (!templateFileId || !fields) {
    return res.status(400).json({ error: 'Missing templateFileId or fields' });
  }
  
  // Check if PDFtk is available
  const pdftkAvailable = await checkPDFtk();
  if (!pdftkAvailable) {
    return res.status(500).json({ 
      error: 'PDFtk not available', 
      detail: 'PDFtk is required but not installed on this system' 
    });
  }
  
  const drive = getDriveClient();
  const templatePath = path.join(TMP, `template_${templateFileId}.pdf`);
  const outputPath = path.join(TMP, `output_${Date.now()}.pdf`);
  
  try {
    // 0. Inspect template metadata
    const meta = await drive.files.get({
      fileId: templateFileId,
      fields: 'id, name, mimeType, size, owners(emailAddress)',
      supportsAllDrives: true,
    });
    console.log(`📄 Template meta: name=${meta.data.name} mime=${meta.data.mimeType} size=${meta.data.size}`);

    // 1. Download template (export if it's a Google Doc)
    console.log(`📥 Downloading template: ${templateFileId}`);
    if (meta.data.mimeType && meta.data.mimeType.startsWith('application/vnd.google-apps')) {
      // Not a binary PDF on Drive → export as PDF
      const exportRes = await drive.files.export(
        { fileId: templateFileId, mimeType: 'application/pdf' },
        { responseType: 'stream' }
      );
      const writeStream = fs.createWriteStream(templatePath);
      exportRes.data.pipe(writeStream);
      await new Promise((resolve, reject) => {
        writeStream.on('finish', resolve);
        writeStream.on('error', reject);
      });
    } else {
      const templateFile = await drive.files.get(
        { fileId: templateFileId, alt: 'media', supportsAllDrives: true },
        { responseType: 'stream' }
      );
      const writeStream = fs.createWriteStream(templatePath);
      templateFile.data.pipe(writeStream);
      await new Promise((resolve, reject) => {
        writeStream.on('finish', resolve);
        writeStream.on('error', reject);
      });
    }
    console.log('✅ Template downloaded');
    
    const modeStr = String(mode || 'final').toLowerCase();
    if (modeStr === 'copy') {
      // Just pass-through the template to output (sanity check)
      console.log('🧪 COPY mode: uploading template as-is');
      fs.copyFileSync(templatePath, outputPath);
    } else if (modeStr === 'pdftk-copy') {
      // Rewrite via pdftk without filling (checks pdftk write path)
      const cmd = `pdftk "${templatePath}" cat output "${outputPath}"`;
      console.log(`🔧 Running: ${cmd}`);
      await execAsync(cmd);
    } else {
      // 2. Fill PDF with PDFtk
      console.log(`📝 Filling PDF with ${Object.keys(fields).length} fields...`);
      const flatten = modeStr !== 'preview';
      const fm = String(flattenMethod || 'gs').toLowerCase();
      const onNameMap = await buildOnNameMap(templatePath);
      await fillPdfWithPDFtk(templatePath, outputPath, fields, { flatten, flattenMethod: fm, onNameMap });
    }
    
    // 3. Upload to Drive
    const finalName = outputName || `filled_${Date.now()}.pdf`;
    console.log(`📤 Uploading to Drive: ${finalName}`);
    const uploadedFile = await uploadToDrive(
      drive,
      outputPath,
      finalName,
      folderId || OUTPUT_FOLDER_ID
    );
    
    // Cleanup
    try {
      if (fs.existsSync(templatePath)) {
        fs.unlinkSync(templatePath);
      }
      if (fs.existsSync(outputPath)) {
        fs.unlinkSync(outputPath);
      }
    } catch (cleanupError) {
      console.warn(`⚠️ Cleanup failed: ${cleanupError.message}`);
    }
    
    res.json({
      ok: true,
      driveFile: uploadedFile,
      method: 'pdftk',
    });
    
  } catch (error) {
    console.error('❌ Error:', error);
    res.status(500).json({
      error: 'Fill failed',
      detail: error.message,
    });
  }
});

app.listen(PORT, () => {
  console.log(`🚀 PDF Filler PDFtk running on port ${PORT}`);
  console.log(`   Temp directory: ${TMP}`);
  console.log(`   Output folder: ${OUTPUT_FOLDER_ID || 'not set'}`);
});
