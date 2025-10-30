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
// Checkbox/radio values use name objects (/Yes or /Off) and also set /AS
function generateFDF(fields) {
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
    const isOn = /^(on|yes|true)$/i.test(val);
    const isOff = /^(off|no|false)$/i.test(val);

    fdf += '<<\n';
    fdf += `/T (${fieldName})\n`;
    if (isOn || isOff) {
      const name = isOn ? 'Yes' : 'Off';
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

// Fill PDF using PDFtk
async function fillPdfWithPDFtk(templatePath, outputPath, fields) {
  console.log(`📝 Filling PDF with PDFtk...`);
  console.log(`   Template: ${templatePath}`);
  console.log(`   Output: ${outputPath}`);
  console.log(`   Fields: ${Object.keys(fields).length}`);
  
  // Generate FDF content
  const fdfContent = generateFDF(fields);
  const fdfPath = path.join(TMP, `data_${Date.now()}.fdf`);
  fs.writeFileSync(fdfPath, fdfContent, 'utf8');
  
  try {
    // Use PDFtk to fill the form
    // The 'flatten' option converts form fields to static text
    const command = `pdftk "${templatePath}" fill_form "${fdfPath}" output "${outputPath}" flatten`;
    console.log(`🔧 Running: ${command}`);
    
    const { stdout, stderr } = await execAsync(command);
    
    if (stderr) {
      console.warn(`⚠️ PDFtk stderr: ${stderr}`);
    }
    
    console.log(`✅ PDF filled successfully with PDFtk`);
    console.log(`   Output size: ${fs.statSync(outputPath).size} bytes`);
    
    // Clean up FDF file
    fs.unlinkSync(fdfPath);
    
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
  const { templateFileId, fields, outputName, folderId } = req.body;
  
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
    // 1. Download template
    console.log(`📥 Downloading template: ${templateFileId}`);
    const templateFile = await drive.files.get(
      { fileId: templateFileId, alt: 'media' },
      { responseType: 'stream' }
    );
    
    const writeStream = fs.createWriteStream(templatePath);
    templateFile.data.pipe(writeStream);
    
    await new Promise((resolve, reject) => {
      writeStream.on('finish', resolve);
      writeStream.on('error', reject);
    });
    console.log('✅ Template downloaded');
    
    // 2. Fill PDF with PDFtk
    console.log(`📝 Filling PDF with ${Object.keys(fields).length} fields...`);
    await fillPdfWithPDFtk(templatePath, outputPath, fields);
    
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
