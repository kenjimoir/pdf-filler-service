// PDF Filler v2 - Clean implementation with Japanese font support
// Strategy: Manual appearance updates + save with updateFieldAppearances: false

const fs = require('fs');
const os = require('os');
const path = require('path');
const express = require('express');
const cors = require('cors');
const { PDFDocument, PDFName, PDFString } = require('pdf-lib');
const { google } = require('googleapis');
const fontkit = require('@pdf-lib/fontkit');

const PORT = process.env.PORT || 8080;
const TMP = path.join(os.tmpdir(), 'pdf-filler-v2');
const OUTPUT_FOLDER_ID = process.env.OUTPUT_FOLDER_ID;

// Ensure temp directory exists
if (!fs.existsSync(TMP)) {
  fs.mkdirSync(TMP, { recursive: true });
}

// Initialize Google Drive client
function getDriveClient() {
  // Option 1: JSON string in environment variable
  if (process.env.GOOGLE_CREDENTIALS_JSON) {
    const creds = JSON.parse(process.env.GOOGLE_CREDENTIALS_JSON);
    const auth = new google.auth.GoogleAuth({
      credentials: creds,
      scopes: ['https://www.googleapis.com/auth/drive'],
    });
    return google.drive({ version: 'v3', auth });
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

// Load and embed Japanese font
async function loadJapaneseFont(pdfDoc) {
  const fontFileName = process.env.FONT_TTF_PATH || 'fonts/NotoSansCJKjp-Regular.otf';
  
  // Try multiple paths
  const possiblePaths = [
    path.join(__dirname, fontFileName),
    path.join(__dirname, 'fonts', path.basename(fontFileName)),
    path.join(process.cwd(), fontFileName),
    path.join(process.cwd(), 'fonts', path.basename(fontFileName)),
  ];
  
  let fontPath = null;
  for (const tryPath of possiblePaths) {
    if (fs.existsSync(tryPath)) {
      fontPath = tryPath;
      break;
    }
  }
  
  if (!fontPath) {
    throw new Error(`Font file not found. Tried: ${possiblePaths.join(', ')}`);
  }
  
  // Reject TTC files
  if (fontPath.toLowerCase().endsWith('.ttc')) {
    // Try fallback
    const fallback = path.join(__dirname, 'fonts', 'NotoSansCJKjp-Regular.otf');
    if (fs.existsSync(fallback)) {
      fontPath = fallback;
    } else {
      throw new Error('TTC files not supported. Please use OTF or TTF font.');
    }
  }
  
  pdfDoc.registerFontkit(fontkit);
  const fontBytes = fs.readFileSync(fontPath);
  const customFont = await pdfDoc.embedFont(fontBytes);
  
  console.log(`✅ Loaded Japanese font from: ${fontPath}`);
  console.log(`   Font object: ${customFont ? 'valid' : 'invalid'}`);
  return customFont;
}

// Fill PDF with fields
async function fillPdf(srcPath, outPath, fields, customFont) {
  console.log(`📝 Starting PDF fill process...`);
  console.log(`   Template: ${srcPath}`);
  console.log(`   Output: ${outPath}`);
  console.log(`   Fields to fill: ${Object.keys(fields).length}`);
  console.log(`   Custom font: ${customFont ? 'loaded' : 'not loaded'}`);
  
  const bytes = fs.readFileSync(srcPath);
  const pdfDoc = await PDFDocument.load(bytes, { updateFieldAppearances: false });
  
  const form = pdfDoc.getForm();
  const allFields = form.getFields();
  
  console.log(`📋 Found ${allFields.length} form fields in template`);
  
  let filledCount = 0;
  let processedCount = 0;
  
  // Fill fields with timeout protection
  const startTime = Date.now();
  const MAX_PROCESSING_TIME = 30000; // 30 seconds max
  
  for (const field of allFields) {
    // Check timeout
    if (Date.now() - startTime > MAX_PROCESSING_TIME) {
      throw new Error(`PDF processing timeout after ${MAX_PROCESSING_TIME}ms`);
    }
    
    processedCount++;
    const fieldName = field.getName();
    const fieldType = field.constructor.name;
    const value = fields[fieldName];
    
    if (value == null || value === '') {
      if (processedCount % 20 === 0) {
        console.log(`   Processed ${processedCount}/${allFields.length} fields...`);
      }
      continue;
    }
    
    try {
      if (fieldType.includes('TextField')) {
        // Text field: simplified approach
        const textValue = String(value);
        
        // Set font first if available
        if (customFont) {
          try {
            field.updateAppearances(customFont);
          } catch (fontError) {
            console.warn(`⚠️ Font update failed for "${fieldName}": ${fontError.message}`);
          }
        }
        
        // Set text
        try {
          field.setText(textValue);
          filledCount++;
        } catch (textError) {
          if (textError.message && textError.message.includes('WinAnsi')) {
            console.warn(`⚠️ WinAnsi error on "${fieldName}" - skipping font update`);
            // Try without font update
            if (customFont) {
              try {
                field.updateAppearances(); // No font parameter
                field.setText(textValue);
                filledCount++;
              } catch (retryError) {
                console.warn(`⚠️ Retry failed for "${fieldName}": ${retryError.message}`);
              }
            }
          } else {
            throw textError;
          }
        }
        
      } else if (fieldType.includes('CheckBox')) {
        // Checkbox: check/uncheck then update appearance (no font - uses ZapfDingbats)
        const shouldCheck = (
          value === 'on' ||
          value === 'yes' ||
          value === 'true' ||
          value === '1' ||
          value === 'はい' ||
          String(value).toUpperCase() === 'TRUE'
        );
        
        if (shouldCheck) {
          field.check();
        } else {
          field.uncheck();
        }
        
        // Update appearance WITHOUT font (checkboxes use ZapfDingbats)
        field.updateAppearances();
        filledCount++;
        
      } else if (fieldType.includes('Dropdown')) {
        // Dropdown: select value
        field.select(String(value));
        if (customFont) {
          field.updateAppearances(customFont);
        }
        filledCount++;
      }
    } catch (error) {
      console.warn(`⚠️ Failed to fill field "${fieldName}": ${error.message}`);
    }
  }
  
  // NOTE: We're relying on manual updateAppearances() calls above
  // Setting AcroForm DA might interfere with checkbox rendering
  // So we skip it and trust that updateAppearances(customFont) worked for text fields
  
  // Save with updateFieldAppearances: false to preserve our manual appearances
  // The DA we set above will be used if any fields need default font
  const pdfBytes = await pdfDoc.save({
    updateFieldAppearances: false,  // Don't recalculate - use our manual updates
    useObjectStreams: false,
    addDefaultPage: false,
  });
  
  fs.writeFileSync(outPath, pdfBytes);
  console.log(`✅ Filled ${filledCount} fields, saved to: ${outPath}`);
  
  return { filled: filledCount, size: pdfBytes.length };
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
  
  // Use the verified folderId (already set above as finalFolderId)
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
    requestBody: fileMetadata, // Use requestBody (matches old service format)
    media: media,
    fields: 'id, name, webViewLink, webContentLink',
    supportsAllDrives: true, // Support Shared Drives
  });
  
  return file.data;
}

// HTTP Server
const app = express();
app.use(cors());
app.use(express.json({ limit: '10mb' }));

app.get('/', (_req, res) => {
  res.json({ service: 'PDF Filler v2', status: 'running' });
});

app.get('/health', (_req, res) => {
  res.json({
    ok: true,
    fontPath: process.env.FONT_TTF_PATH || 'fonts/NotoSansCJKjp-Regular.otf',
    outputFolder: OUTPUT_FOLDER_ID || 'not set',
  });
});

app.post('/fill', async (req, res) => {
  const { templateFileId, fields, outputName, folderId } = req.body;
  
  if (!templateFileId || !fields) {
    return res.status(400).json({ error: 'Missing templateFileId or fields' });
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
    
    // 2. Load font (with fallback)
    const pdfDoc = await PDFDocument.load(fs.readFileSync(templatePath));
    let customFont = null;
    
    try {
      customFont = await loadJapaneseFont(pdfDoc);
    } catch (fontError) {
      console.warn(`⚠️ Font loading failed: ${fontError.message}`);
      console.warn(`   Continuing without custom font (may cause WinAnsi errors)`);
    }
    
    // 3. Fill PDF with timeout
    console.log(`📝 Filling PDF with ${Object.keys(fields).length} fields...`);
    
    const fillPromise = fillPdf(templatePath, outputPath, fields, customFont);
    const timeoutPromise = new Promise((_, reject) => {
      setTimeout(() => reject(new Error('PDF fill timeout after 60 seconds')), 60000);
    });
    
    await Promise.race([fillPromise, timeoutPromise]);
    
    // 4. Upload to Drive
    const finalName = outputName || `filled_${Date.now()}.pdf`;
    console.log(`📤 Uploading to Drive: ${finalName}`);
    const uploadedFile = await uploadToDrive(
      drive,
      outputPath,
      finalName,
      folderId || OUTPUT_FOLDER_ID
    );
    
    // Cleanup
    fs.unlinkSync(templatePath);
    fs.unlinkSync(outputPath);
    
    res.json({
      ok: true,
      driveFile: uploadedFile,
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
  console.log(`🚀 PDF Filler v2 running on port ${PORT}`);
});

