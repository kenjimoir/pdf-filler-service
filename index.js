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
  
  if (!customFont) {
    throw new Error(`Failed to embed font - embedFont() returned null/undefined`);
  }
  
  console.log(`✅ Loaded Japanese font from: ${fontPath}`);
  console.log(`   Font object: valid, size: ${fontBytes.length} bytes`);
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
  if (bytes.length === 0) {
    throw new Error(`Template file is empty: ${srcPath}`);
  }
  
  const pdfDoc = await PDFDocument.load(bytes, { updateFieldAppearances: false });
  if (!pdfDoc) {
    throw new Error(`Failed to load PDF document from: ${srcPath}`);
  }
  
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
        // Text field: MANDATORY font application to prevent WinAnsi fallback
        const textValue = String(value);
        
        // Font application is MANDATORY - no fallback allowed
        if (!customFont) {
          console.warn(`⚠️ Skipping text field "${fieldName}" - no custom font available`);
          continue; // Skip this field entirely
        }
        
        try {
          // Set font FIRST - this MUST succeed to prevent WinAnsi fallback
          field.updateAppearances(customFont);
          console.log(`✅ Applied custom font to text field "${fieldName}"`);
        } catch (fontError) {
          console.warn(`⚠️ Skipping text field "${fieldName}" - font application failed: ${fontError.message}`);
          continue; // Skip this field entirely - don't allow WinAnsi fallback
        }
        
        // Now set text - font is guaranteed to be applied
        try {
          field.setText(textValue);
          filledCount++;
          console.log(`✅ Set text field "${fieldName}" to: "${textValue}"`);
        } catch (textError) {
          console.warn(`⚠️ Skipping text field "${fieldName}" - text setting failed: ${textError.message}`);
          // Skip this field - don't retry without font
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
        
        // Don't call updateAppearances on checkboxes - they use ZapfDingbats, not custom font
        filledCount++;
        
      } else if (fieldType.includes('Dropdown')) {
        // Dropdown: select value
        field.select(String(value));
        // Don't call updateAppearances on dropdowns - they use default fonts
        filledCount++;
      }
    } catch (error) {
      console.warn(`⚠️ Failed to fill field "${fieldName}": ${error.message}`);
    }
  }
  
  // Set AcroForm default appearance to use custom font
  // This ensures updateFieldAppearances: true uses the correct font
  if (customFont) {
    try {
      const acroForm = pdfDoc.catalog.get(PDFName.of('AcroForm'));
      if (acroForm) {
        // Get the actual font reference name from the embedded font
        // The font reference is typically F0, F1, etc. based on order of embedding
        const fontRef = customFont.ref || 'F0'; // Fallback to F0 if ref not available
        const daString = `/${fontRef} 12 Tf 0 g`; // Font reference, size 12, color black
        acroForm.set(PDFName.of('DA'), PDFString.of(daString));
        console.log(`✅ Set AcroForm default appearance to use custom font: ${daString}`);
      }
    } catch (daError) {
      console.warn(`⚠️ Failed to set AcroForm DA: ${daError.message}`);
    }
  }
  
  // Save with updateFieldAppearances: true to generate appearance streams
  // The AcroForm DA should ensure it uses our custom font
  const pdfBytes = await pdfDoc.save({
    updateFieldAppearances: true,  // Generate appearance streams for visual display
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
    
    // 2. Load font (reuse the PDF document from fillPdf)
    let customFont = null;
    
    // Load PDF first to pass to font loading
    const templateBytes = fs.readFileSync(templatePath);
    if (templateBytes.length === 0) {
      throw new Error(`Template file is empty: ${templatePath}`);
    }
    
    const pdfDoc = await PDFDocument.load(templateBytes, { updateFieldAppearances: false });
    
    // Check if we have any text fields that might contain Japanese text
    const form = pdfDoc.getForm();
    const allFields = form.getFields();
    const hasTextFields = allFields.some(field => field.constructor.name.includes('TextField'));
    
    if (hasTextFields) {
      // Font loading is MANDATORY for text fields - fail if it doesn't work
      try {
        customFont = await loadJapaneseFont(pdfDoc);
        console.log('✅ Custom font loaded successfully');
      } catch (fontError) {
        throw new Error(`Font loading is mandatory for Japanese text: ${fontError.message}`);
      }
    } else {
      console.log('ℹ️ No text fields found - skipping font loading');
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
    
    // Cleanup (with error handling)
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

