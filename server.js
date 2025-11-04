import express from 'express';
import { PDFDocument, StandardFonts } from 'pdf-lib';
import fontkit from '@pdf-lib/fontkit';
import { google } from 'googleapis';
import { readFileSync } from 'fs';
import { fileURLToPath } from 'url';
import { dirname, join } from 'path';
import { existsSync } from 'fs';
import dotenv from 'dotenv';

// Load environment variables from .env file (for local development)
dotenv.config();

const __filename = fileURLToPath(import.meta.url);
const __dirname = dirname(__filename);

const app = express();
const PORT = process.env.PORT || 3000;

// JSON body parser (limit 20mb for large PDFs)
app.use(express.json({ limit: '20mb' }));

// Check if string contains CJK characters
function containsCJK(str) {
  if (!str) return false;
  return /[\u4e00-\u9fff\u3400-\u4dbf\u3040-\u309f\u30a0-\u30ff\uac00-\ud7af]/.test(str);
}

// Load fonts
let cjkFont = null;
let latinFont = null;
let zapfDingbatsFont = null;

async function loadFonts() {
  try {
    // Load CJK font (required)
    // Try multiple possible file names (both .otf and .ttf formats)
    const possibleCjkNames = [
      'NotoSansCJKjp-Regular.otf',
      'NotoSansJP-Regular.otf',
      'NotoSansJP-Regular.ttf',  // TTF format support
      'NotoSansCJK-Regular.otf',
      'NotoSansCJK-Regular.ttf',  // TTF format support
    ];
    
    let cjkFontPath = null;
    for (const name of possibleCjkNames) {
      const path = join(__dirname, 'fonts', name);
      if (existsSync(path)) {
        cjkFontPath = path;
        break;
      }
    }
    
    if (!cjkFontPath) {
      throw new Error(`Required CJK font not found. Please place one of these files in fonts/ directory: ${possibleCjkNames.join(', ')}`);
    }
    
    const cjkFontBytes = readFileSync(cjkFontPath);
    console.log('Loaded CJK font:', cjkFontPath);

    // Load Latin font (optional)
    const latinFontPath = join(__dirname, 'fonts', 'NotoSans-Regular.ttf');
    if (existsSync(latinFontPath)) {
      const latinFontBytes = readFileSync(latinFontPath);
      console.log('Loaded Latin font:', latinFontPath);
      latinFont = { bytes: latinFontBytes };
    } else {
      console.log('Latin font not found (optional), will use CJK font for all text');
    }

    cjkFont = { bytes: cjkFontBytes };
  } catch (error) {
    console.error('Error loading fonts:', error);
    throw error;
  }
}

// Initialize Google Drive API
let drive = null;

async function initializeDrive() {
  try {
    const serviceAccountBase64 = process.env.GOOGLE_SERVICE_ACCOUNT_BASE64;
    if (!serviceAccountBase64) {
      throw new Error('GOOGLE_SERVICE_ACCOUNT_BASE64 environment variable is required');
    }

    const serviceAccountJson = JSON.parse(
      Buffer.from(serviceAccountBase64, 'base64').toString('utf-8')
    );

    const auth = new google.auth.JWT({
      email: serviceAccountJson.client_email,
      key: serviceAccountJson.private_key,
      scopes: ['https://www.googleapis.com/auth/drive'],
    });

    drive = google.drive({ version: 'v3', auth });
    console.log('Google Drive API initialized');
  } catch (error) {
    console.error('Error initializing Google Drive:', error);
    throw error;
  }
}

// Bearer token authentication middleware
function authenticateBearerToken(req, res, next) {
  const authHeader = req.headers.authorization;
  const expectedToken = process.env.API_BEARER_TOKEN;

  if (!expectedToken) {
    console.error('API_BEARER_TOKEN environment variable is not set');
    return res.status(500).json({ ok: false, error: 'Server configuration error' });
  }

  if (!authHeader || !authHeader.startsWith('Bearer ')) {
    return res.status(401).json({ ok: false, error: 'Missing or invalid Authorization header' });
  }

  const token = authHeader.substring(7);
  if (token !== expectedToken) {
    return res.status(401).json({ ok: false, error: 'Invalid bearer token' });
  }

  next();
}

// Health check endpoint
app.get('/health', (req, res) => {
  res.json({ ok: true });
});

// Set field value and update appearance
async function setFieldValue(form, field, value, font, zapfFont) {
  try {
    const fieldType = field.constructor.name;

    if (fieldType === 'PDFCheckBox') {
      // Checkbox handling
      const boolValue = value === true || value === 'true' || value === 'TRUE' || 
                       value === 'on' || value === 'On' || value === 'ON' ||
                       value === 'yes' || value === 'Yes' || value === 'YES' ||
                       String(value).toLowerCase() === 'true';
      
      if (boolValue) {
        // Try to check with export value first
        try {
          field.check(value);
        } catch (e) {
          // Fallback to simple check()
          field.check();
        }
      } else {
        field.uncheck();
      }
      
      // Update appearance with ZapfDingbats
      if (zapfFont) {
        await field.updateAppearances(zapfFont);
      }
    } else if (fieldType === 'PDFRadioGroup') {
      // Radio button handling
      const stringValue = String(value || '').trim();
      if (stringValue) {
        try {
          field.select(stringValue);
        } catch (e) {
          // Try lowercase
          try {
            field.select(stringValue.toLowerCase());
          } catch (e2) {
            // Try uppercase
            try {
              field.select(stringValue.toUpperCase());
            } catch (e3) {
              console.warn(`Could not set radio value "${stringValue}" for field "${field.getName()}"`);
            }
          }
        }
      }
      
      // Update appearance with ZapfDingbats
      if (zapfFont) {
        await field.updateAppearances(zapfFont);
      }
    } else if (fieldType === 'PDFTextField') {
      // Text field handling
      const stringValue = String(value || '');
      field.setText(stringValue);
      
      // Update appearance with provided font
      if (font) {
        await field.updateAppearances(font);
      }
    } else {
      // Other field types
      const stringValue = String(value || '');
      
      try {
        if (typeof field.setText === 'function') {
          field.setText(stringValue);
        }
      } catch (e) {
        console.warn(`Could not set value for field "${field.getName()}"`);
      }
      
      // Update appearance with provided font
      if (font) {
        try {
          await field.updateAppearances(font);
        } catch (e) {
          console.warn(`Could not update appearance for field "${field.getName()}"`);
        }
      }
    }
  } catch (error) {
    console.warn(`Error setting field value: ${error.message}`);
  }
}

// Main PDF fill endpoint
app.post('/fill', authenticateBearerToken, async (req, res) => {
  try {
    // Support both formats: {templateId, output: {name, folderId}, fields}
    // and {templateFileId, outputName, folderId, fields} (from Code.gs)
    let templateId, outputName, folderId, fields;

    if (req.body.templateId && req.body.output) {
      // New format from requirements
      templateId = req.body.templateId;
      outputName = req.body.output.name;
      folderId = req.body.output.folderId;
      fields = req.body.fields || {};
    } else if (req.body.templateFileId) {
      // Existing format from Code.gs
      templateId = req.body.templateFileId;
      outputName = req.body.outputName;
      folderId = req.body.folderId;
      fields = req.body.fields || {};
    } else {
      return res.status(400).json({ 
        ok: false, 
        error: 'Missing required fields: templateId/templateFileId, output/outputName, folderId, fields' 
      });
    }

    if (!templateId) {
      return res.status(400).json({ ok: false, error: 'templateId is required' });
    }

    if (!outputName) {
      return res.status(400).json({ ok: false, error: 'outputName is required' });
    }

    // Use default folder if not specified
    if (!folderId) {
      folderId = process.env.DEFAULT_OUTPUT_FOLDER_ID || '';
      if (!folderId) {
        return res.status(400).json({ ok: false, error: 'folderId is required' });
      }
    }

    console.log(`Processing PDF fill request: templateId=${templateId}, outputName=${outputName}`);

    // 1. Download template PDF from Google Drive
    console.log('Downloading template PDF...');
    const templateResponse = await drive.files.get(
      { fileId: templateId, alt: 'media' },
      { responseType: 'arraybuffer' }
    );
    const templateBytes = new Uint8Array(templateResponse.data);

    // 2. Load PDF with pdf-lib
    console.log('Loading PDF document...');
    const pdfDoc = await PDFDocument.load(templateBytes);
    pdfDoc.registerFontkit(fontkit);

    // 3. Embed fonts
    const cjkFontEmbedded = await pdfDoc.embedFont(cjkFont.bytes);
    const zapfDingbatsFontEmbedded = await pdfDoc.embedStandardFont(StandardFonts.ZapfDingbats);
    
    let latinFontEmbedded = null;
    if (latinFont) {
      latinFontEmbedded = await pdfDoc.embedFont(latinFont.bytes);
    }

    // 4. Get form and fill fields
    console.log('Filling form fields...');
    const form = pdfDoc.getForm();

    // Process each field
    for (const [fieldName, value] of Object.entries(fields)) {
      try {
        const field = form.getTextField(fieldName) || 
                     form.getCheckBox(fieldName) || 
                     form.getRadioGroup(fieldName) ||
                     form.getDropdown(fieldName);

        if (field) {
          // Determine font based on field type
          let font = cjkFontEmbedded;
          let zapfFont = null;

          const fieldType = field.constructor.name;
          if (fieldType === 'PDFCheckBox' || fieldType === 'PDFRadioGroup') {
            zapfFont = zapfDingbatsFontEmbedded;
          } else if (fieldType === 'PDFTextField') {
            // Use Latin font if available and value doesn't contain CJK
            if (latinFontEmbedded && !containsCJK(String(value))) {
              font = latinFontEmbedded;
            }
          }

          await setFieldValue(form, field, value, font, zapfFont);
        } else {
          console.warn(`Field not found: ${fieldName}`);
        }
      } catch (error) {
        console.warn(`Error processing field "${fieldName}": ${error.message}`);
      }
    }

    // 5. Update all field appearances with CJK font as fallback
    console.log('Updating all field appearances...');
    try {
      await form.updateFieldAppearances(cjkFontEmbedded);
    } catch (error) {
      console.warn(`Error updating field appearances: ${error.message}`);
    }

    // 6. Save PDF (non-flattened, compatibility-focused)
    console.log('Saving PDF...');
    const pdfBytes = await pdfDoc.save({ 
      updateFieldAppearances: false,  // Already updated manually
      useObjectStreams: false  // Better compatibility with older viewers
    });

    // 7. Upload to Google Drive
    console.log('Uploading to Google Drive...');
    const fileMetadata = {
      name: outputName,
      parents: folderId ? [folderId] : undefined,
    };

    const media = {
      mimeType: 'application/pdf',
      body: Buffer.from(pdfBytes),
    };

    const uploadResponse = await drive.files.create({
      requestBody: fileMetadata,
      media: media,
      fields: 'id, name, webViewLink',
    });

    const uploadedFile = uploadResponse.data;

    console.log(`PDF uploaded successfully: ${uploadedFile.id}`);

    // 8. Return response
    res.json({
      ok: true,
      file: {
        id: uploadedFile.id,
        name: uploadedFile.name,
        webViewLink: uploadedFile.webViewLink,
      },
      driveFile: {
        id: uploadedFile.id,
        name: uploadedFile.name,
        parents: folderId ? [folderId] : [],
        webViewLink: uploadedFile.webViewLink,
      },
    });

  } catch (error) {
    console.error('Error processing PDF fill request:', error);
    // Don't expose stack trace to client
    const errorMessage = error.message || 'Internal server error';
    res.status(500).json({ 
      ok: false, 
      error: errorMessage 
    });
  }
});

// Initialize and start server
async function startServer() {
  try {
    await loadFonts();
    await initializeDrive();
    
    app.listen(PORT, () => {
      console.log(`PDF filler service running on port ${PORT}`);
    });
  } catch (error) {
    console.error('Failed to start server:', error);
    process.exit(1);
  }
}

startServer();

