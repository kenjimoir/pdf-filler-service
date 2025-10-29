const express = require('express');
const { PDFDocument, PDFName, PDFString, PDFBool } = require('pdf-lib');
const { google } = require('googleapis');
const fs = require('fs');
const path = require('path');

const app = express();
const PORT = process.env.PORT || 3000;
const ROOT = process.env.ROOT || __dirname;

// Simple logging function
const log = (...args) => console.log(new Date().toISOString(), '-', ...args);

// Google Drive client
function getDriveClient() {
  let credentials = null;
  if (process.env.GOOGLE_CREDENTIALS_JSON) {
    try { 
      credentials = JSON.parse(process.env.GOOGLE_CREDENTIALS_JSON); 
    } catch (e) {
      log('ERROR parsing GOOGLE_CREDENTIALS_JSON:', e && e.message);
    }
  }
  const auth = new google.auth.GoogleAuth({
    credentials,
    scopes: ['https://www.googleapis.com/auth/drive'],
  });
  return google.drive({ version: 'v3', auth });
}

app.use(express.json({ limit: '50mb' }));

// Simple field value resolver
function resolveValue(fieldName, fields) {
  return fields[fieldName] || '';
}

// Normalize yes/no values
function normalizeYesNo(value) {
  if (!value) return 'no';
  const str = String(value).toLowerCase().trim();
  if (str === 'はい' || str === 'yes' || str === 'on' || str === 'true' || str === '1') return 'yes';
  return 'no';
}


app.post('/fill', async (req, res) => {
  try {
    log('Received PDF fill request');
    
    const { templateFileId, fields, outputName, folderId, mode, accessToken } = req.body;
    if (!templateFileId || !fields) {
      return res.status(400).json({ error: 'Missing templateFileId or fields' });
    }

    log('Fields received:', Object.keys(fields).length);
    log('Template file ID:', templateFileId);
    log('Output name:', outputName);
    log('Folder ID:', folderId);

    // Download template from Google Drive
    log('Downloading template from Google Drive...');
    const drive = getDriveClient();
    const templateResponse = await drive.files.get({
      fileId: templateFileId,
      alt: 'media',
      supportsAllDrives: true,
    }, {
      responseType: 'arraybuffer',
    });
    
    const templateBytes = templateResponse.data;
    log('Template downloaded, size:', templateBytes.byteLength);

    // Load PDF
    const pdfDoc = await PDFDocument.load(templateBytes);
    const form = pdfDoc.getForm();
    const allFields = form.getFields();
    
    log('PDF loaded, total fields:', allFields.length);

    let filled = 0;

    // Process each field
    for (const field of allFields) {
      const fieldName = field.getName();
      const fieldValue = resolveValue(fieldName, fields);
      
      if (fieldValue === '') continue;

      try {
        // Handle text fields
        if (field.constructor.name.includes('TextField')) {
          field.setText(String(fieldValue));
          filled++;
          log(`✅ Text field ${fieldName}: "${fieldValue}"`);
        }
        // Handle checkboxes
        else if (field.constructor.name.includes('CheckBox')) {
          // Handle explicit checkbox fields (CoverageValue_日, CoverageValue_月, TravelerSex_男性, TravelerSex_女性)
          if (fieldName === 'CoverageValue_日' || fieldName === 'CoverageValue_月' || 
              fieldName === 'TravelerSex_男性' || fieldName === 'TravelerSex_女性') {
            const shouldCheck = fieldValue === 'on';
            
            if (shouldCheck) {
              field.check();
              log(`✅ Checkbox ${fieldName}: checked`);
            } else {
              field.uncheck();
              log(`❌ Checkbox ${fieldName}: unchecked`);
            }
          }
          // Handle other checkboxes (8 questions, phone type, same as traveler)
          else {
            const shouldCheck = normalizeYesNo(fieldValue) === 'yes';
            
            if (shouldCheck) {
              field.check();
              log(`✅ Checkbox ${fieldName}: checked`);
            } else {
              field.uncheck();
              log(`❌ Checkbox ${fieldName}: unchecked`);
            }
          }
          filled++;
        }
      } catch (error) {
        log(`⚠️ Error processing field ${fieldName}:`, error.message);
      }
    }

    log(`Processed ${filled} fields successfully`);

    // Save PDF (keep interactive)
    const pdfBytes = await pdfDoc.save({
      useObjectStreams: false,
      addDefaultPage: false,
    });

    log('PDF saved, size:', pdfBytes.length);

    // Upload to Google Drive
    log('Starting upload to Google Drive...');
    const drive = getDriveClient();
    const fileMetadata = {
      name: outputName || 'filled.pdf',
      parents: folderId ? [folderId] : undefined,
    };
    
    const uploadResult = await drive.files.create({
      requestBody: fileMetadata,
      media: {
        mimeType: 'application/pdf',
        body: Buffer.from(pdfBytes),
      },
      fields: 'id, name, webViewLink',
      supportsAllDrives: true,
    });
    
    log('PDF uploaded successfully, file ID:', uploadResult.data.id);

    res.json({
      success: true,
      fileId: uploadResult.data.id,
      fieldsProcessed: filled,
    });

  } catch (error) {
    log('ERROR:', error.message);
    res.status(500).json({
      error: 'Fill failed',
      detail: error.message,
    });
  }
});

app.get('/health', (req, res) => {
  res.json({ status: 'ok', timestamp: new Date().toISOString() });
});

app.listen(PORT, () => {
  log(`PDF Filler Service running on port ${PORT}`);
});
