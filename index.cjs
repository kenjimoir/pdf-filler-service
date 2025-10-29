const express = require('express');
const { PDFDocument, PDFName, PDFString, PDFBool } = require('pdf-lib');
const fs = require('fs');
const path = require('path');

const app = express();
const PORT = process.env.PORT || 3000;
const ROOT = process.env.ROOT || __dirname;

// Simple logging function
const log = (...args) => console.log(new Date().toISOString(), '-', ...args);

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

// Create multipart body for Google Drive upload
function createMultipartBody(fileBytes, fileName, folderId) {
  const boundary = '----WebKitFormBoundary' + Math.random().toString(16).substr(2);
  const metadata = {
    name: fileName,
    parents: folderId ? [folderId] : undefined,
  };
  
  const body = [
    `--${boundary}`,
    'Content-Disposition: form-data; name="metadata"',
    'Content-Type: application/json',
    '',
    JSON.stringify(metadata),
    `--${boundary}`,
    'Content-Disposition: form-data; name="file"',
    'Content-Type: application/pdf',
    '',
    fileBytes,
    `--${boundary}--`,
  ].join('\r\n');
  
  return {
    body: Buffer.from(body),
    headers: {
      'Content-Type': `multipart/related; boundary=${boundary}`,
    },
  };
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
    const templateResponse = await fetch(`https://www.googleapis.com/drive/v3/files/${templateFileId}?alt=media`, {
      headers: {
        'Authorization': `Bearer ${accessToken}`,
      },
    });
    
    if (!templateResponse.ok) {
      throw new Error(`Failed to download template: ${templateResponse.status}`);
    }
    const templateBytes = await templateResponse.arrayBuffer();
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
    const multipartData = createMultipartBody(pdfBytes, outputName || 'filled.pdf', folderId);
    const uploadResponse = await fetch('https://www.googleapis.com/upload/drive/v3/files?uploadType=multipart', {
      method: 'POST',
      headers: {
        'Authorization': `Bearer ${accessToken}`,
        ...multipartData.headers,
      },
      body: multipartData.body,
    });

    if (!uploadResponse.ok) {
      const errorText = await uploadResponse.text();
      throw new Error(`Upload failed: ${uploadResponse.status} - ${errorText}`);
    }

    const uploadResult = await uploadResponse.json();
    log('PDF uploaded successfully, file ID:', uploadResult.id);

    res.json({
      success: true,
      fileId: uploadResult.id,
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
