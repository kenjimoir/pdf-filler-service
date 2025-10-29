// index.cjs — PDF fill → upload to Drive (Shared Drives OK)
// deps: express, cors, pdf-lib, googleapis, @pdf-lib/fontkit
// runs on Render (uses process.env.PORT) and writes temp files in os.tmpdir()

const fs = require('fs');
const os = require('os');
const path = require('path');
const express = require('express');
const cors = require('cors');
const { PDFDocument, rgb, degrees, PDFName, PDFBool, PDFDict, PDFString, PDFNumber } = require('pdf-lib');
const { google } = require('googleapis');
const fontkit = require('@pdf-lib/fontkit');

const PORT = process.env.PORT || 8080;
const TMP = path.join(os.tmpdir(), 'pdf-filler');
ensureDir(TMP);
const ROOT = process.cwd();

const OUTPUT_FOLDER_ID = process.env.OUTPUT_FOLDER_ID || '';

// === new flags ===
const RESPECT_TEMPLATE_APPEARANCE =
  String(process.env.RESPECT_TEMPLATE_APPEARANCE || 'false').toLowerCase() === 'true';
const FORCE_BURN_IN =
  String(process.env.FORCE_BURN_IN || 'true').toLowerCase() === 'true';

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

  // 1) font - Use system font that supports Japanese
  try { pdfDoc.registerFontkit(fontkit); } catch (_) {}
  let customFont = null;
  
  // Try to use system fonts that support Japanese
  const systemFonts = [
    '/System/Library/Fonts/Helvetica.ttc', // macOS
    '/System/Library/Fonts/Arial.ttf',     // macOS
    'C:/Windows/Fonts/arial.ttf',         // Windows
    'C:/Windows/Fonts/msgothic.ttc',       // Windows (Japanese)
    '/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf', // Linux
  ];
  
  for (const fontPath of systemFonts) {
    try {
      if (fs.existsSync(fontPath)) {
        const fontBytes = fs.readFileSync(fontPath);
        customFont = await pdfDoc.embedFont(fontBytes);
        log('Using system font:', fontPath);
        break;
      }
    } catch (e) {
      log('Font embed failed for', fontPath, e.message);
    }
  }
  
  // If no system font worked, try to embed Helvetica (fallback)
  if (!customFont) {
    try {
      customFont = await pdfDoc.embedFont('Helvetica');
      log('Using Helvetica fallback font');
    } catch (e) {
      log('All font embedding failed, using PDF default font:', e.message);
    }
  }

  // 2) AcroForm default appearance (only if custom font is available)
  if (customFont) {
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

    // Set default appearance with custom font
    // Note: updateFieldAppearances: false on save preserves template auto-sizing
    acroForm.set(PDFName.of('DA'), PDFString.of('/F0 12 Tf 0 g'));
    // NeedAppearances not set - let template's original value remain (avoid PDFBool API issues)
    
    log('Helvetica font embedding enabled');
  } else {
    log('Using PDF default font (no custom font embedding)');
  }

  // 3) fill
  const form = pdfDoc.getForm();
  const allFields = form.getFields();
  const valueBy = buildAliasView(fields);
  let filled = 0;

  // Re-enable export value checking with robust error handling
  const USE_EXPORT_VALUE_CHECKING = true;
  
  // Collect checkboxes by name for batch processing
  const checkboxGroups = {};

  // Debug: Log all field names and export values to help identify the correct field names
  log('===== ALL PDF FIELD NAMES AND EXPORT VALUES =====');
  for (const f of allFields) {
    const name = f.getName ? f.getName() : '';
    const ctor = f.constructor && f.constructor.name || '';
    let exportValues = '';
    if (ctor.includes('Check') && f.getExportValues) {
      try {
        const exports = f.getExportValues();
        exportValues = ` [exportValues: ${exports.join(', ')}]`;
      } catch (e) {
        exportValues = ` [exportValues: error - ${e.message}]`;
      }
    }
    log(`Field: "${name}" (${ctor})${exportValues}`);
  }
  log('===== END FIELD NAMES =====');

  for (const f of allFields) {
    const name = f.getName ? f.getName() : '';
    const ctor = f.constructor && f.constructor.name || '';
    const valRaw = resolveValue(name, fields, valueBy);

    // Debug specific fields
    if (name === 'CoverageValue' || name === 'DestinationOtherText') {
      log(`Debug field ${name}: valRaw="${valRaw}", ctor="${ctor}"`);
    }

    // Collect checkboxes for batch processing
    if (ctor.includes('Check')) {
      // Debug: Log all checkbox fields
      log(`Found checkbox: name="${name}", valRaw="${valRaw}"`);
      
      if (!checkboxGroups[name]) {
        checkboxGroups[name] = [];
      }
      checkboxGroups[name].push(f);
      continue; // Skip individual processing for now
    }

    if (ctor.includes('Text')) {
      if (valRaw != null && valRaw !== '') {
        f.setText(String(valRaw));
        // Update appearances to use Japanese-compatible font
        // Setting updateFieldAppearances: false on save will preserve template auto-sizing
        if (customFont) {
          try { 
            f.updateAppearances(customFont);
            const isAddressField = name.includes('Address') || name.includes('住所') || name.includes('FullAddress');
            if (isAddressField) {
              log(`Updated appearance for address field "${name}" with auto-sizing preserved`);
            }
          } catch (_) {}
        }
        filled++;
        if (name === 'DestinationOtherText') {
          log(`✅ Set text field ${name} to: "${valRaw}"`);
        }
      } else {
        if (name === 'DestinationOtherText') {
          log(`❌ Text field ${name} has empty/null value: "${valRaw}"`);
        }
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
              if (!RESPECT_TEMPLATE_APPEARANCE && customFont) {
                try { f.updateAppearances(customFont); } catch (_) {}
              }
              filled++;
              log(`✅ Checked checkbox ${n} (era field with value ${numericValue})`);
            } else {
              f.uncheck();
              if (!RESPECT_TEMPLATE_APPEARANCE && customFont) {
                try { f.updateAppearances(customFont); } catch (_) {}
              }
              log(`❌ Unchecked checkbox ${n} (not era field or wrong value)`);
            }
          } catch (e) {
            f.uncheck();
            log(`❌ Error with checkbox ${n}:`, e.message);
          }
        } else {
          // Check for PhoneType, TravelerSex, and SameAsTraveler checkboxes (Japanese text values)
          const isPhoneTypeField = n.includes('PhoneType') || n.includes('電話');
          const isTravelerSexField = n.includes('TravelerSex') || n.includes('Sex') || n.includes('性別');
          const isSameAsTravelerField = n.includes('SameAsTraveler') || n.includes('同一') || n.includes('Same');
          const phoneTypeValue = String(single).trim();
          const travelerSexValue = String(single).trim();
          const sameAsTravelerValue = String(single).trim();
          
          if (isPhoneTypeField) {
            log(`PhoneType checkbox ${n}: value="${phoneTypeValue}"`);
            
            // Check if this field should be checked based on the value
            const shouldCheck = (
              // Handle explicit field names (Option 1 approach)
              (n === 'PhoneType_自宅' && phoneTypeValue === 'yes') ||
              (n === 'PhoneType_勤務先' && phoneTypeValue === 'yes') ||
              (n === 'PhoneType_携帯' && phoneTypeValue === 'yes') ||
              // Handle TravelerPhoneType explicit field names
              (n === 'TravelerPhoneType_自宅' && phoneTypeValue === 'yes') ||
              (n === 'TravelerPhoneType_勤務先' && phoneTypeValue === 'yes') ||
              (n === 'TravelerPhoneType_携帯' && phoneTypeValue === 'yes') ||
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
              if (!RESPECT_TEMPLATE_APPEARANCE && customFont) {
                try { f.updateAppearances(customFont); } catch (_) {}
              }
              filled++;
              log(`✅ Checked PhoneType ${n} (value: ${phoneTypeValue})`);
            } else {
              f.uncheck();
              if (!RESPECT_TEMPLATE_APPEARANCE && customFont) {
                try { f.updateAppearances(customFont); } catch (_) {}
              }
              log(`❌ Unchecked PhoneType checkbox ${n} (no match for value: ${phoneTypeValue})`);
            }
          } else if (isTravelerSexField) {
            log(`TravelerSex checkbox ${n}: value="${travelerSexValue}"`);
            
            // Handle gender checkboxes with same field name but different export values
            if (USE_EXPORT_VALUE_CHECKING) {
            try {
              // Get the export value of this specific checkbox - improved method
              let exportValue = '';
              try {
                // Method 1: Try getExportValues() function
                if (f.getExportValues && typeof f.getExportValues === 'function') {
                  const exports = f.getExportValues();
                  exportValue = exports && exports.length > 0 ? exports[0] : '';
                  log(`Gender Method 1 (getExportValues): ${exportValue}`);
                }
                
                // Method 2: Try exportValues property
                if (!exportValue && f.exportValues && f.exportValues.length > 0) {
                  exportValue = f.exportValues[0];
                  log(`Gender Method 2 (exportValues): ${exportValue}`);
                }
                
                // Method 3: Try options property
                if (!exportValue && f.options && f.options.length > 0) {
                  exportValue = f.options[0];
                  log(`Gender Method 3 (options): ${exportValue}`);
                }
                
                // Method 4: Try to get from widget properties
                if (!exportValue) {
                  try {
                    const widgets = f.acroField && f.acroField.getWidgets ? f.acroField.getWidgets() : [];
                    if (widgets.length > 0) {
                      const widget = widgets[0];
                      if (widget && widget.getOnValue) {
                        const onValue = widget.getOnValue();
                        if (onValue && onValue.asString) {
                          exportValue = onValue.asString();
                          log(`Gender Method 4 (widget OnValue): ${exportValue}`);
                        }
                      }
                    }
                  } catch (widgetError) {
                    log(`Gender Method 4 failed: ${widgetError.message}`);
                  }
                }
                
                // Method 5: Try to get from field properties
                if (!exportValue) {
                  try {
                    if (f.acroField && f.acroField.getOnValue) {
                      const onValue = f.acroField.getOnValue();
                      if (onValue && onValue.asString) {
                        exportValue = onValue.asString();
                        log(`Gender Method 5 (acroField OnValue): ${exportValue}`);
                      }
                    }
                  } catch (acroError) {
                    log(`Gender Method 5 failed: ${acroError.message}`);
                  }
                }
                
              } catch (exportError) {
                log(`Warning: Could not get export value for ${n}: ${exportError.message}`);
                exportValue = '';
              }
              
              log(`TravelerSex checkbox: fieldName="${n}", exportValue="${exportValue}", inputValue="${travelerSexValue}"`);
              
              // Check if this checkbox's export value matches the input value
              if (exportValue === travelerSexValue) {
                f.check();
                if (!RESPECT_TEMPLATE_APPEARANCE) {
                  try { f.updateAppearances(customFont); } catch (_) {}
                }
                filled++;
                log(`✅ Checked TravelerSex checkbox (exportValue: ${exportValue} matches input: ${travelerSexValue})`);
                continue;
              } else {
                f.uncheck();
                if (!RESPECT_TEMPLATE_APPEARANCE) {
                  try { f.updateAppearances(customFont); } catch (_) {}
                }
                log(`❌ Unchecked TravelerSex checkbox (exportValue: ${exportValue} does not match input: ${travelerSexValue})`);
                continue;
              }
            } catch (e) {
              log(`Error handling TravelerSex checkbox export values: ${e.message}`);
            }
            } else {
              // Fallback to simple checkbox logic when export value checking is disabled
              log(`TravelerSex checkbox (fallback mode): fieldName="${n}", inputValue="${travelerSexValue}"`);
            }
            
            // Fallback to old behavior if export value handling fails
            const shouldCheck = (
              // Handle explicit field names (if you rename them)
              (n === 'TravelerSex_男性' && travelerSexValue === 'on') ||
              (n === 'TravelerSex_女性' && travelerSexValue === 'on') ||
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
              if (!RESPECT_TEMPLATE_APPEARANCE && customFont) {
                try { f.updateAppearances(customFont); } catch (_) {}
              }
              filled++;
              log(`✅ Checked TravelerSex ${n} (value: ${travelerSexValue})`);
            } else {
              f.uncheck();
              if (!RESPECT_TEMPLATE_APPEARANCE && customFont) {
                try { f.updateAppearances(customFont); } catch (_) {}
              }
              log(`❌ Unchecked TravelerSex checkbox ${n} (no match for value: ${travelerSexValue})`);
            }
          } else if (isSameAsTravelerField) {
            log(`SameAsTraveler checkbox ${n}: value="${sameAsTravelerValue}"`);
            
            // Debug: Check what value is actually being sent
            log(`SameAsTraveler debug - field name: "${n}", value: "${sameAsTravelerValue}", type: ${typeof sameAsTravelerValue}`);
            
            // Check if this field should be checked based on the value
            const shouldCheck = (
              // Handle various checkbox value formats
              sameAsTravelerValue === 'on' ||
              sameAsTravelerValue === 'true' ||
              sameAsTravelerValue === '1' ||
              sameAsTravelerValue === 'yes' ||
              sameAsTravelerValue === 'はい' ||
              sameAsTravelerValue === 'checked' ||
              sameAsTravelerValue === 'SameAsTraveler' ||  // Sometimes the field name itself is sent
              sameAsTravelerValue === 'on' ||
              sameAsTravelerValue === 'true'
            );
            
            log(`SameAsTraveler shouldCheck: ${shouldCheck} for value: "${sameAsTravelerValue}"`);
            
            if (shouldCheck) {
              // Simple checkbox logic for SameAsTraveler
              f.check();
              if (!RESPECT_TEMPLATE_APPEARANCE && customFont) {
                try { f.updateAppearances(customFont); } catch (_) {}
              }
              filled++;
              log(`✅ Checked SameAsTraveler ${n} (value: ${sameAsTravelerValue})`);
            } else {
              f.uncheck();
              if (!RESPECT_TEMPLATE_APPEARANCE && customFont) {
                try { f.updateAppearances(customFont); } catch (_) {}
              }
              log(`❌ Unchecked SameAsTraveler checkbox ${n} (no match for value: ${sameAsTravelerValue})`);
            }
          } else {
            // Check for yes/no checkbox fields (TreatmentNow_yes, TreatmentNow_no, etc.)
            const yesFieldMatch = n.match(/^(.+)_(yes|はい)$/);
            const noFieldMatch = n.match(/^(.+)_(no|いいえ)$/);
            
            if (yesFieldMatch || noFieldMatch) {
              const match = yesFieldMatch || noFieldMatch;
              const baseName = match[1];
              const fieldType = match[2];
              
              // Check if we have a value for the corresponding yes/no field
              const oppositeType = fieldType.includes('yes') || fieldType.includes('はい') ? 
                (n.replace(/_yes$|_はい$/, '_no').replace(/_yes$/, '_いいえ')) : 
                (n.replace(/_no$|_いいえ$/, '_yes').replace(/_no$/, '_はい'));
              
              const fieldValue = String(single).trim();
              const shouldCheck = fieldValue === 'on' || fieldValue === 'はい' || fieldValue === 'yes' || fieldValue === 'true';
              
              log(`Yes/No checkbox ${n}: baseName="${baseName}", fieldType="${fieldType}", fieldValue="${fieldValue}", shouldCheck=${shouldCheck}`);
              
              if (shouldCheck) {
                f.check();
                if (!RESPECT_TEMPLATE_APPEARANCE && customFont) {
                  try { f.updateAppearances(customFont); } catch (_) {}
                }
                filled++;
                log(`✅ Checked ${n}`);
              } else {
                f.uncheck();
                if (!RESPECT_TEMPLATE_APPEARANCE && customFont) {
                  try { f.updateAppearances(customFont); } catch (_) {}
                }
                log(`❌ Unchecked ${n}`);
            }
          } else {
            // For other non-numeric values, use the existing yes/no logic
            const yn = normalizeYesNo(single);
            if (yn === 'yes' || yn === 'on' || yn === '1' || yn === 'true') f.check();
            else f.uncheck();
            filled++;
            }
          }
        }
        continue;
      }
    }
  }

  // Process checkbox groups - NEW APPROACH: Use explicit field names
  log('===== PROCESSING CHECKBOX GROUPS =====');
  for (const [groupName, checkboxes] of Object.entries(checkboxGroups)) {
    log(`Processing checkbox group: ${groupName} (${checkboxes.length} checkboxes)`);
    
    // Handle explicit checkbox field names (new approach)
    if (groupName === 'CoverageValue_日') {
      const shouldCheck = fields.CoverageValue_日 === 'on';
      log(`CoverageValue_日: shouldCheck=${shouldCheck}`);
      for (const checkbox of checkboxes) {
        if (shouldCheck) {
          checkbox.check();
          if (!RESPECT_TEMPLATE_APPEARANCE) {
            try { checkbox.updateAppearances(customFont); } catch (_) {}
          }
          filled++;
          log(`✅ Checked CoverageValue_日 checkbox`);
        } else {
          checkbox.uncheck();
          if (!RESPECT_TEMPLATE_APPEARANCE) {
            try { checkbox.updateAppearances(customFont); } catch (_) {}
          }
          log(`❌ Unchecked CoverageValue_日 checkbox`);
        }
      }
    } else if (groupName === 'CoverageValue_月') {
      const shouldCheck = fields.CoverageValue_月 === 'on';
      log(`CoverageValue_月: shouldCheck=${shouldCheck}`);
      for (const checkbox of checkboxes) {
        if (shouldCheck) {
          checkbox.check();
          if (!RESPECT_TEMPLATE_APPEARANCE) {
            try { checkbox.updateAppearances(customFont); } catch (_) {}
          }
          filled++;
          log(`✅ Checked CoverageValue_月 checkbox`);
        } else {
          checkbox.uncheck();
          if (!RESPECT_TEMPLATE_APPEARANCE) {
            try { checkbox.updateAppearances(customFont); } catch (_) {}
          }
          log(`❌ Unchecked CoverageValue_月 checkbox`);
        }
      }
    } else if (groupName === 'TravelerSex_男性') {
      const shouldCheck = fields.TravelerSex_男性 === 'on';
      log(`TravelerSex_男性: shouldCheck=${shouldCheck}`);
      for (const checkbox of checkboxes) {
        if (shouldCheck) {
          checkbox.check();
          if (!RESPECT_TEMPLATE_APPEARANCE) {
            try { checkbox.updateAppearances(customFont); } catch (_) {}
          }
          filled++;
          log(`✅ Checked TravelerSex_男性 checkbox`);
        } else {
          checkbox.uncheck();
          if (!RESPECT_TEMPLATE_APPEARANCE) {
            try { checkbox.updateAppearances(customFont); } catch (_) {}
          }
          log(`❌ Unchecked TravelerSex_男性 checkbox`);
        }
      }
    } else if (groupName === 'TravelerSex_女性') {
      const shouldCheck = fields.TravelerSex_女性 === 'on';
      log(`TravelerSex_女性: shouldCheck=${shouldCheck}`);
      for (const checkbox of checkboxes) {
        if (shouldCheck) {
          checkbox.check();
          if (!RESPECT_TEMPLATE_APPEARANCE) {
            try { checkbox.updateAppearances(customFont); } catch (_) {}
          }
          filled++;
          log(`✅ Checked TravelerSex_女性 checkbox`);
        } else {
          checkbox.uncheck();
          if (!RESPECT_TEMPLATE_APPEARANCE) {
            try { checkbox.updateAppearances(customFont); } catch (_) {}
          }
          log(`❌ Unchecked TravelerSex_女性 checkbox`);
        }
      }
    } else if (groupName.startsWith('Emergency') && (groupName.endsWith('Thousand') || groupName.includes('Thousand'))) {
      // Handle Emergency Return checkboxes (TRUE/FALSE strings)
      const fieldValue = String(fields[groupName] || '').toUpperCase();
      const shouldCheck = fieldValue === 'TRUE' || fieldValue === 'YES' || fieldValue === 'ON' || fieldValue === '1';
      log(`${groupName}: shouldCheck=${shouldCheck} (value="${fields[groupName]}")`);
      
      for (const checkbox of checkboxes) {
        if (shouldCheck) {
          checkbox.check();
          if (!RESPECT_TEMPLATE_APPEARANCE) {
            try { checkbox.updateAppearances(customFont); } catch (_) {}
          }
          filled++;
          log(`✅ Checked ${groupName} checkbox`);
        } else {
          checkbox.uncheck();
          if (!RESPECT_TEMPLATE_APPEARANCE) {
            try { checkbox.updateAppearances(customFont); } catch (_) {}
          }
          log(`❌ Unchecked ${groupName} checkbox`);
        }
      }
    } else if (groupName.includes('BirthEra') || groupName === 'ApplicantBirthEra19or20' || groupName === 'TravelerBirthEra19or20') {
      // Handle Birth Era checkboxes (numeric values "19" or "20")
      const fieldValue = String(fields[groupName] || '');
      const shouldCheck = fieldValue === '19' || fieldValue === '20';
      log(`${groupName}: shouldCheck=${shouldCheck} (value="${fieldValue}")`);
      
      for (const checkbox of checkboxes) {
        if (shouldCheck) {
          checkbox.check();
          if (!RESPECT_TEMPLATE_APPEARANCE) {
            try { checkbox.updateAppearances(customFont); } catch (_) {}
          }
          filled++;
          log(`✅ Checked ${groupName} checkbox`);
        } else {
          checkbox.uncheck();
          if (!RESPECT_TEMPLATE_APPEARANCE) {
            try { checkbox.updateAppearances(customFont); } catch (_) {}
          }
          log(`❌ Unchecked ${groupName} checkbox`);
        }
      }
    } else if (groupName === 'Q5_DestinationRegion') {
      // Handle Destination Region - check the corresponding region checkbox
      const regionValue = String(fields[groupName] || '');
      const regionCheckboxName = `DestinationRegion_${regionValue}`;
      
      // We need to find and check the corresponding checkbox in the checkboxGroups
      if (checkboxGroups[regionCheckboxName] && checkboxGroups[regionCheckboxName].length > 0) {
        log(`Q5_DestinationRegion: Found region="${regionValue}", checking ${regionCheckboxName}`);
        for (const checkbox of checkboxGroups[regionCheckboxName]) {
          checkbox.check();
          if (!RESPECT_TEMPLATE_APPEARANCE) {
            try { checkbox.updateAppearances(customFont); } catch (_) {}
          }
          filled++;
          log(`✅ Checked ${regionCheckboxName} checkbox`);
        }
      } else {
        log(`Q5_DestinationRegion: Region checkbox ${regionCheckboxName} not found`);
        // Uncheck all DestinationRegion checkboxes
        for (const checkbox of checkboxes) {
          checkbox.uncheck();
          if (!RESPECT_TEMPLATE_APPEARANCE) {
            try { checkbox.updateAppearances(customFont); } catch (_) {}
          }
        }
        log(`❌ Unchecked Q5_DestinationRegion checkbox (region not found)`);
      }
    } else {
      // Fallback for all other checkbox fields (like the 8 questions)
      const fieldValue = fields[groupName];
      const shouldCheck = fieldValue === 'on' || fieldValue === 'はい' || fieldValue === 'yes' || fieldValue === 'true' || String(fieldValue).toUpperCase() === 'TRUE';
      log(`${groupName}: shouldCheck=${shouldCheck} (value="${fieldValue}")`);
      
      for (const checkbox of checkboxes) {
        if (shouldCheck) {
          checkbox.check();
          if (!RESPECT_TEMPLATE_APPEARANCE) {
            try { checkbox.updateAppearances(customFont); } catch (_) {}
          }
          filled++;
          log(`✅ Checked ${groupName} checkbox`);
        } else {
          checkbox.uncheck();
          if (!RESPECT_TEMPLATE_APPEARANCE) {
            try { checkbox.updateAppearances(customFont); } catch (_) {}
          }
          log(`❌ Unchecked ${groupName} checkbox`);
        }
      }
    }
  }
  log('===== END CHECKBOX GROUPS =====');

  // 4) Burn-in (optional) - Re-enabled with improved text handling for consistency
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

        // Skip burn-in for problematic fields that have auto-sizing in template
        if (name.includes('代理店') || name.includes('Agent') || name.includes('Code') || 
            name.includes('扱者') || name.includes('仲立人') || name.includes('コード') ||
            name.includes('Address') || name.includes('住所') || name.includes('FullAddress')) {
          log(`Skipping burn-in for field "${name}" to preserve template auto-sizing`);
          continue;
        }
        
        // Debug: Log all field names being processed for burn-in
        if (raw && raw.length > 10) {
          log(`Burn-in processing field "${name}" with text: "${raw}"`);
        }

        const widgets = (f.acroField && f.acroField.getWidgets) ? f.acroField.getWidgets() : [];
        for (const w of widgets) {
          const page = w.getPage && w.getPage();
          if (!page) continue;
          const rect = w.getRectangle && w.getRectangle();
          if (!rect) continue;
          const text = String(raw);
          const padding = 1; // Reduced padding to give more space
          let size = Math.min(10, rect.height - 2 * padding); // Start with smaller size
          const maxWidth = rect.width - 2 * padding;
          
          // Improved text fitting algorithm for consistency
          let displayText = text;
          let textWidth = customFont.widthOfTextAtSize(text, size);
          
          // More conservative font size reduction to prevent cutoff
          while (size > 8 && textWidth > maxWidth) {
            size -= 0.25; // Smaller increments for better precision
            textWidth = customFont.widthOfTextAtSize(text, size);
          }
          
          // If still too wide, use smaller font size instead of truncation
          if (textWidth > maxWidth) {
            size = Math.max(6, size - 1); // Ensure minimum readable size
            displayText = text; // Keep full text, don't truncate
          }

          // Debug logging for problematic fields
          if (name.includes('代理店') || name.includes('Agent') || name.includes('Code')) {
            log(`Burn-in field "${name}": original="${text}", display="${displayText}", size=${size}, rect=${rect.width}x${rect.height}`);
          }

          page.drawText(displayText, {
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
  } else {
    log('Burn-in disabled to prevent text cutoff issues - using template auto-sizing');
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

  // 6) save - Enhanced for consistent rendering across all devices
  let outBytes;
  try {
    // Enhanced save options for maximum consistency and font embedding
    log('Saving PDF with enhanced settings for cross-device consistency...');
    outBytes = await pdfDoc.save({
      useObjectStreams: false,
      addDefaultPage: false,
      updateFieldAppearances: false, // Preserve template auto-sizing - only fields we explicitly updated will have new appearances
      objectsPerTick: 50,
      // Additional options for consistency and font embedding
      compress: true,
      // Force font subsetting for better compatibility
      subset: true,
    });
    log('PDF saved successfully with enhanced consistency settings');
  } catch (e) {
    log('Enhanced save failed, trying standard options:', e.message);
    try {
      // Standard save options
      outBytes = await pdfDoc.save({
        useObjectStreams: false,
        addDefaultPage: false,
        updateFieldAppearances: false, // Preserve template auto-sizing
      });
      log('PDF saved with standard options');
    } catch (e2) {
      log('Standard save failed, trying minimal options:', e2.message);
      try {
        // Minimal save options as last resort
    outBytes = await pdfDoc.save();
        log('PDF saved with minimal options');
      } catch (e3) {
        log('All save attempts failed:', e3.message);
        throw e3;
      }
    }
  }

  fs.writeFileSync(outPath, outBytes);
  return { outPath, filled, size: outBytes.length };
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
    log('===== PDF FILLER SERVICE REQUEST RECEIVED =====');
    log('Request body keys:', Object.keys(req.body || {}));
    
    const { templateFileId, fields, outputName, folderId, mode, watermarkText } = req.body || {};
    if (!templateFileId) {
      log('ERROR: templateFileId is required');
      return res.status(400).json({ error: 'templateFileId is required' });
    }

    log('Starting PDF fill process...');
    log('Template file ID:', templateFileId);
    log('Output name:', outputName);
    log('Folder ID:', folderId);

    const tmpTemplate = path.join(TMP, `template_${templateFileId}.pdf`);
    log('Downloading template file...');
    await downloadDriveFile(templateFileId, tmpTemplate);
    log('Template file downloaded successfully');

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
    log('Starting PDF fill process with watermark:', wm);
    const result = await fillPdf(tmpTemplate, outPath, fields || {}, { watermarkText: wm });
    log(`PDF fill process completed successfully`);
    log(`Filled PDF -> ${result.outPath} (${result.size} bytes, fields filled: ${result.filled})`);

    log('Starting upload to Google Drive...');
    const uploaded = await uploadToDrive(result.outPath, outName, folderId);
    log('Upload to Google Drive completed successfully');
    log('Uploaded to Drive:', uploaded);

    log('===== PDF FILLER SERVICE RESPONSE =====');
    const response = { ok: true, filledCount: result.filled, driveFile: uploaded, webViewLink: uploaded.webViewLink };
    log('Response:', JSON.stringify(response));
    log('===== END RESPONSE =====');

    res.json(response);
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
