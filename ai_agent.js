// ============================================================
//  ICF-SL  School-Based ITN Distribution — Google Apps Script
//  Handles:
//    1. Form data submission → Google Sheet (human-readable headers)
//    2. Claude AI agent proxy (Anthropic API)
//    3. getData action for Analysis dashboard
// ============================================================

const CONFIG = {
  SHEET_ID:          '1cXlYiTMzcRP1BCj9mt1JXoK_pjgWbRtDEEQUPMg2HPs',
  ANTHROPIC_API_KEY: 'PASTE_YOUR_KEY_HERE',   // ← replace this with your real sk-ant-api03-... key
  ANTHROPIC_MODEL:   'claude-sonnet-4-20250514',
  SHEET_NAME:        'Submissions',            // alias used by getColumns route
  SHEET_NAME_DATA:   'Submissions',
  SHEET_NAME_ITN:    'ITN Movement',
  SHEET_NAME_PHU:    'PHU Receipts',
  SHEET_NAME_LOG:    'AI_Log',
  SHEET_NAME_ATT:    'Attendance & Payment',
  SHEET_NAME_SBD:    'SBD_Records',
  SHEET_NAME_DTAG:   'Device Tags',
  SHEET_NAME_DTRACK: 'Device Tracking',
  MAX_TOKENS:        1200
};

// ── FIELD MAP: variable name → human-readable label ──────────
// Order here = column order in the sheet
const FIELD_MAP = [
  // Meta
  // Location
  { field: 'district',            label: 'District'                     },
  { field: 'chiefdom',            label: 'Chiefdom'                     },
  { field: 'facility',            label: 'Health Facility (PHU)'        },
  { field: 'community',           label: 'Community / Village'          },
  { field: 'school_name',         label: 'School Name'                  },
  { field: 'school_enrollment',   label: 'School Enrollment'            },
  { field: 'emis_number',         label: 'EMIS Number'                  },
  { field: 'school_status',       label: 'School Status'                },
  { field: 'is_new_school',       label: 'New School?'                  },
  // School profile
  { field: 'head_teacher',        label: 'Head Teacher Name'            },
  { field: 'head_teacher_phone',  label: 'Head Teacher Phone'           },
  { field: 'distribution_date',   label: 'Distribution Date'            },
  // ITNs received
  { field: 'itns_received',       label: 'Total ITNs Received'          },
  { field: 'itn_type',           label: 'ITN Type'                     },
  // Class 1
  { field: 'c1_teacher_name',     label: 'Class 1 Teacher Name'         },
  { field: 'c1_teacher_phone',    label: 'Class 1 Teacher Phone'        },
  { field: 'c1_boys',             label: 'Class 1 — Boys Enrolled'      },
  { field: 'c1_boys_itn',         label: 'Class 1 — Boys Received ITN'  },
  { field: 'c1_girls',            label: 'Class 1 — Girls Enrolled'     },
  { field: 'c1_girls_itn',        label: 'Class 1 — Girls Received ITN' },
  // Class 2
  { field: 'c2_teacher_name',     label: 'Class 2 Teacher Name'         },
  { field: 'c2_teacher_phone',    label: 'Class 2 Teacher Phone'        },
  { field: 'c2_boys',             label: 'Class 2 — Boys Enrolled'      },
  { field: 'c2_boys_itn',         label: 'Class 2 — Boys Received ITN'  },
  { field: 'c2_girls',            label: 'Class 2 — Girls Enrolled'     },
  { field: 'c2_girls_itn',        label: 'Class 2 — Girls Received ITN' },
  // Class 3
  { field: 'c3_teacher_name',     label: 'Class 3 Teacher Name'         },
  { field: 'c3_teacher_phone',    label: 'Class 3 Teacher Phone'        },
  { field: 'c3_boys',             label: 'Class 3 — Boys Enrolled'      },
  { field: 'c3_boys_itn',         label: 'Class 3 — Boys Received ITN'  },
  { field: 'c3_girls',            label: 'Class 3 — Girls Enrolled'     },
  { field: 'c3_girls_itn',        label: 'Class 3 — Girls Received ITN' },
  // Class 4
  { field: 'c4_teacher_name',     label: 'Class 4 Teacher Name'         },
  { field: 'c4_teacher_phone',    label: 'Class 4 Teacher Phone'        },
  { field: 'c4_boys',             label: 'Class 4 — Boys Enrolled'      },
  { field: 'c4_boys_itn',         label: 'Class 4 — Boys Received ITN'  },
  { field: 'c4_girls',            label: 'Class 4 — Girls Enrolled'     },
  { field: 'c4_girls_itn',        label: 'Class 4 — Girls Received ITN' },
  // Class 5
  { field: 'c5_teacher_name',     label: 'Class 5 Teacher Name'         },
  { field: 'c5_teacher_phone',    label: 'Class 5 Teacher Phone'        },
  { field: 'c5_boys',             label: 'Class 5 — Boys Enrolled'      },
  { field: 'c5_boys_itn',         label: 'Class 5 — Boys Received ITN'  },
  { field: 'c5_girls',            label: 'Class 5 — Girls Enrolled'     },
  { field: 'c5_girls_itn',        label: 'Class 5 — Girls Received ITN' },
  // Computed totals
  { field: 'total_boys',          label: 'Total Boys Enrolled'          },
  { field: 'total_girls',         label: 'Total Girls Enrolled'         },
  { field: 'total_pupils',        label: 'Total Pupils Enrolled'        },
  { field: 'total_boys_itn',      label: 'Total Boys Received ITN'      },
  { field: 'total_girls_itn',     label: 'Total Girls Received ITN'     },
  { field: 'total_itn',           label: 'Total ITNs Distributed'       },
  { field: 'itns_remaining',      label: 'ITNs Remaining'               },
  { field: 'prop_boys',           label: 'Proportion Boys (%)'          },
  { field: 'prop_girls',          label: 'Proportion Girls (%)'         },
  { field: 'coverage_boys',       label: 'Boys ITN Coverage (%)'        },
  { field: 'coverage_girls',      label: 'Girls ITN Coverage (%)'       },
  { field: 'coverage_total',      label: 'Overall ITN Coverage (%)'     },
  // Team & GPS
  { field: 'survey_date',         label: 'Survey Date'                  },
  { field: 'gps_lat',             label: 'GPS Latitude'                 },
  { field: 'gps_lng',             label: 'GPS Longitude'                },
  { field: 'gps_acc',             label: 'GPS Accuracy (m)'             },
  { field: 'team1_name',          label: 'Health Staff Name'            },
  { field: 'team1_phone',         label: 'Health Staff Phone'           },
  { field: 'team1_signature',     label: 'Health Staff Signed?'         },
  { field: 'team2_name',          label: 'Teacher Name'                 },
  { field: 'team2_phone',         label: 'Teacher Phone'                },
  { field: 'team2_signature',     label: 'Teacher Signed?'              }
];

// ── ENTRY POINTS ─────────────────────────────────────────────
function doGet(e) {
  const action = (e && e.parameter && e.parameter.action) || '';

  if (action === 'ping') {
    return jsonResponse({ status: 'ok', message: 'ICF-SL GAS Backend live' });
  }

  if (action === 'checkAttDuplicate') {
    const name  = (e.parameter.name  || '').trim().toLowerCase();
    const phone = (e.parameter.phone || '').trim();
    const date  = (e.parameter.date  || '').trim();
    return checkAttDuplicate(name, phone, date);
  }

  if (action === 'getColumns') {
    return getSheetColumns(e.parameter.sheet || CONFIG.SHEET_NAME);
  }
  if (action === 'count') {
    return jsonResponse({ count: getSubmissionCount() });
  }
  if (action === 'getData') {
    // Try cache first (60s TTL) — avoids slow sheet reads on repeat fetches
    try {
      const cache = CacheService.getScriptCache();
      const cached = cache.get('submissions_data');
      if (cached) {
        return jsonResponse({ success: true, rows: JSON.parse(cached), cached: true });
      }
    } catch(e) {}
    const rows = getAllSubmissionsAsObjects();
    // Cache for 60 seconds
    try {
      const cache = CacheService.getScriptCache();
      const json = JSON.stringify(rows);
      if (json.length < 100000) cache.put('submissions_data', json, 60);
    } catch(e) {}
    return jsonResponse({ success: true, rows: rows });
  }
  if (action === 'getAllSubmissions') {
    const rows = getAllSubmissionsAsObjects();
    return jsonResponse({ status: 'ok', submissions: rows, count: rows.length });
  }
  if (action === 'checkDuplicate') {
    return jsonResponse(checkDuplicateInSheet(e.parameter));
  }
  if (action === 'checkPHUDispatch') {
    return jsonResponse(checkPHUDispatch(e.parameter));
  }
  if (action === 'getPHUReceipts') {
    return jsonResponse(getPHUReceipts());
  }
  if (action === 'getDispatch') {
    return jsonResponse(getDispatchById(e.parameter.id || ''));
  }
  if (action === 'getAllDispatches') {
    return jsonResponse(getAllDispatches());
  }
  if (action === 'getAllReceipts') {
    return jsonResponse(getAllReceipts());
  }
  if (action === 'getDeviceTracking') {
    return jsonResponse(getDeviceTrackingRecords());
  }
  return jsonResponse({ status: 'ok' });
}

function doPost(e) {
  try {
    let body = {};
    if (e && e.parameter && e.parameter.payload) {
      body = JSON.parse(e.parameter.payload);
    } else if (e && e.postData && e.postData.contents) {
      body = JSON.parse(e.postData.contents);
    }

    const action = body.action || 'submit';
    if (action === 'submit')           return handleSubmission(body);
    if (action === 'itn_movement')     return handleITNMovement(body);
    if (action === 'ai_query')         return handleAIQuery(body);
    if (action === 'savePHUReceipt')   return handlePHUReceipt(body);
    if (action === 'saveAssessmentGrade') return handleAssessmentGrade(body);
    if (action === 'saveMonitoring')       return handleMonitoring(body);
    if (action === 'id_card')               return handleIDCardSubmission(body);
    if (action === 'device_tag')             return handleDeviceTagSubmission(body);
    if (action === 'device_register')        return handleDeviceTracking(body);
    if (action === 'device_return')          return handleDeviceTracking(body);

    // Attendance — app sends action2 field
    const action2 = body.action2 || '';
    if (action2 === 'handleSignInOut') return handleAttendance(body);

    return jsonResponse({ success: false, error: 'Unknown action: ' + action });
  } catch (err) {
    return jsonResponse({ success: false, error: err.message });
  }
}

function jsonResponse(data) {
  return ContentService
    .createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
}

// ── ITN MOVEMENT SUBMISSION ───────────────────────────────────
// Handles records from itn_movement.html (DMS to PHU dispatches)
function handleITNMovement(data) {
  try {
    const ss  = SpreadsheetApp.openById(CONFIG.SHEET_ID);
    let sheet = ss.getSheetByName(CONFIG.SHEET_NAME_ITN);

    if (!sheet) {
      sheet = ss.insertSheet(CONFIG.SHEET_NAME_ITN);
      const headers = [
        'Dispatch ID', 'Timestamp',
        'Staff Name', 'Staff Username', 'Staff District',
        'Destination District', 'Chiefdom', 'Health Facility (PHU)',
        'Dual-AI ITNs',
        'Conveyor Name', 'Conveyor Username', 'Conveyor Phone',
        'Vehicle Plate', 'Status'
      ];
      sheet.appendRow(headers);
      const hdr = sheet.getRange(1, 1, 1, headers.length);
      hdr.setBackground('#004080').setFontColor('#ffffff')
         .setFontWeight('bold').setFontFamily('Arial').setFontSize(10);
      sheet.setFrozenRows(1);
      sheet.setColumnWidths(1, headers.length, 160);
    }

    const row = [
      data.dispatch_id           || '',
      data.timestamp             || new Date().toISOString(),
      data.staff_name            || '',
      data.staff_username        || '',
      data.staff_district        || '',
      data.destination_district  || '',
      data.chiefdom              || '',
      data.phu                   || '',
      parseInt(data.total_qty || data.dualai_qty) || 0,
      data.driver_name           || data.conveyor_name   || '',
      data.driver_username       || data.conveyor_username|| '',
      data.driver_phone          || data.conveyor_phone  || '',
      data.vehicle               || '',
      data.status                || 'dispatched'
    ];

    sheet.appendRow(row);
    const lastRow = sheet.getLastRow();
    if (lastRow % 2 === 0) {
      sheet.getRange(lastRow, 1, 1, row.length).setBackground('#f0f6ff');
    }
    sheet.autoResizeColumns(1, row.length);
    return jsonResponse({ success: true, message: 'ITN movement recorded', row: lastRow });

  } catch(err) {
    return jsonResponse({ success: false, error: err.message });
  }
}

// ── FORM DATA SUBMISSION ───────────────────────────────────────
function handleSubmission(data) {
  try {
    const ss  = SpreadsheetApp.openById(CONFIG.SHEET_ID);
    let sheet = ss.getSheetByName(CONFIG.SHEET_NAME_DATA);

    if (!sheet) {
      sheet = ss.insertSheet(CONFIG.SHEET_NAME_DATA);
      writeHeaders(sheet);
    } else if (sheet.getLastRow() === 0) {
      writeHeaders(sheet);
    }

    // Map field values in FIELD_MAP order
    const row = FIELD_MAP.map(({ field }) => {
      // Try field name first, then _val suffix variant (form uses both patterns)
      const val = data[field] !== undefined ? data[field]
                : data[field + '_val'] !== undefined ? data[field + '_val']
                : '';
      // Store base64 signatures as YES/NO flag (saves space)
      if (field.includes('signature') && val && String(val).length > 100) return 'YES';
      // Computed totals: if empty, try recalculating from class data
      return val !== undefined ? val : '';
    });

    sheet.appendRow(row);

    try {
      sheet.autoResizeColumns(1, Math.min(FIELD_MAP.length, 50));
    } catch(e) {}

    // Clear cache so next getData fetch returns fresh data
    try { CacheService.getScriptCache().remove('submissions_data'); } catch(e) {}
    return jsonResponse({
      success: true,
      message: 'Submission saved',
      row:     sheet.getLastRow()
    });
  } catch (err) {
    return jsonResponse({ success: false, error: err.message });
  }
}

function writeHeaders(sheet) {
  const labels = FIELD_MAP.map(f => f.label);
  sheet.appendRow(labels);

  const headerRange = sheet.getRange(1, 1, 1, labels.length);
  headerRange.setBackground('#004080');
  headerRange.setFontColor('#ffffff');
  headerRange.setFontWeight('bold');
  headerRange.setFontFamily('Arial');
  headerRange.setFontSize(10);
  sheet.setFrozenRows(1);

  // Extra wide for label columns
  sheet.setColumnWidth(1, 180);  // Submission Date
  sheet.setColumnWidth(9, 200);  // School Name
  sheet.setColumnWidth(10, 180); // Head Teacher
  sheet.setColumnWidth(7, 200);  // PHU
}

function getSubmissionCount() {
  try {
    const ss    = SpreadsheetApp.openById(CONFIG.SHEET_ID);
    const sheet = ss.getSheetByName(CONFIG.SHEET_NAME_DATA);
    if (!sheet) return 0;
    return Math.max(0, sheet.getLastRow() - 1);
  } catch(e) { return 0; }
}

// ── checkDuplicate: returns {exists, timestamp, submitted_by} ──
// Called by the browser before allowing submission.
// Matches on ALL 6 geo-hierarchy fields (case-insensitive).

// ── getAllDispatches — all ITN Movement rows ─────────────────
function getAllDispatches() {
  try {
    const ss    = SpreadsheetApp.openById(CONFIG.SHEET_ID);
    const sheet = ss.getSheetByName(CONFIG.SHEET_NAME_ITN);
    if (!sheet || sheet.getLastRow() <= 1) return [];
    const data    = sheet.getDataRange().getValues();
    const headers = data[0].map(h => String(h||'').trim().toLowerCase());

    const iDistrict = headers.indexOf('destination district');
    const iChiefdom = headers.indexOf('chiefdom');
    const iPHU      = headers.indexOf('health facility (phu)');
    const iDispId   = headers.indexOf('dispatch id');
    const iDualAI   = headers.indexOf('dual-ai itns');
    const iStatus   = headers.indexOf('status');

    return data.slice(1).filter(r => r[iPHU]).map(r => ({
      dispatch_id: String(r[iDispId]  || ''),
      district:    String(r[iDistrict]|| '').trim(),
      chiefdom:    String(r[iChiefdom]|| '').trim(),
      phu:         String(r[iPHU]     || '').trim(),
      dual_ai_qty: r[iDualAI] || 0,
      status:      String(r[iStatus]  || '')
    }));
  } catch(e) { return []; }
}

// ── getAllReceipts — all PHU Receipts rows ───────────────────
function getAllReceipts() {
  try {
    const ss    = SpreadsheetApp.openById(CONFIG.SHEET_ID);
    const sheet = ss.getSheetByName(CONFIG.SHEET_NAME_PHU);
    if (!sheet || sheet.getLastRow() <= 1) return [];
    const data    = sheet.getDataRange().getValues();
    const headers = data[0].map(h => String(h||'').trim().toLowerCase());

    const iDistrict = headers.indexOf('district');
    const iChiefdom = headers.indexOf('chiefdom');
    const iPHU      = headers.indexOf('phu');
    const iRecId    = headers.indexOf('receipt id');
    const iDispId   = headers.indexOf('dispatch id');
    const iReceived = headers.indexOf('received dual-ai itns');

    return data.slice(1).filter(r => r[iPHU]).map(r => ({
      receipt_id:  String(r[iRecId]   || ''),
      dispatch_id: String(r[iDispId]  || ''),
      district:    String(r[iDistrict]|| '').trim(),
      chiefdom:    String(r[iChiefdom]|| '').trim(),
      phu:         String(r[iPHU]     || '').trim(),
      received_qty:r[iReceived] || 0
    }));
  } catch(e) { return []; }
}

// ── CHECK PHU DISPATCH (ITN Movement sheet) ─────────────────────────────────
// Called via GET ?action=checkPHUDispatch&district=X&chiefdom=Y&phu=Z
// Returns { found: true/false, ...dispatch fields } 
function checkPHUDispatch(params) {
  try {
    const district = String(params.district || '').trim().toLowerCase();
    const chiefdom = String(params.chiefdom || '').trim().toLowerCase();
    const phu      = String(params.phu      || '').trim().toLowerCase();

    if (!district || !chiefdom || !phu) return { found: false };

    const ss    = SpreadsheetApp.openById(CONFIG.SHEET_ID);
    const sheet = ss.getSheetByName(CONFIG.SHEET_NAME_ITN);
    if (!sheet || sheet.getLastRow() <= 1) return { found: false };

    const data    = sheet.getDataRange().getValues();
    const headers = data[0].map(h => String(h).trim().toLowerCase());

    // Find column positions by header name
    const col = name => headers.indexOf(name);
    const iDispatchId = col('dispatch id');
    const iTimestamp  = col('timestamp');
    const iStaff      = col('dms staff name');
    const iDistrict   = col('destination district');
    const iChiefdom   = col('chiefdom');
    const iPHU        = col('health facility (phu)');
    const iDualAI     = col('dual-ai itns');
    const iTotal      = col('total itns');
    const iDriver     = col('driver name');
    const iVehicle    = col('vehicle plate');

    const lc = s => String(s || '').trim().toLowerCase();

    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      if (lc(row[iDistrict]) === district &&
          lc(row[iChiefdom]) === chiefdom &&
          lc(row[iPHU])      === phu) {
        return {
          found:        true,
          dispatch_id:  row[iDispatchId] || '',
          timestamp:    row[iTimestamp]  ? new Date(row[iTimestamp]).toISOString() : '',
          staff_name:   row[iStaff]      || '',
          chiefdom:     row[iChiefdom]   || '',
          phu:          row[iPHU]        || '',
          dual_ai_qty:  row[iDualAI]     || 0,
          total_qty:    row[iTotal]      || 0,
          driver_name:  row[iDriver]     || '',
          vehicle:      row[iVehicle]    || ''
        };
      }
    }
    return { found: false };

  } catch(err) {
    return { found: false, error: err.message };
  }
}

function checkDuplicateInSheet(params) {
  try {
    const ss    = SpreadsheetApp.openById(CONFIG.SHEET_ID);
    const sheet = ss.getSheetByName(CONFIG.SHEET_NAME_DATA);
    if (!sheet || sheet.getLastRow() <= 1) return { exists: false };

    const data   = sheet.getDataRange().getValues();
    const labels = data[0].map(h => String(h).trim());

    // Find column indices for the four hierarchy fields (no section)
    const fieldCols = {
      district:  FIELD_MAP.findIndex(f => f.field === 'district'),
      chiefdom:  FIELD_MAP.findIndex(f => f.field === 'chiefdom'),
      facility:  FIELD_MAP.findIndex(f => f.field === 'facility'),
      community: FIELD_MAP.findIndex(f => f.field === 'community'),
      school:    FIELD_MAP.findIndex(f => f.field === 'school_name')
    };

    // Map FIELD_MAP index → actual sheet column index (via label matching)
    function sheetCol(fieldIdx) {
      if (fieldIdx < 0) return -1;
      const lbl = FIELD_MAP[fieldIdx].label;
      return labels.indexOf(lbl);
    }

    const colD  = sheetCol(fieldCols.district);
    const colC  = sheetCol(fieldCols.chiefdom);
    const colF  = sheetCol(fieldCols.facility);
    const colCo = sheetCol(fieldCols.community);
    const colSc = sheetCol(fieldCols.school);
    const colTs = sheetCol(FIELD_MAP.findIndex(f => f.field === 'timestamp'));
    const colBy = sheetCol(FIELD_MAP.findIndex(f => f.field === 'submitted_by'));

    const lc = s => String(s||'').trim().toLowerCase();

    const qD  = lc(params.district);
    const qC  = lc(params.chiefdom);
    const qF  = lc(params.facility);
    const qCo = lc(params.community);
    const qSc = lc(params.school);

    const rows = data.slice(1);
    for (let i = 0; i < rows.length; i++) {
      const row = rows[i];
      if (
        lc(row[colD])  === qD  &&
        lc(row[colC])  === qC  &&
        lc(row[colF])  === qF  &&
        lc(row[colCo]) === qCo &&
        lc(row[colSc]) === qSc
      ) {
        return {
          exists:       true,
          timestamp:    colTs >= 0 ? String(row[colTs]) : '',
          submitted_by: colBy >= 0 ? String(row[colBy]) : ''
        };
      }
    }

    return { exists: false };

  } catch(e) {
    Logger.log('checkDuplicateInSheet error: ' + e.message);
    return { exists: false, error: e.message };
  }
}

// ── getData: returns rows keyed by FIELD names (not labels) ──
// This lets the browser analysis dashboard work with field keys.
function getAllSubmissionsAsObjects() {
  try {
    const ss    = SpreadsheetApp.openById(CONFIG.SHEET_ID);
    const sheet = ss.getSheetByName(CONFIG.SHEET_NAME_DATA);
    if (!sheet || sheet.getLastRow() <= 1) return [];

    const data    = sheet.getDataRange().getValues();
    const labels  = data[0].map(h => String(h).trim());
    const rows    = data.slice(1).filter(r => r.some(c => c !== ''));

    // Build a label→field reverse map
    const labelToField = {};
    FIELD_MAP.forEach(({ field, label }) => { labelToField[label] = field; });

    return rows.map(row => {
      const obj = {};
      labels.forEach((lbl, i) => {
        const fieldName = labelToField[lbl] || lbl; // fallback to label if not found
        obj[fieldName] = row[i] !== undefined ? String(row[i]) : '';
      });
      return obj;
    });
  } catch(e) {
    Logger.log('getAllSubmissionsAsObjects error: ' + e.message);
    return [];
  }
}

// ── 2. AI AGENT PROXY ─────────────────────────────────────────
function handleAIQuery(body) {
  try {
    const userMessage    = body.message    || '';
    const history        = body.history    || [];
    const sessionContext = body.context    || '';

    if (!userMessage) {
      return jsonResponse({ success: false, error: 'No message provided' });
    }
    if (!CONFIG.ANTHROPIC_API_KEY || CONFIG.ANTHROPIC_API_KEY.trim() === '') {
      return jsonResponse({ success: false, error: 'Anthropic API key not configured in GAS. Add your key to CONFIG.ANTHROPIC_API_KEY.' });
    }

    const sheetContext = getAllSubmissionsAsText();
    const dataContext  = sheetContext +
      (sessionContext && sessionContext.trim().length > 50
        ? '\n\n=== ADDITIONAL SESSION DATA (not yet synced to sheet) ===\n' + sessionContext
        : '');

    const messages = [];
    history.forEach(h => {
      if (h.role && h.content) messages.push({ role: h.role, content: h.content });
    });
    messages.push({ role: 'user', content: userMessage });

    const payload = {
      model:      CONFIG.ANTHROPIC_MODEL,
      max_tokens: CONFIG.MAX_TOKENS,
      system:     buildSystemPrompt(dataContext),
      messages:   messages
    };

    const options = {
      method:      'post',
      contentType: 'application/json',
      headers: {
        'x-api-key':         CONFIG.ANTHROPIC_API_KEY,
        'anthropic-version': '2023-06-01'
      },
      payload:            JSON.stringify(payload),
      muteHttpExceptions: true
    };

    const response     = UrlFetchApp.fetch('https://api.anthropic.com/v1/messages', options);
    const responseCode = response.getResponseCode();
    const responseText = response.getContentText();

    if (responseCode !== 200) {
      let errMsg = 'Anthropic API error (' + responseCode + ')';
      try { errMsg = JSON.parse(responseText).error?.message || errMsg; } catch(e) {}
      return jsonResponse({ success: false, error: errMsg });
    }

    const result = JSON.parse(responseText);
    const reply  = (result.content || [])
      .filter(b => b.type === 'text')
      .map(b => b.text)
      .join('');

    logAIQuery(userMessage, reply);

    return jsonResponse({ success: true, reply, usage: result.usage || {} });

  } catch (err) {
    return jsonResponse({ success: false, error: err.message });
  }
}

function buildSystemPrompt(dataContext) {
  return `You are the ICF Data Agent, an expert AI assistant embedded inside the ICF-SL School-Based ITN Distribution PWA for Sierra Leone, built by Informatics Consultancy Firm Sierra Leone (ICF-SL).

You have DIRECT ACCESS to all distribution data from the Google Sheet. The full dataset is provided below.

Your role:
- Answer any question about the distribution data quickly and accurately
- Perform calculations: totals, averages, coverage rates, rankings, comparisons, gender disaggregation
- Highlight insights, anomalies, or patterns
- Be concise but thorough — use bullet points and **bold** for key figures
- Respond in the same language the user writes in
- Only use data that is explicitly provided — do not invent figures

${dataContext ? 'LIVE DATA SNAPSHOT:\n' + dataContext : 'NOTE: No submissions in the sheet yet. Inform the user.'}`;
}

function logAIQuery(question, answer) {
  try {
    const ss    = SpreadsheetApp.openById(CONFIG.SHEET_ID);
    let sheet   = ss.getSheetByName(CONFIG.SHEET_NAME_LOG);
    if (!sheet) {
      sheet = ss.insertSheet(CONFIG.SHEET_NAME_LOG);
      sheet.appendRow(['Timestamp', 'Question', 'Answer (truncated)']);
      const hdr = sheet.getRange(1, 1, 1, 3);
      hdr.setBackground('#1a1a2e'); hdr.setFontColor('#fff'); hdr.setFontWeight('bold');
      sheet.setFrozenRows(1);
    }
    sheet.appendRow([
      new Date().toISOString(),
      question.substring(0, 500),
      answer.substring(0, 1000)
    ]);
  } catch(e) { /* non-critical */ }
}

// ── 3. ALL SUBMISSIONS AS TEXT (for AI context) ──────────────
function getAllSubmissionsAsText() {
  try {
    const ss    = SpreadsheetApp.openById(CONFIG.SHEET_ID);
    const sheet = ss.getSheetByName(CONFIG.SHEET_NAME_DATA);
    if (!sheet || sheet.getLastRow() <= 1) return 'No submissions in the Google Sheet yet.';

    const data   = sheet.getDataRange().getValues();
    const labels = data[0].map(h => String(h).trim());
    const rows   = data.slice(1).filter(r => r.some(c => c !== ''));
    if (!rows.length) return 'No submissions yet.';

    // Build label→field reverse map for lookup
    const labelToField = {};
    FIELD_MAP.forEach(({ field, label }) => { labelToField[label] = field; });

    function col(r, fieldName) {
      // find by field name via label
      const lbl = FIELD_MAP.find(f => f.field === fieldName)?.label;
      if (!lbl) return '';
      const i = labels.indexOf(lbl);
      return i >= 0 ? String(r[i]).trim() : '';
    }
    function num(r, fieldName) { return parseInt(col(r, fieldName)) || 0; }

    let totalPupils=0, totalITN=0, totalBoys=0, totalGirls=0,
        totalBoysITN=0, totalGirlsITN=0, totalReceived=0, totalRemaining=0;
    const byDistrict={}, byChiefdom={}, bySubmitter={};
    const schoolLines=[];

    rows.forEach((row, idx) => {
      const tp  = num(row,'total_pupils'),    ti  = num(row,'total_itn');
      const tb  = num(row,'total_boys'),      tg  = num(row,'total_girls');
      const tbi = num(row,'total_boys_itn'),  tgi = num(row,'total_girls_itn');
      const rec = num(row,'itns_received'),   rem = num(row,'itns_remaining');
      const cov = num(row,'coverage_total');
      const dist   = col(row,'district')    || 'Unknown';
      const chief  = col(row,'chiefdom')    || 'Unknown';
      const subBy  = col(row,'submitted_by')|| 'Unknown';
      const school = col(row,'school_name') || '—';
      const comm   = col(row,'community')   || '—';
      const sDate  = col(row,'distribution_date') || '—';
      const itnTypes = [
        col(row,'itn_type')||'Dual-AI'
      ].filter(Boolean).join(',') || '—';

      totalPupils   += tp; totalITN      += ti;
      totalBoys     += tb; totalGirls    += tg;
      totalBoysITN  += tbi; totalGirlsITN+= tgi;
      totalReceived += rec; totalRemaining+= rem;

      if (!byDistrict[dist]) byDistrict[dist] = {schools:0,pupils:0,itn:0,received:0};
      byDistrict[dist].schools++; byDistrict[dist].pupils+=tp;
      byDistrict[dist].itn+=ti;  byDistrict[dist].received+=rec;

      const ck = dist+'/'+chief;
      if (!byChiefdom[ck]) byChiefdom[ck] = {schools:0,pupils:0,itn:0};
      byChiefdom[ck].schools++; byChiefdom[ck].pupils+=tp; byChiefdom[ck].itn+=ti;

      if (!bySubmitter[subBy]) bySubmitter[subBy] = {count:0,pupils:0,itn:0};
      bySubmitter[subBy].count++; bySubmitter[subBy].pupils+=tp; bySubmitter[subBy].itn+=ti;

      const cLine = [1,2,3,4,5].map(c => {
        const cb =num(row,'c'+c+'_boys'),   cg =num(row,'c'+c+'_girls');
        const cbi=num(row,'c'+c+'_boys_itn'),cgi=num(row,'c'+c+'_girls_itn');
        return 'C'+c+':'+cb+'B/'+cg+'G('+cbi+'/'+cgi+'ITN)';
      }).join(' ');

      schoolLines.push(
        '[' + (idx+1) + '] ' + school + ' | ' + comm + ', ' + chief + ', ' + dist + '\n' +
        '    Date:'+sDate+' | SubmittedBy:'+subBy+' | Types:'+itnTypes+'\n' +
        '    Pupils:'+tp+'('+tb+'B/'+tg+'G) | Received:'+rec+' | Distributed:'+ti+' | Remaining:'+rem+' | Coverage:'+cov+'%\n' +
        '    BoysITN:'+tbi+'('+num(row,'coverage_boys')+'%) GirlsITN:'+tgi+'('+num(row,'coverage_girls')+'%)\n' +
        '    ' + cLine
      );
    });

    const ov  = totalPupils>0 ? Math.round((totalITN/totalPupils)*100) : 0;
    const bc  = totalBoys>0   ? Math.round((totalBoysITN/totalBoys)*100) : 0;
    const gc  = totalGirls>0  ? Math.round((totalGirlsITN/totalGirls)*100) : 0;
    const avg = rows.length>0 ? Math.round(totalPupils/rows.length) : 0;

    let out = '=== ICF-SL ITN DISTRIBUTION — GOOGLE SHEET DATA (' + new Date().toLocaleDateString() + ') ===\n\n';
    out += 'TOTALS\n';
    out += '  Schools submitted : ' + rows.length + '\n';
    out += '  Avg enrollment    : ' + avg + ' pupils/school\n';
    out += '  Total pupils      : ' + totalPupils + ' (Boys:' + totalBoys + ' Girls:' + totalGirls + ')\n';
    out += '  ITNs received     : ' + totalReceived + '\n';
    out += '  ITNs distributed  : ' + totalITN + '\n';
    out += '  ITNs remaining    : ' + totalRemaining + '\n';
    out += '  Overall coverage  : ' + ov + '%\n';
    out += '  Boys coverage     : ' + bc + '%\n';
    out += '  Girls coverage    : ' + gc + '%\n\n';

    out += 'BY SUBMITTER\n';
    Object.entries(bySubmitter)
      .sort((a,b) => b[1].count - a[1].count)
      .forEach(([name,v]) => {
        out += '  '+name+': '+v.count+' schools, '+v.pupils+' pupils, '+(v.pupils>0?Math.round((v.itn/v.pupils)*100):0)+'% coverage\n';
      });

    out += '\nBY DISTRICT\n';
    Object.entries(byDistrict)
      .sort((a,b) => b[1].schools - a[1].schools)
      .forEach(([d,v]) => {
        const c = v.pupils>0?Math.round((v.itn/v.pupils)*100):0;
        out += '  '+d+': '+v.schools+' schools, '+v.pupils+' pupils, '+v.received+' received, '+v.itn+' distributed, '+c+'% coverage\n';
      });

    out += '\nBY CHIEFDOM\n';
    Object.entries(byChiefdom)
      .sort((a,b) => b[1].schools - a[1].schools)
      .forEach(([ck,v]) => {
        const c = v.pupils>0?Math.round((v.itn/v.pupils)*100):0;
        out += '  '+ck+': '+v.schools+' schools, '+v.pupils+' pupils, '+c+'% coverage\n';
      });

    out += '\nALL SCHOOL RECORDS\n' + schoolLines.join('\n');
    return out;

  } catch(e) {
    return 'Error reading sheet: ' + e.message;
  }
}


// ================================================================
// PHU RECEIPT HANDLERS — itn_received.html + itn_reconciliation
// ================================================================

function handlePHUReceipt(data) {
  try {
    const ss    = SpreadsheetApp.openById(CONFIG.SHEET_ID);
    let sheet   = ss.getSheetByName(CONFIG.SHEET_NAME_PHU);

    if (!sheet) {
      sheet = ss.insertSheet(CONFIG.SHEET_NAME_PHU);
      const headers = [
        'Receipt ID','Dispatch ID','Confirmed At','Confirmed Date','Confirmed Time',
        'District','Chiefdom','PHU',
        'Conveyor Name','Conveyor Phone','Conveyor Username','Vehicle Plate','Dispatch Date',
        'Expected Dual-AI ITNs','Received Dual-AI ITNs',
        'Variance','Quantity Match',
        'Discrepancy Reason','Discrepancy Notes',
        'PHU Staff Name','PHU Staff Phone','PHU Staff Username',
        'Conveyor Confirmed','Conveyor Held Accountable'
      ];
      sheet.appendRow(headers);
      const hdr = sheet.getRange(1, 1, 1, headers.length);
      hdr.setBackground('#004080').setFontColor('#ffffff')
         .setFontWeight('bold').setFontFamily('Arial');
      sheet.setFrozenRows(1);
      sheet.setColumnWidths(1, headers.length, 160);
    }

    // Duplicate check — skip if receipt ID already exists
    if (data.id && sheet.getLastRow() > 1) {
      const ids = sheet.getRange(2, 1, sheet.getLastRow()-1, 1).getValues().flat().map(String);
      if (ids.includes(String(data.id))) {
        return jsonResponse({ success: true, id: data.id, duplicate: true, message: 'Receipt already recorded' });
      }
    }

    sheet.appendRow([
      data.id                    || '',
      data.dispatchId            || '',
      data.confirmedAt           || new Date().toISOString(),
      data.confirmedDate         || '',
      data.confirmedTime         || '',
      data.district              || '',
      data.chiefdom              || '',
      data.phu                   || '',
      data.conveyorName          || data.driverName     || '',
      data.conveyorPhone         || data.driverPhone    || '',
      data.conveyorUsername      || data.driverUsername || '',
      data.vehiclePlate          || '',
      data.dispatchDate          || '',
      parseInt(data.expectedTotal) || 0,
      parseInt(data.receivedTotal) || 0,
      parseInt(data.variance)      || 0,
      data.quantityMatch           ? 'Yes' : 'No',
      data.discrepancyReason       || '',
      data.discrepancyNotes        || '',
      data.phuStaffName            || '',
      data.phuStaffPhone           || '',
      data.phuStaffUsername        || '',
      data.conveyorConfirmed || data.driverConfirmed ? 'Yes' : 'No',
      data.conveyorHeldAccountable || data.driverHeldAccountable ? 'Yes' : 'No'
    ]);

    sheet.autoResizeColumns(1, sheet.getLastColumn());
    return jsonResponse({ success: true, id: data.id });
  } catch (err) {
    return jsonResponse({ success: false, error: err.message });
  }
}

function getPHUReceipts() {
  try {
    const ss    = SpreadsheetApp.openById(CONFIG.SHEET_ID);
    const sheet = ss.getSheetByName(CONFIG.SHEET_NAME_PHU);
    if (!sheet || sheet.getLastRow() < 2) return [];

    const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    const data    = sheet.getRange(2, 1, sheet.getLastRow() - 1, sheet.getLastColumn()).getValues();

    return data.map(row => {
      const obj = {};
      headers.forEach((h, i) => { obj[h] = row[i]; });
      // Friendly keys for reconciliation matching
      obj.district      = (obj['District']       || '').trim();
      obj.chiefdom      = (obj['Chiefdom']        || '').trim();
      obj.phu           = (obj['PHU']             || '').trim();
      obj.receivedTotal = parseInt(obj['Received Total']) || 0;
      obj.expectedTotal = parseInt(obj['Expected Total']) || 0;
      obj.variance      = parseInt(obj['Variance'])       || 0;
      obj.quantityMatch = obj['Quantity Match'] === 'Yes';
      return obj;
    });
  } catch (err) {
    return [];
  }
}

// ================================================================
// GET SINGLE DISPATCH BY ID — used by itn_received.html for
// QR verification and manual dispatch ID lookup
// ================================================================
function handleMonitoring(data) {
  try {
    const ss    = SpreadsheetApp.openById(CONFIG.SHEET_ID);
    const name  = 'Monitoring';
    let sheet   = ss.getSheetByName(name);
    if (!sheet) {
      sheet = ss.insertSheet(name);
      const headers = ['Timestamp','Date','District','Chiefdom','Facility','Community','School','Urban/Rural',
        'Monitor Name','Organisation',
        'All team present','Locally recruited','Members trained','Waiting area','Tables','Crowd control','HW communicating',
        'Devices at PHU','Safe storage PHU','Charging area PHU','Devices at school','Devices charged','Device challenges','Challenge type',
        'ITN storage PHU','Stocksheet used','Stocksheet updated','ITN in device','Net movement arranged','School communicated','Nets at DP','Restock plan',
        'Following distribution tasks','Tracking system','Non-register pupils served',
        'Data accurate','Data aligns stock','Enrollment tally',
        'Health messages','Key ITN messages','Flash cards used',
        'Good practices','Problems observed','Problem details',
        'Beneficiary satisfied','Dissatisfaction reason','Nets received','Expected nets','ITN message received','ITN message content','Other channel info','Other channel medium',
        'Monitor Remarks','Datetime'];
      sheet.appendRow(headers);
      sheet.getRange(1,1,1,headers.length).setFontWeight('bold').setBackground('#004080').setFontColor('#ffffff');
    }
    sheet.appendRow([
      new Date(), data.visit_date||'', data.district||'', data.chiefdom||'', data.facility||'',
      data.community||'', data.school_name||'', data.urban_rural||'',
      data.monitor_name||'', data.monitor_org||'',
      data.q1_1||'', data.q1_2||'', data.q1_3||'', data.q1_4||'', data.q1_5||'', data.q1_6||'', data.q1_7||'',
      data.q2_1||'', data.q2_2||'', data.q2_3||'', data.q2_4||'', data.q2_5||'', data.q2_6||'', data.q2_7||'',
      data.q3_1||'', data.q3_2||'', data.q3_3||'', data.q3_4||'', data.q3_5||'', data.q3_6||'', data.q3_7||'', data.q3_8||'',
      data.q4_1||'', data.q4_2||'', data.q4_3||'',
      data.q5_1||'', data.q5_2||'', data.q5_3||'',
      data.q6_1||'', data.q6_2||'', data.q6_3||'',
      data.q7_1||'', data.q7_2||'', data.q7_3||'',
      data.q8_1||'', data.q8_2||'', data.q8_3||'', data.q8_4||'', data.q8_5||'', data.q8_6||'', data.q8_7||'', data.q8_8||'',
      data.monitor_remarks||'', data.datetime||''
    ]);
    return jsonResponse({ status: 'ok' });
  } catch(e) {
    return jsonResponse({ status: 'error', message: e.toString() });
  }
}

// ── ATTENDANCE ───────────────────────────────────────────────
function handleAssessmentGrade(data) {
  try {
    const ss    = SpreadsheetApp.openById(CONFIG.SHEET_ID);
    const name  = 'Assessment Grades';
    let sheet   = ss.getSheetByName(name);
    if (!sheet) {
      sheet = ss.insertSheet(name);
      sheet.appendRow(['Timestamp','Name','Role','Score %','Correct','Total','Time Used','Result','Timed Out','Date/Time']);
      sheet.getRange(1,1,1,10).setFontWeight('bold').setBackground('#004080').setFontColor('#ffffff');
    }
    sheet.appendRow([
      new Date(),
      data.name        || '',
      data.role        || '',
      data.pct         || 0,
      data.correct     || 0,
      data.total       || 50,
      data.timeUsed    || '',
      data.pass ? 'PASS' : 'FAIL',
      data.timedOut ? 'Yes' : 'No',
      data.datetime    || ''
    ]);
    return jsonResponse({ status: 'ok' });
  } catch(e) {
    return jsonResponse({ status: 'error', message: e.toString() });
  }
}

// ── handleSignInOut: Attendance Sign In / Sign Out ───────────
// Fields from attendance_payment.html:
//   action, code (SBD ID), name, tel, role, district, date, timestamp
function handleAttendance(data) {
  try {
    const ss = SpreadsheetApp.openById(CONFIG.SHEET_ID);
    let sheet = ss.getSheetByName(CONFIG.SHEET_NAME_ATT);
    if (!sheet) {
      sheet = ss.insertSheet(CONFIG.SHEET_NAME_ATT);
      const headers = [
        'Timestamp', 'Action', 'SBD ID/Code', 'Name',
        'Telephone', 'Role', 'District', 'Date'
      ];
      sheet.appendRow(headers);
      const hdr = sheet.getRange(1, 1, 1, headers.length);
      hdr.setBackground('#c8991a').setFontColor('#ffffff')
         .setFontWeight('bold').setFontFamily('Arial').setFontSize(10);
      sheet.setFrozenRows(1);
      sheet.setColumnWidth(1, 170); // Timestamp
      sheet.setColumnWidth(3, 160); // SBD ID
      sheet.setColumnWidth(4, 200); // Name
    }
    sheet.appendRow([
      data.timestamp || new Date().toISOString(),
      data.action    || '',   // SIGN_IN or SIGN_OUT
      data.code      || '',   // SBD ID e.g. SBD-079123456
      data.name      || '',
      data.tel       || data.phone || '',   // telephone
      data.role      || '',
      data.district  || '',
      data.date      || ''
    ]);
    return jsonResponse({ status: 'ok', message: (data.action||'Record') + ' saved for ' + (data.name||data.code||'unknown') });
  } catch(e) {
    return jsonResponse({ status: 'error', message: e.toString() });
  }
}

// ── checkAttDuplicate: block same person on same date ────────
function checkAttDuplicate(name, phone, date) {
  try {
    const ss    = SpreadsheetApp.openById(CONFIG.SHEET_ID);
    const sheet = ss.getSheetByName(CONFIG.SHEET_NAME_ATT);
    if (!sheet) return jsonResponse({ duplicate: false });

    const rows = sheet.getDataRange().getValues();
    // Health Staff and Teachers are stored as JSON in cols 9 and 10 (index 8,9)
    // Date is col 7 (index 6)
    for (var i = 1; i < rows.length; i++) {
      var rowDate = String(rows[i][6] || '').trim();
      if (rowDate !== date) continue;

      // Parse health staff and teachers JSON
      var staffJSON   = String(rows[i][8] || '[]');
      var teacherJSON = String(rows[i][9] || '[]');
      var allStaff = [];
      try { allStaff = allStaff.concat(JSON.parse(staffJSON));   } catch(e){}
      try { allStaff = allStaff.concat(JSON.parse(teacherJSON)); } catch(e){}

      for (var j = 0; j < allStaff.length; j++) {
        var s = allStaff[j];
        var sName  = (s.name  || '').trim().toLowerCase();
        var sPhone = (s.tel || s.phone || '').trim(); // tel field in new version
        // Match by phone if provided, else by name
        var matchPhone = phone && sPhone && sPhone === phone;
        var matchName  = name  && sName  && sName  === name;
        if (matchPhone || matchName) {
          return jsonResponse({
            duplicate: true,
            message: (s.name || 'This person') + ' already has an attendance record for ' + date
          });
        }
      }
    }
    return jsonResponse({ duplicate: false });
  } catch(e) {
    return jsonResponse({ duplicate: false, error: e.toString() });
  }
}

// ── getSheetColumns: returns current column headers of any sheet ──
function getSheetColumns(sheetName) {
  try {
    const ss    = SpreadsheetApp.openById(CONFIG.SHEET_ID);
    const sheet = ss.getSheetByName(sheetName);
    if (!sheet) return jsonResponse({ status: 'error', message: 'Sheet not found: ' + sheetName });
    const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    return jsonResponse({ status: 'ok', sheet: sheetName, columns: headers, count: headers.length });
  } catch(e) {
    return jsonResponse({ status: 'error', message: e.toString() });
  }
}

function getDispatchById(dispatchId) {
  if (!dispatchId) return { error: 'No dispatch ID provided' };

  try {
    const ss    = SpreadsheetApp.openById(CONFIG.SHEET_ID);
    const sheet = ss.getSheetByName(CONFIG.SHEET_NAME_ITN);
    if (!sheet || sheet.getLastRow() < 2) return { error: 'ITN Movement sheet not found or empty' };

    const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    const data    = sheet.getRange(2, 1, sheet.getLastRow() - 1, sheet.getLastColumn()).getValues();

    // Find the row where Dispatch ID matches
    const idIdx = headers.findIndex(h => h.toString().toLowerCase().replace(/\s/g,'_') === 'dispatch_id'
                                      || h.toString() === 'Dispatch ID');

    for (const row of data) {
      const rowId = (row[idIdx] || '').toString().trim().toUpperCase();
      if (rowId === dispatchId.trim().toUpperCase()) {
        // Build object from headers
        const obj = {};
        headers.forEach((h, i) => {
          const key = h.toString().toLowerCase().replace(/\s+/g,'_').replace(/[^a-z0-9_]/g,'');
          obj[key] = row[i];
        });
        // Friendly field aliases for itn_received.html
        obj.dispatch_id      = obj.dispatch_id    || dispatchId;
        obj.district         = obj.destination_district || obj.district || '';
        obj.chiefdom         = obj.chiefdom        || '';
        obj.phu              = obj.health_facility_phu || obj.phu || '';
        obj.dual_ai_qty      = parseInt(obj['dual-ai itns']) || parseInt(obj.dual_ai_qty) || parseInt(obj.total_qty) || 0;
        obj.total_qty        = parseInt(obj.total_itns) || parseInt(obj.total_qty) || 0;
        obj.driver_name      = obj.driver_name     || obj.conveyor_name || '';
        obj.driver_username  = obj.driver_username || obj.conveyor_username || '';
        obj.vehicle          = obj.vehicle_plate   || obj.vehicle || '';
        return obj;
      }
    }

    return { error: 'Dispatch ID not found: ' + dispatchId };

  } catch(err) {
    return { error: err.message };
  }
}

// ── handleIDCardSubmission: saves SBD ID card records ────────────
function handleIDCardSubmission(data) {
  try {
    const ss = SpreadsheetApp.openById(CONFIG.SHEET_ID);
    let sheet = ss.getSheetByName(CONFIG.SHEET_NAME_SBD);

    if (!sheet) {
      sheet = ss.insertSheet(CONFIG.SHEET_NAME_SBD);
      const headers = [
        'Timestamp','Full Name','Telephone','ID Number',
        'Role','District','Chiefdom','Facility',
        'Valid Until','Issued By','Submitted At'
      ];
      sheet.appendRow(headers);
      const hdr = sheet.getRange(1, 1, 1, headers.length);
      hdr.setBackground('#004080').setFontColor('#ffffff').setFontWeight('bold').setFontFamily('Arial').setFontSize(11);
      sheet.setFrozenRows(1);
      sheet.setColumnWidths(1, headers.length, 160);
    }

    sheet.appendRow([
      new Date(),
      data.name       || '',
      data.telephone  || '',
      data.idNumber   || '',
      data.role       || '',
      data.district   || '',
      data.chiefdom   || '',
      data.facility   || '',
      data.validUntil || '',
      data.issuedBy   || '',
      data.timestamp  || ''
    ]);

    sheet.autoResizeColumns(1, sheet.getLastColumn());
    return jsonResponse({ status: 'success', message: 'Record saved.' });
  } catch(e) {
    return jsonResponse({ status: 'error', message: e.toString() });
  }
}

// ══════════════════════════════════════════════════════════════
// ══════════════════════════════════════════════════════════════
// SHEET DEFINITIONS — single source of truth for all columns
// ══════════════════════════════════════════════════════════════
function getSheetDefinitions(ss) {
  return {
    // 1. Submissions (Distribution)
    [CONFIG.SHEET_NAME_DATA]: FIELD_MAP.map(function(f) { return f.label; }),

    // 2. ITN Movement
    [CONFIG.SHEET_NAME_ITN]: [
      'Dispatch ID','Timestamp',
      'Staff Name','Staff Username','Staff District',
      'Destination District','Chiefdom','Health Facility (PHU)',
      'Dual-AI ITNs',
      'Conveyor Name','Conveyor Username','Conveyor Phone',
      'Vehicle Plate','Status'
    ],

    // 3. PHU Receipts (ITN Received)
    [CONFIG.SHEET_NAME_PHU]: [
      'Receipt ID','Dispatch ID','Confirmed At','Confirmed Date','Confirmed Time',
      'District','Chiefdom','PHU',
      'Conveyor Name','Conveyor Phone','Conveyor Username','Vehicle Plate','Dispatch Date',
      'Expected Dual-AI ITNs','Received Dual-AI ITNs',
      'Variance','Quantity Match',
      'Discrepancy Reason','Discrepancy Notes',
      'PHU Staff Name','PHU Staff Phone','PHU Staff Username',
      'Conveyor Confirmed','Conveyor Held Accountable'
    ],

    // 4. Attendance & Payment
    [CONFIG.SHEET_NAME_ATT]: [
      'Timestamp','Action','SBD ID/Code','Name',
      'Telephone','Role','District','Chiefdom','Health Facility','Date',
      'GPS','GPS Accuracy (m)'
    ],

    // 5. Monitoring
    'Monitoring': [
      'Timestamp','Date','District','Chiefdom','Facility','Community','School','Urban/Rural',
      'Monitor Name','Organisation',
      'All team present','Locally recruited','Members trained','Waiting area','Tables',
      'Crowd control','HW communicating',
      'Devices at PHU','Safe storage PHU','Charging area PHU','Devices at school',
      'Devices charged','Device challenges','Challenge type',
      'ITN storage PHU','Stocksheet used','Stocksheet updated','ITN in device',
      'Net movement arranged','School communicated','Nets at DP','Restock plan',
      'Following distribution tasks','Tracking system','Non-register pupils served',
      'Data accurate','Data aligns stock','Enrollment tally',
      'Health messages','Key ITN messages','Flash cards used',
      'Good practices','Problems observed','Problem details',
      'Beneficiary satisfied','Dissatisfaction reason','Nets received','Expected nets',
      'ITN message received','ITN message content','Other channel info','Other channel medium',
      'Monitor Remarks','Datetime'
    ],

    // 6. SBD Records (ID Cards)
    [CONFIG.SHEET_NAME_SBD]: [
      'Timestamp','Full Name','Telephone','ID Number',
      'Role','District','Chiefdom','Facility',
      'Valid Until','Issued By','Submitted At'
    ],

    // 7. Device Tags
    [CONFIG.SHEET_NAME_DTAG]: [
      'Timestamp','Year','ID Number','Full Name','Telephone',
      'Role','District','Chiefdom','Health Facility','Generated At'
    ],

    // 8. Device Tracking
    [CONFIG.SHEET_NAME_DTRACK]: [
      'Timestamp','Status','Device ID','Full Name','Telephone',
      'Role','District','Chiefdom','Health Facility','Year',
      'Issued At','Returned At','Submitted At'
    ]
  };
}

// ══════════════════════════════════════════════════════════════
// updateSheetHeaders — ADDS missing columns (safe — no data lost)
// Run manually from Apps Script editor: updateSheetHeaders()
// ══════════════════════════════════════════════════════════════
function updateSheetHeaders() {
  const ss   = SpreadsheetApp.openById(CONFIG.SHEET_ID);
  const defs = getSheetDefinitions(ss);
  const log  = [];

  Object.keys(defs).forEach(function(sheetName) {
    const expectedHeaders = defs[sheetName];
    let sheet = ss.getSheetByName(sheetName);

    // Create sheet if it doesn't exist
    if (!sheet) {
      sheet = ss.insertSheet(sheetName);
      sheet.appendRow(expectedHeaders);
      styleHeader(sheet, expectedHeaders.length);
      log.push('✅ Created: ' + sheetName + ' (' + expectedHeaders.length + ' cols)');
      return;
    }

    // Add headers if sheet is empty
    if (sheet.getLastRow() === 0) {
      sheet.appendRow(expectedHeaders);
      styleHeader(sheet, expectedHeaders.length);
      log.push('✅ Headers added to empty sheet: ' + sheetName);
      return;
    }

    // Add any missing columns
    const existing = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0]
      .map(function(h) { return String(h).trim(); });
    let added = 0;
    expectedHeaders.forEach(function(col) {
      if (existing.indexOf(col) === -1) {
        const newCol = sheet.getLastColumn() + 1;
        const cell = sheet.getRange(1, newCol);
        cell.setValue(col);
        cell.setBackground('#004080').setFontColor('#ffffff')
            .setFontWeight('bold').setFontFamily('Arial').setFontSize(10);
        sheet.setColumnWidth(newCol, 160);
        existing.push(col);
        added++;
      }
    });
    log.push(added > 0
      ? '➕ ' + sheetName + ': added ' + added + ' column(s)'
      : '✓  ' + sheetName + ': up to date');
  });

  Logger.log('=== updateSheetHeaders ===\n' + log.join('\n'));
  try {
    SpreadsheetApp.getUi().alert('Sheet Headers Updated', log.join('\n'), SpreadsheetApp.getUi().ButtonSet.OK);
  } catch(e) {}
}

// ══════════════════════════════════════════════════════════════
// resetSheetHeaders — REBUILDS header row for a sheet
// USE WITH CAUTION — rewrites row 1, data rows are untouched
// Run manually: resetSheetHeaders('Sheet Name')
// Or reset all: resetAllSheetHeaders()
// ══════════════════════════════════════════════════════════════
function resetSheetHeaders(sheetName) {
  const ss   = SpreadsheetApp.openById(CONFIG.SHEET_ID);
  const defs = getSheetDefinitions(ss);
  const log  = [];

  function doReset(name) {
    const expectedHeaders = defs[name];
    if (!expectedHeaders) { log.push('❌ Unknown sheet: ' + name); return; }
    let sheet = ss.getSheetByName(name);
    if (!sheet) {
      sheet = ss.insertSheet(name);
      sheet.appendRow(expectedHeaders);
      styleHeader(sheet, expectedHeaders.length);
      log.push('✅ Created new sheet: ' + name);
      return;
    }
    // Clear only row 1, rewrite it
    const lastCol = Math.max(sheet.getLastColumn(), expectedHeaders.length);
    sheet.getRange(1, 1, 1, lastCol).clearContent().clearFormat();
    expectedHeaders.forEach(function(col, i) {
      const cell = sheet.getRange(1, i + 1);
      cell.setValue(col);
      cell.setBackground('#004080').setFontColor('#ffffff')
          .setFontWeight('bold').setFontFamily('Arial').setFontSize(10);
      sheet.setColumnWidth(i + 1, 160);
    });
    sheet.setFrozenRows(1);
    log.push('🔄 Reset headers: ' + name + ' (' + expectedHeaders.length + ' cols)');
  }

  if (sheetName) {
    doReset(sheetName);
  }

  Logger.log(log.join('\n'));
  try {
    SpreadsheetApp.getUi().alert('Headers Reset', log.join('\n'), SpreadsheetApp.getUi().ButtonSet.OK);
  } catch(e) {}
}

function resetAllSheetHeaders() {
  const ss   = SpreadsheetApp.openById(CONFIG.SHEET_ID);
  const defs = getSheetDefinitions(ss);
  const log  = [];

  Object.keys(defs).forEach(function(name) {
    const expectedHeaders = defs[name];
    let sheet = ss.getSheetByName(name);
    if (!sheet) {
      sheet = ss.insertSheet(name);
      sheet.appendRow(expectedHeaders);
      styleHeader(sheet, expectedHeaders.length);
      log.push('✅ Created: ' + name);
      return;
    }
    const lastCol = Math.max(sheet.getLastColumn(), expectedHeaders.length);
    sheet.getRange(1, 1, 1, lastCol).clearContent().clearFormat();
    expectedHeaders.forEach(function(col, i) {
      const cell = sheet.getRange(1, i + 1);
      cell.setValue(col);
      cell.setBackground('#004080').setFontColor('#ffffff')
          .setFontWeight('bold').setFontFamily('Arial').setFontSize(10);
      sheet.setColumnWidth(i + 1, 160);
    });
    sheet.setFrozenRows(1);
    log.push('🔄 Reset: ' + name);
  });

  Logger.log('=== resetAllSheetHeaders ===\n' + log.join('\n'));
  try {
    SpreadsheetApp.getUi().alert('All Headers Reset', log.join('\n'), SpreadsheetApp.getUi().ButtonSet.OK);
  } catch(e) {}
}

function styleHeader(sheet, numCols) {
  const hdr = sheet.getRange(1, 1, 1, numCols);
  hdr.setBackground('#004080').setFontColor('#ffffff')
     .setFontWeight('bold').setFontFamily('Arial').setFontSize(10);
  sheet.setFrozenRows(1);
  sheet.setColumnWidths(1, numCols, 160);
}

// ── handleDeviceTagSubmission — saves device tag generation event ─
function handleDeviceTagSubmission(data) {
  try {
    const ss = SpreadsheetApp.openById(CONFIG.SHEET_ID);
    let sheet = ss.getSheetByName(CONFIG.SHEET_NAME_DTAG);
    if (!sheet) {
      sheet = ss.insertSheet(CONFIG.SHEET_NAME_DTAG);
      const headers = [
        'Timestamp', 'Year', 'ID Number', 'Full Name', 'Telephone',
        'Role', 'District', 'Chiefdom', 'Health Facility', 'Generated At'
      ];
      sheet.appendRow(headers);
      const hdr = sheet.getRange(1, 1, 1, headers.length);
      hdr.setBackground('#004080').setFontColor('#ffffff')
         .setFontWeight('bold').setFontFamily('Arial').setFontSize(10);
      sheet.setFrozenRows(1);
      sheet.setColumnWidths(1, headers.length, 150);
    }
    sheet.appendRow([
      new Date(),
      data.year       || '',
      data.id         || data.code || '',
      data.name       || '',
      data.tel        || '',
      data.role       || '',
      data.district   || '',
      data.chiefdom   || '',
      data.facility   || '',
      data.generated_at || new Date().toISOString()
    ]);
    sheet.autoResizeColumns(1, sheet.getLastColumn());
    return jsonResponse({ success: true, message: 'Device tag saved.' });
  } catch(e) {
    return jsonResponse({ success: false, error: e.toString() });
  }
}

// ── handleDeviceTracking — saves device register/return events ─────
function handleDeviceTracking(data) {
  try {
    const ss = SpreadsheetApp.openById(CONFIG.SHEET_ID);
    let sheet = ss.getSheetByName(CONFIG.SHEET_NAME_DTRACK);
    if (!sheet) {
      sheet = ss.insertSheet(CONFIG.SHEET_NAME_DTRACK);
      const headers = [
        'Timestamp', 'Action', 'Device ID', 'Full Name',
        'Telephone', 'Role', 'District', 'Chiefdom',
        'Health Facility', 'Year', 'Submitted At'
      ];
      sheet.appendRow(headers);
      const hdr = sheet.getRange(1, 1, 1, headers.length);
      hdr.setBackground('#004080').setFontColor('#ffffff')
         .setFontWeight('bold').setFontFamily('Arial').setFontSize(10);
      sheet.setFrozenRows(1);
      sheet.setColumnWidths(1, headers.length, 160);
    }
    sheet.appendRow([
      new Date(),
      data.action    || '',
      data.code      || '',
      data.name      || '',
      data.phone     || '',
      data.role      || '',
      data.district  || '',
      data.chiefdom  || '',
      data.facility  || '',
      data.year      || '',
      data.timestamp || new Date().toISOString()
    ]);
    sheet.autoResizeColumns(1, sheet.getLastColumn());
    return jsonResponse({ success: true, message: (data.action || 'Event') + ' recorded.' });
  } catch(e) {
    return jsonResponse({ success: false, error: e.toString() });
  }
}

// ── getDeviceTrackingRecords — returns all device tracking rows ──
function getDeviceTrackingRecords() {
  try {
    const ss    = SpreadsheetApp.openById(CONFIG.SHEET_ID);
    const sheet = ss.getSheetByName(CONFIG.SHEET_NAME_DTRACK);
    if (!sheet || sheet.getLastRow() < 2) return { status:'ok', records:[] };

    const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0]
      .map(function(h){ return String(h).trim(); });
    const rows = sheet.getRange(2, 1, sheet.getLastRow()-1, sheet.getLastColumn()).getValues();

    const records = rows.map(function(row) {
      var obj = {};
      headers.forEach(function(h, i){ obj[h] = row[i] !== undefined ? String(row[i]) : ''; });
      return obj;
    }).filter(function(r){ return r['Device ID'] || r['code']; });

    return { status:'ok', records: records };
  } catch(e) {
    return { status:'error', error: e.toString() };
  }
}

// ══════════════════════════════════════════════════════════════
// TEST FUNCTIONS — run manually from Apps Script editor
// ══════════════════════════════════════════════════════════════

// Test attendance sign-in submission
function testAttendanceSignIn() {
  const result = handleAttendance({
    action:    'SIGN_IN',
    code:      'SBD-079123456',
    name:      'Test User',
    tel:       '079123456',
    phone:     '079123456',
    role:      'PHU Staff',
    district:  'Bo',
    chiefdom:  'Bo Town',
    facility:  'Bo District Hospital',
    date:      new Date().toLocaleDateString('en-GB'),
    timestamp: new Date().toISOString(),
    gps:       '7.9643,−11.7383',
    gps_acc:   '10'
  });
  Logger.log('testAttendanceSignIn: ' + JSON.stringify(result));
  try {
    SpreadsheetApp.getUi().alert('Attendance Test', JSON.stringify(result), SpreadsheetApp.getUi().ButtonSet.OK);
  } catch(e) {}
}

// Test attendance sign-out submission
function testAttendanceSignOut() {
  const result = handleAttendance({
    action:    'SIGN_OUT',
    code:      'SBD-079123456',
    name:      'Test User',
    tel:       '079123456',
    phone:     '079123456',
    role:      'PHU Staff',
    district:  'Bo',
    chiefdom:  'Bo Town',
    facility:  'Bo District Hospital',
    date:      new Date().toLocaleDateString('en-GB'),
    timestamp: new Date().toISOString(),
    gps:       '7.9643,−11.7383',
    gps_acc:   '10'
  });
  Logger.log('testAttendanceSignOut: ' + JSON.stringify(result));
  try {
    SpreadsheetApp.getUi().alert('Attendance Test', JSON.stringify(result), SpreadsheetApp.getUi().ButtonSet.OK);
  } catch(e) {}
}

// Test doPost routing for handleSignInOut
function testAttendanceRouting() {
  const mockEvent = {
    postData: {
      contents: JSON.stringify({
        action:    'handleSignInOut',
        code:      'SBD-079999999',
        name:      'Routing Test',
        phone:     '079999999',
        role:      'Teacher',
        district:  'Kenema',
        chiefdom:  'Kenema Town',
        date:      new Date().toLocaleDateString('en-GB'),
        timestamp: new Date().toISOString()
      })
    }
  };
  const result = doPost(mockEvent);
  Logger.log('testAttendanceRouting: ' + result.getContent());
  try {
    SpreadsheetApp.getUi().alert('Routing Test', result.getContent(), SpreadsheetApp.getUi().ButtonSet.OK);
  } catch(e) {}
}
