/**
 * 03_Importer.gs
 * Reads vendor usage files (.xlsx) from inbox folders, parses, dedupes,
 * appends to _ledger. Handles multi-cost-center assignments via pipe codes.
 *
 * ============================================================
 * EXAMPLE / DEMONSTRATION SCHEMA
 * ============================================================
 * The column-name maps below (VENDOR_A_HEADER_MAP, VENDOR_B_HEADER_MAP)
 * are illustrative placeholders showing the shape of a typical
 * ride-hailing or logistics vendor export. If you fork this repo, replace
 * both maps with your actual vendor export columns. The rest of the
 * importer logic (dedup, attribution, archive) is provider-agnostic.
 */

const VENDOR_A_HEADER_MAP = {
  // Left = column header in vendor's export
  // Right = canonical field in _ledger
  'Transaction ID': 'transaction_id',
  'Booking Time': 'booking_time',
  'Pickup Time': 'pickup_time',
  'Completion Time': 'completion_time',
  'Employee Name': 'employee_name',
  'Employee ID': 'employee_raw_id',
  'Group': 'group_raw',
  'Service Type': 'service_type',
  'Pickup Location': 'pickup_location',
  'Pickup City': 'pickup_city',
  'Destination': 'destination',
  'Destination City': 'destination_city',
  'Distance (Km)': 'distance_km',
  'Total Amount': 'total_amount',
  'Base Fare': 'base_fare',
  'Platform Fee': 'platform_fee',
  'Toll Fee': 'toll_fee',
  'Parking Fee': 'parking_fee',
  'Other Fee': 'other_fee',
  'Voucher Discount': 'voucher_discount',
  'Trip Reason': 'trip_reason',
  'Trip Reason Detail': 'trip_reason_detail'
};

const VENDOR_B_HEADER_MAP = {
  'BOOKING_CODE': 'transaction_id',
  'TRANSACTION_TIME': 'booking_time',
  'CREATION_TIME': 'pickup_time',
  'COMPLETION_TIME': 'completion_time',
  'EMPLOYEE_NAME': 'employee_name',
  'EMPLOYEE_ID': 'employee_raw_id',
  'GROUP_NAME': 'group_raw',
  'VERTICAL': 'vertical',
  'SERVICE TYPE': 'service_type',
  'CITY': 'pickup_city',
  'PICK_UP_ADDRESS': 'pickup_location',
  'DROP_OFF_ADDRESS': 'destination',
  'DISTANCE_IN_KM': 'distance_km',
  'AMOUNT': 'total_amount'
};

// ============================================================
// MENU ENTRY
// ============================================================

function menu_processUsageUploads() {
  const ui = SpreadsheetApp.getUi();
  if (!isDriveApiEnabled_()) {
    ui.alert('Drive API not enabled. Apps Script editor → Services (+) → Drive API → v2.');
    return;
  }
  try {
    const result = processUsageUploads();
    let msg = 'Done.\n\n' +
      '📥 Vendor A: ' + result.vendorA.filesProcessed + ' file(s), ' + result.vendorA.rowsAppended + ' new rows, ' + result.vendorA.rowsSkipped + ' duplicates\n' +
      '📥 Vendor B: ' + result.vendorB.filesProcessed + ' file(s), ' + result.vendorB.rowsAppended + ' new rows, ' + result.vendorB.rowsSkipped + ' duplicates\n';
    if (result.unmappedAdded > 0) msg += '\n🚨 ' + result.unmappedAdded + ' new unmapped IDs flagged.';
    if (result.errors.length > 0) msg += '\n\n⚠️ Errors:\n' + result.errors.map(e => '• ' + e).join('\n');
    if (result.vendorA.filesProcessed === 0 && result.vendorB.filesProcessed === 0) {
      msg = 'No files found in inbox folders.';
    }
    ui.alert('Import complete', msg, ui.ButtonSet.OK);
  } catch (e) {
    ui.alert('Import failed', e.message, ui.ButtonSet.OK);
  }
}

function menu_verifyDriveApi() {
  const ui = SpreadsheetApp.getUi();
  if (isDriveApiEnabled_()) ui.alert('✅ Advanced Drive API is enabled.');
  else ui.alert('Drive API not enabled. Services (+) → Drive API → v2.');
}

function isDriveApiEnabled_() {
  try { Drive.Files.list({ maxResults: 1 }); return true; } catch (e) { return false; }
}

function menu_resetLedger() {
  const ui = SpreadsheetApp.getUi();
  if (ui.alert('💣 RESET ledger', 'Delete all transaction rows? Master sheets are preserved.', ui.ButtonSet.YES_NO) !== ui.Button.YES) return;
  const confirm2 = ui.prompt('Final confirmation', 'Type DELETE to confirm:', ui.ButtonSet.OK_CANCEL);
  if (confirm2.getSelectedButton() !== ui.Button.OK || confirm2.getResponseText().trim() !== 'DELETE') {
    ui.alert('Cancelled.'); return;
  }

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let deleted = 0;
  deleted += clearLedgerLikeSheet_(ss.getSheetByName(CONFIG.SHEET_LEDGER));
  clearLedgerLikeSheet_(ss.getSheetByName(CONFIG.SHEET_UNMAPPED));

  ui.alert('✅ Ledger reset (' + deleted + ' rows). Drop usage files and run "Process new uploads".');
}

function clearLedgerLikeSheet_(sh) {
  if (!sh) return 0;
  const lastRow = sh.getLastRow();
  const maxCols = sh.getMaxColumns();
  if (lastRow <= 1) return 0;

  const frozenRows = sh.getFrozenRows();
  if (frozenRows > 0) sh.setFrozenRows(0);
  sh.deleteRows(2, lastRow - 1);
  sh.insertRowsAfter(1, 1);

  const newRow = sh.getRange(2, 1, 1, maxCols);
  newRow.clearFormat();
  newRow.clearContent();
  newRow.clearDataValidations();
  newRow.clearNote();
  newRow.setBackground(null).setFontWeight('normal').setFontColor(null)
    .setFontStyle('normal').setFontFamily(null).setFontSize(10)
    .setHorizontalAlignment(null).setVerticalAlignment(null).setNumberFormat(null);

  if (frozenRows > 0) sh.setFrozenRows(frozenRows);
  return lastRow - 1;
}

// ============================================================
// MAIN PROCESS
// ============================================================

function processUsageUploads() {
  const props = PropertiesService.getScriptProperties();
  const errors = [];
  const result = {
    vendorA: { filesProcessed: 0, rowsAppended: 0, rowsSkipped: 0 },
    vendorB: { filesProcessed: 0, rowsAppended: 0, rowsSkipped: 0 },
    unmappedAdded: 0,
    errors: errors
  };

  processVendorFolder_(
    'VendorA', props.getProperty(CONFIG.PROP_FOLDER_INBOX_VENDOR_A),
    props.getProperty(CONFIG.PROP_FOLDER_ARCHIVE_VENDOR_A),
    importVendorAFile_, result.vendorA, errors, result
  );
  processVendorFolder_(
    'VendorB', props.getProperty(CONFIG.PROP_FOLDER_INBOX_VENDOR_B),
    props.getProperty(CONFIG.PROP_FOLDER_ARCHIVE_VENDOR_B),
    importVendorBFile_, result.vendorB, errors, result
  );

  return result;
}

function processVendorFolder_(vendorLabel, inboxId, archiveId, importFn, tally, errors, result) {
  if (!inboxId || !archiveId) return;
  const inbox = DriveApp.getFolderById(inboxId);
  const archive = DriveApp.getFolderById(archiveId);
  const files = inbox.getFiles();
  while (files.hasNext()) {
    const file = files.next();
    try {
      const stats = importFn(file);
      tally.filesProcessed += 1;
      tally.rowsAppended += stats.appended;
      tally.rowsSkipped += stats.skipped;
      result.unmappedAdded += stats.unmappedAdded;
      try {
        archiveFile_(file, archive);
      } catch (archiveErr) {
        errors.push(vendorLabel + ' "' + file.getName() + '" imported OK but archive failed: ' + archiveErr.message);
      }
    } catch (e) {
      errors.push(vendorLabel + ' "' + file.getName() + '": ' + e.message);
    }
  }
}

// ============================================================
// XLSX → temp Sheet conversion
// ============================================================

function readXlsxAsTempSheet_(file) {
  const blob = file.getBlob();
  const tempName = '__tmp_import_' + Utilities.getUuid().substring(0, 8);

  let tempFile = null;
  let lastErr = null;
  for (let attempt = 1; attempt <= 3; attempt++) {
    try {
      if (attempt > 1) Utilities.sleep(attempt * 1500);
      tempFile = Drive.Files.insert({
        title: tempName,
        mimeType: 'application/vnd.google-apps.spreadsheet'
      }, blob, { convert: true });
      break;
    } catch (e) { lastErr = e; }
  }

  if (!tempFile) {
    throw new Error(
      'Failed to convert "' + file.getName() + '" after 3 attempts. ' +
      'Workaround: open in Drive → "Open with → Google Sheets" → re-upload converted file. ' +
      'Original: ' + (lastErr && lastErr.message)
    );
  }

  return {
    spreadsheet: SpreadsheetApp.openById(tempFile.id),
    cleanup: function() {
      try { DriveApp.getFileById(tempFile.id).setTrashed(true); } catch (e) {}
    }
  };
}

// ============================================================
// VENDOR A IMPORTER
// ============================================================

function importVendorAFile_(file) {
  const fileName = file.getName();
  const handle = readXlsxAsTempSheet_(file);
  try {
    const rows = extractVendorARows_(handle.spreadsheet, fileName);
    const enriched = rows.map(r => {
      const attrib = attributeBranch_(r.employee_raw_id);
      r.canonical_branch_code = attrib.code || '';
      r.imported_at = formatJakartaIsoString_(new Date());
      r.source_file = fileName;
      // Example rule: service types starting with "Delivery" go to Logistik vertical.
      // Adapt to your real classification.
      r.vertical = String(r.service_type || '').match(/^Delivery/i) ? 'Logistik' : 'Transport';
      r.booking_time = normalizeTimestampToText_(r.booking_time);
      r.pickup_time = normalizeTimestampToText_(r.pickup_time);
      r.completion_time = normalizeTimestampToText_(r.completion_time);
      r.period = derivePeriodFromText_(r.booking_time);
      return { row: r, isNewUnmapped: attrib.isNew };
    });

    const unmappedRows = enriched.filter(e => e.isNewUnmapped).map(e => e.row);
    const unmappedAdded = trackUnmapped_(unmappedRows, 'VendorA');
    const stats = appendToLedger_(enriched.map(e => e.row));
    return { appended: stats.appended, skipped: stats.skipped, unmappedAdded };
  } finally { handle.cleanup(); }
}

function extractVendorARows_(spreadsheet, fileName) {
  return extractRows_(spreadsheet, fileName, VENDOR_A_HEADER_MAP, 'VendorA',
    ['distance_km', 'total_amount', 'base_fare', 'platform_fee', 'toll_fee', 'parking_fee', 'other_fee', 'voucher_discount']);
}

// ============================================================
// VENDOR B IMPORTER
// ============================================================

function importVendorBFile_(file) {
  const fileName = file.getName();
  const handle = readXlsxAsTempSheet_(file);
  try {
    const rows = extractVendorBRows_(handle.spreadsheet, fileName);
    const enriched = rows.map(r => {
      const attrib = attributeBranch_(r.employee_raw_id);
      r.canonical_branch_code = attrib.code || '';
      r.imported_at = formatJakartaIsoString_(new Date());
      r.source_file = fileName;
      const v = String(r.vertical || '').toLowerCase();
      r.vertical = v === 'logistics' ? 'Logistik' : 'Transport';
      // Vendor B quirk: some fields use UTC despite lacking a Z suffix,
      // others use local time with a misleading Z. Adjust flags to your data.
      r.booking_time = normalizeTimestampToText_(r.booking_time, false);
      r.pickup_time = normalizeTimestampToText_(r.pickup_time, false);
      r.completion_time = normalizeTimestampToText_(r.completion_time, true);
      r.period = derivePeriodFromText_(r.booking_time);
      return { row: r, isNewUnmapped: attrib.isNew };
    });

    const unmappedRows = enriched.filter(e => e.isNewUnmapped).map(e => e.row);
    const unmappedAdded = trackUnmapped_(unmappedRows, 'VendorB');
    const stats = appendToLedger_(enriched.map(e => e.row));
    return { appended: stats.appended, skipped: stats.skipped, unmappedAdded };
  } finally { handle.cleanup(); }
}

function extractVendorBRows_(spreadsheet, fileName) {
  return extractRows_(spreadsheet, fileName, VENDOR_B_HEADER_MAP, 'VendorB',
    ['distance_km', 'total_amount']);
}

// ============================================================
// SHARED EXTRACTION + HELPERS
// ============================================================

function extractRows_(spreadsheet, fileName, headerMap, providerLabel, numericKeys) {
  const sheets = spreadsheet.getSheets();
  let dataSheet = null, headerLen = 0, columnMap = null;
  for (const sh of sheets) {
    if (sh.getLastRow() < 2) continue;
    const firstRow = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0];
    const map = mapHeaders_(firstRow, headerMap);
    if (map.transaction_id !== undefined) {
      dataSheet = sh; headerLen = firstRow.length; columnMap = map; break;
    }
  }
  if (!dataSheet) throw new Error('No ' + providerLabel + ' data sheet found in "' + fileName + '".');

  const lastRow = dataSheet.getLastRow();
  if (lastRow < 2) return [];

  const data = dataSheet.getRange(2, 1, lastRow - 1, headerLen).getDisplayValues();
  const rows = [];
  for (const row of data) {
    if (!row[columnMap.transaction_id]) continue;
    const obj = { provider: providerLabel };
    for (const canonicalKey in columnMap) {
      let val = row[columnMap[canonicalKey]];
      if (numericKeys.indexOf(canonicalKey) >= 0) val = parseDisplayNumber_(val);
      obj[canonicalKey] = val;
    }
    rows.push(obj);
  }
  return rows;
}

function mapHeaders_(headerRow, headerMap) {
  const result = {};
  headerRow.forEach((header, idx) => {
    const trimmed = String(header).trim();
    if (headerMap[trimmed] !== undefined) result[headerMap[trimmed]] = idx;
  });
  return result;
}

function normalizeTimestampToText_(value, treatZAsJakarta) {
  if (!value) return '';
  const s = String(value).trim();
  if (!s) return '';

  let m = s.match(/^(\d{4})-(\d{2})-(\d{2})[T\s](\d{2}):(\d{2}):(\d{2})(?:\.\d+)?(Z|[+-]\d{2}:?\d{2})?$/);
  if (m) {
    let year = parseInt(m[1]), month = parseInt(m[2]), day = parseInt(m[3]);
    let hour = parseInt(m[4]), minute = parseInt(m[5]), second = parseInt(m[6]);
    const tz = m[7];
    if (tz === 'Z' && !treatZAsJakarta) {
      hour += 7;
      if (hour >= 24) {
        hour -= 24;
        const d = new Date(Date.UTC(year, month - 1, day)); d.setUTCDate(d.getUTCDate() + 1);
        year = d.getUTCFullYear(); month = d.getUTCMonth() + 1; day = d.getUTCDate();
      }
    }
    return year + '-' + pad2_(month) + '-' + pad2_(day) + ' ' + pad2_(hour) + '.' + pad2_(minute) + '.' + pad2_(second);
  }

  m = s.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})\s+(\d{1,2})[:.](\d{2})[:.](\d{2})\s*(AM|PM)?$/i);
  if (m) {
    const month = parseInt(m[1]), day = parseInt(m[2]), year = parseInt(m[3]);
    let hour = parseInt(m[4]);
    const minute = parseInt(m[5]), second = parseInt(m[6]), ampm = m[7];
    if (ampm) {
      if (ampm.toUpperCase() === 'PM' && hour < 12) hour += 12;
      if (ampm.toUpperCase() === 'AM' && hour === 12) hour = 0;
    }
    return year + '-' + pad2_(month) + '-' + pad2_(day) + ' ' + pad2_(hour) + '.' + pad2_(minute) + '.' + pad2_(second);
  }

  const d = new Date(s);
  if (!isNaN(d.getTime())) return formatJakartaIsoString_(d);
  return s;
}

function pad2_(n) { return String(n).padStart(2, '0'); }

function formatJakartaIsoString_(date) {
  return Utilities.formatDate(date, 'Asia/Jakarta', 'yyyy-MM-dd HH.mm.ss');
}

function derivePeriodFromText_(timestampText) {
  if (!timestampText) return '';
  const m = String(timestampText).match(/^(\d{4})-(\d{2})/);
  return m ? m[1] + '-' + m[2] : '';
}

function parseDisplayNumber_(s) {
  if (s === null || s === undefined || s === '') return '';
  let cleaned = String(s).trim().replace(/[^0-9.,\-]/g, '');
  const lastDot = cleaned.lastIndexOf('.'), lastComma = cleaned.lastIndexOf(',');
  if (lastDot >= 0 && lastComma >= 0) {
    if (lastDot > lastComma) cleaned = cleaned.replace(/,/g, '');
    else cleaned = cleaned.replace(/\./g, '').replace(',', '.');
  } else if (lastComma >= 0) {
    const afterComma = cleaned.length - lastComma - 1;
    if (afterComma === 3 && cleaned.replace(/[^0-9]/g, '').length > 3) cleaned = cleaned.replace(/,/g, '');
    else cleaned = cleaned.replace(',', '.');
  }
  const n = parseFloat(cleaned);
  return isNaN(n) ? '' : n;
}

function appendToLedger_(rows) {
  if (rows.length === 0) return { appended: 0, skipped: 0 };
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG.SHEET_LEDGER);
  if (!sh) throw new Error('Ledger sheet not found.');

  const existingIds = new Set();
  const lastRow = sh.getLastRow();
  if (lastRow > 1) {
    sh.getRange(2, 1, lastRow - 1, 1).getValues().forEach(r => {
      if (r[0]) existingIds.add(String(r[0]));
    });
  }

  const newRows = rows.filter(r => r.transaction_id && !existingIds.has(String(r.transaction_id)));
  if (newRows.length > 0) {
    const valuesToWrite = newRows.map(r => CONFIG.HEADERS_LEDGER.map(h => {
      const v = r[h];
      return (v === undefined || v === null) ? '' : v;
    }));
    const startRow = lastRow + 1;
    sh.getRange(startRow, 1, valuesToWrite.length, CONFIG.HEADERS_LEDGER.length).setValues(valuesToWrite);
    const headers = sh.getRange(1, 1, 1, CONFIG.HEADERS_LEDGER.length).getValues()[0];
    ['booking_time', 'pickup_time', 'completion_time', 'imported_at', 'period'].forEach(col => {
      const idx = headers.indexOf(col);
      if (idx >= 0) sh.getRange(startRow, idx + 1, valuesToWrite.length, 1).setNumberFormat('@');
    });
  }
  return { appended: newRows.length, skipped: rows.length - newRows.length };
}

function archiveFile_(file, archiveFolder) {
  const ts = Utilities.formatDate(new Date(), 'Asia/Jakarta', 'yyyyMMdd_HHmmss');
  const newName = ts + '__' + file.getName();
  let lastErr = null;
  for (let attempt = 1; attempt <= 3; attempt++) {
    try {
      if (attempt > 1) Utilities.sleep(attempt * 1500);
      file.makeCopy(newName, archiveFolder); file.setTrashed(true); return;
    } catch (e) { lastErr = e; }
  }
  throw new Error('Failed to archive "' + file.getName() + '" after 3 attempts. Original: ' + (lastErr && lastErr.message));
}

function menu_recoverArchive() {
  const ui = SpreadsheetApp.getUi();
  const resp = ui.prompt('Recover archived files', 'Number of recent files per provider (1-10):', ui.ButtonSet.OK_CANCEL);
  if (resp.getSelectedButton() !== ui.Button.OK) return;
  const n = parseInt(resp.getResponseText().trim() || '1');
  if (isNaN(n) || n < 1 || n > 10) { ui.alert('Invalid.'); return; }

  const props = PropertiesService.getScriptProperties();
  const rA = recoverFromArchive_(props, CONFIG.PROP_FOLDER_INBOX_VENDOR_A, CONFIG.PROP_FOLDER_ARCHIVE_VENDOR_A, n);
  const rB = recoverFromArchive_(props, CONFIG.PROP_FOLDER_INBOX_VENDOR_B, CONFIG.PROP_FOLDER_ARCHIVE_VENDOR_B, n);
  const rI = recoverFromArchive_(props, CONFIG.PROP_FOLDER_INBOX_INVOICES, CONFIG.PROP_FOLDER_ARCHIVE_INVOICES, n);

  ui.alert('Recovery complete',
    'Vendor A: ' + rA.recovered.length + '\nVendor B: ' + rB.recovered.length + '\nInvoices: ' + rI.recovered.length,
    ui.ButtonSet.OK);
}

function recoverFromArchive_(props, inboxPropKey, archivePropKey, n) {
  const inboxId = props.getProperty(inboxPropKey);
  const archiveId = props.getProperty(archivePropKey);
  if (!inboxId || !archiveId) return { recovered: [] };
  const inbox = DriveApp.getFolderById(inboxId);
  const archive = DriveApp.getFolderById(archiveId);
  const files = [];
  const it = archive.getFiles();
  while (it.hasNext()) { const f = it.next(); files.push({ f, updatedAt: f.getLastUpdated() }); }
  files.sort((a, b) => b.updatedAt - a.updatedAt);
  const recovered = [];
  for (const entry of files.slice(0, n)) {
    const original = entry.f.getName().replace(/^\d{8}_\d{6}__/, '');
    try {
      entry.f.makeCopy(original, inbox);
      entry.f.setTrashed(true);
      recovered.push(original);
    } catch (e) {}
  }
  return { recovered };
}
