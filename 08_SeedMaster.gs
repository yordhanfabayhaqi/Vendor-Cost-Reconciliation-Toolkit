/**
 * 08_SeedMaster.gs
 * Auto-populates _master_cost_centers from unique canonical codes seen in ledger.
 * Also includes a one-time fix for legacy date-typed period columns.
 */

function menu_seedMasterBranches() {
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const branchSh = ss.getSheetByName(CONFIG.SHEET_BRANCHES);
  if (!branchSh) { ui.alert('Master cost centers sheet not found.'); return; }
  const ledgerSh = ss.getSheetByName(CONFIG.SHEET_LEDGER);
  if (!ledgerSh || ledgerSh.getLastRow() < 2) { ui.alert('Ledger empty.'); return; }

  const existingCount = branchSh.getLastRow() - 1;
  if (existingCount > 0 && ui.alert('Auto-populate', 'Master has ' + existingCount + ' rows. This ADDS new codes only. Continue?', ui.ButtonSet.YES_NO) !== ui.Button.YES) return;

  const result = seedMasterBranches_();
  ui.alert('Done', 'Codes found: ' + result.codesFound + '\nAdded: ' + result.added + '\nBreakdown:\n' +
    '• CC (Self default): ' + result.byType.CC + '\n' +
    '• Individual: ' + result.byType.Individual + '\n' +
    '• Trainer: ' + result.byType.Trainer + '\n' +
    '• Other: ' + result.byType.Other + '\n\n' +
    'NEXT: manually correct type, name, active_from, and eligibility flags in _master_cost_centers.',
    ui.ButtonSet.OK);
  ss.setActiveSheet(branchSh);
}

function seedMasterBranches_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const branchSh = ss.getSheetByName(CONFIG.SHEET_BRANCHES);
  const ledgerSh = ss.getSheetByName(CONFIG.SHEET_LEDGER);

  const existing = new Set();
  if (branchSh.getLastRow() > 1) {
    branchSh.getRange(2, 1, branchSh.getLastRow() - 1, 1).getValues().forEach(r => {
      if (r[0]) existing.add(String(r[0]).trim());
    });
  }

  const lh = ledgerSh.getRange(1, 1, 1, ledgerSh.getLastColumn()).getValues()[0];
  const idxCode = lh.indexOf('canonical_branch_code');
  const idxName = lh.indexOf('employee_name');
  const codesInLedger = new Map();
  ledgerSh.getRange(2, 1, ledgerSh.getLastRow() - 1, ledgerSh.getLastColumn()).getValues().forEach(row => {
    const code = String(row[idxCode] || '').trim();
    if (!code || code === 'HQ' || code.includes('|')) return;
    if (!codesInLedger.has(code)) {
      codesInLedger.set(code, { firstName: row[idxName] || '', count: 1 });
    } else {
      codesInLedger.get(code).count += 1;
    }
  });

  const sorted = Array.from(codesInLedger.keys()).sort((a, b) => {
    const aCC = a.match(/^CC-?\s*(\d+)$/i), bCC = b.match(/^CC-?\s*(\d+)$/i);
    if (aCC && bCC) return parseInt(aCC[1]) - parseInt(bCC[1]);
    if (aCC) return -1;
    if (bCC) return 1;
    return a.localeCompare(b);
  });

  const newRows = [];
  const byType = { CC: 0, Individual: 0, Trainer: 0, Other: 0 };
  for (const code of sorted) {
    if (existing.has(code)) continue;
    const info = codesInLedger.get(code);
    let type = 'Other', name = code, prorata = false, mgmt = false;
    if (/^CC-?\d+$/i.test(code)) {
      type = 'Self'; name = 'Cost Center ' + code;
      prorata = true; mgmt = true; byType.CC += 1;
    } else if (/Individual-?\d+/i.test(code)) {
      type = 'Contractor';
      name = info.firstName || code;
      byType.Individual += 1;
    } else if (code === 'TC') {
      type = 'Trainer'; name = 'Trainer / Training Center';
      byType.Trainer += 1;
    } else {
      byType.Other += 1;
    }
    newRows.push([code, name, type, '2024-01-01', '', prorata, mgmt,
      'Auto-seeded from ledger (' + info.count + ' txns)']);
  }

  if (newRows.length > 0) {
    branchSh.getRange(branchSh.getLastRow() + 1, 1, newRows.length, CONFIG.HEADERS_BRANCHES.length).setValues(newRows);
  }

  return { codesFound: codesInLedger.size, added: newRows.length, byType };
}

function menu_fixPeriodColumns() {
  const ui = SpreadsheetApp.getUi();
  if (ui.alert('Fix period columns', 'Converts _ledger and _invoices period columns from Date to text. Continue?', ui.ButtonSet.YES_NO) !== ui.Button.YES) return;

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let total = 0;
  total += fixPeriodColumn_(ss.getSheetByName(CONFIG.SHEET_LEDGER), 'period');
  total += fixPeriodColumn_(ss.getSheetByName(CONFIG.SHEET_INVOICES), 'period');
  ui.alert('Fixed ' + total + ' cells.');
}

function fixPeriodColumn_(sh, headerName) {
  if (!sh || sh.getLastRow() < 2) return 0;
  const headers = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0];
  const idx = headers.indexOf(headerName);
  if (idx < 0) return 0;
  const range = sh.getRange(2, idx + 1, sh.getLastRow() - 1, 1);
  const values = range.getValues();
  const newValues = values.map(([v]) => {
    if (v === '' || v === null || v === undefined) return [''];
    if (v instanceof Date) return [Utilities.formatDate(v, 'Asia/Jakarta', 'yyyy-MM')];
    if (typeof v === 'number') {
      const d = new Date(Math.round((v - 25569) * 86400 * 1000));
      return [Utilities.formatDate(d, 'Asia/Jakarta', 'yyyy-MM')];
    }
    return [String(v).trim()];
  });
  range.setNumberFormat('@');
  range.setValues(newValues);
  return newValues.length;
}
