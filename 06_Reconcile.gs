/**
 * 06_Reconcile.gs
 * Reconciles ledger sums against invoice totals per period+provider+vertical.
 * Surfaces boundary transactions when gaps exist.
 */

const EXCLUDED_HEADER = 'excluded';

function menu_reconcilePeriod() {
  const ui = SpreadsheetApp.getUi();
  const resp = ui.prompt('Reconcile period', 'Enter YYYY-MM (e.g. 2026-03):', ui.ButtonSet.OK_CANCEL);
  if (resp.getSelectedButton() !== ui.Button.OK) return;
  const period = resp.getResponseText().trim();
  if (!/^\d{4}-\d{2}$/.test(period)) { ui.alert('Invalid format.'); return; }

  ensureExcludedColumn_();
  const result = reconcilePeriod_(period);
  let msg = 'Reconciliation report for ' + period + '\n\n';
  for (const seg of result.segments) {
    const status = seg.gap === 0 ? '✅' : '❌';
    msg += status + ' ' + seg.provider + ' ' + seg.vertical + ' (' + seg.type + ')\n';
    msg += '   Ledger:  ' + seg.ledger_sum.toLocaleString('en-US') + '\n';
    msg += '   Invoice: ' + seg.invoice_amount.toLocaleString('en-US') + '\n';
    if (seg.gap !== 0) msg += '   Gap: ' + Math.abs(seg.gap).toLocaleString('en-US') + '\n';
    msg += '\n';
  }
  if (result.missingInvoices.length > 0) msg += '⚠️ Missing invoices for:\n' + result.missingInvoices.map(m => '   • ' + m).join('\n');
  ui.alert('Reconciliation', msg, ui.ButtonSet.OK);
}

function ensureExcludedColumn_() {
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG.SHEET_LEDGER);
  if (!sh) throw new Error('Ledger not found.');
  const headers = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0];
  if (headers.indexOf(EXCLUDED_HEADER) === -1) {
    const newCol = sh.getLastColumn() + 1;
    sh.getRange(1, newCol).setValue(EXCLUDED_HEADER).setFontWeight('bold').setBackground('#f0f0f0');
  }
}

function reconcilePeriod_(period) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ledgerSh = ss.getSheetByName(CONFIG.SHEET_LEDGER);
  const headers = ledgerSh.getRange(1, 1, 1, ledgerSh.getLastColumn()).getValues()[0];
  const idx = {
    period: headers.indexOf('period'), provider: headers.indexOf('provider'),
    vertical: headers.indexOf('vertical'), amount: headers.indexOf('total_amount'),
    excluded: headers.indexOf(EXCLUDED_HEADER)
  };

  const invSh = ss.getSheetByName(CONFIG.SHEET_INVOICES);
  const invData = invSh.getLastRow() > 1
    ? invSh.getRange(2, 1, invSh.getLastRow() - 1, CONFIG.HEADERS_INVOICES.length).getValues() : [];

  const ledgerSums = new Map();
  if (ledgerSh.getLastRow() > 1) {
    ledgerSh.getRange(2, 1, ledgerSh.getLastRow() - 1, ledgerSh.getLastColumn()).getValues().forEach(row => {
      if (String(row[idx.period]).trim() !== period) return;
      if (idx.excluded >= 0 && row[idx.excluded] === true) return;
      const key = row[idx.provider] + '|' + row[idx.vertical];
      ledgerSums.set(key, (ledgerSums.get(key) || 0) + (Number(row[idx.amount]) || 0));
    });
  }

  const segments = [], missingInvoices = [], seen = new Set();
  for (const inv of invData) {
    if (String(inv[0]).trim() !== period || inv[4] !== 'cost') continue;
    const key = inv[1] + '|' + inv[2];
    if (seen.has(key)) continue;
    seen.add(key);
    let ledgerSum;
    if (inv[2] === 'Combined') {
      ledgerSum = (ledgerSums.get(inv[1] + '|Transport') || 0) + (ledgerSums.get(inv[1] + '|Logistik') || 0);
    } else {
      ledgerSum = ledgerSums.get(inv[1] + '|' + inv[2]) || 0;
    }
    segments.push({
      provider: inv[1], vertical: inv[2], type: inv[4],
      ledger_sum: ledgerSum, invoice_amount: Number(inv[5]) || 0,
      gap: ledgerSum - (Number(inv[5]) || 0)
    });
  }

  for (const [key, sum] of ledgerSums.entries()) {
    const [provider, vertical] = key.split('|');
    const matched = segments.some(s => s.provider === provider && (s.vertical === vertical || s.vertical === 'Combined'));
    if (!matched && sum > 0) missingInvoices.push(provider + ' ' + vertical + ' (' + sum.toLocaleString('en-US') + ')');
  }

  return { segments, missingInvoices, hasGaps: segments.some(s => s.gap !== 0) };
}

function menu_resolveBoundary() {
  const ui = SpreadsheetApp.getUi();
  ui.alert('Not yet available in this template. See README.', ui.ButtonSet.OK);
}

function menu_showReconStatus() {
  const ui = SpreadsheetApp.getUi();
  const ledgerSh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG.SHEET_LEDGER);
  if (!ledgerSh || ledgerSh.getLastRow() < 2) { ui.alert('Ledger empty.'); return; }
  const headers = ledgerSh.getRange(1, 1, 1, ledgerSh.getLastColumn()).getValues()[0];
  const idxPeriod = headers.indexOf('period');
  const periods = new Set();
  ledgerSh.getRange(2, 1, ledgerSh.getLastRow() - 1, ledgerSh.getLastColumn()).getValues().forEach(row => {
    if (row[idxPeriod]) periods.add(String(row[idxPeriod]).trim());
  });
  ensureExcludedColumn_();
  let msg = 'Periods in ledger:\n\n';
  for (const p of Array.from(periods).sort().reverse().slice(0, 12)) {
    const r = reconcilePeriod_(p);
    const status = r.segments.length === 0 ? '⏳ no invoice' : (r.hasGaps ? '❌ gaps' : '✅ reconciled');
    msg += p + ': ' + status + '\n';
  }
  ui.alert('Reconciliation status', msg, ui.ButtonSet.OK);
}
