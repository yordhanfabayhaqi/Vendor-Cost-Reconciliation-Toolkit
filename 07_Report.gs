/**
 * 07_Report.gs
 * Generates per-cost-center report grouped by invoice number.
 * Each invoice = one column = one downstream submission.
 * Multi-cost-center codes (pipe-separated) are split evenly.
 */

function menu_generateReport() {
  const ui = SpreadsheetApp.getUi();
  const resp = ui.prompt('Generate Report', 'Enter YYYY-MM:', ui.ButtonSet.OK_CANCEL);
  if (resp.getSelectedButton() !== ui.Button.OK) return;
  const period = resp.getResponseText().trim();
  if (!/^\d{4}-\d{2}$/.test(period)) { ui.alert('Invalid format.'); return; }

  try {
    const result = generateReport_(period);
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    ss.setActiveSheet(ss.getSheetByName(result.sheetName));
    let msg = '✅ Report: ' + result.sheetName + '\n\n';
    result.invoiceColumns.forEach(c => msg += '• ' + c.label + ': ' + Math.round(c.totalDistributed).toLocaleString('en-US') + '\n');
    if (result.warnings.length > 0) msg += '\n⚠️\n' + result.warnings.map(w => '• ' + w).join('\n');
    ui.alert('Report ready', msg, ui.ButtonSet.OK);
  } catch (e) {
    ui.alert('Failed', e.message, ui.ButtonSet.OK);
  }
}

function generateReport_(period) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  // Load master cost centers
  const branchSh = ss.getSheetByName(CONFIG.SHEET_BRANCHES);
  if (!branchSh || branchSh.getLastRow() < 2) throw new Error('_master_cost_centers empty.');
  const bh = branchSh.getRange(1, 1, 1, branchSh.getLastColumn()).getValues()[0];
  const idxB = {
    code: bh.indexOf('code'), name: bh.indexOf('name'), type: bh.indexOf('type'),
    activeFrom: bh.indexOf('active_from'), activeTo: bh.indexOf('active_to'),
    prorata: bh.indexOf('prorata_eligible'), mgmt: bh.indexOf('mgmt_fee_eligible')
  };
  const branchRows = branchSh.getRange(2, 1, branchSh.getLastRow() - 1, branchSh.getLastColumn()).getValues();

  const periodStart = new Date(parseInt(period.substring(0, 4)), parseInt(period.substring(5, 7)) - 1, 1);
  const periodEnd = new Date(periodStart.getFullYear(), periodStart.getMonth() + 1, 0);

  const activeBranches = [];
  for (const row of branchRows) {
    const code = String(row[idxB.code] || '').trim();
    if (!code) continue;
    const af = parseDateCell_(row[idxB.activeFrom]);
    const at = parseDateCell_(row[idxB.activeTo]);
    if (af && af > periodEnd) continue;
    if (at && at < periodStart) continue;
    activeBranches.push({
      code, name: String(row[idxB.name] || code).trim(),
      type: String(row[idxB.type] || '').trim(),
      prorata_eligible: row[idxB.prorata] === true || String(row[idxB.prorata]).toUpperCase() === 'TRUE',
      mgmt_fee_eligible: row[idxB.mgmt] === true || String(row[idxB.mgmt]).toUpperCase() === 'TRUE'
    });
  }

  // Load ledger
  const ledgerSh = ss.getSheetByName(CONFIG.SHEET_LEDGER);
  const lh = ledgerSh.getRange(1, 1, 1, ledgerSh.getLastColumn()).getValues()[0];
  const idxL = {
    period: lh.indexOf('period'), provider: lh.indexOf('provider'),
    code: lh.indexOf('canonical_branch_code'), employee: lh.indexOf('employee_name'),
    amount: lh.indexOf('total_amount'), excluded: lh.indexOf('excluded')
  };

  const sumByCodeProvider = new Map();
  const employeeNameByCode = new Map();
  let unmappedSum = 0;
  const hqByProvider = new Map();

  if (ledgerSh.getLastRow() > 1) {
    ledgerSh.getRange(2, 1, ledgerSh.getLastRow() - 1, ledgerSh.getLastColumn()).getValues().forEach(row => {
      if (String(row[idxL.period]).trim() !== period) return;
      if (idxL.excluded >= 0 && row[idxL.excluded] === true) return;
      const code = String(row[idxL.code] || '').trim();
      const provider = String(row[idxL.provider] || '').trim();
      const amt = Number(row[idxL.amount]) || 0;
      if (!code) { unmappedSum += amt; return; }
      if (code === 'HQ') { hqByProvider.set(provider, (hqByProvider.get(provider) || 0) + amt); return; }
      const codes = code.includes('|') ? code.split('|').map(c => c.trim()).filter(c => c) : [code];
      const splitAmt = amt / codes.length;
      for (const c of codes) {
        sumByCodeProvider.set(c + '|' + provider, (sumByCodeProvider.get(c + '|' + provider) || 0) + splitAmt);
        if (!employeeNameByCode.has(c) && row[idxL.employee]) employeeNameByCode.set(c, row[idxL.employee]);
      }
    });
  }

  // Load invoices, group by invoice number
  const invSh = ss.getSheetByName(CONFIG.SHEET_INVOICES);
  const invData = invSh && invSh.getLastRow() > 1
    ? invSh.getRange(2, 1, invSh.getLastRow() - 1, CONFIG.HEADERS_INVOICES.length).getValues() : [];

  const invoicesByNumber = new Map();
  for (const row of invData) {
    if (String(row[0]).trim() !== period) continue;
    const invNo = String(row[3]).trim();
    if (!invoicesByNumber.has(invNo)) {
      invoicesByNumber.set(invNo, {
        invoiceNumber: invNo, provider: String(row[1]).trim(),
        types: new Set(), verticals: new Set(),
        costAmount: 0, mgmtFeeWithPpn: 0
      });
    }
    const inv = invoicesByNumber.get(invNo);
    inv.types.add(String(row[4]).trim());
    inv.verticals.add(String(row[2]).trim());
    if (row[4] === 'cost') inv.costAmount += Number(row[5]) || 0;
    if (row[4] === 'mgmt_fee') inv.mgmtFeeWithPpn += Number(row[7]) || 0;
  }

  const invoiceColumns = [];
  invoicesByNumber.forEach(inv => {
    const vLabel = Array.from(inv.verticals).join(' & ');
    let label;
    if (inv.types.size === 1) {
      const t = Array.from(inv.types)[0];
      label = inv.provider + ' ' + vLabel + ' ' + (t === 'cost' ? 'Cost' : 'Mgmt Fee');
    } else {
      label = inv.provider + ' ' + vLabel + ' Cost & Mgmt Fee';
    }
    invoiceColumns.push({
      invoiceNumber: inv.invoiceNumber, provider: inv.provider, types: inv.types,
      label, costAmount: inv.costAmount, mgmtFeeWithPpn: inv.mgmtFeeWithPpn
    });
  });

  invoiceColumns.sort((a, b) => {
    const provOrder = { 'VendorA': 1, 'VendorB': 2 };
    const ap = provOrder[a.provider] || 9, bp = provOrder[b.provider] || 9;
    if (ap !== bp) return ap - bp;
    const score = c => c.types.has('cost') && c.types.has('mgmt_fee') ? 3 : c.types.has('cost') ? 1 : 2;
    return score(a) - score(b);
  });

  const eligibleProrata = activeBranches.filter(b => b.prorata_eligible);
  const eligibleMgmt = activeBranches.filter(b => b.mgmt_fee_eligible);
  const activeProrataCount = eligibleProrata.length;
  const activeMgmtCount = eligibleMgmt.length;
  if (activeProrataCount === 0) throw new Error('No prorata-eligible cost centers active in ' + period + '.');

  const typeOrder = { 'Self': 1, 'Partner': 2, 'Franchise': 3, 'Contractor': 4, 'Trainer': 5, 'Other': 6 };
  const sortedBranches = activeBranches.slice().sort((a, b) => {
    const ao = typeOrder[a.type] || 9, bo = typeOrder[b.type] || 9;
    if (ao !== bo) return ao - bo;
    return a.code.localeCompare(b.code, undefined, { numeric: true });
  });

  const fixedCols = ['Cost Center', 'Name'];
  const headerRow1 = fixedCols.concat(invoiceColumns.map(c => c.invoiceNumber));
  const headerRow2 = ['', ''].concat(invoiceColumns.map(c => c.label));
  const rows = [headerRow1, headerRow2];
  const colTotals = invoiceColumns.map(() => 0);

  for (const b of sortedBranches) {
    const rowValues = [b.code, b.name];
    let rowHasAnyValue = (b.type === 'Self' || b.type === 'Partner');
    for (let ci = 0; ci < invoiceColumns.length; ci++) {
      const col = invoiceColumns[ci];
      let value = 0;
      const ownCost = sumByCodeProvider.get(b.code + '|' + col.provider) || 0;
      const hqPool = hqByProvider.get(col.provider) || 0;
      const prorataShare = b.prorata_eligible ? hqPool / activeProrataCount : 0;
      const mgmtShare = b.mgmt_fee_eligible && col.mgmtFeeWithPpn > 0 ? col.mgmtFeeWithPpn / activeMgmtCount : 0;
      if (col.types.has('cost') && col.types.has('mgmt_fee')) value = ownCost + prorataShare + mgmtShare;
      else if (col.types.has('cost')) value = ownCost + prorataShare;
      else if (col.types.has('mgmt_fee')) value = mgmtShare;
      if (value > 0) rowHasAnyValue = true;
      rowValues.push(value);
      colTotals[ci] += value;
    }
    if (rowHasAnyValue) rows.push(rowValues);
  }

  rows.push(['TOTAL', '(verification)'].concat(colTotals));

  const warnings = [];
  if (unmappedSum > 0) warnings.push('Unmapped in ledger: ' + unmappedSum.toLocaleString('en-US'));
  const masterCodes = new Set(activeBranches.map(b => b.code));
  sumByCodeProvider.forEach((_, key) => {
    const code = key.split('|')[0];
    if (!masterCodes.has(code)) warnings.push('Code "' + code + '" in ledger but not in master.');
  });

  const mmyy = period.substring(5, 7) + period.substring(2, 4);
  const sheetName = 'Report ' + mmyy;
  let outSh = ss.getSheetByName(sheetName);
  if (outSh) ss.deleteSheet(outSh);
  outSh = ss.insertSheet(sheetName);

  const numCols = headerRow1.length;
  outSh.getRange(1, 1, rows.length, numCols).setValues(rows);
  outSh.getRange(1, 1, 2, numCols).setFontWeight('bold').setBackground('#b6d7a8');
  outSh.getRange(2, 1, 1, numCols).setFontStyle('italic');
  outSh.setFrozenRows(2); outSh.setFrozenColumns(2);
  if (rows.length > 2) outSh.getRange(3, 3, rows.length - 2, numCols - 2).setNumberFormat('#,##0');
  outSh.getRange(rows.length, 1, 1, numCols).setFontWeight('bold').setBackground('#fff2cc');
  outSh.autoResizeColumns(1, numCols);

  return {
    sheetName,
    invoiceColumns: invoiceColumns.map((c, i) => ({ invoiceNumber: c.invoiceNumber, label: c.label, totalDistributed: colTotals[i] })),
    warnings
  };
}

function parseDateCell_(value) {
  if (!value) return null;
  if (value instanceof Date) return value;
  if (typeof value === 'string') {
    const d = new Date(value.trim());
    return isNaN(d.getTime()) ? null : d;
  }
  if (typeof value === 'number') return new Date(Math.round((value - 25569) * 86400 * 1000));
  return null;
}
