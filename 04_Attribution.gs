/**
 * 04_Attribution.gs
 * Maps raw employee IDs to canonical cost-center codes.
 * 
 * Multi-cost-center assignment: raw IDs containing 2+ CC codes
 * (e.g. "CC-012 CC-035") are stored as pipe-separated "CC-012|CC-035".
 * The report layer splits cost evenly across the listed cost centers.
 * 
 * The regex patterns below assume codes of the form "CC-NNN" and
 * "Individual-NNN". Adapt to your own coding scheme.
 */

let _ALIAS_CACHE = null;

function attributeBranch_(employeeRawId) {
  if (!employeeRawId) return { code: null, isNew: false };
  const normalized = String(employeeRawId).trim();
  if (!normalized) return { code: null, isNew: false };

  const aliases = getAliasesMap_();
  if (aliases.has(normalized)) {
    const mapped = aliases.get(normalized);
    if (mapped === '_IGNORED_') return { code: null, isNew: false };
    return { code: mapped, isNew: false };
  }

  const parsed = tryParseBranchCode_(normalized);
  if (parsed) {
    addAlias_(normalized, parsed, 'auto-parsed');
    return { code: parsed, isNew: false };
  }

  return { code: null, isNew: true };
}

function tryParseBranchCode_(raw) {
  const trimmed = String(raw).trim();
  if (!trimmed) return null;

  if (/^HQ$/i.test(trimmed)) return 'HQ';
  if (/Trainer/i.test(trimmed) || /Trainee/i.test(trimmed)) return 'TC';

  // Individual contractor pattern: "Role - Individual-NNN"
  let m = trimmed.match(/^(RoleA|RoleB)\s*-\s*Individual\s*-\s*0*(\d+)$/i);
  if (m) {
    const role = m[1].charAt(0).toUpperCase() + m[1].slice(1).toLowerCase();
    return role + ' - Individual-' + String(m[2]).padStart(3, '0');
  }

  // Multi-CC pattern: two or more "CC-NNN" references
  const ccMatches = trimmed.match(/CC-?\s*0*\d+/gi);
  if (ccMatches && ccMatches.length >= 2) {
    const codes = Array.from(new Set(
      ccMatches.map(s => 'CC-' + String(s.match(/(\d+)/)[1]).padStart(3, '0'))
    )).sort((a, b) => parseInt(a.match(/(\d+)/)[1]) - parseInt(b.match(/(\d+)/)[1]));
    if (codes.length >= 2) return codes.join('|');
  }

  // Single CC
  m = trimmed.match(/CC-?\s*0*(\d+)/i);
  if (m) return 'CC-' + String(m[1]).padStart(3, '0');

  return null;
}

function normalizeMultiCode_(code) {
  const s = String(code || '').trim();
  if (!s) return s;

  if (s.includes('|')) {
    const parts = s.split('|').map(p => p.trim()).filter(p => p.length > 0);
    const normalized = Array.from(new Set(parts.map(p => {
      const m = p.match(/^CC-?\s*0*(\d+)$/i);
      return m ? 'CC-' + String(m[1]).padStart(3, '0') : p;
    })));
    normalized.sort((a, b) => {
      const na = a.match(/CC-?\s*(\d+)/i), nb = b.match(/CC-?\s*(\d+)/i);
      if (na && nb) return parseInt(na[1]) - parseInt(nb[1]);
      return a.localeCompare(b);
    });
    return normalized.length === 1 ? normalized[0] : normalized.join('|');
  }

  const ccMatches = s.match(/CC-?\s*0*\d+/gi);
  if (ccMatches && ccMatches.length >= 2) {
    const codes = Array.from(new Set(
      ccMatches.map(m => 'CC-' + String(m.match(/(\d+)/)[1]).padStart(3, '0'))
    )).sort((a, b) => parseInt(a.match(/(\d+)/)[1]) - parseInt(b.match(/(\d+)/)[1]));
    if (codes.length >= 2) return codes.join('|');
  }

  const singleCc = s.match(/^CC-?\s*0*(\d+)$/i);
  if (singleCc) return 'CC-' + String(singleCc[1]).padStart(3, '0');

  return s;
}

function getAliasesMap_() {
  if (_ALIAS_CACHE !== null) return _ALIAS_CACHE;
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG.SHEET_ALIASES);
  if (!sh) throw new Error('Aliases sheet not found.');
  const map = new Map();
  const lastRow = sh.getLastRow();
  if (lastRow > 1) {
    sh.getRange(2, 1, lastRow - 1, 2).getValues().forEach(row => {
      const raw = String(row[0]).trim();
      const code = normalizeMultiCode_(String(row[1]).trim());
      if (raw && code) map.set(raw, code);
    });
  }
  _ALIAS_CACHE = map;
  return map;
}

function addAlias_(rawId, canonicalCode, notes) {
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG.SHEET_ALIASES);
  if (!sh) throw new Error('Aliases sheet not found.');
  const normalized = normalizeMultiCode_(canonicalCode);
  sh.appendRow([rawId, normalized, new Date(), notes || '']);
  if (_ALIAS_CACHE !== null) _ALIAS_CACHE.set(rawId, normalized);
}

function clearAliasCache_() { _ALIAS_CACHE = null; }

function trackUnmapped_(rows, provider) {
  if (rows.length === 0) return 0;
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG.SHEET_UNMAPPED);
  if (!sh) throw new Error('Unmapped sheet not found.');

  const batchInfo = new Map();
  for (const r of rows) {
    const raw = String(r.employee_raw_id || '').trim();
    if (!raw) continue;
    if (!batchInfo.has(raw)) {
      batchInfo.set(raw, { count: 0, sampleName: r.employee_name || '', sampleAmount: r.total_amount || 0 });
    }
    batchInfo.get(raw).count += 1;
  }

  const existing = new Map();
  const lastRow = sh.getLastRow();
  if (lastRow > 1) {
    sh.getRange(2, 1, lastRow - 1, CONFIG.HEADERS_UNMAPPED.length).getValues().forEach((row, idx) => {
      if (row[0]) existing.set(String(row[0]).trim(), { rowIdx: idx + 2, count: Number(row[4]) || 0 });
    });
  }

  let added = 0;
  const newEntries = [];
  const updates = [];
  batchInfo.forEach((info, raw) => {
    if (existing.has(raw)) {
      updates.push({ rowIdx: existing.get(raw).rowIdx, newCount: existing.get(raw).count + info.count });
    } else {
      newEntries.push([raw, info.sampleName, provider, new Date(), info.count, info.sampleAmount]);
      added += 1;
    }
  });

  if (newEntries.length > 0) sh.getRange(sh.getLastRow() + 1, 1, newEntries.length, CONFIG.HEADERS_UNMAPPED.length).setValues(newEntries);
  for (const u of updates) sh.getRange(u.rowIdx, 5).setValue(u.newCount);
  return added;
}

function menu_reAttributeLedger() {
  const ui = SpreadsheetApp.getUi();
  if (ui.alert('Re-attribute ledger', 'Re-runs branch attribution on all _ledger rows with empty canonical_branch_code. Continue?', ui.ButtonSet.YES_NO) !== ui.Button.YES) return;
  clearAliasCache_();
  const result = reAttributeLedger_();
  ui.alert('Done', 'Scanned: ' + result.scanned + '\nFilled: ' + result.filled + '\nStill unmapped: ' + result.stillUnmapped, ui.ButtonSet.OK);
}

function reAttributeLedger_() {
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG.SHEET_LEDGER);
  if (!sh) throw new Error('Ledger not found.');
  const lastRow = sh.getLastRow();
  if (lastRow < 2) return { scanned: 0, filled: 0, stillUnmapped: 0 };
  const headers = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0];
  const idxRawId = headers.indexOf('employee_raw_id');
  const idxCode = headers.indexOf('canonical_branch_code');
  const data = sh.getRange(2, 1,
