/**
 * 05_InvoiceParser.gs
 * Extracts structured billing data from invoice PDFs using the LLM API.
 * 
 * IMPORTANT: the parser prompt below describes a GENERIC three-format invoice
 * pattern (Vendor A cost passthrough, Vendor A management fee with VAT, and
 * Vendor B combined). Adapt the prompt to your real invoice formats before use.
 */

function menu_processInvoicePdfs() {
  const ui = SpreadsheetApp.getUi();
  try { getConfig_(CONFIG.PROP_LLM_API_KEY); }
  catch (e) { ui.alert('LLM API key not set. Run ⚙️ Setup first.'); return; }

  try {
    const result = processInvoicePdfs_();
    let msg = '📄 Processed: ' + result.filesProcessed + ' PDF(s)\n' +
      '📋 Invoice rows added: ' + result.rowsAdded + '\n' +
      '📋 Skipped duplicates: ' + result.rowsSkipped;
    if (result.errors.length > 0) msg += '\n\n⚠️ Errors:\n' + result.errors.map(e => '• ' + e).join('\n');
    if (result.filesProcessed === 0) msg = 'No PDFs found in invoices inbox.';
    ui.alert('Invoice processing complete', msg, ui.ButtonSet.OK);
  } catch (e) {
    ui.alert('Failed', e.message, ui.ButtonSet.OK);
  }
}

function processInvoicePdfs_() {
  const inboxId = getConfig_(CONFIG.PROP_FOLDER_INBOX_INVOICES);
  const archiveId = getConfig_(CONFIG.PROP_FOLDER_ARCHIVE_INVOICES);
  const inbox = DriveApp.getFolderById(inboxId);
  const archive = DriveApp.getFolderById(archiveId);
  const errors = [];
  let filesProcessed = 0, rowsAdded = 0, rowsSkipped = 0;

  const files = inbox.getFiles();
  while (files.hasNext()) {
    const file = files.next();
    if (!file.getName().toLowerCase().endsWith('.pdf')) continue;
    try {
      const extracted = extractInvoiceFromPdf_(file);
      const stats = appendInvoiceRows_(extracted, file);
      rowsAdded += stats.added; rowsSkipped += stats.skipped; filesProcessed += 1;
      try { archiveFile_(file, archive); }
      catch (archiveErr) { errors.push('"' + file.getName() + '" parsed OK but archive failed.'); }
    } catch (e) { errors.push('"' + file.getName() + '": ' + e.message); }
  }
  return { filesProcessed, rowsAdded, rowsSkipped, errors };
}

function extractInvoiceFromPdf_(file) {
  const blob = file.getBlob();
  const base64Pdf = Utilities.base64Encode(blob.getBytes());

  // Generic prompt — adapt to your real invoice formats.
  const systemPrompt =
    'You extract structured billing data from vendor invoice PDFs. ' +
    'Return ONLY valid JSON, no markdown fences, no commentary. ' +
    'Three example formats:\n' +
    '1. Vendor A cost passthrough — no tax, splits by vertical (Transport / Logistik).\n' +
    '2. Vendor A management fee — with VAT, splits by vertical.\n' +
    '3. Vendor B combined — single reimbursement line + management fee + VAT.\n\n' +
    'Output JSON schema:\n' +
    '{\n' +
    '  "provider": "VendorA" | "VendorB",\n' +
    '  "invoice_number": string,\n' +
    '  "invoice_date": "YYYY-MM-DD",\n' +
    '  "due_date": "YYYY-MM-DD or null",\n' +
    '  "period": "YYYY-MM",\n' +
    '  "type": "cost" | "mgmt_fee" | "combined",\n' +
    '  "lines": [ { "vertical": "Transport"|"Logistik"|"Combined", "amount": number, "ppn": number, "total": number } ]\n' +
    '}';

  const response = UrlFetchApp.fetch(CONFIG.LLM_API_URL, {
    method: 'POST',
    contentType: 'application/json',
    headers: { 'x-api-key': getConfig_(CONFIG.PROP_LLM_API_KEY), 'anthropic-version': CONFIG.LLM_API_VERSION },
    payload: JSON.stringify({
      model: CONFIG.MODEL_CHEAP, max_tokens: 1024, system: systemPrompt,
      messages: [{ role: 'user', content: [
        { type: 'document', source: { type: 'base64', media_type: 'application/pdf', data: base64Pdf } },
        { type: 'text', text: 'Extract this invoice. Return only the JSON.' }
      ]}]
    }),
    muteHttpExceptions: true
  });

  if (response.getResponseCode() !== 200) {
    throw new Error('LLM API HTTP ' + response.getResponseCode() + ': ' + response.getContentText().substring(0, 300));
  }

  const parsed = JSON.parse(response.getContentText());
  let text = parsed.content[0].text.trim().replace(/^```(?:json)?\s*/i, '').replace(/\s*```$/, '').trim();
  return JSON.parse(text);
}

function appendInvoiceRows_(extracted, file) {
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG.SHEET_INVOICES);
  if (!sh) throw new Error('Invoices sheet not found.');

  const existing = new Set();
  const lastRow = sh.getLastRow();
  if (lastRow > 1) {
    sh.getRange(2, 1, lastRow - 1, 5).getValues().forEach(r => {
      existing.add(r[0] + '|' + r[1] + '|' + r[2] + '|' + r[4]);
    });
  }

  const newRows = [];
  const lines = extracted.lines || [];
  for (const line of lines) {
    let type = extracted.type;
    if (extracted.type === 'combined') {
      type = (lines.indexOf(line) === 0) ? 'cost' : 'mgmt_fee';
    }
    const key = extracted.period + '|' + extracted.provider + '|' + (line.vertical || 'Combined') + '|' + type;
    if (existing.has(key)) continue;
    newRows.push([
      extracted.period, extracted.provider, line.vertical || 'Combined',
      extracted.invoice_number, type, line.amount, line.ppn || 0, line.total || (line.amount + (line.ppn || 0)),
      extracted.invoice_date, extracted.due_date || '', file.getName(), file.getId(), new Date()
    ]);
    existing.add(key);
  }

  if (newRows.length > 0) {
    const startRow = sh.getLastRow() + 1;
    sh.getRange(startRow, 1, newRows.length, CONFIG.HEADERS_INVOICES.length).setValues(newRows);
    sh.getRange(startRow, 1, newRows.length, 1).setNumberFormat('@');
  }
  return { added: newRows.length, skipped: lines.length - newRows.length };
}
