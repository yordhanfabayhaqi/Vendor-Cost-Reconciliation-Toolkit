/**
 * 02_Setup.gs
 * Setup functions: config prompts, folder creation, master sheet init,
 * API key management, diagnostic + status displays.
 * All functions are idempotent — safe to run repeatedly.
 */

// ============================================================
// CONFIG PROMPTS
// ============================================================

function menu_setup() {
  const ui = SpreadsheetApp.getUi();
  const items = [
    { key: CONFIG.PROP_LLM_API_KEY,      label: 'LLM API key', mask: true },
    { key: CONFIG.PROP_EMAIL_RECIPIENTS, label: 'Notification email recipient(s), comma-separated' },
    { key: CONFIG.PROP_EMAIL_FROM_NAME,  label: 'Sender display name for notifications' }
  ];

  for (const item of items) {
    const current = PropertiesService.getScriptProperties().getProperty(item.key) || '(not set)';
    const shown = item.mask ? maskSecret_(current) : current;
    const response = ui.prompt(
      '⚙️ Setup: ' + item.label,
      'Current value: ' + shown + '\n\nEnter new value, or leave blank to keep current:',
      ui.ButtonSet.OK_CANCEL
    );
    if (response.getSelectedButton() !== ui.Button.OK) {
      ui.alert('Setup cancelled. Any changes so far are already saved.');
      return;
    }
    const val = response.getResponseText().trim();
    if (val) saveConfig_(item.key, val);
  }

  ui.alert(
    '⚙️ Setup complete',
    'Runtime config saved. Next steps:\n\n' +
    '1. ⚙️ Setup → Set parent folder for system\n' +
    '2. ⚙️ Setup → Initialize folder structure\n' +
    '3. ⚙️ Setup → Initialize master sheets\n' +
    '4. ⚙️ Setup → Test LLM API key',
    ui.ButtonSet.OK
  );
}

// ============================================================
// FOLDER STRUCTURE
// ============================================================

function menu_initFolders() {
  const ui = SpreadsheetApp.getUi();
  try {
    const result = setupFolders();
    ui.alert(
      'Folder structure ready',
      'Created or verified these folders in: "' + result.parentName + '"\n\n' +
      '📂 ' + CONFIG.FOLDER_ROOT_INBOX + '/\n' +
      '   ├─ ' + CONFIG.FOLDER_VENDOR_A_USAGE + '/\n' +
      '   ├─ ' + CONFIG.FOLDER_VENDOR_B_USAGE + '/\n' +
      '   └─ ' + CONFIG.FOLDER_INVOICES + '/\n\n' +
      '📂 ' + CONFIG.FOLDER_ROOT_ARCHIVE + '/\n' +
      '   ├─ ' + CONFIG.FOLDER_VENDOR_A_USAGE + '/\n' +
      '   ├─ ' + CONFIG.FOLDER_VENDOR_B_USAGE + '/\n' +
      '   └─ ' + CONFIG.FOLDER_INVOICES + '/\n\n' +
      'Folder IDs stored in Script Properties.',
      ui.ButtonSet.OK
    );
    log_('INFO', 'menu_initFolders', 'Folders initialized');
  } catch (e) {
    ui.alert('Error setting up folders', e.message + '\n\n' + (e.stack || ''), ui.ButtonSet.OK);
    log_('ERROR', 'menu_initFolders', e.message);
  }
}

function setupFolders() {
  const props = PropertiesService.getScriptProperties();
  let parent;

  const explicitParentId = props.getProperty(CONFIG.PROP_PARENT_FOLDER_ID);
  if (explicitParentId) {
    try {
      parent = DriveApp.getFolderById(explicitParentId);
    } catch (e) {
      throw new Error(
        'Cannot access stored parent folder (ID: ' + explicitParentId + '). ' +
        'Use ⚙️ Setup → "Set parent folder for system" to set a new one.\n\n' +
        'Original: ' + e.message
      );
    }
  } else {
    try {
      const file = DriveApp.getFileById(SpreadsheetApp.getActiveSpreadsheet().getId());
      const parents = file.getParents();
      if (!parents.hasNext()) {
        throw new Error('Spreadsheet has no parent folder. Move it into a folder first.');
      }
      parent = parents.next();
    } catch (e) {
      throw new Error(
        'Cannot read spreadsheet\'s parent folder. Use ⚙️ Setup → "Set parent folder for system".\n\n' +
        'Original: ' + e.message
      );
    }
  }

  let inboxRoot, archiveRoot;
  try {
    inboxRoot = getOrCreateFolder_(parent, CONFIG.FOLDER_ROOT_INBOX);
    archiveRoot = getOrCreateFolder_(parent, CONFIG.FOLDER_ROOT_ARCHIVE);
  } catch (e) {
    throw new Error(
      'Cannot create folders inside "' + parent.getName() + '". Check edit permissions.\n\n' +
      'Original: ' + e.message
    );
  }

  const inboxA = getOrCreateFolder_(inboxRoot, CONFIG.FOLDER_VENDOR_A_USAGE);
  const inboxB = getOrCreateFolder_(inboxRoot, CONFIG.FOLDER_VENDOR_B_USAGE);
  const inboxInv = getOrCreateFolder_(inboxRoot, CONFIG.FOLDER_INVOICES);
  const archiveA = getOrCreateFolder_(archiveRoot, CONFIG.FOLDER_VENDOR_A_USAGE);
  const archiveB = getOrCreateFolder_(archiveRoot, CONFIG.FOLDER_VENDOR_B_USAGE);
  const archiveInv = getOrCreateFolder_(archiveRoot, CONFIG.FOLDER_INVOICES);

  props.setProperties({
    [CONFIG.PROP_FOLDER_INBOX_VENDOR_A]: inboxA.getId(),
    [CONFIG.PROP_FOLDER_INBOX_VENDOR_B]: inboxB.getId(),
    [CONFIG.PROP_FOLDER_INBOX_INVOICES]: inboxInv.getId(),
    [CONFIG.PROP_FOLDER_ARCHIVE_VENDOR_A]: archiveA.getId(),
    [CONFIG.PROP_FOLDER_ARCHIVE_VENDOR_B]: archiveB.getId(),
    [CONFIG.PROP_FOLDER_ARCHIVE_INVOICES]: archiveInv.getId()
  });

  return {
    parentName: parent.getName(),
    inboxRootId: inboxRoot.getId(),
    archiveRootId: archiveRoot.getId()
  };
}

function getOrCreateFolder_(parent, name) {
  const existing = parent.getFoldersByName(name);
  if (existing.hasNext()) return existing.next();

  let lastErr;
  for (let attempt = 1; attempt <= 3; attempt++) {
    try {
      if (attempt > 1) Utilities.sleep(attempt * 500);
      return parent.createFolder(name);
    } catch (e) {
      lastErr = e;
    }
  }
  throw new Error(
    'Failed to create folder "' + name + '" inside "' + parent.getName() + '" after 3 attempts. ' +
    'Original: ' + (lastErr && lastErr.message)
  );
}

function menu_setParentFolder() {
  const ui = SpreadsheetApp.getUi();
  const response = ui.prompt(
    'Set parent folder',
    'Paste the URL or ID of a Drive folder where you have full edit access. ' +
    'The system will create inbox and archive subfolders inside it.',
    ui.ButtonSet.OK_CANCEL
  );
  if (response.getSelectedButton() !== ui.Button.OK) return;

  let input = response.getResponseText().trim();
  if (!input) { ui.alert('Empty input.'); return; }

  const urlMatch = input.match(/\/folders\/([a-zA-Z0-9_-]+)/);
  if (urlMatch) input = urlMatch[1];

  try {
    const folder = DriveApp.getFolderById(input);
    saveConfig_(CONFIG.PROP_PARENT_FOLDER_ID, input);
    ui.alert('✅ Parent folder set: "' + folder.getName() + '"\n\nNow run ⚙️ Setup → Initialize folder structure.');
  } catch (e) {
    ui.alert('❌ Cannot access folder. Check the ID/URL and permissions.\n\n' + e.message);
  }
}

function menu_diagnoseFolder() {
  const ui = SpreadsheetApp.getUi();
  const parentId = getConfig_(CONFIG.PROP_PARENT_FOLDER_ID, false);
  if (!parentId) { ui.alert('No parent folder set.'); return; }

  const lines = [];
  lines.push('Running as: ' + (Session.getActiveUser().getEmail() || '(unknown)'));
  lines.push('Parent folder ID: ' + parentId);
  lines.push('');

  let folder;
  try {
    folder = DriveApp.getFolderById(parentId);
    lines.push('✅ Can read folder: "' + folder.getName() + '"');
  } catch (e) {
    lines.push('❌ Cannot read folder: ' + e.message);
    ui.alert('Diagnostic', lines.join('\n'), ui.ButtonSet.OK);
    return;
  }

  try {
    const owner = folder.getOwner();
    lines.push('Owner: ' + (owner ? owner.getEmail() : '(no owner — likely Shared Drive)'));
  } catch (e) {
    lines.push('Owner: (cannot read)');
  }

  lines.push('');
  lines.push('--- Testing folder creation ---');
  try {
    const testName = '_test_' + Date.now();
    const test = folder.createFolder(testName);
    lines.push('✅ Created test folder');
    test.setTrashed(true);
    lines.push('✅ Trashed test folder');
    lines.push('');
    lines.push('VERDICT: Permissions are fine.');
  } catch (e) {
    lines.push('❌ Cannot create folder: ' + e.message);
  }

  ui.alert('Folder diagnostic', lines.join('\n'), ui.ButtonSet.OK);
}

function menu_openInboxFolder() {
  const ui = SpreadsheetApp.getUi();
  const inboxA = getConfig_(CONFIG.PROP_FOLDER_INBOX_VENDOR_A, false);
  if (!inboxA) { ui.alert('Folders not set up yet. Run ⚙️ Setup → Initialize folder structure.'); return; }
  const folder = DriveApp.getFolderById(inboxA).getParents().next();
  ui.alert('Inbox folder', 'Open this URL to upload files:\n\n' + folder.getUrl(), ui.ButtonSet.OK);
}

// ============================================================
// LLM API KEY MANAGEMENT
// ============================================================

function menu_testApiKey() {
  const ui = SpreadsheetApp.getUi();
  const key = getConfig_(CONFIG.PROP_LLM_API_KEY, false);
  if (!key) { ui.alert('No API key stored. Run ⚙️ Setup first.'); return; }

  try {
    const response = UrlFetchApp.fetch(CONFIG.LLM_API_URL, {
      method: 'POST',
      contentType: 'application/json',
      headers: {
        'x-api-key': key,
        'anthropic-version': CONFIG.LLM_API_VERSION
      },
      payload: JSON.stringify({
        model: CONFIG.MODEL_CHEAP,
        max_tokens: 20,
        messages: [{ role: 'user', content: 'Reply with exactly the word: pong' }]
      }),
      muteHttpExceptions: true
    });

    const code = response.getResponseCode();
    const body = response.getContentText();

    if (code === 200) {
      const parsed = JSON.parse(body);
      const reply = (parsed.content && parsed.content[0] && parsed.content[0].text) || '(no text)';
      ui.alert('✅ API key works', 'Response code: ' + code + '\nReply: ' + reply.trim(), ui.ButtonSet.OK);
    } else {
      ui.alert('❌ API call failed', 'HTTP ' + code + '\n\n' + body.substring(0, 500), ui.ButtonSet.OK);
    }
  } catch (e) {
    ui.alert('Error calling LLM API', e.message, ui.ButtonSet.OK);
  }
}

function menu_clearApiKey() {
  const ui = SpreadsheetApp.getUi();
  const confirm = ui.alert('Clear API key', 'Remove the stored LLM API key?', ui.ButtonSet.YES_NO);
  if (confirm !== ui.Button.YES) return;
  PropertiesService.getScriptProperties().deleteProperty(CONFIG.PROP_LLM_API_KEY);
  ui.alert('API key cleared.');
}

// ============================================================
// MASTER SHEETS
// ============================================================

function menu_initMasterSheets() {
  const ui = SpreadsheetApp.getUi();
  try {
    const created = initMasterSheets();
    if (created.length === 0) {
      ui.alert('All master sheets already exist.');
    } else {
      ui.alert('Master sheets initialized', 'Created:\n' + created.map(n => '• ' + n).join('\n'), ui.ButtonSet.OK);
    }
  } catch (e) {
    ui.alert('Error', e.message, ui.ButtonSet.OK);
  }
}

function initMasterSheets() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const created = [];
  const specs = [
    { name: CONFIG.SHEET_BRANCHES, headers: CONFIG.HEADERS_BRANCHES },
    { name: CONFIG.SHEET_ALIASES, headers: CONFIG.HEADERS_ALIASES },
    { name: CONFIG.SHEET_INVOICES, headers: CONFIG.HEADERS_INVOICES },
    { name: CONFIG.SHEET_LEDGER, headers: CONFIG.HEADERS_LEDGER },
    { name: CONFIG.SHEET_UNMAPPED, headers: CONFIG.HEADERS_UNMAPPED },
    { name: CONFIG.SHEET_LOG, headers: CONFIG.HEADERS_LOG }
  ];

  specs.forEach(spec => {
    if (!ss.getSheetByName(spec.name)) {
      const sh = ss.insertSheet(spec.name);
      sh.getRange(1, 1, 1, spec.headers.length).setValues([spec.headers]);
      sh.getRange(1, 1, 1, spec.headers.length).setFontWeight('bold').setBackground('#f0f0f0');
      sh.setFrozenRows(1);
      sh.autoResizeColumns(1, spec.headers.length);
      created.push(spec.name);
    }
  });
  return created;
}

// ============================================================
// STATUS + LOG
// ============================================================

function menu_showConfigStatus() {
  const ui = SpreadsheetApp.getUi();
  const props = PropertiesService.getScriptProperties();
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  const apiKey = props.getProperty(CONFIG.PROP_LLM_API_KEY);
  const emails = props.getProperty(CONFIG.PROP_EMAIL_RECIPIENTS) || '(not set)';

  const folderKeys = [
    ['Vendor A inbox', CONFIG.PROP_FOLDER_INBOX_VENDOR_A],
    ['Vendor B inbox', CONFIG.PROP_FOLDER_INBOX_VENDOR_B],
    ['Invoices inbox', CONFIG.PROP_FOLDER_INBOX_INVOICES],
    ['Vendor A archive', CONFIG.PROP_FOLDER_ARCHIVE_VENDOR_A],
    ['Vendor B archive', CONFIG.PROP_FOLDER_ARCHIVE_VENDOR_B],
    ['Invoices archive', CONFIG.PROP_FOLDER_ARCHIVE_INVOICES]
  ];
  const folderStatus = folderKeys
    .map(([label, key]) => (props.getProperty(key) ? '✅' : '❌') + ' ' + label)
    .join('\n');

  const sheets = [
    CONFIG.SHEET_BRANCHES, CONFIG.SHEET_ALIASES, CONFIG.SHEET_INVOICES,
    CONFIG.SHEET_LEDGER, CONFIG.SHEET_UNMAPPED, CONFIG.SHEET_LOG
  ];
  const sheetStatus = sheets.map(n => (ss.getSheetByName(n) ? '✅' : '❌') + ' ' + n).join('\n');

  ui.alert(
    'Config status',
    '📧 Notifications: ' + emails + '\n\n' +
    '🔑 LLM API key: ' + (apiKey ? '✅ Set (' + maskSecret_(apiKey) + ')' : '❌ Not set') + '\n\n' +
    '📂 Drive folders:\n' + folderStatus + '\n\n' +
    '📋 System sheets:\n' + sheetStatus,
    ui.ButtonSet.OK
  );
}

function log_(level, fn, message) {
  Logger.log('[' + level + '] ' + fn + ': ' + message);
  try {
    const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG.SHEET_LOG);
    if (sh) sh.appendRow([new Date(), level, fn, message]);
  } catch (e) {
    Logger.log('log_ failed: ' + e.message);
  }
}
