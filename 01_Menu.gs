/**
 * 01_Menu.gs
 * Builds the toolkit menu on spreadsheet open.
 */

function onOpen() {
  const ui = SpreadsheetApp.getUi();

  ui.createMenu('Vendor Cost Toolkit')
    .addItem('📖 How to Run Me — a Tutorial', 'menu_showTutorial')
    .addSeparator()
    .addSubMenu(ui.createMenu('⚙️ Setup')
      .addItem('🔧 Initialize folder structure', 'menu_initFolders')
      .addItem('📌 Set parent folder for system', 'menu_setParentFolder')
      .addItem('🩺 Diagnose parent folder', 'menu_diagnoseFolder')
      .addItem('📋 Initialize master sheets', 'menu_initMasterSheets')
      .addItem('🌱 Auto-populate cost centers from ledger', 'menu_seedMasterBranches')
      .addItem('🔧 Fix period column types (run once)', 'menu_fixPeriodColumns')
      .addSeparator()
      .addItem('🔑 Configure runtime settings (keys, IDs, recipients)', 'menu_setup')
      .addItem('🔑 Test LLM API key', 'menu_testApiKey')
      .addItem('🔑 Clear LLM API key', 'menu_clearApiKey')
      .addSeparator()
      .addItem('🔌 Verify Drive API enabled', 'menu_verifyDriveApi')
      .addItem('ℹ️ Show config status', 'menu_showConfigStatus')
      .addItem('📂 Open inbox folder', 'menu_openInboxFolder'))
    .addSeparator()
    .addSubMenu(ui.createMenu('📥 Import')
      .addItem('📥 Process new uploads', 'menu_processUsageUploads')
      .addItem('🧾 Process invoice PDFs', 'menu_processInvoicePdfs')
      .addItem('🔁 Re-attribute ledger (after adding aliases)', 'menu_reAttributeLedger')
      .addItem('🚨 Show unmapped employee IDs', 'menu_showUnmapped')
      .addItem('🛠️ Bulk resolve unmapped (interactive)', 'menu_bulkResolveUnmapped')
      .addItem('✅ Apply bulk resolutions', 'menu_applyBulkResolutions')
      .addSeparator()
      .addItem('💣 RESET ledger (delete all transactions)', 'menu_resetLedger')
      .addSeparator()
      .addItem('♻️ Recover last archived files to inbox', 'menu_recoverArchive'))
    .addSeparator()
    .addSubMenu(ui.createMenu('🔍 Reconcile')
      .addItem('🎯 Reconcile period...', 'menu_reconcilePeriod')
      .addItem('✏️ Resolve boundary transactions...', 'menu_resolveBoundary')
      .addItem('📋 Show reconciliation status', 'menu_showReconStatus'))
    .addSeparator()
    .addSubMenu(ui.createMenu('📊 Report')
      .addItem('📄 Generate Report for period...', 'menu_generateReport'))
    .addToUi();
}
