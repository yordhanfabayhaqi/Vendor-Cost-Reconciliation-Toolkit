/**
 * 09_Tutorial.gs
 * In-app monthly tutorial for downstream operators.
 * Generic — replace vendor names, email senders, and submission link
 * placeholders with your organization's specifics.
 */

function menu_showTutorial() {
  const html = HtmlService.createHtmlOutput(buildTutorialHtml_())
    .setWidth(720).setHeight(640).setTitle('How to Run Me — Monthly Tutorial');
  SpreadsheetApp.getUi().showModalDialog(html, 'How to Run Me');
}

function buildTutorialHtml_() {
  return `
<!DOCTYPE html>
<html>
<head><base target="_blank"><style>
body { font-family: system-ui, sans-serif; font-size: 14px; line-height: 1.5; padding: 16px 24px; }
h1 { font-size: 18px; color: #0d652d; margin: 0 0 4px 0; }
.step { border-left: 3px solid #0d652d; padding: 10px 14px; margin: 14px 0; background: #f6fbf7; border-radius: 0 6px 6px 0; }
.step-title { font-weight: 600; color: #0d652d; margin-bottom: 6px; }
code { background: #eee; padding: 1px 5px; border-radius: 3px; font-size: 12px; }
.pill { display: inline-block; background: #0d652d; color: white; padding: 1px 8px; border-radius: 12px; font-size: 11px; font-weight: 600; margin-right: 4px; }
.callout { background: #fff8e1; border: 1px solid #ffd54f; padding: 8px 12px; margin: 10px 0; font-size: 12px; border-radius: 4px; }
.toc { background: #f0f4ff; padding: 10px 16px; border-radius: 4px; font-size: 12px; margin-bottom: 16px; }
</style></head>
<body>

<h1>How to Run This Monthly</h1>
<div style="color:#666;font-size:13px;margin-bottom:18px">For operators running the vendor cost reconciliation.</div>

<div class="toc">
  <strong>3 phases:</strong><br>
  📥 Collect invoices + usage files<br>
  ⚙️ Process in this spreadsheet<br>
  📤 Submit results downstream
</div>

<h2>📥 Phase 1: Collect</h2>

<div class="step">
  <div class="step-title"><span class="pill">Step 1</span> Vendor A management fee invoice</div>
  From: <code>[vendor-a-billing@example.com]</code><br>
  Subject pattern: <code>Invoice [number] - VendorA - [Month] [Year]</code><br>
  Download the PDF attachment.
</div>

<div class="step">
  <div class="step-title"><span class="pill">Step 2</span> Vendor A cost passthrough</div>
  From: <code>[vendor-a-billing@example.com]</code><br>
  Subject pattern: <code>Transfer request [number] - VendorA - [Month] [Year]</code><br>
  Download the PDF attachment.
</div>

<div class="step">
  <div class="step-title"><span class="pill">Step 3</span> Vendor B invoice</div>
  From: <code>[vendor-b-billing@example.com]</code><br>
  Read the PDF to identify the billing period (may not be stated in email).<br>
  Download the PDF.
</div>

<div class="step">
  <div class="step-title"><span class="pill">Step 4</span> Vendor A + B usage detail files</div>
  Download the period-scoped usage exports from each vendor's admin portal.
</div>

<h2>⚙️ Phase 2: Process</h2>

<div class="step">
  <div class="step-title"><span class="pill">Step 5</span> Drop files in inbox folders</div>
  Google Drive → your parent folder → <code>VendorCost_Inbox/</code>:
  <ul>
    <li><code>vendor_a_usage/</code> → Vendor A xlsx</li>
    <li><code>vendor_b_usage/</code> → Vendor B xlsx</li>
    <li><code>invoices/</code> → all 3 PDFs</li>
  </ul>
</div>

<div class="step">
  <div class="step-title"><span class="pill">Step 6</span> Menu commands</div>
  Menu → <code>Vendor Cost Toolkit</code>:
  <ol>
    <li>📥 Import → Process new uploads</li>
    <li>📥 Import → Process invoice PDFs</li>
    <li>🔍 Reconcile → Reconcile period... (verify all ✅)</li>
    <li>📊 Report → Generate Report for period...</li>
  </ol>
</div>

<div class="step">
  <div class="step-title"><span class="pill">Step 7</span> Resolve unmapped (if any)</div>
  Menu → Import → Bulk resolve unmapped → fill canonical codes → Apply.
  <div class="callout">Multi-cost-center IDs like "CC-012 CC-035" get split evenly.</div>
</div>

<h2>📤 Phase 3: Submit</h2>

<div class="step">
  <div class="step-title"><span class="pill">Step 8</span> Submit to your downstream system</div>
  Open your expense management system: <code>https://example.com/expense-form</code><br><br>
  Submit ONCE per invoice column in the report. Match each column's total against the invoice PDF before final submission.
</div>

<hr>
<div style="font-size:11px;color:#888;text-align:center;margin-top:20px">
  Configure vendor names, email senders, and submission URLs in <code>09_Tutorial.gs</code> to match your organization.
</div>

</body></html>`;
}
