/**
 * ============================================================
 *  Sales Team Tracking Tool — Google Apps Script
 *  Bound to: Sales_Tracker spreadsheet
 *  Author:   [Your Name]
 *  Version:  1.0
 * ============================================================
 */

// ── CONSTANTS ─────────────────────────────────────────────
const SETTINGS_SHEET = "Settings";
const DASHBOARD_SHEET = "Dashboard";
const EMAIL_CELL      = "B3";   // Recipient email in Settings
const SUBJECT_CELL    = "B4";   // Email subject in Settings

// ── MENU ──────────────────────────────────────────────────
/**
 * Adds custom "Admin Tools" menu on spreadsheet open.
 */
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu("Admin Tools")
    .addItem("Send Report Now", "sendWeeklyReport")
    .addSeparator()
    .addItem("Install Weekly Auto-Trigger", "installWeeklyTrigger")
    .addItem("Remove All Triggers", "removeTriggers")
    .addToUi();
}

// ── PDF GENERATION & EMAILING ────────────────────────────
/**
 * Converts the Dashboard tab to a PDF and emails it.
 * Email address and subject are read from the Settings tab.
 */
function sendWeeklyReport() {
  const ss      = SpreadsheetApp.getActiveSpreadsheet();
  const settings = ss.getSheetByName(SETTINGS_SHEET);

  // Pull config from Settings tab (no hardcoded values in script)
  const recipientEmail = settings.getRange(EMAIL_CELL).getValue();
  const emailSubject   = settings.getRange(SUBJECT_CELL).getValue();

  if (!recipientEmail) {
    SpreadsheetApp.getUi().alert("⚠️ No recipient email found in Settings tab (cell B3).");
    return;
  }

  // Generate PDF of Dashboard sheet only
  const pdfBlob = exportDashboardAsPdf(ss);
  const dateStr = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy-MM-dd");

  // Compose and send email
  GmailApp.sendEmail(
    recipientEmail,
    emailSubject,
    "",   // plain-text body (empty — PDF is the content)
    {
      htmlBody: buildEmailBody(dateStr),
      attachments: [pdfBlob.setName(`Sales_Report_${dateStr}.pdf`)],
      name: "Sales Reporting Bot"
    }
  );

  Logger.log(`Report sent to ${recipientEmail} on ${dateStr}`);
  SpreadsheetApp.getActiveSpreadsheet().toast(
    `✅ Report emailed to ${recipientEmail}`, "Success", 5
  );
}

/**
 * Exports only the Dashboard sheet as a PDF blob.
 * @param {Spreadsheet} ss - Active spreadsheet.
 * @returns {Blob} PDF blob.
 */
function exportDashboardAsPdf(ss) {
  const dashSheet = ss.getSheetByName(DASHBOARD_SHEET);
  const ssId      = ss.getId();
  const sheetId   = dashSheet.getSheetId();

  // URL parameters for PDF export (Dashboard tab only, A4 landscape, fit to page)
  const url = `https://docs.google.com/spreadsheets/d/${ssId}/export`
    + `?format=pdf`
    + `&size=A4`
    + `&portrait=false`
    + `&fitw=true`
    + `&fith=true`
    + `&top_margin=0.5`
    + `&bottom_margin=0.5`
    + `&left_margin=0.5`
    + `&right_margin=0.5`
    + `&sheetnames=false`
    + `&printtitle=false`
    + `&pagenumbers=false`
    + `&gridlines=false`
    + `&fzr=false`
    + `&gid=${sheetId}`;   // Only export this sheet

  const token    = ScriptApp.getOAuthToken();
  const response = UrlFetchApp.fetch(url, {
    headers: { Authorization: `Bearer ${token}` }
  });

  return response.getBlob().setContentType("application/pdf");
}

/**
 * Builds a simple HTML email body.
 * @param {string} dateStr - Formatted date string.
 * @returns {string} HTML string.
 */
function buildEmailBody(dateStr) {
  return `
    <div style="font-family:Arial,sans-serif;max-width:600px;margin:0 auto">
      <div style="background:#1A2744;padding:20px;border-radius:8px 8px 0 0">
        <h2 style="color:#fff;margin:0">📊 Weekly Sales Report</h2>
        <p style="color:#93C5FD;margin:4px 0 0">Generated: ${dateStr}</p>
      </div>
      <div style="background:#F3F6FC;padding:20px;border-radius:0 0 8px 8px">
        <p style="color:#1E293B">Please find attached the weekly Sales Performance Dashboard PDF.</p>
        <p style="color:#64748B;font-size:12px">
          This is an automated report. Do not reply to this email.<br>
          To update the recipient or subject, edit the <strong>Settings</strong> tab.
        </p>
      </div>
    </div>
  `;
}

// ── TRIGGERS ─────────────────────────────────────────────
/**
 * Installs a weekly time-driven trigger: every Friday at 14:00.
 * Safe to call multiple times — removes existing triggers first.
 */
function installWeeklyTrigger() {
  removeTriggers();   // clean slate

  ScriptApp.newTrigger("sendWeeklyReport")
    .timeBased()
    .onWeekDay(ScriptApp.WeekDay.FRIDAY)
    .atHour(14)          // 14:00 (2 PM)
    .nearMinute(0)
    .create();

  SpreadsheetApp.getActiveSpreadsheet().toast(
    "✅ Weekly trigger set: Every Friday at 2:00 PM", "Trigger Installed", 5
  );
  Logger.log("Weekly Friday 14:00 trigger installed.");
}

/**
 * Removes all project triggers for sendWeeklyReport.
 */
function removeTriggers() {
  ScriptApp.getProjectTriggers()
    .filter(t => t.getHandlerFunction() === "sendWeeklyReport")
    .forEach(t => ScriptApp.deleteTrigger(t));
}

// ── DASHBOARD FILTER (QUERY ALTERNATIVE) ─────────────────
/**
 * Refreshes the At-Risk deals table on the Dashboard.
 * Called manually or from onOpen if desired.
 * Uses JavaScript to build the data array, then writes it.
 */
function refreshAtRiskTable() {
  const ss        = SpreadsheetApp.getActiveSpreadsheet();
  const dashboard = ss.getSheetByName(DASHBOARD_SHEET);
  const repNames  = ["Rep_1", "Rep_2", "Rep_3"];
  const today     = new Date();
  today.setHours(0, 0, 0, 0);

  const atRisk = [];

  repNames.forEach(repName => {
    const sheet = ss.getSheetByName(repName);
    if (!sheet) return;

    const data = sheet.getRange("A3:G200").getValues();
    data.forEach(row => {
      const [deal, client, closeDate, value, stage, , weightedForecast] = row;
      if (!deal) return;                                   // skip empty rows
      if (stage === "Closed-Won (100%)" || stage === "Closed-Lost (0%)") return;

      const close = new Date(closeDate);
      if (close < today) {
        atRisk.push([repName, deal, client, closeDate, value, stage, weightedForecast]);
      }
    });
  });

  // Write to Dashboard starting at row 20
  const startRow = 20;
  const clearRange = dashboard.getRange(startRow, 2, 30, 7);
  clearRange.clearContent().clearFormat();

  if (atRisk.length === 0) {
    const noRisk = dashboard.getRange(startRow, 2, 1, 7);
    noRisk.merge();
    noRisk.setValue("✅ No at-risk deals at this time.");
    noRisk.setFontStyle("italic").setFontColor("#16A34A");
    return;
  }

  const outputRange = dashboard.getRange(startRow, 2, atRisk.length, 7);
  outputRange.setValues(atRisk);

  // Highlight in light red
  outputRange.setBackground("#FEE2E2");
  outputRange.setFontColor("#991B1B");
  outputRange.setBorder(true, true, true, true, true, true, "#DC2626",
                        SpreadsheetApp.BorderStyle.SOLID);

  // Format date & currency columns
  dashboard.getRange(startRow, 5, atRisk.length, 1).setNumberFormat("YYYY-MM-DD");
  dashboard.getRange(startRow, 6, atRisk.length, 1).setNumberFormat("$#,##0");
  dashboard.getRange(startRow, 8, atRisk.length, 1).setNumberFormat("$#,##0");
}
