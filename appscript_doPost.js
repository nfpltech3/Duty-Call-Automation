function doPost(e) {
  var lock = LockService.getScriptLock();
  var lockAcquired = false;

  // ── CONFIG: Add or remove recipient emails here ──
  var RECIPIENTS = [
    "crm5@nagarkot.co.in"
    // "operations@nagarkot.co.in",
    // "manager@nagarkot.co.in"
  ];

  try {
    lock.waitLock(30000);
    lockAcquired = true;

    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Validation Log");
    var data = JSON.parse(e.postData.contents);

    if (!data.rows || data.rows.length === 0) {
      return ContentService.createTextOutput("No rows").setMimeType(ContentService.MimeType.TEXT);
    }

    // 1. Append all rows to the sheet
    sheet.getRange(sheet.getLastRow() + 1, 1, data.rows.length, data.rows[0].length).setValues(data.rows);
    SpreadsheetApp.flush();

    // 2. Build and send the HIGH RISK email alert
    sendHighRiskAlert_(data.rows, RECIPIENTS);

    return ContentService.createTextOutput("Success").setMimeType(ContentService.MimeType.TEXT);

  } catch (error) {
    return ContentService.createTextOutput("Error: " + error.toString()).setMimeType(ContentService.MimeType.TEXT);
  } finally {
    if (lockAcquired) {
      lock.releaseLock();
    }
  }
}


/**
 * Sends a single HTML email containing a detailed table of all HIGH RISK rows.
 * Column mapping (0-indexed) from the Python payload:
 *   0: Timestamp, 1: Importer, 2: BE No, 3: Job No,
 *   4: BE Ass Value, 5: CL Ass Value, 6: Ass Diff,
 *   7: BE Duty, 8: CL Duty, 9: Duty Diff,
 *   10: BE Scripts, 11: CL Foregone, 12: Foregone Diff,
 *   13: Penalty, 14: Fine, 15: Interest, 16: PFI Remark
 */
function sendHighRiskAlert_(rows, recipients) {
  if (!rows || rows.length === 0) return;

  var count = rows.length;
  var subject = "⚠️ HIGH RISK Alert — " + count + " entr" + (count === 1 ? "y" : "ies") + " detected | Duty Call Automation";

  var TABLE_COL_COUNT = 13; // Columns 0-12 go into the main table

  // ── Build the HTML email body ──
  var html = "";
  html += '<div style="font-family:Segoe UI,Arial,sans-serif; max-width:850px; margin:0 auto; border:1px solid #E0E0E0; border-radius:8px; overflow:hidden; background:#FFF; box-shadow: 0 4px 6px rgba(0,0,0,0.05);">';

  // Header Banner
  html += '<div style="background:#FFF; border-bottom:2px solid #ECEFF1; padding:20px 24px;">';
  html += '<h2 style="margin:0; color:#263238; font-size:18px; font-weight:600;">';
  html += '<span style="color:#C62828; margin-right:8px;">⚠️</span> HIGH RISK Validation Alert';
  html += '</h2>';
  html += '<p style="margin:6px 0 0; color:#607D8B; font-size:13px;">' + count + ' discrepanc' + (count === 1 ? 'y' : 'ies') + ' exceeded the ₹500 threshold</p>';
  html += '</div>';

  // Table Container
  html += '<div style="overflow-x:auto;">';
  html += '<table style="width:100%; border-collapse:collapse; font-size:13px; line-height:1.4;">';

  // Table Header Row
  html += '<tr style="background:#F8F9FA;">';
  var headers = ["Timestamp", "Importer", "BE No", "Job No", "BE Ass Val", "CL Ass Val", "Ass Diff", "BE Duty", "CL Duty", "Duty Diff", "BE Scripts", "CL Foregone", "Foregone Diff"];
  for (var h = 0; h < headers.length; h++) {
    var thStyle = 'padding:10px 8px; text-align:left; border-bottom:2px solid #CFD8DC; font-size:11px; color:#546E7A; font-weight:600;';
    if (h === 4 || h === 7 || h === 10) {
      thStyle += ' border-left:2.5px solid #B0BEC5;';
    }
    if (h === 6 || h === 9 || h === 12) {
      thStyle += ' background:#FFFBFB;';
    }
    html += '<th style="' + thStyle + '">' + headers[h] + '</th>';
  }
  html += '</tr>';

  // Data rows
  for (var i = 0; i < rows.length; i++) {
    var bgColor = (i % 2 === 0) ? "#FFFFFF" : "#FAFAFA";
    html += '<tr style="background:' + bgColor + ';">';
    for (var j = 0; j < TABLE_COL_COUNT; j++) {
      var cellVal = (j < rows[i].length && rows[i][j] !== null && rows[i][j] !== undefined) ? rows[i][j].toString() : "";
      var cellStyle = "padding:8px; border-bottom:1px solid #ECEFF1; color:#37474F;";
      if (j === 4 || j === 7 || j === 10) {
        cellStyle += " border-left:2.5px solid #B0BEC5;";
      }
      if (j === 6 || j === 9 || j === 12) {
        cellStyle += " background:#FFFBFB;";
        if (cellVal !== "" && cellVal !== "0" && cellVal !== "0.00" && cellVal !== "0.0") {
          cellStyle += " color:#C62828; font-weight:bold; font-size:13px;";
        }
      }
      html += '<td style="' + cellStyle + '">' + cellVal + '</td>';
    }
    html += '</tr>';

    // PFI Remark row (column index 16)
    var pfiRemark = (rows[i].length > 16 && rows[i][16]) ? rows[i][16].toString() : "";
    if (pfiRemark) {
      html += '<tr style="background:' + bgColor + ';">';
      html += '<td colspan="' + TABLE_COL_COUNT + '" style="padding:10px 12px; border-bottom:1px solid #ECEFF1; font-size:13.5px; color:#E65100; font-style:italic;">';
      html += '&#9432; ' + pfiRemark;
      
      // Parse values for bolding
      var p = (rows[i].length > 13) ? parseFloat(rows[i][13]) || 0 : 0;
      var f = (rows[i].length > 14) ? parseFloat(rows[i][14]) || 0 : 0;
      var n = (rows[i].length > 15) ? parseFloat(rows[i][15]) || 0 : 0;
      
      var pStr = p > 0 ? '<b style="color:#C62828;">' + p + '</b>' : '0';
      var fStr = f > 0 ? '<b style="color:#C62828;">' + f + '</b>' : '0';
      var nStr = n > 0 ? '<b style="color:#C62828;">' + n + '</b>' : '0';

      html += ' <span style="color:#78909C; font-size:12px; font-style:normal; margin-left:8px;">(Penalty: ' + pStr + ' | Fine: ' + fStr + ' | Interest: ' + nStr + ')</span>';
      html += '</td></tr>';
    }
  }

  html += '</table></div>';

  // Footer
  html += '<div style="background:#FAFAFA; border-top:1px solid #ECEFF1; padding:12px 24px; font-size:10px !important; color:#90A4AE; mso-line-height-rule:exactly; line-height:14px;">';
  html += 'This is an automated alert from Duty Call Automation';
  html += '</div>';
  html += '</div>';

  // ── Send ──
  GmailApp.sendEmail(
    recipients.join(","),
    subject,
    "",
    {
      htmlBody: html
    }
  );
}

