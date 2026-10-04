/**
 * ============================================================================
 * WORKSTATION ATTENDANCE SYSTEM - BACKEND SCRIPT (Google Apps Script)
 * ============================================================================
 * Features:
 *  - Permanent, cryptographically signed (HMAC-SHA256) QR Code generation
 *  - Automated QR badge email delivery with inline images & attachments
 *  - Real-time scanner backend with race-condition prevention (LockService)
 *  - Time-window validation & duplicate check-in prevention
 *  - Automated & manual daily report generation with rich KPI summaries
 *  - Scheduled time-driven report email triggers
 * ============================================================================
 */

// ----------------------------------------------------------------------------
// GLOBAL CONSTANTS & CONFIGURATION
// ----------------------------------------------------------------------------
var SHEET_MEMBERS = "Members";
var SHEET_ATTENDANCE = "Attendance_Log";
var SHEET_CONFIG = "Config";
var REPORT_PREFIX = "Report_";
var PROP_SECRET_KEY = "WORKSTATION_HMAC_SECRET_KEY";
var SCRIPT_TIMEZONE = "Asia/Kolkata";

/**
 * Custom UI Menu when the Spreadsheet is opened.
 */
function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu("🏢 Workstation Attendance")
    .addItem("🚀 Initialize / Refresh System Setup", "setup")
    .addItem("✨ Populate Sample Demo Data", "populateDemoData")
    .addItem("🔄 Clear System Cache / Refresh Settings", "clearSystemCache")
    .addSeparator()
    .addItem("📧 Send QR Codes to Members", "sendQrCodesToMembers")
    .addItem("⚡ Email Provider Setup (Brevo / Resend / Gmail)", "showEmailConfigDialog")
    .addItem("📁 Export All QR Badges to Google Drive", "exportAllQrsToDrive")
    .addSeparator()
    .addItem("📊 Generate Today's Report", "generateTodayReportMenu")
    .addItem("⏰ Setup Daily Report Trigger", "setupTriggers")
    .addSeparator()
    .addItem("🧪 Run Test Scan (Simulation)", "testScanMember")
    .addItem("🌐 Get Web App Scanner URL", "showScannerUrlDialog")
    .addToUi();
}

// ----------------------------------------------------------------------------
// 1. SYSTEM INITIALIZATION (setup)
// ----------------------------------------------------------------------------

/**
 * Initializes sheets, sample members, formatting, and secure HMAC keys.
 */
function setup() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  
  // 1. Ensure HMAC secret key exists
  getOrCreateSecretKey();
  
  // 2. Setup Config Sheet
  var configSheet = getOrCreateSheet(ss, SHEET_CONFIG);
  if (configSheet.getLastRow() === 0) {
    var adminEmail = getAdminEmailSafe();
    var defaultConfigs = [
      ["Setting", "Value", "Description"],
      ["Workstation Name", "Main Innovation Lab", "Name of the facility / workstation"],
      ["Check-in Start Time", "08:00", "Start of check-in window (HH:mm 24-hr format)"],
      ["Check-in End Time", "20:00", "End of check-in window (HH:mm 24-hr format)"],
      ["Report Generation Time", "21:00", "Daily automated report trigger time (HH:mm 24-hr format)"],
      ["Admin Email", adminEmail, "Recipient email for daily attendance summaries"],
      ["Allowed Re-scan", "No", "Allow duplicate scans on the same day (Yes/No)"],
      ["Email Provider", "Gmail (Default)", "Email service: 'Gmail (Default)', 'Brevo', or 'Resend'"],
      ["Email API Key", "", "API Key for Brevo or Resend (bypasses Gmail 100/day limit)"],
      ["Sender Email", adminEmail, "Sender email address (verified in Brevo/Resend)"],
      ["Sender Name", "Workstation Admin", "Display name for outgoing QR badge emails"]
    ];
    configSheet.getRange(1, 1, defaultConfigs.length, 3).setValues(defaultConfigs);
    formatHeaderRow(configSheet, 1, 3, "#1A73E8");
    configSheet.setColumnWidth(1, 200);
    configSheet.setColumnWidth(2, 220);
    configSheet.setColumnWidth(3, 360);
  } else {
    // Ensure new email provider settings exist in already-created Config sheets
    var existingConfigs = configSheet.getDataRange().getValues();
    var existingKeys = {};
    for (var k = 1; k < existingConfigs.length; k++) {
      existingKeys[String(existingConfigs[k][0]).trim()] = true;
    }
    var newSettings = [];
    var adminEmailSafe = getAdminEmailSafe();
    if (!existingKeys["Email Provider"]) {
      newSettings.push(["Email Provider", "Gmail (Default)", "Email service: 'Gmail (Default)', 'Brevo', or 'Resend'"]);
    }
    if (!existingKeys["Email API Key"]) {
      newSettings.push(["Email API Key", "", "API Key for Brevo or Resend (bypasses Gmail 100/day limit)"]);
    }
    if (!existingKeys["Sender Email"]) {
      newSettings.push(["Sender Email", adminEmailSafe, "Sender email address (verified in Brevo/Resend)"]);
    }
    if (!existingKeys["Sender Name"]) {
      newSettings.push(["Sender Name", "Workstation Admin", "Display name for outgoing QR badge emails"]);
    }
    if (newSettings.length > 0) {
      configSheet.getRange(configSheet.getLastRow() + 1, 1, newSettings.length, 3).setValues(newSettings);
    }
  }

  // 3. Setup Members Sheet
  var membersSheet = getOrCreateSheet(ss, SHEET_MEMBERS);
  if (membersSheet.getLastRow() === 0) {
    var memberHeaders = [["Member ID", "Name", "Email", "Token", "QR Link", "Email Sent"]];
    membersSheet.getRange(1, 1, 1, 6).setValues(memberHeaders);
    formatHeaderRow(membersSheet, 1, 6, "#0F9D58");

    // Sample Members for instant testing
    var sampleMembers = [
      ["WS-001", "Aarav Sharma", "aarav.sharma@example.com", "", "", "No"],
      ["WS-002", "Priya Patel", "priya.patel@example.com", "", "", "No"],
      ["WS-003", "Rahul Verma", "rahul.verma@example.com", "", "", "No"],
      ["WS-004", "Ananya Iyer", "ananya.iyer@example.com", "", "", "No"],
      ["WS-005", "Vikram Singh", "vikram.singh@example.com", "", "", "No"]
    ];
    membersSheet.getRange(2, 1, sampleMembers.length, 6).setValues(sampleMembers);
  }

  // Generate tokens & QR links for any member lacking them
  generateMissingTokensAndQrs(membersSheet);
  
  // Format Members Sheet Columns
  membersSheet.setColumnWidth(1, 120);
  membersSheet.setColumnWidth(2, 160);
  membersSheet.setColumnWidth(3, 220);
  membersSheet.setColumnWidth(4, 260);
  membersSheet.setColumnWidth(5, 300);
  membersSheet.setColumnWidth(6, 110);
  membersSheet.setFrozenRows(1);

  // 4. Setup Attendance Log Sheet
  var logSheet = getOrCreateSheet(ss, SHEET_ATTENDANCE);
  if (logSheet.getLastRow() === 0) {
    var logHeaders = [["Timestamp", "Date", "Member ID", "Name", "Status", "Scanner / Device Info"]];
    logSheet.getRange(1, 1, 1, 6).setValues(logHeaders);
    formatHeaderRow(logSheet, 1, 6, "#673AB7");
    logSheet.setColumnWidth(1, 180);
    logSheet.setColumnWidth(2, 120);
    logSheet.setColumnWidth(3, 120);
    logSheet.setColumnWidth(4, 180);
    logSheet.setColumnWidth(5, 100);
    logSheet.setColumnWidth(6, 220);
    logSheet.setFrozenRows(1);
  }

  SpreadsheetApp.getActiveSpreadsheet().toast("System setup initialized successfully!", "✅ Setup Complete", 5);
}

// ----------------------------------------------------------------------------
// 2. CRYPTOGRAPHIC TOKEN & QR CODE LOGIC
// ----------------------------------------------------------------------------

/**
 * Retrieves the permanent secret key from Script Properties or generates a new one.
 * @return {string} Secret key.
 */
function getOrCreateSecretKey() {
  var props = PropertiesService.getScriptProperties();
  var secret = props.getProperty(PROP_SECRET_KEY);
  if (!secret) {
    secret = Utilities.getUuid() + "-" + Utilities.getUuid() + "-" + new Date().getTime();
    props.setProperty(PROP_SECRET_KEY, secret);
  }
  return secret;
}

/**
 * Generates an HMAC-SHA256 signature token for a Member ID.
 * @param {string} memberId - e.g. "WS-001"
 * @return {string} Hex-encoded HMAC-SHA256 token.
 */
function generateMemberToken(memberId) {
  if (!memberId) return "";
  var cleanId = String(memberId).trim().toUpperCase();
  var secret = getOrCreateSecretKey();
  var signatureBytes = Utilities.computeHmacSha256Signature(cleanId, secret);
  
  // Convert byte array to hexadecimal string
  var hexToken = signatureBytes.map(function(byte) {
    var b = byte < 0 ? byte + 256 : byte;
    return ("0" + b.toString(16)).slice(-2);
  }).join("");
  
  return hexToken;
}

/**
 * Verifies if the provided token matches the expected HMAC-SHA256 for the Member ID.
 * @param {string} memberId
 * @param {string} token
 * @return {boolean}
 */
function verifyMemberToken(memberId, token) {
  if (!memberId || !token) return false;
  var expectedToken = generateMemberToken(memberId);
  return expectedToken.toLowerCase() === String(token).trim().toLowerCase();
}

/**
 * Generates the QR image URL using api.qrserver.com.
 * QR content format: "MemberID|Token"
 * @param {string} memberId
 * @param {string} token
 * @return {string} QR Image URL
 */
function generateQrCodeUrl(memberId, token) {
  var qrContent = String(memberId).trim().toUpperCase() + "|" + String(token).trim();
  var encoded = encodeURIComponent(qrContent);
  return "https://api.qrserver.com/v1/create-qr-code/?size=300x300&margin=10&data=" + encoded;
}

/**
 * Scans Members sheet and fills missing Tokens and QR URLs.
 */
function generateMissingTokensAndQrs(sheet) {
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return;
  
  var data = sheet.getRange(2, 1, lastRow - 1, 6).getValues();
  var updated = false;
  
  for (var i = 0; i < data.length; i++) {
    var memberId = data[i][0];
    var token = data[i][3];
    var qrLink = data[i][4];
    
    if (memberId) {
      if (!token) {
        token = generateMemberToken(memberId);
        data[i][3] = token;
        updated = true;
      }
      if (!qrLink || qrLink.indexOf("api.qrserver.com") === -1) {
        data[i][4] = generateQrCodeUrl(memberId, token);
        updated = true;
      }
    }
  }
  
  if (updated) {
    sheet.getRange(2, 1, data.length, 6).setValues(data);
  }
}

// ----------------------------------------------------------------------------
// 3. FEATURE 1: SEND QR CODES BY EMAIL (MULTI-PROVIDER: GMAIL, BREVO, RESEND)
// ----------------------------------------------------------------------------

/**
 * Sends permanent QR ID badges to all members where Email Sent is not "Yes".
 * Supports Gmail (default 100/day limit), Brevo API (300/day free), and Resend API (3,000/mo free).
 */
function sendQrCodesToMembers() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SHEET_MEMBERS);
  if (!sheet) {
    SpreadsheetApp.getUi().alert("Error: Members sheet not found. Please run Setup first.");
    return;
  }
  
  generateMissingTokensAndQrs(sheet);
  
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) {
    SpreadsheetApp.getUi().alert("No members found in Members sheet.");
    return;
  }
  
  var config = getConfigMap();
  var workstationName = config["Workstation Name"] || "Workstation";
  var provider = String(config["Email Provider"] || "Gmail").trim().toLowerCase();
  var apiKey = String(config["Email API Key"] || "").trim();
  var senderEmail = String(config["Sender Email"] || getAdminEmailSafe()).trim();
  var senderName = String(config["Sender Name"] || (workstationName + " Admin")).trim();
  
  var isBrevo = provider.indexOf("brevo") !== -1 || apiKey.indexOf("xkeysib-") === 0;
  var isResend = provider.indexOf("resend") !== -1 || apiKey.indexOf("re_") === 0;
  
  // Validation for API services
  if ((isBrevo || isResend) && !apiKey) {
    SpreadsheetApp.getUi().alert(
      "🔑 API Key Required",
      "You have selected '" + (isBrevo ? "Brevo" : "Resend") + "' as your email provider, but no API Key was found.\n\n" +
      "Please go to menu:\n" +
      "🏢 Workstation Attendance > ⚡ Email Provider Setup\n" +
      "and paste your free API key.",
      SpreadsheetApp.getUi().ButtonSet.OK
    );
    return;
  }
  
  // Validation for Gmail quota
  if (!isBrevo && !isResend) {
    var quotaRemaining = MailApp.getRemainingDailyQuota();
    if (quotaRemaining <= 0) {
      SpreadsheetApp.getUi().alert(
        "⏳ Gmail Daily Quota Reached",
        "Your Google Account's daily email sending quota has been reached for today.\n\n" +
        "🚀 To send more emails immediately today:\n" +
        "1. Create a free account on Brevo (300 free emails/day) or Resend (3,000/mo free).\n" +
        "2. Click '🏢 Workstation Attendance' > '⚡ Email Provider Setup' to add your free key.",
        SpreadsheetApp.getUi().ButtonSet.OK
      );
      return;
    }
  }
  
  var data = sheet.getRange(2, 1, lastRow - 1, 6).getValues();
  var sentCount = 0;
  var skipCount = 0;
  var errorCount = 0;
  var errors = [];
  var providerLabel = isBrevo ? "Brevo API" : (isResend ? "Resend API" : "Gmail");
  
  for (var i = 0; i < data.length; i++) {
    var memberId = data[i][0];
    var name = data[i][1];
    var email = data[i][2];
    var token = data[i][3];
    var qrUrl = data[i][4];
    var emailSent = String(data[i][5]).trim().toLowerCase();
    
    if (emailSent === "yes") {
      skipCount++;
      continue;
    }
    
    if (!email || email.indexOf("@") === -1) {
      errorCount++;
      errors.push("Row " + (i + 2) + " (" + name + "): Invalid email address");
      continue;
    }
    
    try {
      // Fetch QR Code Blob
      var response = UrlFetchApp.fetch(qrUrl, { muteHttpExceptions: true });
      if (response.getResponseCode() !== 200) {
        throw new Error("Failed to fetch QR image from API (HTTP " + response.getResponseCode() + ")");
      }
      var qrBlob = response.getBlob().setName("QR_ID_" + memberId + ".png");
      var subject = "🎫 Your Permanent Workstation Access QR Badge - " + workstationName;
      
      if (isBrevo) {
        // Send via Brevo HTTP API
        var htmlBody = buildQrEmailHtml(name, memberId, workstationName, false, qrUrl);
        sendViaBrevoApi(apiKey, senderEmail, senderName, email, name, subject, htmlBody, qrBlob, memberId);
      } else if (isResend) {
        // Send via Resend HTTP API
        var htmlBody = buildQrEmailHtml(name, memberId, workstationName, false, qrUrl);
        sendViaResendApi(apiKey, senderEmail, senderName, email, subject, htmlBody, qrBlob, memberId);
      } else {
        // Send via Google Gmail App
        var htmlBody = buildQrEmailHtml(name, memberId, workstationName, true, qrUrl);
        MailApp.sendEmail({
          to: email,
          subject: subject,
          htmlBody: htmlBody,
          inlineImages: {
            qrImage: qrBlob
          },
          attachments: [qrBlob]
        });
      }
      
      // Update Email Sent status
      data[i][5] = "Yes";
      sentCount++;
      
    } catch (e) {
      errorCount++;
      errors.push("Row " + (i + 2) + " (" + name + "): " + e.message);
    }
  }
  
  // Persist updated "Email Sent" status back to sheet
  sheet.getRange(2, 1, data.length, 6).setValues(data);
  
  var msg = "Provider: " + providerLabel + "\n\n" +
            "• Sent successfully: " + sentCount + "\n" +
            "• Already sent (skipped): " + skipCount + "\n" +
            "• Errors: " + errorCount;
            
  if (errors.length > 0) {
    msg += "\n\nError details:\n" + errors.slice(0, 5).join("\n");
  }
  
  SpreadsheetApp.getUi().alert("📧 QR Badges Dispatch Complete", msg, SpreadsheetApp.getUi().ButtonSet.OK);
}

/**
 * Dispatches an email via Brevo's REST API (v3).
 * Brevo Free tier allows 300 emails/day with no credit card required.
 */
function sendViaBrevoApi(apiKey, senderEmail, senderName, toEmail, toName, subject, htmlContent, qrBlob, memberId) {
  var url = "https://api.brevo.com/v3/smtp/email";
  var base64Img = Utilities.base64Encode(qrBlob.getBytes());
  
  var payload = {
    sender: {
      name: senderName || "Workstation Attendance",
      email: senderEmail
    },
    to: [
      {
        email: toEmail,
        name: toName || toEmail
      }
    ],
    subject: subject,
    htmlContent: htmlContent,
    attachment: [
      {
        content: base64Img,
        name: "QR_ID_" + memberId + ".png"
      }
    ]
  };

  var options = {
    method: "post",
    headers: {
      "api-key": apiKey,
      "Content-Type": "application/json",
      "Accept": "application/json"
    },
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  };

  var response = UrlFetchApp.fetch(url, options);
  var code = response.getResponseCode();
  if (code < 200 || code >= 300) {
    var err = response.getContentText();
    throw new Error("Brevo Error (" + code + "): " + err);
  }
  return JSON.parse(response.getContentText());
}

/**
 * Dispatches an email via Resend's REST API.
 * Resend Free tier provides 3,000 emails/month free.
 */
function sendViaResendApi(apiKey, senderEmail, senderName, toEmail, subject, htmlContent, qrBlob, memberId) {
  var url = "https://api.resend.com/emails";
  var base64Img = Utilities.base64Encode(qrBlob.getBytes());
  
  var fromAddress = senderName ? (senderName + " <" + senderEmail + ">") : senderEmail;
  var payload = {
    from: fromAddress,
    to: [toEmail],
    subject: subject,
    html: htmlContent,
    attachments: [
      {
        filename: "QR_ID_" + memberId + ".png",
        content: base64Img
      }
    ]
  };

  var options = {
    method: "post",
    headers: {
      "Authorization": "Bearer " + apiKey,
      "Content-Type": "application/json"
    },
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  };

  var response = UrlFetchApp.fetch(url, options);
  var code = response.getResponseCode();
  if (code < 200 || code >= 300) {
    var err = response.getContentText();
    throw new Error("Resend Error (" + code + "): " + err);
  }
  return JSON.parse(response.getContentText());
}

/**
 * Builds a styled HTML email template for the QR badge.
 */
function buildQrEmailHtml(name, memberId, workstationName, isInlineCid, qrImageUrl) {
  var imageSrc = isInlineCid ? "cid:qrImage" : escapeHtml(qrImageUrl);
  
  return '<!DOCTYPE html>' +
  '<html>' +
  '<head><meta charset="utf-8"></head>' +
  '<body style="margin:0; padding:0; background-color:#F4F6F9; font-family:\'Segoe UI\', Arial, sans-serif; color:#333333;">' +
  '  <table width="100%" cellpadding="0" cellspacing="0" style="background-color:#F4F6F9; padding:20px 0;">' +
  '    <tr>' +
  '      <td align="center">' +
  '        <table width="520" cellpadding="0" cellspacing="0" style="background-color:#ffffff; border-radius:12px; box-shadow:0 4px 15px rgba(0,0,0,0.08); overflow:hidden; border:1px solid #E2E8F0;">' +
  '          <!-- Header -->' +
  '          <tr>' +
  '            <td style="background: linear-gradient(135deg, #1E3A8A 0%, #3B82F6 100%); padding:24px 30px; text-align:center; color:#ffffff;">' +
  '              <h1 style="margin:0; font-size:22px; font-weight:700; letter-spacing:0.5px;">🏢 ' + escapeHtml(workstationName) + '</h1>' +
  '              <p style="margin:6px 0 0 0; font-size:14px; opacity:0.9;">Official Workstation Access Pass</p>' +
  '            </td>' +
  '          </tr>' +
  '          <!-- Body -->' +
  '          <tr>' +
  '            <td style="padding:28px 32px; text-align:center;">' +
  '              <p style="font-size:16px; margin:0 0 16px 0; color:#1F2937; text-align:left;">Hello <strong>' + escapeHtml(name) + '</strong>,</p>' +
  '              <p style="font-size:14px; line-height:1.6; margin:0 0 20px 0; color:#4B5563; text-align:left;">' +
  '                Here is your official, permanent workstation QR code. This functions as your digital ID badge for automatic check-in at the entrance.' +
  '              </p>' +
  '              <!-- QR Badge Card -->' +
  '              <div style="background-color:#F8FAFC; border:2px dashed #CBD5E1; border-radius:10px; padding:20px; display:inline-block; margin:10px auto;">' +
  '                <img src="' + imageSrc + '" width="220" height="220" alt="Workstation QR Code" style="display:block; margin:0 auto; border-radius:6px; background:#fff; padding:6px; box-shadow:0 2px 6px rgba(0,0,0,0.05);" />' +
  '                <div style="margin-top:12px; font-size:15px; font-weight:700; color:#1E3A8A; letter-spacing:1px;">ID: ' + escapeHtml(memberId) + '</div>' +
  '                <div style="font-size:12px; color:#64748B; margin-top:2px;">' + escapeHtml(name) + '</div>' +
  '              </div>' +
  '              <!-- Important Notice Box -->' +
  '              <div style="background-color:#FEF3C7; border-left:4px solid #F59E0B; padding:12px 16px; border-radius:4px; text-align:left; margin-top:22px;">' +
  '                <div style="font-weight:700; color:#92400E; font-size:13px; margin-bottom:4px;">🛡️ IMPORTANT INSTRUCTIONS</div>' +
  '                <ul style="margin:0; padding-left:18px; color:#78350F; font-size:12px; line-height:1.5;">' +
  '                  <li><strong>Keep this QR safe:</strong> Save it on your phone or print it.</li>' +
  '                  <li><strong>Permanent ID:</strong> This QR code never expires and never changes.</li>' +
  '                  <li><strong>Fast Check-in:</strong> Scan this code in front of the kiosk scanner upon entering.</li>' +
  '                </ul>' +
  '              </div>' +
  '            </td>' +
  '          </tr>' +
  '          <!-- Footer -->' +
  '          <tr>' +
  '            <td style="background-color:#F1F5F9; padding:16px 30px; text-align:center; font-size:12px; color:#64748B; border-top:1px solid #E2E8F0;">' +
  '              Automated Workstation Attendance System &bull; ' + escapeHtml(workstationName) + '<br/>' +
  '              <em>The attached image file can be saved directly to your phone photos.</em>' +
  '            </td>' +
  '          </tr>' +
  '        </table>' +
  '      </td>' +
  '    </tr>' +
  '  </table>' +
  '</body>' +
  '</html>';
}

/**
 * Exports all member QR codes as individual PNG images inside a new Google Drive folder.
 */
function exportAllQrsToDrive() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SHEET_MEMBERS);
  if (!sheet) {
    SpreadsheetApp.getUi().alert("Members sheet not found. Please run Setup first.");
    return;
  }
  
  generateMissingTokensAndQrs(sheet);
  
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) {
    SpreadsheetApp.getUi().alert("No members found in Members sheet.");
    return;
  }
  
  var config = getConfigMap();
  var workstationName = config["Workstation Name"] || "Workstation";
  var timeStampStr = Utilities.formatDate(new Date(), SCRIPT_TIMEZONE, "yyyy-MM-dd_HHmm");
  var folderName = "QR_Badges_" + workstationName.replace(/[^a-zA-Z0-9_-]/g, "_") + "_" + timeStampStr;
  
  var folder = DriveApp.createFolder(folderName);
  var data = sheet.getRange(2, 1, lastRow - 1, 6).getValues();
  var count = 0;
  var errors = 0;
  
  SpreadsheetApp.getActiveSpreadsheet().toast("Generating and saving QR badges to Google Drive...", "⏳ Exporting QRs", 10);
  
  for (var i = 0; i < data.length; i++) {
    var memberId = data[i][0];
    var name = (data[i][1] || "Member").toString().replace(/[^a-zA-Z0-9_-]/g, "_");
    var qrUrl = data[i][4];
    
    if (memberId && qrUrl) {
      try {
        var resp = UrlFetchApp.fetch(qrUrl, { muteHttpExceptions: true });
        if (resp.getResponseCode() === 200) {
          var fileName = memberId + "_" + name + ".png";
          var blob = resp.getBlob().setName(fileName);
          folder.createFile(blob);
          count++;
        }
      } catch (e) {
        errors++;
      }
    }
  }
  
  var html = 
    '<div style="font-family:Arial,sans-serif; padding:16px;">' +
    '  <h3 style="color:#1E3A8A; margin-top:0;">✅ QR Badges Exported to Drive</h3>' +
    '  <p style="font-size:13px; color:#333;"><strong>' + count + '</strong> QR badges were generated and saved in Google Drive:</p>' +
    '  <p style="font-size:13px; color:#555;">📁 <strong>' + escapeHtml(folder.getName()) + '</strong></p>' +
    '  <p style="margin:20px 0;"><a href="' + folder.getUrl() + '" target="_blank" style="background:#2563EB; color:#fff; padding:10px 16px; text-decoration:none; border-radius:6px; font-weight:bold; display:inline-block;">Open Folder in Google Drive &rarr;</a></p>' +
    '  <p style="font-size:11px; color:#888;">Tip: You can download the entire folder as a ZIP file from Google Drive, or share the folder link directly with your members!</p>' +
    '</div>';
    
  var dialog = HtmlService.createHtmlOutput(html).setWidth(450).setHeight(230);
  SpreadsheetApp.getUi().showModalDialog(dialog, "📁 Google Drive Export");
}

/**
 * Modal dialog for easy email provider configuration (Gmail / Brevo / Resend).
 */
function showEmailConfigDialog() {
  var config = getConfigMap();
  var currentProvider = config["Email Provider"] || "Gmail (Default)";
  var currentApiKey = config["Email API Key"] || "";
  var currentSenderEmail = config["Sender Email"] || getAdminEmailSafe();
  var currentSenderName = config["Sender Name"] || "Workstation Admin";
  var quotaRemaining = MailApp.getRemainingDailyQuota();
  
  var html = 
    '<!DOCTYPE html>' +
    '<html><head><style>' +
    '  body { font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, sans-serif; padding:16px 20px; color:#1e293b; background:#f8fafc; margin:0; font-size:13px; }' +
    '  h2 { font-size:16px; color:#0f172a; margin:0 0 6px 0; font-weight:700; }' +
    '  p.desc { font-size:12px; color:#64748b; margin:0 0 16px 0; }' +
    '  .form-group { margin-bottom:12px; }' +
    '  label { display:block; font-weight:600; font-size:12px; margin-bottom:4px; color:#334155; }' +
    '  select, input[type="text"] { width:100%; padding:8px 10px; border:1px solid #cbd5e1; border-radius:6px; font-size:12px; box-sizing:border-box; background:#fff; }' +
    '  select:focus, input[type="text"]:focus { outline:none; border-color:#3b82f6; box-shadow:0 0 0 3px rgba(59,130,246,0.15); }' +
    '  .badge-quota { display:inline-block; font-size:11px; padding:3px 8px; border-radius:12px; background:#e0f2fe; color:#0369a1; font-weight:600; margin-bottom:12px; }' +
    '  .actions { display:flex; justify-content:flex-end; gap:8px; margin-top:18px; }' +
    '  button { padding:8px 16px; border-radius:6px; font-weight:600; font-size:12px; cursor:pointer; border:none; }' +
    '  .btn-primary { background:#2563eb; color:#fff; }' +
    '  .btn-primary:hover { background:#1d4ed8; }' +
    '  .btn-secondary { background:#e2e8f0; color:#334155; }' +
    '  .link-help { font-size:11px; color:#2563eb; text-decoration:none; margin-top:3px; display:inline-block; }' +
    '</style></head><body>' +
    '  <h2>⚡ Email Provider Configuration</h2>' +
    '  <p class="desc">Send QR badges via Gmail or connect a free email API for sending 300+ emails in 1 day.</p>' +
    '  <div class="badge-quota">Gmail Remaining Quota: ' + quotaRemaining + ' emails</div>' +
    '  <form id="cfgForm">' +
    '    <div class="form-group">' +
    '      <label>Email Provider</label>' +
    '      <select id="prov" onchange="toggleApiHelp()">' +
    '        <option value="Gmail (Default)"' + (currentProvider.indexOf("Gmail") !== -1 ? ' selected' : '') + '>Gmail (Default 100/day free)</option>' +
    '        <option value="Brevo"' + (currentProvider.indexOf("Brevo") !== -1 ? ' selected' : '') + '>Brevo / Sendinblue (300/day 100% Free)</option>' +
    '        <option value="Resend"' + (currentProvider.indexOf("Resend") !== -1 ? ' selected' : '') + '>Resend (3,000/month Free)</option>' +
    '      </select>' +
    '    </div>' +
    '    <div class="form-group">' +
    '      <label>Email API Key (Required for Brevo / Resend)</label>' +
    '      <input type="text" id="apikey" placeholder="e.g. xkeysib-... or re_..." value="' + escapeHtml(currentApiKey) + '">' +
    '      <a id="apiHelp" class="link-help" target="_blank" href="https://app.brevo.com/settings/keys/api">👉 Get free Brevo API Key (Instant)</a>' +
    '    </div>' +
    '    <div class="form-group">' +
    '      <label>Sender Email Address</label>' +
    '      <input type="text" id="senderMail" value="' + escapeHtml(currentSenderEmail) + '">' +
    '    </div>' +
    '    <div class="form-group">' +
    '      <label>Sender Display Name</label>' +
    '      <input type="text" id="senderName" value="' + escapeHtml(currentSenderName) + '">' +
    '    </div>' +
    '    <div class="actions">' +
    '      <button type="button" class="btn-secondary" onclick="google.script.host.close()">Cancel</button>' +
    '      <button type="button" class="btn-primary" onclick="saveData()">Save Settings</button>' +
    '    </div>' +
    '  </form>' +
    '  <script>' +
    '    function toggleApiHelp() {' +
    '      var p = document.getElementById("prov").value;' +
    '      var h = document.getElementById("apiHelp");' +
    '      if (p === "Brevo") { h.href = "https://app.brevo.com/settings/keys/api"; h.innerText = "👉 Get Free Brevo API Key (Instant)"; h.style.display="inline-block"; }' +
    '      else if (p === "Resend") { h.href = "https://resend.com/api-keys"; h.innerText = "👉 Get Free Resend API Key (Instant)"; h.style.display="inline-block"; }' +
    '      else { h.style.display="none"; }' +
    '    }' +
    '    toggleApiHelp();' +
    '    function saveData() {' +
    '      var prov = document.getElementById("prov").value;' +
    '      var key = document.getElementById("apikey").value.trim();' +
    '      var email = document.getElementById("senderMail").value.trim();' +
    '      var name = document.getElementById("senderName").value.trim();' +
    '      google.script.run.withSuccessHandler(function() {' +
    '        google.script.host.close();' +
    '      }).saveEmailConfig(prov, key, email, name);' +
    '    }' +
    '  </script>' +
    '</body></html>';
    
  var dialog = HtmlService.createHtmlOutput(html).setWidth(460).setHeight(380);
  SpreadsheetApp.getUi().showModalDialog(dialog, "⚡ Email Provider Setup");
}

/**
 * Saves email configuration directly into Config sheet.
 */
function saveEmailConfig(provider, apiKey, senderEmail, senderName) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SHEET_CONFIG);
  if (!sheet) return;
  
  var data = sheet.getDataRange().getValues();
  var map = {
    "Email Provider": provider,
    "Email API Key": apiKey,
    "Sender Email": senderEmail,
    "Sender Name": senderName
  };
  
  var updatedKeys = {};
  for (var i = 1; i < data.length; i++) {
    var key = String(data[i][0]).trim();
    if (map[key] !== undefined) {
      sheet.getRange(i + 1, 2).setValue(map[key]);
      updatedKeys[key] = true;
    }
  }
  
  // If any key was not already in sheet, append it
  for (var k in map) {
    if (!updatedKeys[k]) {
      sheet.appendRow([k, map[k], "Configured via Email Setup Dialog"]);
    }
  }
  
  clearSystemCache();
  SpreadsheetApp.getActiveSpreadsheet().toast("Email provider configured successfully!", "✅ Saved", 3);
}

// ----------------------------------------------------------------------------
// 4. FEATURE 2 & 3: SCANNER WEB APP (doGet & markAttendance)
// ----------------------------------------------------------------------------

/**
 * Serves the mobile-friendly HTML5 QR Scanner web application.
 */
function doGet(e) {
  if (e && e.parameter && e.parameter.action === "config") {
    var configData = getWorkstationConfig();
    return ContentService.createTextOutput(JSON.stringify(configData))
      .setMimeType(ContentService.MimeType.JSON);
  }
  
  var template = HtmlService.createTemplateFromFile("Scanner");
  return template.evaluate()
    .setTitle("Workstation QR Attendance Scanner")
    .addMetaTag("viewport", "width=device-width, initial-scale=1.0, maximum-scale=1.0, user-scalable=no")
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

/**
 * Handles HTTP POST requests for external / standalone scanner web apps.
 */
function doPost(e) {
  try {
    var params = {};
    if (e.postData && e.postData.contents) {
      try {
        params = JSON.parse(e.postData.contents);
      } catch (parseErr) {
        params = e.parameter || {};
      }
    } else {
      params = e.parameter || {};
    }
    
    var qrData = params.qrData;
    var deviceInfo = params.deviceInfo || "Standalone Kiosk App";
    
    var result = markAttendance(qrData, deviceInfo);
    
    return ContentService.createTextOutput(JSON.stringify(result))
      .setMimeType(ContentService.MimeType.JSON);
  } catch (err) {
    return ContentService.createTextOutput(JSON.stringify({
      success: false,
      status: "ERROR",
      title: "Server Error",
      message: err.toString()
    })).setMimeType(ContentService.MimeType.JSON);
  }
}

/**
 * Helper to fetch config for the web client.
 */
function getWorkstationConfig() {
  var config = getConfigMap();
  return {
    workstationName: config["Workstation Name"] || "Workstation",
    checkInStart: config["Check-in Start Time"] || "08:00",
    checkInEnd: config["Check-in End Time"] || "20:00"
  };
}

/**
 * Retrieves cached dictionary of { memberId: memberName } with 1-hour TTL.
 */
function getMemberMapCached(ss) {
  var cache = CacheService.getScriptCache();
  var cached = cache.get("MEMBERS_MAP");
  if (cached) {
    try {
      return JSON.parse(cached);
    } catch (e) {}
  }
  
  var membersSheet = ss.getSheetByName(SHEET_MEMBERS);
  if (!membersSheet) return {};
  var lastRow = membersSheet.getLastRow();
  if (lastRow < 2) return {};
  
  var data = membersSheet.getRange(2, 1, lastRow - 1, 2).getValues(); // Only Col A & B
  var map = {};
  for (var i = 0; i < data.length; i++) {
    var id = String(data[i][0]).trim().toUpperCase();
    var name = String(data[i][1]).trim();
    if (id) {
      map[id] = name;
    }
  }
  
  try {
    cache.put("MEMBERS_MAP", JSON.stringify(map), 3600); // 1 hour
  } catch (e) {}
  
  return map;
}

/**
 * High-Speed Attendance Marking with In-Memory Caching and Concurrency Locks.
 * @param {string} qrData - Scanned content ("MemberID|Token")
 * @param {string} deviceInfo - Optional scanner details
 * @return {object} JSON status object
 */
function markAttendance(qrData, deviceInfo) {
  // 1. Instant format validation (<1ms)
  if (!qrData || typeof qrData !== "string") {
    return {
      success: false,
      status: "INVALID_FORMAT",
      title: "Invalid QR Code",
      message: "QR code data is empty or unreadable."
    };
  }
  
  var parts = qrData.trim().split("|");
  if (parts.length !== 2) {
    return {
      success: false,
      status: "INVALID_FORMAT",
      title: "Invalid QR Format",
      message: "Unrecognized QR code structure."
    };
  }
  
  var memberId = parts[0].trim().toUpperCase();
  var token = parts[1].trim();
  
  // 2. In-memory cryptographic HMAC verification (<1ms)
  if (!verifyMemberToken(memberId, token)) {
    return {
      success: false,
      status: "INVALID_SIGNATURE",
      title: "Counterfeit QR Code",
      message: "Cryptographic signature failed. Unauthorized or tampered badge."
    };
  }
  
  // 3. Fast Config & Time Window Check (Cached, <2ms)
  var config = getConfigMap();
  var startTime = sanitizeTimeString(config["Check-in Start Time"], "00:00");
  var endTime = sanitizeTimeString(config["Check-in End Time"], "23:59");
  var allowedRescan = String(config["Allowed Re-scan"] || "No").trim().toLowerCase() === "yes";
  
  var now = new Date();
  var currentTimeStr = Utilities.formatDate(now, SCRIPT_TIMEZONE, "HH:mm");
  var todayDateStr = Utilities.formatDate(now, SCRIPT_TIMEZONE, "yyyy-MM-dd");
  var readableTimeStr = Utilities.formatDate(now, SCRIPT_TIMEZONE, "hh:mm:ss a");
  
  if (currentTimeStr < startTime || currentTimeStr > endTime) {
    return {
      success: false,
      status: "OUTSIDE_HOURS",
      title: "Outside Check-in Window",
      memberId: memberId,
      message: "Check-in allowed only between " + startTime + " and " + endTime + ". (Current: " + currentTimeStr + ")"
    };
  }

  // 4. Ultra-Fast Duplicate Check via CacheService (<5ms)
  var cache = CacheService.getScriptCache();
  var cacheKey = "ATTEND_" + todayDateStr + "_" + memberId;
  var cachedScan = cache.get(cacheKey);
  if (cachedScan && !allowedRescan) {
    var partsCache = cachedScan.split("|");
    var memberNameCached = partsCache[0] || memberId;
    var timeCached = partsCache[1] || "";
    return {
      success: false,
      status: "ALREADY_MARKED",
      title: "Already Present",
      name: memberNameCached,
      memberId: memberId,
      time: timeCached,
      message: memberNameCached + ", you are already marked present today" + (timeCached ? " at " + timeCached : "") + "."
    };
  }

  // 5. Concurrency lock & Record Log
  var lock = LockService.getScriptLock();
  try {
    var hasLock = lock.tryLock(4000); // 4-second fast lock
    if (!hasLock) {
      return {
        success: false,
        status: "BUSY",
        title: "System Busy",
        message: "Server is processing another scan. Please try again."
      };
    }
    
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var memberMap = getMemberMapCached(ss);
    var memberName = memberMap[memberId];
    
    if (!memberName) {
      cache.remove("MEMBERS_MAP");
      memberMap = getMemberMapCached(ss);
      memberName = memberMap[memberId];
    }
    
    if (!memberName) {
      return {
        success: false,
        status: "NOT_FOUND",
        title: "Member Not Found",
        message: "Member ID " + memberId + " is not registered in the system."
      };
    }
    
    var logSheet = ss.getSheetByName(SHEET_ATTENDANCE);
    if (!logSheet) {
      logSheet = getOrCreateSheet(ss, SHEET_ATTENDANCE);
    }
    
    var deviceString = deviceInfo || "Web App Kiosk";
    logSheet.appendRow([now, todayDateStr, memberId, memberName, "Present", deviceString]);
    
    // Store in cache for 24 hours so duplicate re-scans return in <10ms
    try {
      cache.put(cacheKey, memberName + "|" + readableTimeStr, 86400);
    } catch (e) {}
    
    return {
      success: true,
      status: "SUCCESS",
      title: "Check-in Successful",
      name: memberName,
      memberId: memberId,
      time: readableTimeStr,
      message: "Welcome, " + memberName + "! Attendance marked at " + readableTimeStr + "."
    };
    
  } catch (err) {
    return {
      success: false,
      status: "ERROR",
      title: "Server Error",
      message: err.toString()
    };
  } finally {
    lock.releaseLock();
  }
}

// ----------------------------------------------------------------------------
// 5. FEATURE 4: DAILY ATTENDANCE REPORT & AUTOMATION
// ----------------------------------------------------------------------------

/**
 * Menu wrapper to generate today's report manually.
 */
function generateTodayReportMenu() {
  var result = generateDailyReport();
  if (result.success) {
    SpreadsheetApp.getUi().alert(
      "📊 Report Generated",
      "Daily Report [" + result.sheetName + "] created successfully!\n\n" +
      "• Total Members: " + result.total + "\n" +
      "• Present: " + result.present + "\n" +
      "• Absent: " + result.absent + "\n" +
      "• Attendance Rate: " + result.percentage + "%",
      SpreadsheetApp.getUi().ButtonSet.OK
    );
  } else {
    SpreadsheetApp.getUi().alert("Error", result.message, SpreadsheetApp.getUi().ButtonSet.OK);
  }
}

/**
 * Generates or updates the daily report sheet (Report_YYYY-MM-DD).
 * @param {string} [targetDateStr] - Format "yyyy-MM-dd". Defaults to today.
 * @return {object} Summary result
 */
function generateDailyReport(targetDateStr) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var today = new Date();
  
  if (!targetDateStr) {
    targetDateStr = Utilities.formatDate(today, SCRIPT_TIMEZONE, "yyyy-MM-dd");
  }
  
  var reportSheetName = REPORT_PREFIX + targetDateStr;
  var membersSheet = ss.getSheetByName(SHEET_MEMBERS);
  var logSheet = ss.getSheetByName(SHEET_ATTENDANCE);
  
  if (!membersSheet) {
    return { success: false, message: "Members sheet does not exist." };
  }
  
  // 1. Fetch all members
  var membersData = membersSheet.getDataRange().getValues();
  if (membersData.length < 2) {
    return { success: false, message: "No members registered in system." };
  }
  
  var membersList = [];
  for (var i = 1; i < membersData.length; i++) {
    var id = String(membersData[i][0]).trim().toUpperCase();
    var name = membersData[i][1];
    var email = membersData[i][2];
    if (id) {
      membersList.push({ id: id, name: name, email: email });
    }
  }
  
  // 2. Fetch all logs for target date
  var attendanceMap = {};
  if (logSheet && logSheet.getLastRow() > 1) {
    var logs = logSheet.getDataRange().getValues();
    for (var j = 1; j < logs.length; j++) {
      var logDate = logs[j][1];
      var logDateFormatted = (logDate instanceof Date)
        ? Utilities.formatDate(logDate, SCRIPT_TIMEZONE, "yyyy-MM-dd")
        : String(logDate).trim();
        
      if (logDateFormatted === targetDateStr) {
        var mId = String(logs[j][2]).trim().toUpperCase();
        var checkInTime = (logs[j][0] instanceof Date)
          ? Utilities.formatDate(logs[j][0], SCRIPT_TIMEZONE, "hh:mm:ss a")
          : String(logs[j][0]);
          
        if (!attendanceMap[mId]) {
          attendanceMap[mId] = {
            status: "Present",
            time: checkInTime
          };
        }
      }
    }
  }
  
  // 3. Compute stats & rows
  var totalMembers = membersList.length;
  var presentCount = 0;
  var reportRows = [];
  
  for (var k = 0; k < membersList.length; k++) {
    var m = membersList[k];
    var isPresent = attendanceMap[m.id];
    if (isPresent) {
      presentCount++;
      reportRows.push([m.id, m.name, m.email, "Present", isPresent.time]);
    } else {
      reportRows.push([m.id, m.name, m.email, "Absent", "-"]);
    }
  }
  
  var absentCount = totalMembers - presentCount;
  var attendancePct = totalMembers > 0 ? Math.round((presentCount / totalMembers) * 100) : 0;
  
  // 4. Create or recreate the Report Sheet
  var reportSheet = ss.getSheetByName(reportSheetName);
  if (reportSheet) {
    reportSheet.clear();
  } else {
    reportSheet = ss.insertSheet(reportSheetName);
  }
  
  // 5. Design Premium KPI Summary Header
  var config = getConfigMap();
  var workstationName = config["Workstation Name"] || "Workstation";
  
  // Title
  reportSheet.getRange("A1:E1").merge()
    .setValue("🏢 " + workstationName.toUpperCase() + " — DAILY ATTENDANCE REPORT")
    .setFontWeight("bold")
    .setFontSize(14)
    .setFontColor("#ffffff")
    .setBackground("#1E3A8A")
    .setHorizontalAlignment("center")
    .setVerticalAlignment("middle");
  reportSheet.setRowHeight(1, 36);
  
  // Date subtitle
  reportSheet.getRange("A2:E2").merge()
    .setValue("Report Date: " + targetDateStr + " | Generated At: " + Utilities.formatDate(today, SCRIPT_TIMEZONE, "yyyy-MM-dd hh:mm:ss a"))
    .setFontSize(10)
    .setFontColor("#4B5563")
    .setBackground("#F1F5F9")
    .setHorizontalAlignment("center");
  reportSheet.setRowHeight(2, 24);
  
  // KPI Cards Row
  reportSheet.getRange("A3").setValue("Total Members: " + totalMembers).setBackground("#E0E7FF").setFontColor("#3730A3").setFontWeight("bold").setHorizontalAlignment("center");
  reportSheet.getRange("B3").setValue("Present: " + presentCount).setBackground("#DCFCE7").setFontColor("#166534").setFontWeight("bold").setHorizontalAlignment("center");
  reportSheet.getRange("C3").setValue("Absent: " + absentCount).setBackground("#FEE2E2").setFontColor("#991B1B").setFontWeight("bold").setHorizontalAlignment("center");
  reportSheet.getRange("D3:E3").merge().setValue("Attendance Rate: " + attendancePct + "%").setBackground("#FEF3C7").setFontColor("#92400E").setFontWeight("bold").setHorizontalAlignment("center");
  reportSheet.setRowHeight(3, 28);
  
  // Spacer row
  reportSheet.setRowHeight(4, 12);
  
  // Table Header
  var tableHeaders = [["Member ID", "Member Name", "Email Address", "Attendance Status", "Check-in Time"]];
  reportSheet.getRange(5, 1, 1, 5).setValues(tableHeaders);
  formatHeaderRow(reportSheet, 5, 5, "#334155");
  reportSheet.setRowHeight(5, 28);
  
  // Write Data Rows
  if (reportRows.length > 0) {
    reportSheet.getRange(6, 1, reportRows.length, 5).setValues(reportRows);
    
    // Conditional Row Styling
    for (var r = 0; r < reportRows.length; r++) {
      var rowNum = 6 + r;
      var status = reportRows[r][3];
      var range = reportSheet.getRange(rowNum, 1, 1, 5);
      
      if (status === "Present") {
        range.setBackground("#F0FDF4");
        reportSheet.getRange(rowNum, 4).setFontColor("#15803D").setFontWeight("bold");
      } else {
        range.setBackground("#FEF2F2");
        reportSheet.getRange(rowNum, 4).setFontColor("#B91C1C").setFontWeight("bold");
        reportSheet.getRange(rowNum, 5).setFontColor("#9CA3AF");
      }
      reportSheet.setRowHeight(rowNum, 24);
    }
  }
  
  // Formatting & Alignment
  reportSheet.setColumnWidth(1, 130);
  reportSheet.setColumnWidth(2, 180);
  reportSheet.setColumnWidth(3, 230);
  reportSheet.setColumnWidth(4, 150);
  reportSheet.setColumnWidth(5, 160);
  reportSheet.getRange("A5:E" + (5 + reportRows.length)).setBorder(true, true, true, true, true, true, "#CBD5E1", SpreadsheetApp.BorderStyle.SOLID);
  reportSheet.setFrozenRows(5);
  
  return {
    success: true,
    sheetName: reportSheetName,
    total: totalMembers,
    present: presentCount,
    absent: absentCount,
    percentage: attendancePct,
    sheetUrl: ss.getUrl() + "#gid=" + reportSheet.getSheetId()
  };
}

/**
 * Scheduled job triggered daily to generate the report and email it to the Admin.
 */
function dailyReportTriggerJob() {
  var result = generateDailyReport();
  if (!result.success) {
    Logger.log("Daily report generation failed: " + result.message);
    return;
  }
  
  var config = getConfigMap();
  var adminEmail = config["Admin Email"];
  var workstationName = config["Workstation Name"] || "Workstation";
  
  if (!adminEmail || adminEmail.indexOf("@") === -1) {
    Logger.log("Admin email not configured. Skipping email dispatch.");
    return;
  }
  
  var subject = "📊 Daily Attendance Summary: " + workstationName + " (" + result.sheetName.replace(REPORT_PREFIX, "") + ")";
  var htmlBody = 
    '<div style="font-family:Arial,sans-serif; max-width:600px; color:#333; line-height:1.6;">' +
    '  <h2 style="color:#1E3A8A; margin-bottom:4px;">' + escapeHtml(workstationName) + ' Attendance Report</h2>' +
    '  <p style="color:#64748B; margin-top:0;">Date: <strong>' + escapeHtml(result.sheetName.replace(REPORT_PREFIX, "")) + '</strong></p>' +
    '  <table style="width:100%; border-collapse:collapse; margin:16px 0; text-align:center;">' +
    '    <tr>' +
    '      <td style="background:#E0E7FF; color:#3730A3; padding:12px; font-weight:bold; border-radius:6px 0 0 6px;">Total<br><span style="font-size:18px;">' + result.total + '</span></td>' +
    '      <td style="background:#DCFCE7; color:#166534; padding:12px; font-weight:bold;">Present<br><span style="font-size:18px;">' + result.present + '</span></td>' +
    '      <td style="background:#FEE2E2; color:#991B1B; padding:12px; font-weight:bold;">Absent<br><span style="font-size:18px;">' + result.absent + '</span></td>' +
    '      <td style="background:#FEF3C7; color:#92400E; padding:12px; font-weight:bold; border-radius:0 6px 6px 0;">Attendance<br><span style="font-size:18px;">' + result.percentage + '%</span></td>' +
    '    </tr>' +
    '  </table>' +
    '  <p>The detailed sheet has been recorded under <strong>' + escapeHtml(result.sheetName) + '</strong>.</p>' +
    '  <p><a href="' + result.sheetUrl + '" style="background:#2563EB; color:#fff; padding:10px 18px; text-decoration:none; border-radius:6px; display:inline-block; font-weight:bold;">Open Full Report in Google Sheets &rarr;</a></p>' +
    '</div>';
    
  MailApp.sendEmail({
    to: adminEmail,
    subject: subject,
    htmlBody: htmlBody
  });
}

/**
 * Sets up the daily automatic time-driven trigger based on Config setting.
 */
function setupTriggers() {
  // Clear any existing report triggers
  var triggers = ScriptApp.getProjectTriggers();
  for (var i = 0; i < triggers.length; i++) {
    if (triggers[i].getHandlerFunction() === "dailyReportTriggerJob") {
      ScriptApp.deleteTrigger(triggers[i]);
    }
  }
  
  var config = getConfigMap();
  var genTime = config["Report Generation Time"] || "21:00";
  var hour = parseInt(genTime.split(":")[0], 10);
  if (isNaN(hour) || hour < 0 || hour > 23) {
    hour = 21;
  }
  
  ScriptApp.newTrigger("dailyReportTriggerJob")
    .timeBased()
    .everyDays(1)
    .atHour(hour)
    .inTimezone(SCRIPT_TIMEZONE)
    .create();
    
  SpreadsheetApp.getUi().alert(
    "⏰ Trigger Configured",
    "Daily report trigger successfully scheduled at ~" + (hour < 10 ? "0" + hour : hour) + ":00 " + SCRIPT_TIMEZONE + " every day.",
    SpreadsheetApp.getUi().ButtonSet.OK
  );
}

// ----------------------------------------------------------------------------
// 6. TESTING & UTILITIES
// ----------------------------------------------------------------------------

/**
 * Simulates a test scan for a member.
 */
function testScanMember() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SHEET_MEMBERS);
  if (!sheet || sheet.getLastRow() < 2) {
    SpreadsheetApp.getUi().alert("Please run Setup first to create sample members.");
    return;
  }
  
  var sampleMemberId = sheet.getRange("A2").getValue();
  var sampleToken = sheet.getRange("D2").getValue();
  
  if (!sampleToken) {
    sampleToken = generateMemberToken(sampleMemberId);
  }
  
  var testQrData = sampleMemberId + "|" + sampleToken;
  var result = markAttendance(testQrData, "Simulator Test Scan");
  
  SpreadsheetApp.getUi().alert(
    "🧪 Test Scan Simulation",
    "Scanned Payload: " + testQrData + "\n\n" +
    "Result Status: " + result.status + "\n" +
    "Message: " + result.message,
    SpreadsheetApp.getUi().ButtonSet.OK
  );
}

/**
 * Displays the deployed web app URL or instructions to deploy.
 */
function showScannerUrlDialog() {
  var url = ScriptApp.getService().getUrl();
  var ui = SpreadsheetApp.getUi();
  
  if (!url) {
    ui.alert(
      "🌐 Web App Not Deployed Yet",
      "To obtain the scanner link:\n" +
      "1. Click 'Deploy' > 'New deployment' (top right of Apps Script editor).\n" +
      "2. Select type: 'Web app'.\n" +
      "3. Execute as: 'Me' (your account).\n" +
      "4. Who has access: 'Anyone'.\n" +
      "5. Copy the /exec URL and open it on your mobile device or kiosk tablet.",
      ui.ButtonSet.OK
    );
  } else {
    var htmlOutput = HtmlService.createHtmlOutput(
      '<div style="font-family:Arial,sans-serif; padding:16px; text-align:center;">' +
      '  <h3 style="color:#1E3A8A; margin-top:0;">📱 Workstation Scanner URL</h3>' +
      '  <p style="font-size:13px; color:#555;">Open this URL on any smartphone or entrance kiosk device:</p>' +
      '  <input type="text" readonly value="' + url + '" style="width:100%; padding:10px; font-size:12px; border:1px solid #ccc; border-radius:4px; box-sizing:border-box; margin-bottom:12px;" onclick="this.select();">' +
      '  <p style="font-size:11px; color:#888;">Tip: Print this URL as a QR code and paste it near the entrance for fast kiosk access!</p>' +
      '</div>'
    ).setWidth(420).setHeight(200);
    ui.showModalDialog(htmlOutput, "Scanner Web App");
  }
}

// ----------------------------------------------------------------------------
// 7. HELPER FUNCTIONS
// ----------------------------------------------------------------------------

function getOrCreateSheet(ss, sheetName) {
  var sheet = ss.getSheetByName(sheetName);
  if (!sheet) {
    sheet = ss.insertSheet(sheetName);
  }
  return sheet;
}

function formatHeaderRow(sheet, rowNum, numCols, bgColor) {
  var header = sheet.getRange(rowNum, 1, 1, numCols);
  header.setBackground(bgColor)
    .setFontColor("#FFFFFF")
    .setFontWeight("bold")
    .setFontSize(10)
    .setVerticalAlignment("middle");
  sheet.setRowHeight(rowNum, 32);
}

/**
 * Clears in-memory script cache when config or members change.
 */
function clearSystemCache() {
  var cache = CacheService.getScriptCache();
  cache.remove("APP_CONFIG_MAP");
  cache.remove("MEMBERS_MAP");
  try {
    SpreadsheetApp.getActiveSpreadsheet().toast("System cache refreshed successfully!", "🔄 Cache Cleared", 3);
  } catch (e) {}
}

/**
 * Sanitizes any Date object, ISO string, or time string into standard 'HH:mm' 24-hr format.
 */
function sanitizeTimeString(val, defaultVal) {
  if (!val) return defaultVal || "00:00";
  if (val instanceof Date) {
    return Utilities.formatDate(val, SCRIPT_TIMEZONE, "HH:mm");
  }
  var s = String(val).trim();
  if (s.indexOf("T") !== -1 || s.indexOf("1899") !== -1 || s.indexOf("Z") !== -1) {
    try {
      var d = new Date(s);
      if (!isNaN(d.getTime())) {
        return Utilities.formatDate(d, SCRIPT_TIMEZONE, "HH:mm");
      }
    } catch (e) {}
  }
  var match = s.match(/^(\d{1,2}):(\d{2})/);
  if (match) {
    var h = match[1].length === 1 ? "0" + match[1] : match[1];
    var m = match[2];
    return h + ":" + m;
  }
  return defaultVal || "00:00";
}

function getConfigMap() {
  var cache = CacheService.getScriptCache();
  var cached = cache.get("APP_CONFIG_MAP");
  if (cached) {
    try {
      var parsed = JSON.parse(cached);
      if (parsed["Check-in Start Time"]) parsed["Check-in Start Time"] = sanitizeTimeString(parsed["Check-in Start Time"], "00:00");
      if (parsed["Check-in End Time"]) parsed["Check-in End Time"] = sanitizeTimeString(parsed["Check-in End Time"], "23:59");
      return parsed;
    } catch (e) {}
  }
  
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SHEET_CONFIG);
  var config = {};
  if (!sheet) return config;
  
  var data = sheet.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    var key = String(data[i][0]).trim();
    var val = data[i][1];
    if (key) {
      if (val instanceof Date) {
        val = Utilities.formatDate(val, SCRIPT_TIMEZONE, "HH:mm");
      } else if (typeof val === "string" && key.indexOf("Time") !== -1) {
        val = sanitizeTimeString(val, val);
      }
      config[key] = val;
    }
  }
  
  if (!config["Check-in Start Time"]) config["Check-in Start Time"] = "00:00";
  if (!config["Check-in End Time"]) config["Check-in End Time"] = "23:59";
  
  try {
    cache.put("APP_CONFIG_MAP", JSON.stringify(config), 3600); // Cache for 1 hour
  } catch (e) {}
  
  return config;
}

function escapeHtml(text) {
  if (!text) return "";
  return String(text)
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;")
    .replace(/'/g, "&#039;");
}

function getAdminEmailSafe() {
  try {
    var email = Session.getActiveUser().getEmail();
    if (email && email.indexOf("@") !== -1) return email;
  } catch (e) {}
  try {
    var effectiveEmail = Session.getEffectiveUser().getEmail();
    if (effectiveEmail && effectiveEmail.indexOf("@") !== -1) return effectiveEmail;
  } catch (e) {}
  return "admin@example.com";
}

/**
 * Populates 10 realistic sample members and mock attendance logs for instant testing.
 */
function populateDemoData() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  setup(); // ensure sheets & secret key exist
  
  var membersSheet = ss.getSheetByName(SHEET_MEMBERS);
  var logSheet = ss.getSheetByName(SHEET_ATTENDANCE);
  
  // 10 Sample Members
  var demoMembers = [
    ["WS-001", "Aarav Sharma", "aarav.sharma@example.com", "", "", "No"],
    ["WS-002", "Priya Patel", "priya.patel@example.com", "", "", "No"],
    ["WS-003", "Rahul Verma", "rahul.verma@example.com", "", "", "No"],
    ["WS-004", "Ananya Iyer", "ananya.iyer@example.com", "", "", "No"],
    ["WS-005", "Vikram Singh", "vikram.singh@example.com", "", "", "No"],
    ["WS-006", "Sneha Kulkarni", "sneha.kulkarni@example.com", "", "", "No"],
    ["WS-007", "Rohan Mehta", "rohan.mehta@example.com", "", "", "No"],
    ["WS-008", "Neha Gupta", "neha.gupta@example.com", "", "", "No"],
    ["WS-009", "Karthik Raja", "karthik.raja@example.com", "", "", "No"],
    ["WS-010", "Pooja Deshmukh", "pooja.deshmukh@example.com", "", "", "No"]
  ];
  
  // Clear existing member data rows (keep header)
  var lastRow = membersSheet.getLastRow();
  if (lastRow > 1) {
    membersSheet.getRange(2, 1, lastRow - 1, 6).clearContent();
  }
  
  // Populate new members
  membersSheet.getRange(2, 1, demoMembers.length, 6).setValues(demoMembers);
  generateMissingTokensAndQrs(membersSheet);
  
  // Add 6 realistic attendance logs for today
  var now = new Date();
  var todayDateStr = Utilities.formatDate(now, SCRIPT_TIMEZONE, "yyyy-MM-dd");
  
  var sampleLogs = [
    [new Date(now.getTime() - 3600000 * 3.5), todayDateStr, "WS-001", "Aarav Sharma", "Present", "Entrance Kiosk 1"],
    [new Date(now.getTime() - 3600000 * 3.0), todayDateStr, "WS-002", "Priya Patel", "Present", "Entrance Kiosk 1"],
    [new Date(now.getTime() - 3600000 * 2.2), todayDateStr, "WS-003", "Rahul Verma", "Present", "Entrance Kiosk 2"],
    [new Date(now.getTime() - 3600000 * 1.8), todayDateStr, "WS-004", "Ananya Iyer", "Present", "Mobile Web Scanner"],
    [new Date(now.getTime() - 3600000 * 1.2), todayDateStr, "WS-005", "Vikram Singh", "Present", "Entrance Kiosk 1"],
    [new Date(now.getTime() - 3600000 * 0.4), todayDateStr, "WS-006", "Sneha Kulkarni", "Present", "Entrance Kiosk 2"]
  ];
  
  var logLastRow = logSheet.getLastRow();
  if (logLastRow > 1) {
    logSheet.getRange(2, 1, logLastRow - 1, 6).clearContent();
  }
  logSheet.getRange(2, 1, sampleLogs.length, 6).setValues(sampleLogs);
  
  SpreadsheetApp.getActiveSpreadsheet().toast(
    "10 Members & 6 Check-in Logs populated!",
    "✨ Demo Data Loaded",
    5
  );
}
