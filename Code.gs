// ==========================================
// CONFIGURATION
// ==========================================
const SHEET_ID = '1E2z4VJ_iframYJyQgMN_0rKRo4lGIZHbyMZo3eyy4Q4'; 
const DRIVE_FOLDER_ID = '1cALybDaiQab1PC3tf-LjF1v9hGfCiFSO'; 
const OFFICER_EMAIL = 'ray01@anglosingapore.ac.th'; // The officer who gets the first notification

/**
 * ⚠️ REQUIRED FIRST STEP ⚠️
 * Select 'setupPermissions' from the dropdown at the top and click 'Run'.
 * This forces Google to ask for the FULL necessary permissions (Drive, Sheets, Email).
 */
function setupPermissions() {
  // This explicitly tells Google we need WRITE access to Drive
  let dummy = DriveApp.createFile('dummy.txt', 'dummy content');
  dummy.setTrashed(true); // delete it immediately
  
  SpreadsheetApp.openById(SHEET_ID);
  MailApp.getRemainingDailyQuota();
  Logger.log("Permissions successfully granted! You MUST now deploy a NEW version of the Web App.");
}

/**
 * Serves the HTML pages.
 * If URL has ?page=officer, serves Dashboard.html. Otherwise serves Index.html (the form).
 */
function doGet(e) {
  let page = e.parameter.page;
  if (page === 'officer') {
    return HtmlService.createTemplateFromFile('Dashboard')
      .evaluate()
      .setTitle('Officer Dashboard - Maintenance')
      .addMetaTag('viewport', 'width=device-width, initial-scale=1');
  } else {
    return HtmlService.createTemplateFromFile('Index')
      .evaluate()
      .setTitle('C64 Facilities and Maintenance Report Form')
      .addMetaTag('viewport', 'width=device-width, initial-scale=1');
  }
}

/**
 * Processes the form submission from Index.html
 */
function processForm(data) {
  try {
    const sheet = SpreadsheetApp.openById(SHEET_ID).getSheets()[0]; // Assumes data goes to the first tab
    
    // 1. Process File Upload to Google Drive
    let fileUrl = '';
    let fileName = '';
    if (data.fileData && data.fileName) {
      const folder = DriveApp.getFolderById(DRIVE_FOLDER_ID);
      // Remove the data:image/jpeg;base64, prefix
      const base64Data = data.fileData.split(',')[1];
      const decodedData = Utilities.base64Decode(base64Data);
      const blob = Utilities.newBlob(decodedData, data.fileMimeType, data.fileName);
      const file = folder.createFile(blob);
      // Ensure the image can be viewed by anyone with the link
      // Using try-catch because some Google Workspace domains (like schools) block public sharing
      try {
        file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
      } catch (e) {
        try {
          // Fallback: Share with anyone in the school domain
          file.setSharing(DriveApp.Access.DOMAIN_WITH_LINK, DriveApp.Permission.VIEW);
        } catch (e2) {
          Logger.log("Domain restrictions prevented sharing update: " + e2);
        }
      }
      fileUrl = file.getUrl();
      fileName = file.getName();
    }
    
    const submissionId = 'REQ-' + new Date().getTime().toString().slice(-6);
    const submissionTime = new Date();
    const status = 'Action needed';
    
    // 2. Append to Google Sheet
    sheet.appendRow([
      submissionId,
      submissionTime,
      data.email,
      data.staffName,
      data.location,
      data.altLocation,
      data.type,
      data.priority,
      data.details,
      fileUrl,
      status
    ]);
    
    // 3. Prepare Email Content
    const imageText = fileName ? `[${fileName}](${fileUrl})` : 'None';
    
    const summaryHtml = `
      <p><b>Request Summary</b></p>
      <p><b>Location:</b> ${data.location}</p>
      <p><b>Category:</b> ${data.type}</p>
      <p><b>Priority:</b> ${data.priority}</p>
      <p><b>Description:</b> ${data.details}</p>
      <p><b>Image:</b> <a href="${fileUrl}">${fileName || 'None'}</a></p>
    `;
    
    // Get the script URL for the Review button
    const scriptUrl = ScriptApp.getService().getUrl();
    const officerLink = `${scriptUrl}?page=officer`;
    
    // 4. Send Email to Officer
    const officerSubject = `New Maintenance Request: ${data.type}`;
    const officerBodyHtml = `
      ${summaryHtml}
      <br>
      <a href="${officerLink}" style="background-color: #3b82f6; color: white; padding: 10px 20px; text-decoration: none; border-radius: 5px; display: inline-block;">Review for submission</a>
    `;
    
    MailApp.sendEmail({
      to: OFFICER_EMAIL,
      subject: officerSubject,
      body: 'Please enable HTML to view this email.', // Required field
      htmlBody: officerBodyHtml
    });
    
    // Trim emails to prevent 'Invalid argument' from trailing spaces
    const recipientEmail = (data.email || '').toString().trim();
    
    // 5. Send Email to Sender
    const senderSubject = `Received: Maintenance Request`;
    const senderBodyHtml = `
      <p>Dear ${data.staffName},</p>
      <p>We have successfully received your request and it has been forwarded to the Facilities Team for review.</p>
      ${summaryHtml}
      <p>Thank you for helping us maintain a safe and comfortable environment.</p>
      <br>
      <p>Best regards,<br>Facilities & Maintenance Team<br>Anglo Singapore International School</p>
    `;
    
    if (recipientEmail) {
      MailApp.sendEmail({
        to: recipientEmail,
        subject: senderSubject,
        body: 'Please enable HTML to view this email.', // Required field
        htmlBody: senderBodyHtml
      });
    }
    
    return { success: true, submissionId: submissionId };
  } catch (error) {
    Logger.log("Error in processForm: " + error.toString());
    return { success: false, error: error.toString() };
  }
}

/**
 * Retrieves submissions for the Officer Dashboard
 */
function getSubmissions() {
  const sheet = SpreadsheetApp.openById(SHEET_ID).getSheets()[0];
  const data = sheet.getDataRange().getDisplayValues();
  
  // Skip header (assuming row 1 is headers)
  const rows = data.slice(1);
  return rows.map((row, index) => {
    return {
      rowIndex: index + 2, // 1-based index, +1 for skipping header
      submissionId: row[0],
      time: row[1],
      email: row[2],
      staffName: row[3],
      location: row[4],
      altLocation: row[5],
      type: row[6],
      priority: row[7],
      details: row[8],
      fileUrl: row[9],
      status: row[10]
    };
  }).reverse(); // Show newest first
}

/**
 * Updates status from Dashboard and sends notification email
 */
function updateStatus(rowIndex, newStatus) {
  try {
    const sheet = SpreadsheetApp.openById(SHEET_ID).getSheets()[0];
    
    // Update status column (Column K = 11)
    sheet.getRange(rowIndex, 11).setValue(newStatus);
    
    // Fetch row data to send email
    const row = sheet.getRange(rowIndex, 1, 1, 11).getDisplayValues()[0];
    const email = row[2];
    const staffName = row[3];
    const location = row[4];
    const type = row[6];
    const priority = row[7];
    const details = row[8];
    const fileUrl = row[9];
    
    const summaryHtml = `
      <p><b>Request Summary</b></p>
      <p><b>Location:</b> ${location}</p>
      <p><b>Category:</b> ${type}</p>
      <p><b>Priority:</b> ${priority}</p>
      <p><b>Description:</b> ${details}</p>
      <p><b>Image:</b> ${fileUrl ? `<a href="${fileUrl}">View Image</a>` : 'None'}</p>
    `;
    
    let subject = `Update: Maintenance Request - ${newStatus}`;
    let intro = '';
    
    if (newStatus === 'Pending') {
      intro = `Your maintenance request is now being reviewed.`;
    } else if (newStatus === 'Ongoing') {
      intro = `Work has started on your maintenance request.`;
    } else if (newStatus === 'Closed') {
      intro = `Your maintenance request has been completed.`;
    } else {
      intro = `The status of your maintenance request has been updated to: ${newStatus}`;
    }
    
    const senderBodyHtml = `
      <p>Dear ${staffName},</p>
      <p>${intro}</p>
      ${summaryHtml}
      <p>Thank you for helping us maintain a safe and comfortable environment.</p>
      <br>
      <p>Best regards,<br>Facilities & Maintenance Team<br>Anglo Singapore International School</p>
    `;
    
    MailApp.sendEmail({
      to: email,
      cc: OFFICER_EMAIL,
      subject: subject,
      body: 'Please enable HTML to view this email.', // Required field
      htmlBody: senderBodyHtml
    });
    
    return { success: true };
  } catch (error) {
    Logger.log("Error in updateStatus: " + error.toString());
    return { success: false, error: error.toString() };
  }
}
