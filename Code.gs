// ==========================================
// CONFIGURATION
// ==========================================
const SHEET_ID = '1E2z4VJ_iframYJyQgMN_0rKRo4lGIZHbyMZo3eyy4Q4'; 
const DRIVE_FOLDER_ID = '1cALybDaiQab1PC3tf-LjF1v9hGfCiFSO'; 
const OFFICER_EMAIL = 'raymond.b@anglosingapore.ac.th'; // The officer who gets the first notification

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
 * ?page=officer  → Dashboard.html (Task Board / Request Manager)
 * ?page=analytics → Analytics.html (Looker Studio embed & data export)
 * (default)       → Index.html (the public submission form)
 */
function doGet(e) {
  let page = e.parameter.page;
  if (page === 'officer') {
    return HtmlService.createTemplateFromFile('Dashboard')
      .evaluate()
      .setTitle('Task Board - Maintenance Requests')
      .addMetaTag('viewport', 'width=device-width, initial-scale=1');
  } else if (page === 'analytics') {
    return HtmlService.createTemplateFromFile('Analytics')
      .evaluate()
      .setTitle('Analytics & Reports - Maintenance')
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
    
    const lastRow = sheet.getLastRow();
    let nextNum = 1;
    
    if (lastRow > 1) {
      const lastId = sheet.getRange(lastRow, 1).getValue().toString();
      const match = lastId.match(/REQ-FMRF-(\d+)/);
      if (match) {
        nextNum = parseInt(match[1], 10) + 1;
      } else {
        nextNum = lastRow; // Fallback if previous ID doesn't match format
      }
    }
    
    const paddedNum = nextNum.toString().padStart(4, '0');
    const submissionId = `REQ-FMRF-${paddedNum}`;
    const submissionTime = Utilities.formatDate(new Date(), 'Asia/Bangkok', 'M/d/yyyy HH:mm:ss');
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
      status,
      '', // Remarks column (L)
      ''  // Assigned To column (M)
    ]);
    
    // 3. Prepare Email Content
    const imageText = fileName ? `[${fileName}](${fileUrl})` : 'None';
    
    const statusBanner = `
      <div style="background-color: #fef3c7; color: #92400e; padding: 10px; border-radius: 5px; font-weight: bold; text-align: center; margin-bottom: 15px;">
        Status: Action needed
      </div>
    `;

    const summaryHtml = `
      <table style="width: 100%; max-width: 600px; border-collapse: collapse; font-family: sans-serif; font-size: 14px; border: 1px solid #e5e7eb;">
        <tr style="background-color: #f9fafb;">
          <td style="padding: 12px; font-weight: bold; width: 30%; border: 1px solid #e5e7eb;">Issue Number</td>
          <td style="padding: 12px; border: 1px solid #e5e7eb;">${submissionId}</td>
        </tr>
        <tr style="background-color: #ffffff;">
          <td style="padding: 12px; font-weight: bold; border: 1px solid #e5e7eb;">Requested by</td>
          <td style="padding: 12px; border: 1px solid #e5e7eb;"><a href="mailto:${data.email}">${data.email}</a></td>
        </tr>
        <tr style="background-color: #f9fafb;">
          <td style="padding: 12px; font-weight: bold; border: 1px solid #e5e7eb;">Location</td>
          <td style="padding: 12px; border: 1px solid #e5e7eb;">${data.location}</td>
        </tr>
        <tr style="background-color: #ffffff;">
          <td style="padding: 12px; font-weight: bold; border: 1px solid #e5e7eb;">Issue</td>
          <td style="padding: 12px; border: 1px solid #e5e7eb;">${data.type}</td>
        </tr>
        <tr style="background-color: #f9fafb;">
          <td style="padding: 12px; font-weight: bold; border: 1px solid #e5e7eb;">Priority</td>
          <td style="padding: 12px; border: 1px solid #e5e7eb;">${data.priority}</td>
        </tr>
        <tr style="background-color: #ffffff;">
          <td style="padding: 12px; font-weight: bold; border: 1px solid #e5e7eb;">Issue Details</td>
          <td style="padding: 12px; border: 1px solid #e5e7eb;">${data.details}</td>
        </tr>
        <tr style="background-color: #f9fafb;">
          <td style="padding: 12px; font-weight: bold; border: 1px solid #e5e7eb;">Image</td>
          <td style="padding: 12px; border: 1px solid #e5e7eb;">${fileName ? `<a href="${fileUrl}">${fileName}</a>` : 'None'}</td>
        </tr>
      </table>
    `;
    
    // Get the script URL for the Review button
    const scriptUrl = ScriptApp.getService().getUrl();
    const officerLink = `${scriptUrl}?page=officer`;
    
    // 4. Send Email to Officer
    const officerSubject = `Maintenance Request [${submissionId}]`;
    const officerBodyHtml = `
      ${statusBanner}
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
    const senderSubject = `Maintenance Request [${submissionId}]`;
    const senderBodyHtml = `
      ${statusBanner}
      <p>Dear ${data.staffName},</p>
      <p>We have successfully received your request and it has been forwarded to the Facilities Team for review.</p>
      ${summaryHtml}
      <p>Thank you for helping us maintain a safe and comfortable environment.</p>
      <br>
      <p><b>Best regards,<br>Facilities & Maintenance Team<br>Anglo Singapore International School</b></p>
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
      status: row[10],
      remarks: row[11] || '',
      assignedTo: row[12] || ''
    };
  }).reverse(); // Show newest first
}

/**
 * Saves remarks silently to the sheet
 */
function saveRemarks(rowIndex, remarks) {
  try {
    const sheet = SpreadsheetApp.openById(SHEET_ID).getSheets()[0];
    // Server-side guard: reject saves on completed tickets
    const currentStatus = sheet.getRange(rowIndex, 11).getValue().toString().trim();
    if (currentStatus === 'Completed') {
      return { success: false, error: 'Cannot update remarks: ticket is already Completed.' };
    }
    sheet.getRange(rowIndex, 12).setValue(remarks);
    return { success: true };
  } catch (error) {
    Logger.log("Error in saveRemarks: " + error.toString());
    return { success: false, error: error.toString() };
  }
}

/**
 * Saves the assigned-to email to Column M (13) in the sheet
 */
function saveAssignment(rowIndex, assignedEmail) {
  try {
    const sheet = SpreadsheetApp.openById(SHEET_ID).getSheets()[0];
    // Server-side guard: reject saves when ticket is In Progress or Completed
    const currentStatus = sheet.getRange(rowIndex, 11).getValue().toString().trim();
    if (currentStatus === 'In Progress' || currentStatus === 'Completed') {
      return { success: false, error: 'Cannot change assignment: ticket is already ' + currentStatus + '.' };
    }
    sheet.getRange(rowIndex, 13).setValue(assignedEmail);
    return { success: true };
  } catch (error) {
    Logger.log("Error in saveAssignment: " + error.toString());
    return { success: false, error: error.toString() };
  }
}

/**
 * Updates status from Dashboard and sends notification email
 */
function updateStatus(rowIndex, newStatus, remarks) {
  try {
    const sheet = SpreadsheetApp.openById(SHEET_ID).getSheets()[0];
    
    // Server-side guard: prevent backward status changes
    const statusOrder = ['Action needed', 'Under Review', 'In Progress', 'Completed'];
    const currentStatus = sheet.getRange(rowIndex, 11).getValue().toString().trim();
    const currentIdx = statusOrder.indexOf(currentStatus);
    const newIdx = statusOrder.indexOf(newStatus);
    if (currentIdx !== -1 && newIdx !== -1 && newIdx <= currentIdx) {
      return { success: false, error: 'Cannot move status backward. Current: ' + currentStatus + ', Requested: ' + newStatus };
    }
    
    // Update status column (Column K = 11)
    sheet.getRange(rowIndex, 11).setValue(newStatus);
    
    // Update remarks if provided
    if (remarks !== undefined) {
      sheet.getRange(rowIndex, 12).setValue(remarks);
    }
    
    // Fetch row data to send email (now 13 columns to include Assigned To)
    const row = sheet.getRange(rowIndex, 1, 1, 13).getDisplayValues()[0];
    const submissionId = row[0];
    const email = row[2];
    const staffName = row[3];
    const location = row[4];
    const type = row[6];
    const priority = row[7];
    const details = row[8];
    const fileUrl = row[9];
    const finalRemarks = row[11];
    const assignedTo = (row[12] || '').trim();
    
    const summaryHtml = `
      <table style="width: 100%; max-width: 600px; border-collapse: collapse; font-family: sans-serif; font-size: 14px; border: 1px solid #e5e7eb; margin-top: 15px;">
        <tr style="background-color: #f9fafb;">
          <td style="padding: 12px; font-weight: bold; width: 30%; border: 1px solid #e5e7eb;">Issue Number</td>
          <td style="padding: 12px; border: 1px solid #e5e7eb;">${submissionId}</td>
        </tr>
        <tr style="background-color: #ffffff;">
          <td style="padding: 12px; font-weight: bold; border: 1px solid #e5e7eb;">Requested by</td>
          <td style="padding: 12px; border: 1px solid #e5e7eb;"><a href="mailto:${email}">${email}</a></td>
        </tr>
        <tr style="background-color: #f9fafb;">
          <td style="padding: 12px; font-weight: bold; border: 1px solid #e5e7eb;">Location</td>
          <td style="padding: 12px; border: 1px solid #e5e7eb;">${location}</td>
        </tr>
        <tr style="background-color: #ffffff;">
          <td style="padding: 12px; font-weight: bold; border: 1px solid #e5e7eb;">Issue</td>
          <td style="padding: 12px; border: 1px solid #e5e7eb;">${type}</td>
        </tr>
        <tr style="background-color: #f9fafb;">
          <td style="padding: 12px; font-weight: bold; border: 1px solid #e5e7eb;">Priority</td>
          <td style="padding: 12px; border: 1px solid #e5e7eb;">${priority}</td>
        </tr>
        <tr style="background-color: #ffffff;">
          <td style="padding: 12px; font-weight: bold; border: 1px solid #e5e7eb;">Issue Details</td>
          <td style="padding: 12px; border: 1px solid #e5e7eb;">${details}</td>
        </tr>
        <tr style="background-color: #f9fafb;">
          <td style="padding: 12px; font-weight: bold; border: 1px solid #e5e7eb;">Remarks</td>
          <td style="padding: 12px; border: 1px solid #e5e7eb;">${finalRemarks || 'None'}</td>
        </tr>
        <tr style="background-color: #ffffff;">
          <td style="padding: 12px; font-weight: bold; border: 1px solid #e5e7eb;">Assigned To</td>
          <td style="padding: 12px; border: 1px solid #e5e7eb;">${assignedTo || 'Not yet assigned'}</td>
        </tr>
        <tr style="background-color: #f9fafb;">
          <td style="padding: 12px; font-weight: bold; border: 1px solid #e5e7eb;">Image</td>
          <td style="padding: 12px; border: 1px solid #e5e7eb;">${fileUrl ? `<a href="${fileUrl}">View Image</a>` : 'None'}</td>
        </tr>
      </table>
    `;
    
    let subject = `Maintenance Request [${submissionId}]`;
    let intro = '';
    
    let statusColor = '#3730a3';
    let statusBg = '#e0e7ff';

    if (newStatus === 'Under Review') {
      intro = `Your maintenance request is now being reviewed.`;
      statusColor = '#991b1b'; statusBg = '#fee2e2'; // Red
    } else if (newStatus === 'In Progress') {
      intro = `Work has started on your maintenance request.`;
      statusColor = '#1e40af'; statusBg = '#dbeafe'; // Blue
    } else if (newStatus === 'Completed') {
      intro = `Your maintenance request has been completed.`;
      statusColor = '#166534'; statusBg = '#dcfce3'; // Green
    } else {
      intro = `The status of your maintenance request has been updated to: ${newStatus}`;
    }
    
    const statusBanner = `<h3 style="color: #374151; font-family: sans-serif; border-bottom: 1px solid #e5e7eb; padding-bottom: 10px;">Update: Maintenance Request - <span style="background-color: ${statusBg}; color: ${statusColor}; padding: 4px 10px; border-radius: 6px; font-weight: bold; display: inline-block;">${newStatus}</span></h3>`;
    
    const senderBodyHtml = `
      ${statusBanner}
      <p style="font-family: sans-serif;">Dear ${staffName},</p>
      <p style="font-family: sans-serif;">${intro}</p>
      ${summaryHtml}
      <p>Thank you for helping us maintain a safe and comfortable environment.</p>
      <br>
      <p><b>Best regards,<br>Facilities & Maintenance Team<br>Anglo Singapore International School</b></p>
    `;
    
    // Build CC list: always include the officer, plus the assigned person if set
    let ccList = OFFICER_EMAIL;
    if (assignedTo && assignedTo !== OFFICER_EMAIL) {
      ccList = OFFICER_EMAIL + ',' + assignedTo;
    }
    
    MailApp.sendEmail({
      to: email,
      cc: ccList,
      subject: subject,
      body: 'Please enable HTML to view this email.',
      htmlBody: senderBodyHtml
    });
    
    return { success: true };
  } catch (error) {
    Logger.log("Error in updateStatus: " + error.toString());
    return { success: false, error: error.toString() };
  }
}

/**
 * Retrieves aggregated analytics data for the Analytics page charts and export table.
 */
function getAnalyticsData() {
  const sheet = SpreadsheetApp.openById(SHEET_ID).getSheets()[0];
  const data = sheet.getDataRange().getDisplayValues();
  const rows = data.slice(1); // skip headers

  // Counts
  const statusCounts = {};
  const locationCounts = {};
  const priorityCounts = {};
  const dailyCounts = {};
  const weeklyCounts = {};
  const monthlyCounts = {};

  // Raw rows for export table
  const tableRows = [];

  rows.forEach(row => {
    const submissionId = row[0];
    const time = row[1];
    const email = row[2];
    const staffName = row[3];
    const location = row[4];
    const type = row[6];
    const priority = row[7];
    const details = row[8];
    const status = row[10];
    const remarks = row[11] || '';
    const assignedTo = row[12] || '';

    // Status counts
    statusCounts[status] = (statusCounts[status] || 0) + 1;

    // Location counts (take first part before " - " if too long)
    const locKey = location.length > 30 ? location.substring(0, 30) + '...' : location;
    locationCounts[locKey] = (locationCounts[locKey] || 0) + 1;

    // Priority counts
    priorityCounts[priority] = (priorityCounts[priority] || 0) + 1;

    // Date counts
    try {
      const dateObj = new Date(time);
      if (!isNaN(dateObj.getTime())) {
        const year = dateObj.getFullYear();
        const month = String(dateObj.getMonth() + 1).padStart(2, '0');
        const day = String(dateObj.getDate()).padStart(2, '0');
        
        // Daily
        const dayKey = `${year}-${month}-${day}`;
        dailyCounts[dayKey] = (dailyCounts[dayKey] || 0) + 1;
        
        // Monthly
        const monthKey = `${month}/${year}`;
        monthlyCounts[monthKey] = (monthlyCounts[monthKey] || 0) + 1;
        
        // Weekly (approximate ISO week)
        const d = new Date(Date.UTC(year, dateObj.getMonth(), dateObj.getDate()));
        const dayNum = d.getUTCDay() || 7;
        d.setUTCDate(d.getUTCDate() + 4 - dayNum);
        const yearStart = new Date(Date.UTC(d.getUTCFullYear(),0,1));
        const weekNo = Math.ceil((((d - yearStart) / 86400000) + 1)/7);
        const weekKey = `${d.getUTCFullYear()}-W${String(weekNo).padStart(2, '0')}`;
        weeklyCounts[weekKey] = (weeklyCounts[weekKey] || 0) + 1;
      }
    } catch (e) { /* skip bad dates */ }

    // Table data
    tableRows.push({
      submissionId: submissionId,
      time: time,
      staffName: staffName,
      email: email,
      location: location,
      type: type,
      priority: priority,
      details: details,
      status: status,
      remarks: remarks,
      assignedTo: assignedTo
    });
  });

  return {
    totalRequests: rows.length,
    statusCounts: statusCounts,
    locationCounts: locationCounts,
    priorityCounts: priorityCounts,
    dailyCounts: dailyCounts,
    weeklyCounts: weeklyCounts,
    monthlyCounts: monthlyCounts,
    tableRows: tableRows
  };
}
