/**
 * ==========================================
 * 40_Support.gs
 * PURPOSE: Help Desk, File Uploads, and Email Ticketing
 * ==========================================
 */

function showSupportUI() {
  const template = HtmlService.createTemplateFromFile('40_SupportForm');
  const html = template.evaluate()
      .setWidth(450)
      .setHeight(550)
      .setTitle(' ');
  SpreadsheetApp.getUi().showModalDialog(html, ' ');
}

function processSupportTicket(ticketData) {
  try {
    let fileUrl = "No file attached.";

    // 1. Handle File Upload (If provided)
    if (ticketData.fileData && ticketData.fileName) {
      // Create or find the designated Support Folder in Teacher MJ's Drive
      const folderName = "Teacher MJ Support Uploads";
      const folders = DriveApp.getFoldersByName(folderName);
      const folder = folders.hasNext() ? folders.next() : DriveApp.createFolder(folderName);

      // Decode Base64 stream from the frontend
      const contentType = ticketData.fileData.substring(5, ticketData.fileData.indexOf(';'));
      const base64Data = ticketData.fileData.substring(ticketData.fileData.indexOf('base64,') + 7);
      const blob = Utilities.newBlob(Utilities.base64Decode(base64Data), contentType, ticketData.fileName);
      
      const file = folder.createFile(blob);
      fileUrl = file.getUrl();
    }

    // 2. Format and Send the Email
    const supportEmail = "teachermjcbb@gmail.com";
    const subject = "Support Request: Class Record Manager";
    const body = `
New Support Request Received!

ISSUE DESCRIPTION:
${ticketData.description}

ATTACHMENT LINK (Google Drive):
${fileUrl}

--
Sent dynamically from the Teacher MJ Tools System
    `;

    MailApp.sendEmail(supportEmail, subject, body);
    return "Ticket submitted successfully!";
    
  } catch (e) {
    throw new Error("Failed to process ticket: " + e.message);
  }
}