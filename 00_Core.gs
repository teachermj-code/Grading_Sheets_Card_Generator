/**
 * ==========================================
 * 00_Core.gs
 * PURPOSE: System Initialization & Global Constants
 * ==========================================
 */

function onOpen() {
  // 1. Run the security protocol immediately upon opening
  secureSystemTemplates();

  // 2. Build the Top Navigation Menu
  SpreadsheetApp.getUi()
      .createMenu('🗂️ Teacher MJ Tools')
      .addItem('🗃️ Open Premium Dashboard', 'showSidebarUI')
      .addSeparator()
      .addItem('📓 Class Record Manager', 'showClassRecordUI')
      .addItem('🏆 Generate Academic Awards', 'showAwardsUI')
      .addItem('📄 Generate Report Card', 'showReportCardUI')
      .addSeparator()
      .addItem('⚙️ Navigation Control Center', 'showNavigationUI')
      .addSeparator()
      .addItem('💬 Contact Support', 'showSupportUI')
      .addToUi();

  // 3. Automatically launch the Sidebar Dashboard on open
  try {
    showSidebarUI();
  } catch (e) {
    console.error("Auto-launch failed. User must launch via menu.", e);
  }
}

// Global System Constants
const SYSTEM_CONFIG = {
  copyright: '© ' + new Date().getFullYear() + ' CLASSROOM TOOLS BY TEACHER MJ',
  fontPrimary: 'Tahoma'
};

/**
 * Opens the Premium Custom Sidebar Dashboard
 */
function showSidebarUI() {
  const template = HtmlService.createTemplateFromFile('05_Sidebar');
  const html = template.evaluate().setTitle('❖ Teacher MJ Tools');
  SpreadsheetApp.getUi().showSidebar(html);
}

/**
 * Opens the Report Card Generation Modal
 */
function showReportCardUI() {
  const html = HtmlService.createTemplateFromFile('40_ReportCardForm')
      .evaluate()
      .setWidth(400)
      .setHeight(450);
  SpreadsheetApp.getUi().showModalDialog(html, " "); // Blank title for premium look
}

/**
 * Hides and protects core system templates, making sure RC_MASTER_DATA stays invisible.
 */
function secureSystemTemplates() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const templates = [
    "W_Exam", 
    "WO_Exam", 
    "CONSOL GRADE", 
    "HOMEROOM GUIDANCE", 
    "HOMEROOM GUIDANCE LETTER GRADE",
    "RC_MASTER_DATA"
  ];
  
  templates.forEach(name => {
    const sheet = ss.getSheetByName(name);
    if (sheet) {
      // 1. Hide the sheet if it isn't already hidden
      if (!sheet.isSheetHidden()) {
        sheet.hideSheet();
      }
      
      // 2. Check if protection already exists, if not, create it
      let protection = sheet.getProtections(SpreadsheetApp.ProtectionType.SHEET)[0];
      if (!protection) {
        protection = sheet.protect().setDescription('Core System Protection');
      }
      
      // 3. Remove all other editors (Locks it to the owner only)
      const me = Session.getEffectiveUser();
      protection.addEditor(me); // Ensure you don't lock yourself out
      protection.removeEditors(protection.getEditors());
      
      if (protection.canDomainEdit()) {
        protection.setDomainEdit(false);
      }
    }
  });
  
  return "Security Protocol Active.";
}