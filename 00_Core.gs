/**
 * ==========================================
 * 00_Core.gs
 * PURPOSE: System Initialization & Global Constants
 * ==========================================
 */

function onOpen() {
// 1. Run the security protocol to hide system templates
  secureSystemTemplates();

  // 2. Run the enforcement protocol to show and arrange core sheets
  enforceCoreSheets();

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

/**
 * AGGRESSIVE SECURITY WATCHER
 * Triggers every time a user clicks a cell or changes sheets.
 */
function onSelectionChange(e) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const activeSheet = ss.getActiveSheet();
  const activeName = activeSheet.getName();
  
  // List of sheets that must NEVER be seen or clicked by the user
  const forbiddenSheets = [
    "W_Exam", 
    "WO_Exam", 
    "CONSOL GRADE", 
    "HOMEROOM GUIDANCE", 
    "HOMEROOM GUIDANCE LETTER GRADE",
    "RC_MASTER_DATA"
  ];
  
  // If the user is currently looking at a forbidden sheet
  if (forbiddenSheets.includes(activeName)) {
    
    // 1. Instantly hide the forbidden sheet
    activeSheet.hideSheet();
    
    // 2. "Boot" the user back to a safe sheet
    try {
      const safeSheet = ss.getSheetByName("Report Card Setup");
      if (safeSheet) {
        safeSheet.activate();
      }
    } catch (err) {
      // Failsafe if Report Card Setup is missing
    }
    
    // 3. Fire the Premium Modal Overlay instead of a toast
    showSecurityAlertUI();
  }
}

/**
 * Launches the Universal Modal Warning for Access Denied events
 */
function showSecurityAlertUI() {
  const html = HtmlService.createTemplateFromFile('50_SecurityAlert')
      .evaluate()
      .setWidth(450)
      .setHeight(350);
      
  // Pass a blank space " " to hide the default Google title bar for a cleaner look
  SpreadsheetApp.getUi().showModalDialog(html, " "); 
}

/**
 * Enforces the visibility and exact order of the core user sheets.
 */
function enforceCoreSheets() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  // The exact order you want them arranged left-to-right
  const coreSheets = ["Report Card Setup", "RC_Attendance", "Learner's Name"];
  
  coreSheets.forEach((sheetName, index) => {
    const sheet = ss.getSheetByName(sheetName);
    if (sheet) {
      // 1. Force the sheet to be visible
      if (sheet.isSheetHidden()) {
        sheet.showSheet();
      }
      
      // 2. Move the sheet to its strict position
      // moveActiveSheet uses 1-based indexing, so index 0 becomes position 1
      ss.setActiveSheet(sheet);
      ss.moveActiveSheet(index + 1);
    }
  });
  
  // 3. Always land the user safely on the Setup sheet when finished
  const landingSheet = ss.getSheetByName(coreSheets[0]);
  if (landingSheet) {
    landingSheet.activate();
  }
}