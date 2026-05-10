/**
 * ==========================================
 * 00_Core.gs
 * PURPOSE: System Initialization & Global Constants
 * ==========================================
 */

function onOpen() {
  // 1. Build the Top Navigation Menu
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

  // 2. Automatically launch the Sidebar Dashboard on open
  try {
    showSidebarUI();
  } catch (e) {
    // Failsafe: In rare cases where Google delays permissions on load, 
    // it will fail silently and wait for the user to click the menu manually.
    console.error("Auto-launch failed. User must launch via menu.", e);
  }
}

// Global System Constants (Easy to update if school years change)
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

function showReportCardUI() {
  const template = HtmlService.createTemplateFromFile('30_ClassRecordForm'); // Or create a specific RC form
  // For now, we launch the Sidebar since that is your command center
  showSidebarUI(); 
  SpreadsheetApp.getUi().alert("Please use the Premium Sidebar to generate Report Cards.");
}

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


