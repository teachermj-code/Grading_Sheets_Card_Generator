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