/**
 * ==========================================
 * 00_Core.gs
 * PURPOSE: System Initialization & Global Constants
 * ==========================================
 */

function onOpen() {
  SpreadsheetApp.getUi()
      .createMenu('Teacher MJ Tools')
      .addItem('Class Record Manager', 'showClassRecordUI')
      .addItem('Generate Academic Awards', 'showAwardsUI')
      .addSeparator()
      .addItem('Navigation Control Center', 'showNavigationUI')
      .addSeparator()
      .addItem('Contact Support', 'showSupportUI')
      .addToUi();
}

// Global System Constants (Easy to update if school years change)
const SYSTEM_CONFIG = {
  copyright: '© ' + new Date().getFullYear() + ' CLASSROOM TOOLS BY TEACHER MJ',
  fontPrimary: 'Tahoma'
};