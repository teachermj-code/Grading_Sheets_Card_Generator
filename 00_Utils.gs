/**
 * ==========================================
 * 00_Utils.gs
 * PURPOSE: Shared helper functions and templating logic.
 * ==========================================
 */

/**
 * THE MAGIC TEMPLATING ENGINE
 * This takes an external HTML file (like our Universal Modal) 
 * and injects it directly into another HTML file.
 */
function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

/**
 * SHARED DATA MATCHER
 * Fetches the dynamic list of subjects from the Learner's Name sheet.
 */
function getSubjectList() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("Learner's Name");
  if (!sheet) return [];
  
  // Return a flat array, ignoring blank cells
  return sheet.getRange("G4:G18").getValues().flat().filter(s => s !== "");
}