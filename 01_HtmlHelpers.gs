/**
 * ==========================================
 * 01_HtmlHelpers.gs
 * PURPOSE: Universal templating functions for UI elements
 * ==========================================
 */

/**
 * Universal Injector: Allows HTML files to import other HTML/CSS files
 * Usage in HTML: <?!= include('FileName'); ?>
 */
function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}