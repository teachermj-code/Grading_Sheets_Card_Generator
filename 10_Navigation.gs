/**
 * ==========================================
 * 10_Navigation.gs
 * PURPOSE: Workspace Control, View Modes, & Security
 * ==========================================
 */

/**
 * Opens the Navigation Control Center UI
 * Uses 'createTemplateFromFile' to allow injecting the Universal Modal
 */
function showNavigationUI() {
  const template = HtmlService.createTemplateFromFile('10_NavigationForm');
const html = template.evaluate()
      .setWidth(400) // Premium compact width
      .setHeight(650) // Standard height
      .setTitle(' ');
  SpreadsheetApp.getUi().showModalDialog(html, ' ');
}

/**
 * Checks if the system has already been initialized
 */
function checkInitStatus() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const exists = ss.getSheetByName("FINAL CONSOLIDATED") !== null;
  return { exists: exists };
}

function toggleViewMode(mode) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheets = ss.getSheets();
  
  // 1. Pre-check if target quarter exists
  let targetSheet = null;
  if (["1Q", "2Q", "3Q", "4Q"].includes(mode)) {
    targetSheet = ss.getSheetByName(mode + " CONSOL GRADE");
    if (!targetSheet) throw new Error(`The records for ${mode} have not been generated yet. Please generate them in the Class Record Manager first.`);
  } else if (mode === "SUMMARY") {
    targetSheet = ss.getSheetByName("SUMMARY_FILIPINO"); 
  } else if (mode === "GWA") {
    targetSheet = ss.getSheetByName("FINAL CONSOLIDATED");
  }

  // 2. SHOW targets first (to prevent hiding the active sheet error)
  sheets.forEach(sheet => {
    const name = sheet.getName();
    if (name === "Learner's Name") sheet.showSheet(); // ALWAYS VISIBLE
    else if (["1Q", "2Q", "3Q", "4Q"].includes(mode) && name.startsWith(mode + " ")) sheet.showSheet();
    else if (mode === "SUMMARY" && name.startsWith("SUMMARY_")) sheet.showSheet();
    else if (mode === "GWA" && name === "FINAL CONSOLIDATED") sheet.showSheet();
  });

  // 3. Activate a target sheet
  if (targetSheet) {
    targetSheet.showSheet();
    ss.setActiveSheet(targetSheet);
  } else {
    ss.setActiveSheet(ss.getSheetByName("Learner's Name"));
  }

  // 4. HIDE everything else
  sheets.forEach(sheet => {
    const name = sheet.getName();
    if (name === "Learner's Name") return; // Skip hiding
    
    let shouldKeep = false;
    if (["1Q", "2Q", "3Q", "4Q"].includes(mode) && name.startsWith(mode + " ")) shouldKeep = true;
    else if (mode === "SUMMARY" && name.startsWith("SUMMARY_")) shouldKeep = true;
    else if (mode === "GWA" && name === "FINAL CONSOLIDATED") shouldKeep = true;

    if (!shouldKeep) {
      try { sheet.setSheetVisibility(SpreadsheetApp.SheetVisibility.VERY_HIDDEN); } catch(e) { sheet.hideSheet(); }
    }
  });
  return "Success";
}

/**
 * Initializes all Summary and GWA sheets with design and formulas.
 */
function initializeSummaries() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const learnerSheet = ss.getSheetByName("Learner's Name");
  if (!learnerSheet) throw new Error("Please ensure 'Learner's Name' sheet exists.");

  const studentNames = learnerSheet.getRange("B4:B" + learnerSheet.getLastRow()).getValues().filter(n => n[0] !== "");
  
// Standardize Subject Names (Aliases -> Unified Name)
  const normalizeSub = (name) => {
    if (!name) return "";
    
    // Strip all spaces and periods
    const raw = name.toString().toUpperCase().replace(/\s/g, "").replace(/\./g, ""); 
    
    // The Dictionary: Catching all variations and returning YOUR official names
    if (["AP", "ARPAN", "ARALPAN", "ARALINGPANLIPUNAN"].includes(raw)) return "ARALPAN";
    if (["PE", "PHYSICALEDUCATION"].includes(raw)) return "P.E.";
    if (["MATH", "MATHS", "MATHEMATICS"].includes(raw)) return "MATH";
    if (["HG", "HOMEROOM", "HOMEROOMGUIDANCE"].includes(raw)) return "HOMEROOM GUIDANCE";
    
    return name.toString().toUpperCase().trim();
  };

  // Normalize subjects and forcefully append MAPEH if missing from raw data
  let uniqueSubjects = [...new Set(getSubjectList().map(s => normalizeSub(s)))];
  if (!uniqueSubjects.includes("MAPEH")) uniqueSubjects.push("MAPEH");

  // STRICT MASTER ORDER: Forces the Summary tabs to generate in this exact sequence
  const masterOrder = [
    "FILIPINO", "ENGLISH", "MATH", "SCIENCE", "ARALPAN", 
    "GMRC", "HELE", "MAPEH", "MUSIC", "ARTS", "P.E.", "HEALTH", "HOMEROOM GUIDANCE"
  ];

  const sortedSubjects = uniqueSubjects.map(s => [s]).sort((a, b) => {
    let indexA = masterOrder.indexOf(a[0]);
    let indexB = masterOrder.indexOf(b[0]);

    // If a subject isn't found in the master list, push it to the very end
    if (indexA === -1) indexA = 999;
    if (indexB === -1) indexB = 999;

    return indexA - indexB;
  });

  // 1. CREATE SUBJECT SUMMARIES
  sortedSubjects.forEach(sub => {
    const subName = sub[0].toString().toUpperCase();
    const sheetName = "SUMMARY_" + subName;
    let sheet = ss.getSheetByName(sheetName) || ss.insertSheet(sheetName);
    
    sheet.clear();
    try { sheet.setHideGridlines(true); } catch (err) { console.warn("Gridline skipped"); }
    
    const lastRow = studentNames.length + 2;
    sheet.setRowHeights(1, lastRow, 25); // Mandated Row Height 25

    sheet.setColumnWidth(1, 45);  // Col A
    sheet.setColumnWidth(2, 315); // Col B
    [3, 4, 5, 6, 7].forEach(col => sheet.setColumnWidth(col, 75)); // Col C-G

    // Branding: Tahoma 12pt Base
    sheet.getRange(1, 1, lastRow, 7).setFontFamily(SYSTEM_CONFIG.fontPrimary).setVerticalAlignment("middle");

    sheet.getRange("A:A").setHorizontalAlignment("center").setFontWeight("bold").setFontSize(10);
    sheet.getRange("B:B").setHorizontalAlignment("left").setFontSize(12);
    sheet.getRange("C:F").setHorizontalAlignment("center").setFontSize(12);
    sheet.getRange("G:G").setHorizontalAlignment("center").setFontWeight("bold").setFontSize(12);

    sheet.getRange("C1:F1").merge().setValue("QUARTER").setFontWeight("bold").setFontSize(12).setHorizontalAlignment("center");

    const labelRange = sheet.getRange("A2:G2");
    labelRange.setValues([["#", "NAME OF STUDENT", "1ST", "2ND", "3RD", "4TH", "AVG"]])
      .setBackground("#FF00FF").setFontColor("#FFFFFF").setFontWeight("bold").setHorizontalAlignment("center");

    const nameData = studentNames.map((n, i) => [i + 1, n[0]]);
    sheet.getRange(3, 1, nameData.length, 2).setValues(nameData);
    
const quarters = ["1Q", "2Q", "3Q", "4Q"];
    
    // Create a robust wildcard search key based on the normalized subject name
    let searchKey = "*" + subName + "*";
    if (subName === "MATH") searchKey = "*MATH*";
    else if (subName === "ARALPAN") searchKey = "*ARAL*";
    else if (subName === "P.E." || subName === "PE") searchKey = "*P*E*";
    else if (subName === "HOMEROOM GUIDANCE") searchKey = "*HOMEROOM*";

    for (let r = 0; r < studentNames.length; r++) {
      const row = r + 3;
      
      // Use INDIRECT to automatically connect to Quarter sheets the moment they are generated
      quarters.forEach((q, qIdx) => {
        const sheetName = `"${q} CONSOL GRADE"`;
        const formula = `=IFERROR(INDEX(INDIRECT("'" & ${sheetName} & "'!$C$3:$Z"), MATCH($B${row}, INDIRECT("'" & ${sheetName} & "'!$B$3:$B"), 0), MATCH("${searchKey}", INDIRECT("'" & ${sheetName} & "'!$C$2:$Z$2"), 0)), "")`;
        sheet.getRange(row, 3 + qIdx).setFormula(formula);
      });
      
      // Average Column Formula
      sheet.getRange(row, 7).setFormula(`=IF(COUNT(C${row}:F${row})>0, ROUND(AVERAGE(C${row}:F${row}), 2), "")`);
    }

    // Split Number Formatting: Whole numbers for Quarters (C:F), 2 Decimals for Average (G)
    sheet.getRange(3, 3, lastRow - 2, 4).setNumberFormat("0"); 
    sheet.getRange(3, 7, lastRow - 2, 1).setNumberFormat("0.00"); 
    
    sheet.getRange(1, 1, lastRow, 7).setBorder(true, true, true, true, true, true, "black", SpreadsheetApp.BorderStyle.SOLID);
    
    // Protection & Hide
    const protection = sheet.protect().setDescription('Automated Summary Protection');
    protection.removeEditors(protection.getEditors());
    if (protection.canDomainEdit()) protection.setDomainEdit(false);
    try { sheet.setSheetVisibility(SpreadsheetApp.SheetVisibility.VERY_HIDDEN); } catch (e) { sheet.hideSheet(); }
  });

  // 2. CREATE FINAL CONSOLIDATED (Construction Logic)
  let finalSheet = ss.getSheetByName("FINAL CONSOLIDATED") || ss.insertSheet("FINAL CONSOLIDATED");
  try { finalSheet.setHideGridlines(true); } catch (err) {}
    
  finalSheet.setRowHeights(1, studentNames.length + 2, 25);
  finalSheet.setColumnWidth(1, 45); 
  finalSheet.setColumnWidth(2, 315);
  for (let c = 3; c <= 10; c++) { finalSheet.setColumnWidth(c, 90); }
  finalSheet.setColumnWidth(11, 75);
  finalSheet.setColumnWidth(12, 195);

finalSheet.getRange("A1:L1").merge().setFontWeight("bold").setFontSize(14).setHorizontalAlignment("center").setVerticalAlignment("middle");
  finalSheet.getRange("A1").setValue("FINAL CONSOLIDATED GRADES");

  // Dynamically map to the unified Standard Subject Names
  const subMap = ["FILIPINO", "ENGLISH", "MATH", "SCIENCE", "ARALING PANLIPUNAN", "GMRC", "HELE", "MAPEH"];
  const gwaHeaders = [["#", "NAME OF STUDENT", ...subMap, "GWA", "STATUS"]];
  
  finalSheet.getRange("A2:L2").setValues(gwaHeaders)
    .setBackground("#FF00FF").setFontColor("#FFFFFF").setFontWeight("bold")
    .setHorizontalAlignment("center").setVerticalAlignment("middle").setFontSize(12).setFontFamily(SYSTEM_CONFIG.fontPrimary);

  const finalNames = studentNames.map((n, i) => [i + 1, n[0]]);
  finalSheet.getRange(3, 1, finalNames.length, 2).setValues(finalNames);

  subMap.forEach((s, i) => {
    finalSheet.getRange(3, 3 + i, finalNames.length, 1).setFormula(`=IFERROR(VLOOKUP($B3, 'SUMMARY_${s}'!$B$3:$G, 6, FALSE), "")`);
  });

  finalSheet.getRange(3, 11, finalNames.length, 1).setFormula(`=IF(COUNT(C3:J3)>0, ROUND(AVERAGE(C3:J3), 2), "")`);
  finalSheet.getRange(3, 12, finalNames.length, 1).setFormula(`=IFS(K3="","",K3>=98,"With Highest Honors",K3>=95,"With High Honors",K3>=90,"With Honors",K3<90,"")`);

  // Strictly enforce formatting rules
  const dataRangeCount = finalNames.length;
  finalSheet.getRange(3, 1, dataRangeCount, 12).setVerticalAlignment("middle");
  finalSheet.getRange(3, 1, dataRangeCount, 1).setHorizontalAlignment("center"); // Number ID
  finalSheet.getRange(3, 2, dataRangeCount, 1).setHorizontalAlignment("left"); // Student Name
  finalSheet.getRange(3, 3, dataRangeCount, 10).setHorizontalAlignment("center"); // All Grades & Status
  finalSheet.getRange(3, 3, dataRangeCount, 9).setNumberFormat("0.00"); // 2 Decimal Places for Subject Grades & GWA

  finalSheet.getRange(1, 1, studentNames.length + 2, 12).setBorder(true, true, true, true, true, true, "black", SpreadsheetApp.BorderStyle.SOLID);
  try { finalSheet.setSheetVisibility(SpreadsheetApp.SheetVisibility.VERY_HIDDEN); } catch (e) { finalSheet.hideSheet(); }
}

/**
 * LIGHTWEIGHT UPDATE: Called by ClassRecord module to keep GWA Title synced.
 */
function updateConsolidatedHeader(gradeSection) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const finalSheet = ss.getSheetByName("FINAL CONSOLIDATED");
  if (!finalSheet) return;

  const gsName = gradeSection.toUpperCase();
  const fullGS = gsName.startsWith("GRADE ") ? gsName : "GRADE " + gsName;
  finalSheet.getRange("A1").setValue(fullGS + "'S FINAL CONSOLIDATED GRADES");
  finalSheet.getRange("A1").setValue("GRADE & SECTION: " + gsName);
}

/**
 * SNAP-BACK SECURITY: Monitors for unauthorized unhiding
 */
function onEdit(e) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const activeSheet = ss.getActiveSheet();
  const name = activeSheet.getName();
  const forbidden = ["TEMPLATE", "Learner's Name", "SUMMARY_", "FINAL CONSOLIDATED"];
  
  if (forbidden.some(f => name.includes(f))) {
    try { activeSheet.setSheetVisibility(SpreadsheetApp.SheetVisibility.VERY_HIDDEN); } catch(e) { activeSheet.hideSheet(); }
    SpreadsheetApp.getUi().alert("⚠️ SYSTEM SECURITY\n\nThis sheet is protected.");
  }
}