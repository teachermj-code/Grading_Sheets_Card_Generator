/**
 * ==========================================
 * 40_ReportCard.gs
 * PURPOSE: Logic for Report Card Data Merging & PDF Generation
 * ==========================================
 */

const RC_CONFIG = {
  templateMap: {
    "1st Quarter": "1dTZ4VtUiid0gp5Oe5wQUIY_-TK9SyDrwIArGFvl9tvk", // Template A
    "2nd Quarter": "1dTZ4VtUiid0gp5Oe5wQUIY_-TK9SyDrwIArGFvl9tvk", // Template A
    "3rd Quarter": "1dTZ4VtUiid0gp5Oe5wQUIY_-TK9SyDrwIArGFvl9tvk", // Template A
    "4th Quarter": "15EtOUpv2bavC6CCfC1xvZYqpuj9kd4Bcl7r0b2FBa0k"  // Template B
  },
  outputFolderId: "1M5Oa5iXR19P_Sp2o4WrltdLc4sRfVzCs"
};

/**
 * Main Controller for the Sidebar call
 */

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
 * Optimized Controller: Reads exclusively from the RC_MASTER_DATA bridge.
 */
function processReportCards(gradingPeriod) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const masterSheet = ss.getSheetByName("RC_MASTER_DATA");
    
    if (!masterSheet) throw new Error("RC_MASTER_DATA sheet not found! Please run initializeMasterData first.");

    // 1. Get all data from the bridge sheet
    const data = masterSheet.getDataRange().getValues();
    if (data.length < 2) throw new Error("No student data found in RC_MASTER_DATA.");

    const headers = data[0]; // Row 1: The tags (Name, F1, 1, Jul1, etc.)
    
    // 2. Setup Folder & Template
    const parentFolder = DriveApp.getFolderById(RC_CONFIG.outputFolderId);
    const subFolderName = `Report Cards - ${gradingPeriod} - Generated ${new Date().toLocaleDateString()}`;
    const targetFolder = parentFolder.createFolder(subFolderName);
    const templateId = RC_CONFIG.templateMap[gradingPeriod];
    const templateFile = DriveApp.getFileById(templateId);

    // 3. Loop through every student in the MASTER sheet
    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      const studentName = row[0]; // Name is in Column A
      
      if (!studentName) continue;

// 4. Create the "Super Object" of all tags for this student
      let studentData = {};
      headers.forEach((header, index) => {
        if (header) {
          let val = row[index];
          
          // FIX: If the cell is empty or null, make it an empty string, NOT a 0
          if (val === "" || val === null || val === undefined) {
            studentData[header] = ""; 
          } else {
            studentData[header] = val;
          }
        }
      });

      // 5. Generate the PDF using the Single Source Data
      generateSinglePDF(studentData, templateFile, targetFolder, studentName, gradingPeriod);
    }

    return { 
      status: "success", 
      message: `Successfully processed ${data.length - 1} cards into: ${subFolderName}`, 
      folderUrl: targetFolder.getUrl() 
    };

  } catch (e) {
    console.error(e);
    return { status: "error", message: e.message };
  }
}

function generateSinglePDF(dataMap, templateFile, folder, studentName, q) {
  const copy = templateFile.makeCopy(folder);
  const doc = DocumentApp.openById(copy.getId());
  const body = doc.getBody();

// Loop through every tag in the student's data map
  for (let tag in dataMap) {
    let value = dataMap[tag];
    
    // 1. STRICT CHECK: Only replace with blank if the value is actually empty/null.
    // This allows the number 0 to be processed as a valid string.
    if (value === "" || value === null || value === undefined) {
      body.replaceText("<<" + tag + ">>", "");
      continue;
    }

    // 2. FORMAT ACTUAL NUMBERS (Including 0)
    if (typeof value === 'number') {
      if (tag.endsWith('G') || tag === 'GA') {
        // Grades get decimals (e.g., 85.00)
        value = value.toFixed(2); 
      } else {
        // Attendance, Age, and Conduct tags stay as whole numbers
        // This will turn 0 into "0" instead of blank
        value = Math.round(value).toString(); 
      }
    }

    // 3. Final Replacement
    body.replaceText("<<" + tag + ">>", String(value));
  }

  doc.saveAndClose();
  const pdf = DriveApp.getFileById(copy.getId()).getAs("application/pdf");
  folder.createFile(pdf).setName(`${studentName} - ${q}.pdf`);
  DriveApp.getFileById(copy.getId()).setTrashed(true);
}

/**
 * Injects formulas into RC_MASTER_DATA without destroying headers or Columns DC:DE.
 */
function initializeMasterData() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let master = ss.getSheetByName("RC_MASTER_DATA");
  
  if (!master) {
    master = ss.insertSheet("RC_MASTER_DATA");
    SpreadsheetApp.getUi().alert("New RC_MASTER_DATA sheet created. Please paste your headers in Row 1.");
    return;
  }

  const lastRow = master.getMaxRows();
  const lastCol = master.getMaxColumns();

  if (lastRow > 1) {
    // 1. Clear Columns A (1) to DB (106)
    master.getRange(2, 1, lastRow - 1, 106).clear();
    
    // 2. Clear from Column DF (110) to the end of the sheet
    // This leaves DC (107), DD (108), and DE (109) untouched
    if (lastCol >= 110) {
      master.getRange(2, 110, lastRow - 1, lastCol - 109).clear();
    }
  }
  
  // RE-INJECT FORMULAS
  master.getRange("A2").setFormula("=ARRAYFORMULA('Report Card Setup'!A2:J)");
  master.getRange("K2").setFormula("=ARRAYFORMULA(RC_Attendance!B2:AH)");

  const subjects = [
    { range: "AU2", sheet: "SUMMARY_FILIPINO" },
    { range: "AZ2", sheet: "SUMMARY_ENGLISH" },
    { range: "BE2", sheet: "SUMMARY_MATH" },
    { range: "BJ2", sheet: "SUMMARY_SCIENCE" },
    { range: "BO2", sheet: "SUMMARY_ARALPAN" },
    { range: "BT2", sheet: "SUMMARY_GMRC" },
    { range: "BY2", sheet: "SUMMARY_HELE" },
    { range: "CD2", sheet: "SUMMARY_MAPEH" },
    { range: "CI2", sheet: "SUMMARY_MUSIC" },
    { range: "CN2", sheet: "SUMMARY_ARTS" },
    { range: "CS2", sheet: "SUMMARY_P.E." },
    { range: "CX2", sheet: "SUMMARY_HEALTH" }
  ];

  subjects.forEach(sub => {
    master.getRange(sub.range).setFormula(`=IFERROR(ARRAYFORMULA('${sub.sheet}'!C3:G), "")`);
  });

  // Homeroom Guidance starting at DF2
  master.getRange("DF2").setFormula("=IFERROR(ARRAYFORMULA('1Q HOMEROOM GUIDANCE LETTER GRADE'!F6:T), \"\")");
  master.getRange("DU2").setFormula("=IFERROR(ARRAYFORMULA('2Q HOMEROOM GUIDANCE LETTER GRADE'!F6:T), \"\")");
  master.getRange("EJ2").setFormula("=IFERROR(ARRAYFORMULA('3Q HOMEROOM GUIDANCE LETTER GRADE'!F6:T), \"\")");
  master.getRange("EY2").setFormula("=IFERROR(ARRAYFORMULA('4Q HOMEROOM GUIDANCE LETTER GRADE'!F6:T), \"\")");

  master.hideSheet();
  SpreadsheetApp.getUi().alert("Master Bridge Restored. Headers and Columns DC:DE preserved.");
}

/**
 * SECURITY: Automatically re-hides the master sheet if anyone unhides it.
 * This runs every time the spreadsheet is opened.
 */
function enforceRCSheetPrivacy() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const master = ss.getSheetByName("RC_MASTER_DATA");
  if (master && !master.isSheetHidden()) {
    master.hideSheet();
  }
}

/**
 * Checks Google Drive to see if a folder for this Quarter already exists.
 */
function checkIfCardsExist(gradingPeriod) {
  try {
    const parentFolder = DriveApp.getFolderById(RC_CONFIG.outputFolderId);
    // Searches for any folder containing the quarter name
    const folders = parentFolder.searchFolders(`title contains 'Report Cards - ${gradingPeriod}'`);
    
    return { exists: folders.hasNext() };
  } catch (e) {
    console.error(e);
    return { exists: false }; // Defaults to false if Drive is empty or errors out
  }
}