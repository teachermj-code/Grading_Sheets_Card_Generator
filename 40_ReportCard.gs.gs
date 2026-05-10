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

function processReportCards(gradingPeriod) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const qPrefix = gradingPeriod === "1st Quarter" ? "1Q" : gradingPeriod === "2nd Quarter" ? "2Q" : gradingPeriod === "3rd Quarter" ? "3Q" : "4Q";
    
    // 1. Get Data Sources
    const setupSheet = ss.getSheetByName("Report Card Setup");
    const attSheet = ss.getSheetByName("RC_Attendance");
    const consolSheet = ss.getSheetByName(qPrefix + " CONSOL GRADE");

    if (!consolSheet) throw new Error(`Please generate the ${qPrefix} Consol Grade first!`);
    
    const setupData = setupSheet.getDataRange().getValues();
    const attData = attSheet.getDataRange().getValues();
    const consolData = consolSheet.getDataRange().getValues();
    
    const headersSetup = setupData[0];
    const headersAtt = attData[0];
    const headersConsol = consolData[0];

    // 2. Setup Folder
    const parentFolder = DriveApp.getFolderById(RC_CONFIG.outputFolderId);
    const subFolderName = `Report Cards - ${gradingPeriod} (${new Date().toLocaleDateString()})`;
    const targetFolder = parentFolder.createFolder(subFolderName);
    const templateFile = DriveApp.getFileById(RC_CONFIG.templateMap[gradingPeriod]);

    // 3. Loop through students (starting from row 2)
    for (let i = 1; i < setupData.length; i++) {
      let studentData = {};
      const studentName = setupData[i][0];
      if (!studentName) continue;

      // Map Setup Data
      headersSetup.forEach((h, idx) => studentData[h] = setupData[i][idx]);

      // Map Attendance Data (Match by Name)
      const attRow = attData.find(row => row[0] === studentName);
      if (attRow) {
        headersAtt.forEach((h, idx) => studentData[h] = attRow[idx]);
      }

      // Map Consol Grades (Match by Name)
      const consolRow = consolData.find(row => row[1] === studentName); // Name is Col B (index 1)
      if (consolRow) {
        headersConsol.forEach((h, idx) => {
          if (h && h !== "NAME") studentData[h] = consolRow[idx];
        });
      }

      // 4. Create the Doc and Replace Tags
      generateSinglePDF(studentData, templateFile, targetFolder, studentName, gradingPeriod);
    }

    return { status: "success", message: `Generated in: ${subFolderName}`, folderUrl: targetFolder.getUrl() };
  } catch (e) {
    return { status: "error", message: e.message };
  }
}

/**
 * Handles the actual Tag replacement in the Doc
 */
function generateSinglePDF(data, templateFile, folder, studentName, q) {
  const copy = templateFile.makeCopy(folder);
  const doc = DocumentApp.openById(copy.getId());
  const body = doc.getBody();

  // Replace all keys found in the data object
  for (let key in data) {
    const placeholder = `<<${key}>>`;
    let val = data[key] || "";
    // Format numbers to avoid decimals in grades like 90.00000001
    if (typeof val === 'number') val = Math.round(val);
    body.replaceText(placeholder, String(val));
  }

  doc.saveAndClose();
  const pdf = DriveApp.getFileById(copy.getId()).getAs("application/pdf");
  folder.createFile(pdf).setName(`${studentName} - ${q}.pdf`);
  DriveApp.getFileById(copy.getId()).setTrashed(true);
}