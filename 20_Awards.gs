/**
 * ==========================================
 * 20_Awards.gs
 * PURPOSE: Extract honors (>=98) and build PDF certificates.
 * ==========================================
 */

/**
 * Opens the Awards Generator UI
 */
function showAwardsUI() {
  const template = HtmlService.createTemplateFromFile('20_AwardsForm');
  const html = template.evaluate()
      .setWidth(500)
      .setHeight(400)
      .setTitle(' ');
  SpreadsheetApp.getUi().showModalDialog(html, ' ');
}

/**
 * Checks if the PDF already exists in the folder
 */
function checkExistingReport(quarter) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const prefixes = { "First Quarter": "1Q", "Second Quarter": "2Q", "Third Quarter": "3Q", "Fourth Quarter": "4Q" };
  const gradingMap = { "First Quarter": "First Grading", "Second Quarter": "Second Grading", "Third Quarter": "Third Grading", "Fourth Quarter": "Fourth Grading" };
  
  const sheetName = prefixes[quarter] + " CONSOL GRADE";
  const sheet = ss.getSheetByName(sheetName);

  if (!sheet) throw new Error("Could not find '" + sheetName + "'. Please generate records first.");

  const a1Value = sheet.getRange("A1").getValue().toString();
  let cleanGradeSection = a1Value.includes("'S") ? a1Value.split("'S")[0].trim() : "GRADE";
  const headerTitle = cleanGradeSection + "'S " + gradingMap[quarter].toUpperCase() + " ACADEMIC AWARDS";
  const fileName = headerTitle + ".pdf";

  const file = DriveApp.getFileById(ss.getId());
  const folder = file.getParents().hasNext() ? file.getParents().next() : DriveApp.getRootFolder();
  
  const existingFiles = folder.getFilesByName(fileName);
  return { exists: existingFiles.hasNext(), fileName: fileName, quarter: quarter };
}

/**
 * Generates the PDF, skipping empty subjects and overwriting if commanded
 */
function generateAwardsPDF(quarter, overwrite) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const learnerSheet = ss.getSheetByName("Learner's Name");
  
  const prefixes = { "First Quarter": "1Q", "Second Quarter": "2Q", "Third Quarter": "3Q", "Fourth Quarter": "4Q" };
  const gradingMap = { "First Quarter": "First Grading", "Second Quarter": "Second Grading", "Third Quarter": "Third Grading", "Fourth Quarter": "Fourth Grading" };
  
  const sheet = ss.getSheetByName(prefixes[quarter] + " CONSOL GRADE");

  const normalizeSub = (name) => {
    if (!name) return "";
    const raw = name.toString().toUpperCase().replace(/\s/g, "").replace(/\./g, ""); 
    if (["AP", "ARPAN", "ARALPAN", "ARALINGPANLIPUNAN"].includes(raw)) return "ARALING PANLIPUNAN";
    if (["PE", "PHYSICALEDUCATION"].includes(raw)) return "PHYSICAL EDUCATION";
    return name.toString().toUpperCase().trim();
  };

  const a1Value = sheet.getRange("A1").getValue().toString();
  let cleanGradeSection = a1Value.includes("'S") ? a1Value.split("'S")[0].trim() : "GRADE";
  const headerTitle = cleanGradeSection + "'S " + gradingMap[quarter].toUpperCase() + " ACADEMIC AWARDS";
  const fileName = headerTitle + ".pdf";

  const fileId = ss.getId();
  const file = DriveApp.getFileById(fileId);
  const folder = file.getParents().hasNext() ? file.getParents().next() : DriveApp.getRootFolder();

  if (overwrite) {
    const oldFiles = folder.getFilesByName(fileName);
    while (oldFiles.hasNext()) {
      oldFiles.next().setTrashed(true);
    }
  }

  const lastRow = sheet.getLastRow();
  const mainData = sheet.getRange(3, 1, lastRow - 2, 17).getValues(); 
  let mainTableRows = "";
  
  mainData.forEach(row => {
    if (row[1]) {
      mainTableRows += `<tr><td class="center">${row[0]}</td><td class="learner-name">${row[1]}</td><td class="center">${row[15]}</td><td class="center award-text">${row[16]}</td></tr>`;
    }
  });

  const subjectsRow = sheet.getRange("C2:O2").getValues()[0];
  const gradesData = sheet.getRange(3, 2, lastRow - 2, 14).getValues(); 
  const teacherMapping = learnerSheet.getRange("G4:H18").getValues();
  
  const getTeacher = (subName) => {
    const stdName = normalizeSub(subName);
    const match = teacherMapping.find(r => normalizeSub(r[0]) === stdName);
    return match ? match[1] : "Subject Teacher";
  };

  let validSubjects = [];
  subjectsRow.forEach((sub, index) => {
    if (sub && sub !== "") {
      let achievers = [];
      gradesData.forEach(row => {
        const studentName = row[0];
        const grade = row[index + 1];
        if (grade >= 98 && studentName) achievers.push(studentName);
      });
      if (achievers.length > 0) {
        validSubjects.push({ name: sub, achievers: achievers });
      }
    }
  });

  let excellenceHtml = "";
  for (let i = 0; i < validSubjects.length; i += 3) {
    let chunk = [validSubjects[i], validSubjects[i+1], validSubjects[i+2]];

    excellenceHtml += `<tr>`;
    chunk.forEach(subObj => {
      if (subObj) excellenceHtml += `<td class="excel-td"><div class="tch-label">Subject teacher:</div><div class="tch-name">${getTeacher(subObj.name).toUpperCase()}</div></td>`;
      else excellenceHtml += `<td class="excel-td no-border"></td>`;
    });
    excellenceHtml += `</tr><tr>`;
    chunk.forEach(subObj => {
      if (subObj) excellenceHtml += `<td class="excel-td black-row">${normalizeSub(subObj.name)}</td>`;
      else excellenceHtml += `<td class="excel-td no-border"></td>`;
    });
    excellenceHtml += `</tr><tr>`;
    chunk.forEach(subObj => {
      if (subObj) excellenceHtml += `<td class="excel-td val-top student-list learner-name">${subObj.achievers.join("<br>")}</td>`;
      else excellenceHtml += `<td class="excel-td no-border"></td>`;
    });
    excellenceHtml += `</tr>`;
  }

  const htmlContent = `
    <html>
      <head>
        <style>
          @page { size: 8.5in 13in; margin: 0.5in; }
          body { font-family: 'Century Gothic', sans-serif; font-size: 12px; color: #000; margin: 0; }
          .header { text-align: center; margin-bottom: 20px; }
          .title { font-size: 14px; font-weight: bold; text-transform: uppercase; border-bottom: 2px solid #000; display: inline-block; padding-bottom: 4px; }
          .main-table { width: 100%; border-collapse: collapse; table-layout: fixed; margin-bottom: 30px; page-break-inside: auto; }
          .main-table tr { page-break-inside: avoid; page-break-after: auto; }
          .main-table th, .main-table td { border: 1px solid #000; padding: 6px 8px; word-wrap: break-word; font-size: 12px; }
          .main-table th { background-color: #f2f2f2; font-weight: bold; text-transform: uppercase; }
          .center { text-align: center; }
          .award-text { font-weight: bold; }
          .section-label { font-weight: bold; font-size: 13px; margin-bottom: 10px; text-transform: uppercase; page-break-before: auto; }
          .excel-table { width: 100%; border-collapse: collapse; table-layout: fixed; page-break-inside: auto; }
          .excel-table tr { page-break-inside: avoid; page-break-after: auto; }
          .excel-td { border: 1px solid #000; width: 33.33%; padding: 6px; word-wrap: break-word; font-size: 12px; }
          .no-border { border: none !important; }
          .tch-label { font-weight: bold; font-size: 10px; margin-bottom: 2px; }
          .tch-name { font-weight: bold; }
          .black-row { background-color: #000000 !important; color: #ffffff !important; -webkit-print-color-adjust: exact; print-color-adjust: exact; text-align: center; font-size: 14px; font-weight: bold; padding: 5px 0; }
          .val-top { vertical-align: top; }
          .student-list { min-height: 60px; padding-top: 8px; line-height: 1.4; }
          .learner-name { font-family: 'Century Gothic', sans-serif !important; font-size: 12px !important; }
        </style>
      </head>
      <body>
        <div class="header"><div class="title">${headerTitle}</div></div>
        <table class="main-table">
          <thead><tr><th style="width: 40px;">#</th><th>STUDENT NAME</th><th style="width: 80px;">AVERAGE</th><th style="width: 180px;">ACADEMIC AWARD</th></tr></thead>
          <tbody>${mainTableRows}</tbody>
        </table>
        ${excellenceHtml !== "" ? `<div class="section-label">EXCELLENCE IN… (GRADES 98.00-100)</div><table class="excel-table"><tbody>${excellenceHtml}</tbody></table>` : ""}
      </body>
    </html>
  `;

  const blob = Utilities.newBlob(htmlContent, MimeType.HTML).getAs(MimeType.PDF).setName(fileName);
  return folder.createFile(blob).getUrl(); 
}