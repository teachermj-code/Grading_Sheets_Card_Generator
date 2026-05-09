/**
 * ==========================================
 * 30_ClassRecord.gs
 * PURPOSE: Grading sheet generation, Data Archiving, & System Reset
 * ==========================================
 */

function showClassRecordUI() {
  const template = HtmlService.createTemplateFromFile('30_ClassRecordForm');
  const html = template.evaluate()
      .setWidth(850)
      .setHeight(650)
      .setTitle(' ');
  SpreadsheetApp.getUi().showModalDialog(html, ' ');
}

function processBatchForm(data) {
  // Note: Archiving is now handled dynamically by the frontend premium process
  
  // 1. GENERATION LOGIC
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sourceSheet = ss.getSheetByName("Learner's Name");
  const studentNames = sourceSheet.getRange("B4:B18").getValues();
  const subjectsSource = sourceSheet.getRange("G4:G18").getValues().flat();
  const teachersSource = sourceSheet.getRange("H4:H18").getValues().flat();
  const prefixes = { "First Quarter": "1Q", "Second Quarter": "2Q", "Third Quarter": "3Q", "Fourth Quarter": "4Q" };
  const qPrefix = prefixes[data.quarter];

  // Process Subject Sheets
  data.selections.forEach(item => {
    const subjectIndex = subjectsSource.indexOf(item.subject);
    const teacherName = (subjectIndex !== -1) ? teachersSource[subjectIndex] : "N/A";
    const template = ss.getSheetByName(item.hasExam ? "W_Exam" : "WO_Exam");
    const newName = qPrefix + " " + item.subject.toUpperCase();
    
    if (ss.getSheetByName(newName)) ss.deleteSheet(ss.getSheetByName(newName));
    const newSheet = template.copyTo(ss).setName(newName).showSheet();
    template.hideSheet();
    
    newSheet.getRange("A1").setValue(data.quarter.toUpperCase());
    newSheet.getRange("G1").setValue("GRADE & SECTION: " + data.gradeSection.toUpperCase());
    newSheet.getRange("S1").setValue("TEACHER: " + teacherName.toUpperCase());
    newSheet.getRange("AG1").setValue("SUBJECT: " + item.subject.toUpperCase());
    newSheet.getRange("B7:B21").setValues(studentNames).setFontFamily(SYSTEM_CONFIG.fontPrimary);
  });

  // Process Compulsory Sheets
  const compulsory = ["HOMEROOM GUIDANCE", "HOMEROOM GUIDANCE LETTER GRADE", "CONSOL GRADE"];
  compulsory.forEach(name => {
    const template = ss.getSheetByName(name);
    if (!template) return;
    
    const newName = qPrefix + " " + name;
    if (ss.getSheetByName(newName)) ss.deleteSheet(ss.getSheetByName(newName));
    
    const newSheet = template.copyTo(ss).setName(newName).showSheet();
    template.hideSheet();
    
    const nameRange = (name === "CONSOL GRADE") ? "B3:B17" : "B6:B20";
    newSheet.getRange(nameRange).setValues(studentNames).setFontFamily(SYSTEM_CONFIG.fontPrimary);
    
    if (name === "HOMEROOM GUIDANCE") {
      newSheet.getRange("A1").setValue(data.quarter.toUpperCase());
      newSheet.getRange("F1").setValue("GRADE & SECTION: " + data.gradeSection.toUpperCase());
    }
    
    if (name === "CONSOL GRADE") {
      let gsName = data.gradeSection.toUpperCase();
      if (!gsName.startsWith("GRADE ")) gsName = "GRADE " + gsName;
      newSheet.getRange("A1").setValue(gsName + "'S " + data.quarter.toUpperCase() + " CONSOLIDATED GRADES");
      newSheet.getRange("F1").setValue("GRADE & SECTION: " + data.gradeSection.toUpperCase());
    }

    // Formulas & Protection
    if (name === "HOMEROOM GUIDANCE LETTER GRADE" || name === "CONSOL GRADE") {
      newSheet.createTextFinder("'HOMEROOM GUIDANCE'").matchFormulaText(true).replaceAllWith("'" + qPrefix + " HOMEROOM GUIDANCE'");
      
      if (name === "CONSOL GRADE") {
        subjectsSource.forEach(sub => {
          if (sub !== "") {
            const newTarget = "'" + qPrefix + " " + sub.toUpperCase() + "'!";
            newSheet.createTextFinder("'" + sub + "'!").matchFormulaText(true).replaceAllWith(newTarget);
            newSheet.createTextFinder(sub + "!").matchFormulaText(true).replaceAllWith(newTarget);
          }
        });
      }
      const protection = newSheet.protect().setDescription('View Only Formula Sheet');
      protection.setWarningOnly(true);
    }
    
    ss.setActiveSheet(newSheet);
    ss.moveActiveSheet(ss.getNumSheets());
  });

  ss.setActiveSheet(sourceSheet);
  updateConsolidatedHeader(data.gradeSection); // Calls function from 10_Navigation.gs
  secureSystemTemplates(); // Lock down templates immediately
  return "Success";
}

function checkSequence(targetQuarter) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sequence = ["First Quarter", "Second Quarter", "Third Quarter", "Fourth Quarter"];
  const prefixes = ["1Q", "2Q", "3Q", "4Q"];
  
  const targetIndex = sequence.indexOf(targetQuarter);
  if (targetIndex === 0) return { allowed: true, needsConfirm: false }; 
  
  const prevPrefix = prefixes[targetIndex - 1];
  const prevName = sequence[targetIndex - 1];
  const prevConsol = ss.getSheetByName(prevPrefix + " CONSOL GRADE");
  
  if (!prevConsol) {
    return { allowed: false, msg: "Access Denied: You must generate " + prevName + " records before starting " + targetQuarter + "." };
  }
  return { 
    allowed: true, 
    needsConfirm: true, 
    prevName: prevName, 
    prevPrefix: prevPrefix,
    msg: `Generating ${targetQuarter} will automatically HIDE and LOCK all ${prevName} sheets for security. Proceed?` 
  };
}

function getExistingGradeSection() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  // Look through any existing quarter records first
  const possibleSheets = ["1Q CONSOL GRADE", "2Q CONSOL GRADE", "3Q CONSOL GRADE", "4Q CONSOL GRADE", "FINAL CONSOLIDATED"];
  
  for (let name of possibleSheets) {
    const sheet = ss.getSheetByName(name);
    if (sheet) {
      const val = sheet.getRange("F1").getValue().toString();
      if (val && val.includes("GRADE & SECTION:")) {
        return val.replace("GRADE & SECTION:", "").trim();
      }
    }
  }
  return "";
}

function performArchiveStep(prevPrefix) {
  archiveQuarter(prevPrefix);
  return "Archived";
}

function archiveQuarter(prefix) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  ss.getSheets().forEach(s => {
    if (s.getName().startsWith(prefix)) {
      try { s.setSheetVisibility(SpreadsheetApp.SheetVisibility.VERY_HIDDEN); } catch (e) { s.hideSheet(); }
      
      // FIX: Remove existing protections (like 'Warning Only') before locking
      const oldProtections = s.getProtections(SpreadsheetApp.ProtectionType.SHEET);
      oldProtections.forEach(p => p.remove());

      const protection = s.protect().setDescription('Archived Quarter Protection');
      protection.removeEditors(protection.getEditors());
      if (protection.canDomainEdit()) protection.setDomainEdit(false);
    }
  });
}

function secureSystemTemplates() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const templates = ["W_Exam", "WO_Exam", "CONSOL GRADE", "HOMEROOM GUIDANCE", "HOMEROOM GUIDANCE LETTER GRADE"];
  templates.forEach(name => {
    const sheet = ss.getSheetByName(name);
    if (sheet) {
      try { sheet.setSheetVisibility(SpreadsheetApp.SheetVisibility.VERY_HIDDEN); } catch (e) { sheet.hideSheet(); }
      const protection = sheet.protect().setDescription('Core Template Protection');
      protection.removeEditors(protection.getEditors());
      if (protection.canDomainEdit()) protection.setDomainEdit(false);
    }
  });
  return "Security Protocol Active.";
}

function resetSpreadsheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const protectedSheets = ["Learner's Name", "W_Exam", "WO_Exam", "HOMEROOM GUIDANCE", "HOMEROOM GUIDANCE LETTER GRADE", "CONSOL GRADE"];
  let deletedCount = 0;

  ss.getSheets().forEach(sheet => {
    if (!protectedSheets.includes(sheet.getName())) {
      ss.deleteSheet(sheet);
      deletedCount++;
    }
  });
  return "Cleanup complete. Deleted " + deletedCount + " sheets.";
}

function deleteQuarterSheets(quarterName) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const prefixes = { "First Quarter": "1Q ", "Second Quarter": "2Q ", "Third Quarter": "3Q ", "Fourth Quarter": "4Q " };
  const prefix = prefixes[quarterName];
  let count = 0;
  
  ss.getSheets().forEach(sheet => {
    if (sheet.getName().startsWith(prefix)) {
      ss.deleteSheet(sheet);
      count++;
    }
  });
  return "Deleted " + count + " sheets for " + quarterName;
}