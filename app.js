const MAIN_FOLDER_NAME = "Student_Certificates";

function organizeFiles(e) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  const row = e.range.getRow();
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const rowData = sheet.getRange(row, 1, 1, sheet.getLastColumn()).getValues()[0];

  const studentId = rowData[1]; // Column B = Student ID
  if (!studentId) return;

  const mainFolder = getOrCreateMainFolder();
  const studentFolder = getOrCreateStudentFolder(mainFolder, studentId);

  for (let i = 0; i < headers.length; i++) {
    const question = headers[i];
    const value = rowData[i];

    if (!value || !value.toString().includes("drive.google.com")) continue;

    const fileIdMatch = value.match(/[-\w]{25,}/);
    if (!fileIdMatch) continue;

    const file = DriveApp.getFileById(fileIdMatch[0]);

    // Keep original extension (jpg/png/pdf)
    const ext = file.getName().split('.').pop();

    // 👇 EXACT NAME LIKE FORM (clean only unwanted text)
    const cleanName = formatFileName(question);

    // 🗑 Remove older file of same document type
    deleteOldFile(studentFolder, cleanName);

    // Rename + Move
    file.setName(cleanName + "." + ext);
    file.moveTo(studentFolder);
  }
}

function formatFileName(q) {
  return q
    .replace(/Upload.*/i, "")        // remove upload text
    .replace(/\(.*?\)/g, "")         // remove (Transfer Certificate) brackets
    .replace(/\*/g, "")              // remove star
    .trim();                         // KEEP SPACES
}

function deleteOldFile(folder, baseName) {
  const files = folder.getFiles();
  while (files.hasNext()) {
    const f = files.next();
    if (f.getName().startsWith(baseName)) {
      f.setTrashed(true); // delete old version
    }
  }
}

function getOrCreateMainFolder() {
  const folders = DriveApp.getFoldersByName(MAIN_FOLDER_NAME);
  return folders.hasNext() ? folders.next() : DriveApp.createFolder(MAIN_FOLDER_NAME);
}

function getOrCreateStudentFolder(mainFolder, studentId) {
  const folders = mainFolder.getFoldersByName(studentId.toString());
  return folders.hasNext() ? folders.next() : mainFolder.createFolder(studentId.toString());
}

function checkDocumentsDetailed() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const formSheet = ss.getSheetByName("Form responses 1");

  const verifySheet =
    ss.getSheetByName("Verification") || ss.insertSheet("Verification");

  const missingSheet =
    ss.getSheetByName("Missing_Documents") || ss.insertSheet("Missing_Documents");

  verifySheet.clear();
  missingSheet.clear();

  /* ================= VERIFICATION HEADER (UNCHANGED) ================= */
  verifySheet.appendRow([
    "Student ID","Student Photo","Student Signature","Student Aadhar","Student 10th Marksheet","Student 11th Marksheet","Student 12th Marksheet","Student TC",
    "Student Community","Student Nativity","Student Bank Passbook",
    "Parent Photo","Parent Signature","Parent Aadhar","Ration Card",
    "Parent Community","Parent Income","Parent TC","Parent Voter ID",
    "Welfare Board","Parent Bank Front",
    "Parent Bank Last Transaction",
    "Additional Certificate 1",
    "Additional Certificate 2",
    "Additional Certificate 3"
  ]);

  /* ================= MISSING DOCUMENTS HEADER ================= */
  missingSheet.appendRow([
    "Student ID",
    "Missing Documents"
  ]);

  const data = formSheet.getDataRange().getValues();
  const mainFolder = DriveApp.getFoldersByName("Student_Certificates").next();

  const docNames = [
    "Student Photo",
    "Student Signature",
    "Student Aadhar Card",
    "Student 10th Marksheet",
    "Student 11th Marksheet",
    "Student 12th Marksheet",
    "Student Transfer Certificate",
    "Student Community Certificate",
    "Student Nativity Certificate",
    "Student Bank Passbook Front Page",
    "Parent Photo",
    "Parent Signature",
    "Parent Aadhar Card",
    "Parent Ration Smart Card",
    "Parent Community Certificate",
    "Parent Income Certificate",
    "Parent TC Transfer Certificate",
    "Parent Voter ID",
    "Parent Welfare Board ID Nalavariyam",
    "Parent Bank Passbook Front Page",
    "Parent Bank Last Transaction",
    "Additional Certificate 1",
    "Additional Certificate 2",
    "Additional Certificate 3"
  ];

  for (let i = 1; i < data.length; i++) {
    const studentID = data[i][1];
    if (!studentID) continue;

    const verifyRow = [studentID];
    const missingDocs = [];

    const folders = mainFolder.getFoldersByName(studentID.toString());

    if (!folders.hasNext()) {
      verifySheet.appendRow([studentID, "Folder Missing ❌"]);
      missingSheet.appendRow([studentID, "Student Folder Missing ❌"]);
      continue;
    }

    const folder = folders.next();
    const files = folder.getFiles();
    const fileList = [];

    while (files.hasNext()) {
      fileList.push(files.next().getName());
    }

    docNames.forEach(doc => {
      const found = fileList.some(name => name.startsWith(doc));
      verifyRow.push(found ? "✅" : "❌");
      if (!found) missingDocs.push(doc);
    });

    verifySheet.appendRow(verifyRow);

    missingSheet.appendRow([
      studentID,
      missingDocs.length ? missingDocs.join(", ") : "All Documents Uploaded ✅"
    ]);
  }
}
