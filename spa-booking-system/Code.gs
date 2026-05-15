// ============================================================
//  BOOKING FORM — Google Apps Script
//  Paste this entire file into your Google Apps Script editor
//  (script.google.com → New Project → replace all code → Save)
// ============================================================


const SCRIPT_URL = "https://script.google.com/macros/s/AKfycbzlir2B1rdEmchJ56QQ1orNW4mJOdbRNRjWseZ2mfdErGB_oNC8Z5hf6YCrL5Q-cOBS2w/exec"


const SHEET_NAME_MASSAGE    = "Massage Bookings";
const SHEET_NAME_PEDICURE   = "Pedicure & Manicure Bookings";

function doPost(e) {
  try {
    const data       = JSON.parse(e.postData.contents);
    const ss         = SpreadsheetApp.getActiveSpreadsheet();
    const sheetName  = data.formType === "massage" ? SHEET_NAME_MASSAGE : SHEET_NAME_PEDICURE;
    let   sheet      = ss.getSheetByName(sheetName);

    // Create sheet with headers if it doesn't exist yet
    if (!sheet) {
      sheet = ss.insertSheet(sheetName);
      const headers = data.formType === "massage"
        ? ["Timestamp", "Employee Name", "School", "Date", "Time Slot", "Slot #"]
        : ["Timestamp", "Employee Name", "School", "Date", "Service", "Time Slot", "Slot #"];
      sheet.appendRow(headers);
      sheet.getRange(1, 1, 1, headers.length)
        .setFontWeight("bold")
        .setBackground("#E8F5E9");
      sheet.setFrozenRows(1);
    }

    // Write each booked slot as its own row
    const timestamp = new Date().toLocaleString("en-ZA", { timeZone: "Africa/Johannesburg" });
    data.slots.forEach(slot => {
      const row = data.formType === "massage"
        ? [timestamp, data.name, data.school, data.date, slot.time, slot.slotNumber]
        : [timestamp, data.name, data.school, data.date, data.service, slot.time, slot.slotNumber];
      sheet.appendRow(row);
    });

    return ContentService
      .createTextOutput(JSON.stringify({ status: "ok", rows: data.slots.length }))
      .setMimeType(ContentService.MimeType.JSON);

  } catch (err) {
    return ContentService
      .createTextOutput(JSON.stringify({ status: "error", message: err.toString() }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

// Optional: run this once manually to pre-create both sheets
function setup() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  [
    { name: SHEET_NAME_MASSAGE,  headers: ["Timestamp", "Employee Name", "School", "Date", "Time Slot", "Slot #"] },
    { name: SHEET_NAME_PEDICURE, headers: ["Timestamp", "Employee Name", "School", "Date", "Service", "Time Slot", "Slot #"] }
  ].forEach(({ name, headers }) => {
    if (!ss.getSheetByName(name)) {
      const sheet = ss.insertSheet(name);
      sheet.appendRow(headers);
      sheet.getRange(1, 1, 1, headers.length)
        .setFontWeight("bold")
        .setBackground("#E8F5E9");
      sheet.setFrozenRows(1);
    }
  });

  Logger.log("Setup complete — both sheets created.");
}
