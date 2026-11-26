// Kreative Nova — saves every website enquiry into this Google Sheet
// and emails you a notification.
const NOTIFY_EMAIL = "kreativenovaa@gmail.com";   // change if needed

function doPost(e) {
  const p = (e && e.parameter) || {};
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheets()[0];

  if (sheet.getLastRow() === 0) {
    sheet.appendRow(["Date & Time", "Name", "Phone", "Business", "Package", "Project details", "Page"]);
    sheet.getRange(1, 1, 1, 7).setFontWeight("bold").setBackground("#b7ff3c");
    sheet.setFrozenRows(1);
  }

  // Phone saved as text so the + and leading digits are not lost
  const phoneCell = "'" + (p.phone || "");
  sheet.appendRow([new Date(), p.name || "", phoneCell, p.business || "", p.package || "", p.message || "", p.page || ""]);

  try {
    MailApp.sendEmail(NOTIFY_EMAIL, "New enquiry — " + (p.name || "Website"),
      "Name: " + (p.name || "") + "\nPhone: " + (p.phone || "") + "\nBusiness: " + (p.business || "-") +
      "\nPackage: " + (p.package || "") + "\nProject: " + (p.message || "-"));
  } catch (err) {}

  return ContentService.createTextOutput("ok");
}

function doGet() { return ContentService.createTextOutput("Kreative Nova enquiry endpoint is running."); }
