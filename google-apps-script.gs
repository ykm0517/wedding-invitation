function doPost(e) {
  try {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("RSVP") || SpreadsheetApp.getActiveSpreadsheet().insertSheet("RSVP");

    if (sheet.getLastRow() === 0) {
      sheet.appendRow([
        "createdAt",
        "side",
        "attendance",
        "meal",
        "name",
        "phone",
        "guests",
        "guest_names",
        "message",
        "source"
      ]);
    }

    var data = JSON.parse(e.postData.contents || "{}");

    sheet.appendRow([
      data.createdAt || "",
      data.side || "",
      data.attendance || "",
      data.meal || "",
      data.name || "",
      data.phone || "",
      data.guests || "",
      data.guest_names || "",
      data.message || "",
      data.source || ""
    ]);

    return ContentService
      .createTextOutput(JSON.stringify({ status: "success" }))
      .setMimeType(ContentService.MimeType.JSON);
  } catch (error) {
    return ContentService
      .createTextOutput(JSON.stringify({ status: "error", message: String(error) }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

function doGet() {
  return ContentService
    .createTextOutput(JSON.stringify({ status: "ok" }))
    .setMimeType(ContentService.MimeType.JSON);
}
