function SheetListing() {
  let sheets = SpreadsheetApp.getActiveSpreadsheet().getSheets();
  let spreadsheetID = SpreadsheetApp.getActiveSpreadsheet().getId();

  let hyperLinkList = [];
  for (let i = 0; i < sheets.length; i++) {
    let sheetId = sheets[i].getSheetId();
    let sheetName = sheets[i].getSheetName();
    let url =
      "https://docs.google.com/spreadsheets/d/" +
      spreadsheetID +
      "/edit#gid=" +
      sheetId;
    hyperLinkList[i] = ['=HYPERLINK("' + url + '","' + sheetName + '")'];
  }

  let sheet =
    SpreadsheetApp.getActiveSpreadsheet().getSheetByName("シート一覧サブ");
  if (!sheet) {
    sheet = SpreadsheetApp.getActiveSpreadsheet().insertSheet("シート一覧サブ");
  } else {
    sheet.clear();
  }

  let range = sheet.getRange(1, 1, hyperLinkList.length, 1);
  range.setValues(hyperLinkList);
}
