function SheetListing() {
    // スプレッドシート内の全シート取得
    let sheets = SpreadsheetApp.getActiveSpreadsheet().getSheets();
    // スプレッドシートのID取得
    let spreadsheetID = SpreadsheetApp.getActiveSpreadsheet().getId();

    // ハイパーリンク文字列の配列
    let hyperLinkList = [];
    for(let i=0; i<sheets.length; i++) {
        // シートのID取得
        let sheetId = sheets[i].getSheetId();
        // シートの名前取得
        let sheetName = sheets[i].getSheetName();
        // シートのURLからハイパーリンク文字列を組み立て
        let url = "https://docs.google.com/spreadsheets/d/" + spreadsheetID + "/edit#gid=" + sheetId;
        hyperLinkList[i] = [ '=HYPERLINK("' + url + '","' + sheetName + '")' ];
    }

    // 「シート一覧」シートを取得または作成
    let sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("シート一覧サブ");
    if (!sheet) {
        sheet = SpreadsheetApp.getActiveSpreadsheet().insertSheet("シート一覧サブ");
    } else {
        sheet.clear(); // 既存の内容をクリア
    }

    // シートにハイパーリンク文字列を書き込む
    let range = sheet.getRange(1, 1, hyperLinkList.length, 1);
    range.setValues(hyperLinkList);
}