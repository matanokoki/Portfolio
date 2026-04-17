const SHEET_LIST_TARGET_COL = 2; 
const SHEET_LIST_START_ROW = 3; 

function updateSheetListColumnB_Rebuild() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const listSheetName = SHEET_LIST_SHEET_NAME;
  const listColumnIndex = SHEET_LIST_TARGET_COL;
  const startRow = SHEET_LIST_START_ROW; 

  Logger.log(`シートリスト再構築処理 開始 (対象シート: "${listSheetName}", 列: ${listColumnIndex}, 開始行: ${startRow})`);

  const listSheet = ss.getSheetByName(listSheetName);
  if (!listSheet) {
    Logger.log(`エラー: シート "${listSheetName}" が見つかりません。`);
    return;
  }

  const allSheets = ss.getSheets();
  const spreadsheetId = ss.getId();
  let validSheetsInfo = []; 

  Logger.log(`現在の全シートをチェックし、リスト対象シートを選別します...`);
  for (const sheet of allSheets) {
    const sheetName = sheet.getName();
    const trimmedSheetName = sheetName.trim();

    if (!SHEETS_TO_EXCLUDE_FROM_INTEGRATION.includes(trimmedSheetName)) {
      validSheetsInfo.push({
        name: trimmedSheetName,
        id: sheet.getSheetId()
      });
      Logger.log(`  -> 対象シートとして認識: ${trimmedSheetName}`);
    } else {
      Logger.log(`  -> 除外: ${trimmedSheetName} (元名: ${sheetName})`);
    }
  }

  validSheetsInfo.sort((a, b) => a.name.localeCompare(b.name));
  Logger.log(`${validSheetsInfo.length} 件の対象シートを名前順にソートしました。`);

  const newHyperlinkFormulas = validSheetsInfo.map(sheetInfo => {
    const url = `https://docs.google.com/spreadsheets/d/${spreadsheetId}/edit#gid=${sheetInfo.id}`;
    const formula = `=HYPERLINK("${url}","${sheetInfo.name}")`;
    return [formula]; 
  });

  try {
    const lastRow = listSheet.getLastRow();
    if (lastRow >= startRow) { 
      const clearRange = listSheet.getRange(startRow, listColumnIndex, lastRow - startRow + 1, 1);
      Logger.log(`既存リスト範囲 ${clearRange.getA1Notation()} をクリアします。`);
      clearRange.clearContent();
    } else {
      Logger.log("既存リスト範囲にクリア対象行はありません。");
    }

    if (newHyperlinkFormulas.length > 0) {
      const targetRange = listSheet.getRange(startRow, listColumnIndex, newHyperlinkFormulas.length, 1);
      Logger.log(`新しいリスト(${newHyperlinkFormulas.length}件)を範囲 ${targetRange.getA1Notation()} に書き込みます。`);
      targetRange.setFormulas(newHyperlinkFormulas); 
      Logger.log("新リストの書き込み完了。");
    } else {
      Logger.log("書き込むべきシートリンクがありません。");
    }

  } catch (e) {
    Logger.log(`エラー: シートリストのクリアまたは書き込みに失敗: ${e}`);
  }

  Logger.log("シートリスト再構築処理 完了");
}