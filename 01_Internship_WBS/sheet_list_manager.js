// === グローバル定数 ===
// (必要な定数が定義されていること)
// const INTEGRATION_SHEET_NAME = "統合シート";
// const TEMPLATE_SHEET_NAME = "template";
// const USER_INFO_SHEET_NAME = 'タスク状況概観';
// const PROJECT_OVERVIEW_SHEET_NAME = 'プロジェクト状況概観';
// const SHEETS_TO_EXCLUDE_FROM_INTEGRATION = [...]; // 除外リスト (定義済みのはず)
const SHEET_LIST_TARGET_COL = 2; // ★ B列 (1始まり)
const SHEET_LIST_START_ROW = 3; // ★ ヘッダーが1行目にあると仮定し、リストは2行目から開始 ★ (もしB1からリストなら 1 に修正)


/**
 * 「シート一覧」シートのB列を、現在のシート構成に合わせて再構築する
 * (時間駆動トリガーで定期実行を想定)
 */
function updateSheetListColumnB_Rebuild() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const listSheetName = SHEET_LIST_SHEET_NAME;
  const listColumnIndex = SHEET_LIST_TARGET_COL;
  const startRow = SHEET_LIST_START_ROW; // 書き込み開始行

  Logger.log(`シートリスト再構築処理 開始 (対象シート: "${listSheetName}", 列: ${listColumnIndex}, 開始行: ${startRow})`);

  const listSheet = ss.getSheetByName(listSheetName);
  if (!listSheet) {
    Logger.log(`エラー: シート "${listSheetName}" が見つかりません。`);
    return;
  }

  // --- 現在のスプレッドシートから有効なシートリストを取得 ---
  const allSheets = ss.getSheets();
  const spreadsheetId = ss.getId();
  let validSheetsInfo = []; // { name: String, id: String } の配列

  Logger.log(`現在の全シートをチェックし、リスト対象シートを選別します...`);
  for (const sheet of allSheets) {
    const sheetName = sheet.getName();
    const trimmedSheetName = sheetName.trim();

    // 除外対象でないシートのみを収集
    if (!SHEETS_TO_EXCLUDE_FROM_INTEGRATION.includes(trimmedSheetName)) {
      validSheetsInfo.push({
        name: trimmedSheetName, // 表示用にトリム済みを使用
        id: sheet.getSheetId()
      });
      Logger.log(`  -> 対象シートとして認識: ${trimmedSheetName}`);
    } else {
      Logger.log(`  -> 除外: ${trimmedSheetName} (元名: ${sheetName})`);
    }
  }

  // --- シート名でソート ---
  validSheetsInfo.sort((a, b) => a.name.localeCompare(b.name));
  Logger.log(`${validSheetsInfo.length} 件の対象シートを名前順にソートしました。`);

  // --- ハイパーリンク数式リストを作成 ---
  const newHyperlinkFormulas = validSheetsInfo.map(sheetInfo => {
    const url = `https://docs.google.com/spreadsheets/d/${spreadsheetId}/edit#gid=${sheetInfo.id}`;
    const formula = `=HYPERLINK("${url}","${sheetInfo.name}")`;
    return [formula]; // [[formula1], [formula2], ...] の2次元配列にする
  });

  // --- シートへの書き込み ---
  try {
    // まず既存のリスト範囲をクリア (指定された開始行からB列の最後まで)
    const lastRow = listSheet.getLastRow();
    if (lastRow >= startRow) { // クリア対象行がある場合のみ
      const clearRange = listSheet.getRange(startRow, listColumnIndex, lastRow - startRow + 1, 1);
      Logger.log(`既存リスト範囲 ${clearRange.getA1Notation()} をクリアします。`);
      clearRange.clearContent(); // 内容のみクリア
    } else {
      Logger.log("既存リスト範囲にクリア対象行はありません。");
    }

    // 新しいリストを書き込む (リストが空でなければ)
    if (newHyperlinkFormulas.length > 0) {
      const targetRange = listSheet.getRange(startRow, listColumnIndex, newHyperlinkFormulas.length, 1);
      Logger.log(`新しいリスト(${newHyperlinkFormulas.length}件)を範囲 ${targetRange.getA1Notation()} に書き込みます。`);
      targetRange.setFormulas(newHyperlinkFormulas); // 数式として設定
      Logger.log("新リストの書き込み完了。");
    } else {
      Logger.log("書き込むべきシートリンクがありません。");
    }

  } catch (e) {
    Logger.log(`エラー: シートリストのクリアまたは書き込みに失敗: ${e}`);
  }

  Logger.log("シートリスト再構築処理 完了");
}