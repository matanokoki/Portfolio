// ========================================================================
// == 0. スクリプト全体の設定 ===============================================
// ========================================================================

/**
 * @fileoverview
 * スプレッドシートのタスク管理に関する自動化スクリプト。
 * 1. 日付に基づくステータス自動更新（遅延、実施中）
 * 2. 担当者別・種類別（開始予定/遅延）のSlackダイジェスト通知
 * 3. 複数シートのデータを単一シートに統合（表示名とメアドを分離）
 * 4. トリガー管理ユーティリティ
 * 5. ステータス更新テスト用関数
 */

// --- ユーザー情報シート設定 (通知・統合で共通利用) ---
const USER_INFO_SHEET_NAME = 'USER_INFO_SHEET_NAME'; // 表示名(A), Slack ID(B), メールアドレス(C)が記載されたシート
const USER_INFO_NAME_COL = 1;       // 表示名列 (A列=1)
const USER_INFO_SLACK_ID_COL = 2;   // Slack ID列 (B列=2)
const USER_INFO_EMAIL_COL = 3;      // メールアドレス列 (C列=3)

// --- ステータス定義 ---
const FINAL_STATUSES = ['完了', 'キャンセル']; // このステータスは処理スキップ 
const STATUS_DELAYED = '遅延';             
const STATUS_IN_PROGRESS = '実施中';       
const STATUS_NOT_STARTED = '未着手';       

// --- ソースシート(統合元)の列インデックス (1始まり) ---
// ※ カテゴリ列(A列)は削除済み前提
const SOURCE_PROJECT_COL_IDX = 1;      // プロジェクト列 (A列=1)
const SOURCE_PSUB_COL_IDX = 2;         // P-Sub列 (B列=2)
const SOURCE_TASK_COL_IDX = 3;         // タスク内容列 (C列=3)
const SOURCE_MANAGER_NAME_COL_IDX = 4; // 担当者名(表示名/チップ)列 (D列=4)
const SOURCE_START_DATE_COL_IDX = 5;   // 開始日列 (E列=5) 
const SOURCE_DEADLINE_COL_IDX = 6;     // 締め切り日列 (F列=6)
const SOURCE_STATUS_COL_IDX = 7;       // ステータス列 (G列=7)


// ========================================================================
// == 1. 通知機能 (ステータス更新含むダイジェスト形式) =======================
// ========================================================================

// --- 通知用設定 ---
// ★ 通知を送る先の【単一の】Webhook URL
const DIGEST_WEBHOOK_URL = 'DIGEST_WEBHOOK_URL';

// 処理対象シートごとの設定 (列番号は上記定数を使用)
const SHEET_CONFIGS_FOR_DIGEST_NOTIF = {
  // シート名: { sheetIdentifier: '通知用識別子' } の形式
  'SHEET1': { sheetIdentifier: '（SHEET1）' },
  'SHEET2': { sheetIdentifier: '（SHEET2）' },
  'SHEET3':     { sheetIdentifier: '（SHEET3）' },
  'SHEET4':   { sheetIdentifier: '（SHEET4）' },
  'SHEET5':   { sheetIdentifier: '（SHEET5）' },
  // 他のシートがあればここに追加
};

// ========================================================================
// == Slack通知関連ヘルパー ==============================================
// ========================================================================
const SlackHelper = {
  /** 表示名からSlack IDを取得 */
  getSlackId: function(displayName, userInfoSheet) {
    if (!userInfoSheet) { Logger.log("エラー(getSlackId): userInfoSheetがnullです"); return null; }
    if (!displayName || typeof displayName !== 'string') return null;
    const searchName = displayName.trim(); if (searchName === "") return null;
    try {
        const data = userInfoSheet.getDataRange().getValues();
        const nameColIdx = USER_INFO_NAME_COL - 1; // グローバル定数を参照
        const slackIdColIdx = USER_INFO_SLACK_ID_COL - 1; // グローバル定数を参照
        for (let i = 1; i < data.length; i++) {
            const nameInSheet = data[i][nameColIdx];
            if (nameInSheet && typeof nameInSheet === 'string' && nameInSheet.trim() === searchName) {
                const slackId = data[i][slackIdColIdx];
                return (slackId !== undefined && slackId !== null && slackId !== "") ? String(slackId) : null;
            }
        }
        // Logger.log(`Slack ID検索: "${searchName}" のID見つからず`);
        return null;
    } catch (e) { Logger.log(`エラー(getSlackId): ${e}`); return null; }
  },

  /** Slackにメッセージ送信 (アイコン・名前設定込み) */
  send: function(text, webHookUrl) {
    // グローバル定数 SLACK_BOT_USERNAME と SLACK_BOT_ICON_EMOJI を参照
    const botUsername = typeof SLACK_BOT_USERNAME !== 'undefined' ? SLACK_BOT_USERNAME : "BOT_NAME"; // デフォルト名
    const botIcon = typeof SLACK_BOT_ICON_EMOJI !== 'undefined' ? SLACK_BOT_ICON_EMOJI : ":EMOJI:"; // デフォルトアイコン

    if (!webHookUrl || typeof webHookUrl !== 'string' || !webHookUrl.startsWith('https://hooks.slack.com/')) { Logger.log(`エラー: 不正なWebhook URL: ${webHookUrl}`); return; }
    if (!text || typeof text !== 'string') { Logger.log("エラー: 送信メッセージ不正"); return; }
    if (text.length > 3500) { Logger.log(`警告: メッセージ長超過のため短縮`); text = text.substring(0, 3500) + "... (省略)"; }

    // ペイロードに username と icon_emoji を含める
    const payloadObject = {
        text: text,
        username: botUsername,      // 設定したボット名を使用
        icon_emoji: botIcon         // 設定した絵文字アイコンを使用
        // icon_url: SLACK_BOT_ICON_URL // URLを使う場合はこちらを有効化
    };
    const payload = JSON.stringify(payloadObject);

    const options = { method: 'post', contentType: 'application/json', payload: payload, muteHttpExceptions: true };
    try {
        const response = UrlFetchApp.fetch(webHookUrl, options);
        const responseCode = response.getResponseCode();
        if (responseCode === 200) { Logger.log(`Slack通知送信成功 (Name:${botUsername}, Icon:${botIcon})`); }
        else { Logger.log(`Slack通知送信失敗: ${responseCode} ${response.getContentText()}`); }
    } catch (error) { Logger.log(`Slack通知送信中例外: ${error}\n${error.stack}`); }
  }
};

/**
 * メイン関数: タスク状況をチェックし、ダイジェスト通知を作成・送信 (時間駆動トリガーで実行)
 */
function checkTasksAndNotifyDigest() {
  const today = new Date(); today.setHours(0, 0, 0, 0);
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const userInfoSheet = ss.getSheetByName(USER_INFO_SHEET_NAME);
  if (!userInfoSheet) { Logger.log(`エラー: ユーザー情報シート "${USER_INFO_SHEET_NAME}" 不在`); return; }
  Logger.log("タスク状況ダイジェスト作成処理を開始します...");

  let tasksByAssignee = {};

  for (const sheetName in SHEET_CONFIGS_FOR_DIGEST_NOTIF) {
    if (SHEET_CONFIGS_FOR_DIGEST_NOTIF.hasOwnProperty(sheetName)) {
      const config = SHEET_CONFIGS_FOR_DIGEST_NOTIF[sheetName];
      const sheet = ss.getSheetByName(sheetName); if (!sheet) { Logger.log(`シート "${sheetName}" が見つかりません。`); continue; }
      const dataRange = sheet.getDataRange(); if (dataRange.getLastRow() < 2) { continue; }

      try {
          const values = dataRange.offset(1, 0, dataRange.getNumRows() - 1).getValues();
          const statusesToUpdate = [];

          for (let i = 0; i < values.length; i++) {
            const rowData = values[i]; const actualRow = i + 2;
            // ★★★ 修正: SOURCE_STATUS_COL_IDX を使用 ★★★
            let currentStatus = rowData[SOURCE_STATUS_COL_IDX - 1];
            if (FINAL_STATUSES.includes(currentStatus)) { continue; }

            const startDateValue = rowData[SOURCE_START_DATE_COL_IDX - 1];
            const deadlineValue = rowData[SOURCE_DEADLINE_COL_IDX - 1];
            const managerDisplayNameValue = rowData[SOURCE_MANAGER_NAME_COL_IDX - 1]; // 正しい定数名を使用
            const assigneeName = (managerDisplayNameValue !== undefined && managerDisplayNameValue !== null) ? String(managerDisplayNameValue).trim() : "";

            let newStatus = null;
            let deadlineDate = null; if (deadlineValue instanceof Date) { deadlineDate = new Date(deadlineValue); deadlineDate.setHours(0, 0, 0, 0); }
            let startDate = null; if (startDateValue instanceof Date) { startDate = new Date(startDateValue); startDate.setHours(0, 0, 0, 0); }

            // --- ステータス自動更新ロジック ---
            if (deadlineDate !== null && today > deadlineDate && currentStatus !== STATUS_DELAYED) { newStatus = STATUS_DELAYED; }
            else if (newStatus === null && startDate !== null && today >= startDate && (currentStatus === STATUS_NOT_STARTED || currentStatus === "")) { newStatus = STATUS_IN_PROGRESS; }

            if (newStatus !== null) {
                // ★★★ 修正: SOURCE_STATUS_COL_IDX を使用 ★★★
                statusesToUpdate.push({ row: actualRow, col: SOURCE_STATUS_COL_IDX, value: newStatus });
                currentStatus = newStatus;
            }
            // --- ステータス更新ここまで ---

            if (assigneeName === "") continue;

            // --- 担当者別オブジェクトにタスク情報を格納 ---
            if (!tasksByAssignee[assigneeName]) { tasksByAssignee[assigneeName] = { starting: [], overdue: [] }; }
            const taskInfo = { sheet: sheetName, project: rowData[SOURCE_PROJECT_COL_IDX - 1] || '?', task: rowData[SOURCE_TASK_COL_IDX - 1] || '?' };

            // 開始日通知の条件
            if (startDate !== null && startDate.getTime() === today.getTime()) {
              tasksByAssignee[assigneeName].starting.push({ ...taskInfo });
            }
            // 締め切り超過通知の条件 (更新後のステータスで判断)
            if (deadlineDate !== null && today > deadlineDate) {
              tasksByAssignee[assigneeName].overdue.push({ ...taskInfo, deadline: deadlineDate, status: currentStatus });
            }
          } // row loop

          // --- ステータスの一括書き込み (シートごと) ---
          if (statusesToUpdate.length > 0) {
              Logger.log(`シート "${sheetName}" で ${statusesToUpdate.length} 件のステータスを更新します...`);
              statusesToUpdate.forEach(update => {
                  try {
                      // ★★★ 修正: update.col は正しい列番号(SOURCE_STATUS_COL_IDX)を持っているはず ★★★
                      sheet.getRange(update.row, update.col).setValue(update.value);
                  }
                  catch (e) { Logger.log(`エラー: ${sheetName} 行 ${update.row} 更新失敗 ${e}`); }
              });
          }
      } catch (sheetError) {
          Logger.log(`エラー: シート "${sheetName}" の処理中にエラーが発生しました。 ${sheetError}`);
      }
    } // sheet check
  } // sheet loop

  // --- 収集した情報からダイジェストメッセージを作成して送信 ---
  sendDigestNotification(tasksByAssignee, userInfoSheet);

  Logger.log("タスク状況ダイジェスト作成処理が完了しました。");
}


/**
 * ヘルパー関数: ダイジェストメッセージ作成・送信
 */
function sendDigestNotification(tasksByAssignee, userInfoSheet) {
  const assigneeNames = Object.keys(tasksByAssignee);
  if (assigneeNames.length === 0) { Logger.log("通知対象タスクなし (ダイジェスト送信スキップ)"); return; }

  let assigneeSlackIDs = new Map();
  assigneeNames.forEach(name => { const slackId = SlackHelper.getSlackId(name, userInfoSheet); if (slackId) { assigneeSlackIDs.set(name, slackId); }});

  // ▼▼▼ ここから追加 ▼▼▼
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet(); // 現在のスプレッドシートを取得
  const spreadsheetUrl = spreadsheet.getUrl(); // スプレッドシートのURLを取得
  const spreadsheetName = spreadsheet.getName(); // スプレッドシートの名前を取得
  // ▲▲▲ ここまで追加 ▲▲▲

  let messageParts = [];
  const todayStr = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy年MM月dd日');
  messageParts.push(`*${todayStr} のタスク状況サマリー*`);
  let hasContent = false;

  assigneeNames.forEach(assigneeName => {
    const tasks = tasksByAssignee[assigneeName]; const startingTasks = tasks.starting; const overdueTasks = tasks.overdue;
    if (startingTasks.length > 0 || overdueTasks.length > 0) {
      hasContent = true; const slackId = assigneeSlackIDs.get(assigneeName); const mention = slackId ? `<@${slackId}>` : `@${assigneeName}`; messageParts.push(`\n*👤 ${mention}* さん`);
      if (startingTasks.length > 0) { messageParts.push(`　*▼ 本日開始予定 (${startingTasks.length}件)*`); startingTasks.forEach(t => { messageParts.push(`　　・${t.project ? `【${t.project}】` : ''} ${t.task} (${t.sheet})`); }); }
      if (overdueTasks.length > 0) { messageParts.push(`　*▼ 締め切り超過 (${overdueTasks.length}件)*`); overdueTasks.forEach(t => { const dl = Utilities.formatDate(t.deadline, Session.getScriptTimeZone(), 'MM/dd'); messageParts.push(`　　・${t.project ? `【${t.project}】` : ''} ${t.task} (締切:${dl}, 状況:${t.status}, ${t.sheet})`); }); }
    }
  });

  if (hasContent) {
    // ▼▼▼ 修正箇所 ▼▼▼
    messageParts.push(`\n\n---`); // メッセージの区切り線 (任意)
    messageParts.push(`参照スプレッドシート: <${spreadsheetUrl}|${spreadsheetName}>`); // スプレッドシートへのリンクを追加
    // ▲▲▲ 修正箇所 ▲▲▲

    const finalMessage = messageParts.join('\n'); Logger.log("送信ダイジェスト:\n" + finalMessage.substring(0, 1000) + "...");
    SlackHelper.send(finalMessage, DIGEST_WEBHOOK_URL);
  } else { Logger.log("通知内容なし (ダイジェスト送信スキップ)"); }
}


// ========================================================================
// == 2. シート統合 スクリプト ============================================
// ========================================================================


// --- 統合シート列インデックス定義 (0始まり) ---
const INTEG_CATEGORY_COL_IDX = 0;      // A列
const INTEG_PROJECT_COL_IDX = 1;       // B列
const INTEG_PSUB_COL_IDX = 2;          // C列
const INTEG_TASK_COL_IDX = 3;          // D列
const INTEG_MANAGER_NAME_COL_IDX = 4; // E列
const INTEG_START_DATE_COL_IDX = 5;    // F列
const INTEG_END_DATE_COL_IDX = 6;      // G列
const INTEG_STATUS_COL_IDX = 7;        // H列
const INTEG_EMAIL_COL_NUM = 9;





/**
 * メイン関数: 複数シートのデータを統合シートに集約
 */
// ========================================================================
// == 2. シート統合 スクリプト (シート自動検出版) ==========================
// ========================================================================

// --- 統合スクリプト設定 ---
// const SOURCE_SHEET_NAMES_FOR_INTEGRATION = [...] // ← この行は不要なので削除またはコメントアウト

const INTEGRATION_SHEET_NAME = "統合シート";
const PROJECT_OVERVIEW_SHEET_NAME = 'プロジェクト状況概観';
const TEMPLATE_SHEET_NAME = "template"; // ★★★ templateシートの名前を確認・必要なら修正 ★★★

// ★★★ 統合処理から除外するシート名のリスト ★★★
// ※ ここにリストされている名前のシートは統合対象になりません
const SHEETS_TO_EXCLUDE_FROM_INTEGRATION = [
  INTEGRATION_SHEET_NAME,         // "統合シート"
  USER_INFO_SHEET_NAME,           // "タスク状況概観" (グローバル定数)
  PROJECT_OVERVIEW_SHEET_NAME,    // "プロジェクト状況概観" (グローバル定数)
  TEMPLATE_SHEET_NAME,             // "template" シート
  // 他にも統合したくない固定シートがあればここに追加 (例: 'シート一覧', '入力フォーム回答' など)
  'シート一覧',
  'プロジェクト別タイムライン',
  '担当者別タイムライン'
];

// ... (定数定義は変更なし) ...
const TASK_OVERVIEW_HEADER_ROW = 2; 
const TASK_OVERVIEW_FIRST_DATA_ROW = 3; 
const TASK_OVERVIEW_NAME_COL = 1; 
const TASK_OVERVIEW_EMAIL_COL = 3;    // ★★★ この行を追加 ★★★ (C列: メールアドレス)
const TASK_OVERVIEW_SLACK_ID_COL = 2; // (もし定義があれば近くに) B列: Slack ID
const TASK_OVERVIEW_NOT_STARTED_COL = 4; 
const TASK_OVERVIEW_NUM_STATUS_COLS = 5;
const EXPECTED_TASK_OVERVIEW_HEADERS = ["未着手", "実施中", "完了", "遅延", "総計"]; // 必要なら修正
const PROJECT_OVERVIEW_EMAIL_ROW = 1; const PROJECT_OVERVIEW_HEADER_ROW = 2; const PROJECT_OVERVIEW_FIRST_DATA_ROW = 3; const PRJ_OVERVIEW_PROJECT_COL = 2; const PRJ_OVERVIEW_FIRST_DATA_COL = 3;


/**
 * ヘルパー関数: メールアドレス対応表Map作成 (変更なし)
 */
function createEmailMapForIntegration(sheet, nameCol, emailCol) {
  try { if (!sheet) { Logger.log(`エラー(統合): emailMap作成用シート不在 (${USER_INFO_SHEET_NAME})`); return null; } const data = sheet.getDataRange().getValues(); const map = new Map(); const nameIdx = nameCol - 1; const emailIdx = emailCol - 1; for (let i = 1; i < data.length; i++) { const name = data[i][nameIdx]; const email = data[i][emailIdx]; if (name && email && typeof name === 'string' && typeof email === 'string') { map.set(name.trim(), email.trim()); } } if (map.size === 0) { Logger.log(`警告(統合): メールアドレスMap空 (${sheet.getName()})`); } return map; } catch (error) { Logger.log(`エラー(統合): メールアドレスMap作成失敗 (${sheet.getName()}): ${error}`); return null; }
}


/**
 * メイン関数: 複数シートのデータを統合シートに集約 (シート自動検出)
 */
function integrateSheetsImproved() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  Logger.log("シート統合処理開始 (シート自動検出モード)...");
  const userInfoSheetForEmail = ss.getSheetByName(USER_INFO_SHEET_NAME);
  // ★ TASK_OVERVIEW_NAME_COL と TASK_OVERVIEW_EMAIL_COL がグローバル定義されていること
  const emailMap = createEmailMapForIntegration(userInfoSheetForEmail, TASK_OVERVIEW_NAME_COL, TASK_OVERVIEW_EMAIL_COL);
  if (emailMap) { Logger.log(`(統合) メアド対応表読込 (${emailMap.size}件)`); } else { Logger.log(`警告(統合): メアド対応表読込失敗/空`); }

  let integrationSheet = ss.getSheetByName(INTEGRATION_SHEET_NAME);
  const integrationHeader = [ "カテゴリ", "プロジェクト", "P-Sub", "タスク内容", "担当者表示名", "開始日", "終了予定日", "ステータス", "担当者メールアドレス" ];
  let needsHeader = false;
  if (!integrationSheet) { integrationSheet = ss.insertSheet(INTEGRATION_SHEET_NAME); Logger.log(`シート "${INTEGRATION_SHEET_NAME}" 作成`); needsHeader = true; }
  else { try { if (integrationSheet.getRange(1, 1, 1, integrationHeader.length).getValues()[0].every(c => c === "")) needsHeader = true; } catch (e) { needsHeader = true; Logger.log(`既存ヘッダーチェックエラー: ${e}`);}}
  if (needsHeader) { try { integrationSheet.getRange(1, 1, 1, integrationHeader.length).setValues([integrationHeader]); Logger.log(`ヘッダー書込完了`); SpreadsheetApp.flush(); } catch (e) { Logger.log(`エラー: ヘッダー書込失敗 ${e}`); return; }}

  // 統合シートの既存データをクリア (3行目以降を対象に変更)
  try {
    const lr = integrationSheet.getLastRow();
    if (lr >= 2) { // ヘッダー行を除いたデータ行がある場合のみクリア
       integrationSheet.getRange(2, 1, lr - 1, integrationHeader.length).clear({contentsOnly: true});
       Logger.log(`統合シートの既存データ( ${lr - 1} 行)をクリアしました。`);
     } else {
       Logger.log(`統合シートに既存データなし(ヘッダーのみ)。クリアはスキップします。`);
     }
  } catch (e) { Logger.log(`エラー: 既存データクリア失敗 ${e}`); }

  const allData = []; // 統合する全データを格納する配列

  // ★★★ シート自動検出処理 ★★★
  Logger.log(`統合対象シート検索開始... 除外リスト: ${SHEETS_TO_EXCLUDE_FROM_INTEGRATION.join(', ')}`);
  const sheetsToIntegrate = [];
  const allSheets = ss.getSheets();
  for (const sheet of allSheets) {
      const sheetName = sheet.getName();
      if (!SHEETS_TO_EXCLUDE_FROM_INTEGRATION.includes(sheetName)) {
          // 除外リストに含まれていないシートを対象とする
          sheetsToIntegrate.push(sheet); // シートオブジェクトを直接リストに追加
          Logger.log(`  -> 対象シート発見: ${sheetName}`);
      } else {
          Logger.log(`  -> 除外シート: ${sheetName}`);
      }
  }
  Logger.log(`統合元シート処理開始 (自動検出: ${sheetsToIntegrate.length}件)`);

  // ★★★ 自動検出したシートリストでループ ★★★
  for (const sheet of sheetsToIntegrate) {
    const sheetName = sheet.getName();
    let dataRange, numDataRows = 0;
     try {
         dataRange = sheet.getDataRange();
         // ヘッダー行が必ず存在すると仮定 (データ行がなくてもヘッダーはあるはず)
         numDataRows = Math.max(0, dataRange.getNumRows() - 1); // データ行数を計算 (マイナスにならないように)
     } catch (e) { Logger.log(`  シート "${sheetName}" のデータ範囲取得エラー: ${e}。スキップします。`); continue; }

    if (numDataRows <= 0) { Logger.log(`  シート "${sheetName}" にデータ行がありません。スキップします。`); continue; }
    Logger.log(`  シート "${sheetName}" (${numDataRows} 行) を処理中...`);

    try {
        // ヘッダー行を除いた範囲の値を取得
        const values = sheet.getRange(2, 1, numDataRows, sheet.getLastColumn()).getValues();

        for (let i = 0; i < values.length; i++) {
            const row = values[i];
            // 統合シートの列数に合わせて配列初期化
            const processedRow = new Array(integrationHeader.length).fill("");
            try {
                processedRow[INTEG_CATEGORY_COL_IDX] = sheetName; // A列にシート名をカテゴリとして格納

                // ★ ソースシートの列インデックス (0-based) - グローバル定数(1-based)から計算 ★
                // これらの SOURCE_ 定数がグローバルで正しく定義されていることが前提
                const srcProjectCol = SOURCE_PROJECT_COL_IDX - 1;
                const srcPSubCol = SOURCE_PSUB_COL_IDX - 1;
                const srcTaskCol = SOURCE_TASK_COL_IDX - 1;
                const srcManagerNameCol = SOURCE_MANAGER_NAME_COL_IDX - 1;
                const srcStartDateCol = SOURCE_START_DATE_COL_IDX - 1;
                const srcDeadlineCol = SOURCE_DEADLINE_COL_IDX - 1;
                const srcStatusCol = SOURCE_STATUS_COL_IDX - 1;

                // 各列のデータを統合シート用にコピー (範囲外アクセスを防ぐため、列が存在するか確認)
                processedRow[INTEG_PROJECT_COL_IDX] = (srcProjectCol < row.length && row[srcProjectCol] !== undefined && row[srcProjectCol] !== null) ? row[srcProjectCol] : "";
                processedRow[INTEG_PSUB_COL_IDX] = (srcPSubCol < row.length && row[srcPSubCol] !== undefined && row[srcPSubCol] !== null) ? row[srcPSubCol] : "";
                processedRow[INTEG_TASK_COL_IDX] = (srcTaskCol < row.length && row[srcTaskCol] !== undefined && row[srcTaskCol] !== null) ? row[srcTaskCol] : "";

                const managerDisplayNameVal = (srcManagerNameCol < row.length) ? row[srcManagerNameCol] : undefined;
                const managerDisplayName = (managerDisplayNameVal !== undefined && managerDisplayNameVal !== null) ? String(managerDisplayNameVal).trim() : "";
                processedRow[INTEG_MANAGER_NAME_COL_IDX] = managerDisplayName;

                const startDateVal = (srcStartDateCol < row.length) ? row[srcStartDateCol] : undefined;
                processedRow[INTEG_START_DATE_COL_IDX] = (startDateVal instanceof Date) ? startDateVal : "";

                const deadlineVal = (srcDeadlineCol < row.length) ? row[srcDeadlineCol] : undefined;
                processedRow[INTEG_END_DATE_COL_IDX] = (deadlineVal instanceof Date) ? deadlineVal : "";

                const statusVal = (srcStatusCol < row.length) ? row[srcStatusCol] : undefined;
                processedRow[INTEG_STATUS_COL_IDX] = (statusVal !== undefined && statusVal !== null) ? statusVal : "";

                // メールアドレスをMapから検索して格納
                let managerEmail = "";
                if (emailMap && managerDisplayName !== "" && emailMap.has(managerDisplayName)) {
                    managerEmail = emailMap.get(managerDisplayName);
                }
                // ★★★ 正しい定数を使ってEmailを格納 ★★★
                processedRow[INTEG_EMAIL_COL_NUM - 1] = managerEmail; // 0始まりIndex

                allData.push(processedRow);

            } catch (rowError) { Logger.log(`エラー(統合): ${sheetName} ${i + 2}行目処理中 ${rowError} - Data: ${JSON.stringify(row)}`); }
        } // row loop
    } catch (readError) { Logger.log(`エラー(統合): シート "${sheetName}" のデータ読み取り/処理失敗 ${readError}`); }
  } // sheet loop

  // 統合シートへの書き込み
  if (allData.length > 0) {
    try {
      integrationSheet.getRange(2, 1, allData.length, integrationHeader.length).setValues(allData);
      Logger.log(`統合シートに ${allData.length} 件書込`);
    } catch (writeError) { Logger.log(`エラー(統合): シート書込失敗 ${writeError}`); }
  } else { Logger.log("統合データなし"); }

  // 概観シート更新呼び出し
  try {
    updateOverviewSheets();
  } catch (e) { Logger.log(`エラー: 概観シート更新関数の呼び出しに失敗 ${e}`); }

  Logger.log("シート統合処理完了");
} // End of integrateSheetsImproved


// ========================================================================
// == 6. 概観シート更新 スクリプト (プロジェクト再構築版) ===================
// ========================================================================


/**
 * 統合シートのデータに基づき、タスク状況概観とプロジェクト状況概観シートを更新する
 * プロジェクト状況概観は統合シートの内容に基づき毎回再構築する
 */
function updateOverviewSheets() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  Logger.log("概観シート更新処理 開始...");

  // --- 対象シート取得 ---
  const integrationSheet = ss.getSheetByName(INTEGRATION_SHEET_NAME);
  const taskOverviewSheet = ss.getSheetByName(USER_INFO_SHEET_NAME);
  const projectOverviewSheet = ss.getSheetByName(PROJECT_OVERVIEW_SHEET_NAME);

  if (!integrationSheet || !taskOverviewSheet || !projectOverviewSheet) { Logger.log(`エラー: 必要なシートが見つかりません`); return; }

  // --- データ読み込み ---
  let integrationData, taskOverviewAllData, projectOverviewEmailRowData;
  let numProjectAssigneeEmails = 0;
  try {
    const integrationRange = integrationSheet.getDataRange();
    integrationData = (integrationRange.getNumRows() > 1) ? integrationRange.offset(1, 0, integrationRange.getNumRows() - 1).getValues() : [];
    taskOverviewAllData = taskOverviewSheet.getDataRange().getValues();
    const projSheetLastRow = projectOverviewSheet.getLastRow(); const projSheetLastCol = projectOverviewSheet.getLastColumn(); Logger.log(`診断: プロジェクト状況概観 - LastRow: ${projSheetLastRow}, LastCol: ${projSheetLastCol}`);
    if (projSheetLastRow >= PROJECT_OVERVIEW_EMAIL_ROW && projSheetLastCol >= PRJ_OVERVIEW_FIRST_DATA_COL) {
        const emailRowValues = projectOverviewSheet.getRange(PROJECT_OVERVIEW_EMAIL_ROW, PRJ_OVERVIEW_FIRST_DATA_COL, 1, projSheetLastCol - PRJ_OVERVIEW_FIRST_DATA_COL + 1).getValues()[0];
        projectOverviewEmailRowData = emailRowValues; numProjectAssigneeEmails = emailRowValues.filter(String).length; Logger.log(`診断: プロジェクト状況概観 - メールアドレス行から推定される有効な列数: ${numProjectAssigneeEmails}`);
    } else { projectOverviewEmailRowData = []; numProjectAssigneeEmails = 0; Logger.log(`警告: プロジェクト状況概観のメールアドレス行(1行目 C列以降)が存在しないかシートが小さすぎます。`); }
    // !!! プロジェクト状況概観シートからプロジェクトリストを読み込む処理は不要になる !!!
  } catch (e) { Logger.log(`エラー: データ読み込み中に失敗しました。 ${e}`); return; }

  // --- 集計用データ構造初期化 ---
  const taskStatusSummary = {};
  const projectData = {}; // ★ { プロジェクト名: { category: 'カテゴリ名', counts: { email1: 0, ... } } }

  // --- タスク状況概観の担当者リストを初期化 ---
  const taskOverviewUsers = []; const firstDataRowIndex_Task = TASK_OVERVIEW_FIRST_DATA_ROW - 1;
  for (let i = firstDataRowIndex_Task; i < taskOverviewAllData.length; i++) { const assigneeName = taskOverviewAllData[i][TASK_OVERVIEW_NAME_COL - 1]; if (assigneeName && typeof assigneeName === 'string' && assigneeName.trim() !== "") { const trimmedAssignee = assigneeName.trim(); if (!taskStatusSummary[trimmedAssignee]) { taskStatusSummary[trimmedAssignee] = { [STATUS_NOT_STARTED]: 0, [STATUS_IN_PROGRESS]: 0, [FINAL_STATUSES[0]]: 0, [STATUS_DELAYED]: 0, '総計': 0 }; taskOverviewUsers.push(trimmedAssignee); } } }
  Logger.log(`タスク状況概観シートから ${taskOverviewUsers.length} 人の担当者を認識しました。`);

  // --- 統合データを走査して集計 ---
  if (integrationData.length > 0) {
    Logger.log(`統合シートから ${integrationData.length} 行のデータを処理します...`);
    integrationData.forEach((row, index) => {
      const category = (row[INTEG_CATEGORY_COL_IDX] || "").trim(); // ★ カテゴリ取得
      const projectName = (row[INTEG_PROJECT_COL_IDX] || "").trim();
      const assigneeName = (row[INTEG_MANAGER_NAME_COL_IDX] || "").trim();
      const assigneeEmail = (row[INTEG_EMAIL_COL_NUM - 1] || "").trim();
      let status = (row[INTEG_STATUS_COL_IDX] || "");

      // タスク状況概観集計
      if (assigneeName && taskStatusSummary[assigneeName]) { let counted = false; if (status.trim() === STATUS_NOT_STARTED || status === "") { taskStatusSummary[assigneeName][STATUS_NOT_STARTED]++; taskStatusSummary[assigneeName]['総計']++; counted = true; } if (!counted) { const ts = status.trim(); taskStatusSummary[assigneeName]['総計']++; if (ts === STATUS_IN_PROGRESS) taskStatusSummary[assigneeName][STATUS_IN_PROGRESS]++; else if (ts === FINAL_STATUSES[0]) taskStatusSummary[assigneeName][FINAL_STATUSES[0]]++; else if (ts === STATUS_DELAYED) taskStatusSummary[assigneeName][STATUS_DELAYED]++; } }

      // プロジェクト状況概観集計
      if (projectName && assigneeEmail) {
        // プロジェクト情報を初期化 (カテゴリ含む)
        if (!projectData[projectName]) {
          projectData[projectName] = { category: category, counts: {} }; // ★ カテゴリも保存
        }
        // 担当者カウントを初期化・加算
        if (!projectData[projectName].counts[assigneeEmail]) {
          projectData[projectName].counts[assigneeEmail] = 0;
        }
        projectData[projectName].counts[assigneeEmail]++;
      }
    });
    Logger.log("集計処理完了。");
  } else { Logger.log("統合シートに処理対象データがありません。"); }


  // --- 集計結果を書き込むための配列準備 ---
  // 1. タスク状況概観用 (変更なし)
  const taskOverviewOutput = []; const numTaskOverviewUsers = taskOverviewUsers.length;
  for (let i = 0; i < numTaskOverviewUsers; i++) { const userName = taskOverviewUsers[i]; const summary = taskStatusSummary[userName]; taskOverviewOutput.push([ (summary?.[STATUS_NOT_STARTED] || 0), (summary?.[STATUS_IN_PROGRESS] || 0), (summary?.[FINAL_STATUSES[0]] || 0), (summary?.[STATUS_DELAYED] || 0), (summary?.['総計'] || 0) ]); }

  // 2. プロジェクト状況概観用 (★ソート処理を変更★)
  const projectOverviewOutput = [];
  // ★★★ プロジェクト名とカテゴリを含む配列を作成 ★★★
  let projectListForSorting = [];
  for (const projectName in projectData) {
      projectListForSorting.push({
          name: projectName,
          category: projectData[projectName].category || "" // カテゴリがない場合は空文字扱い
      });
  }

  // ★★★ カテゴリ名 -> プロジェクト名でソート ★★★
  projectListForSorting.sort((a, b) => {
      // カテゴリで比較 (空文字は先頭に来るように調整も可能だが、標準のlocaleCompareを使用)
      const categoryCompare = (a.category || "").localeCompare(b.category || "");
      if (categoryCompare !== 0) {
          return categoryCompare; // カテゴリが違えばそれで決定
      }
      // カテゴリが同じならプロジェクト名で比較
      return (a.name || "").localeCompare(b.name || "");
  });

  const numProjectsToWrite = projectListForSorting.length; // ★ソート後のリストの長さ
  const numColsToWriteProjectData = numProjectAssigneeEmails; // C列以降のデータ列数
  Logger.log(`診断: プロジェクト状況概観 - 書き込むプロジェクト数 (ソート後): ${numProjectsToWrite}`);

   if (numProjectsToWrite > 0 && numColsToWriteProjectData > 0) {
      // ★★★ ソート済みリストでループ ★★★
      projectListForSorting.forEach(projectInfo => {
          const projectName = projectInfo.name;
          const category = projectInfo.category;
          const projAggData = projectData[projectName]; // 元の集計データを取得
          const counts = projAggData.counts;
          const outputRow = [category, projectName]; // A列(カテゴリ), B列(プロジェクト名)
          const countDataForRow = new Array(numColsToWriteProjectData).fill(0);

          for (let j = 0; j < numColsToWriteProjectData; j++) { // メールアドレス列数分ループ
              const headerEmail = (projectOverviewEmailRowData[j] || "").trim(); // projectOverviewEmailRowDataは読み込み済み
              if (headerEmail && counts[headerEmail] !== undefined) {
                  countDataForRow[j] = counts[headerEmail];
              }
          }
          outputRow.push(...countDataForRow); // C列以降のカウントデータを追加
          projectOverviewOutput.push(outputRow);
      });
      // --- 総計行は不要 ---
      Logger.log(`診断: プロジェクト状況概観 - 書き込み用配列 projectOverviewOutput の総行数: ${projectOverviewOutput.length}`);
  } else {
       Logger.log("診断: プロジェクト状況概観 - プロジェクト数またはメールアドレス列数が0のため、集計データは作成されません。");
  }


  // --- シートに書き込み ---
  try {
    SpreadsheetApp.flush();

    // --- 1. タスク状況概観 --- (変更なし)
    if (numTaskOverviewUsers > 0) { const rowsToWrite = taskOverviewOutput.length; const taskDataRange = taskOverviewSheet.getRange( TASK_OVERVIEW_FIRST_DATA_ROW, TASK_OVERVIEW_NOT_STARTED_COL, rowsToWrite, TASK_OVERVIEW_NUM_STATUS_COLS ); Logger.log(`タスク状況概観 書き込み範囲 (データ): ${taskDataRange.getA1Notation()}`); taskDataRange.clear({contentsOnly: true}); if (taskOverviewOutput.length === rowsToWrite) { taskDataRange.setValues(taskOverviewOutput); Logger.log(`タスク状況概観シート データ更新完了 (${rowsToWrite}行)`); } else { Logger.log(`警告: タスク状況概観 書き込み配列行数(${taskOverviewOutput.length}) != ユーザー数(${numTaskOverviewUsers})`); } } else { Logger.log("タスク状況概観: 書き込むべきユーザーデータなし。既存データ範囲をクリアします。"); const lastDataRow = taskOverviewSheet.getLastRow(); if(lastDataRow >= TASK_OVERVIEW_FIRST_DATA_ROW) { const clearRange = taskOverviewSheet.getRange(TASK_OVERVIEW_FIRST_DATA_ROW, TASK_OVERVIEW_NOT_STARTED_COL, lastDataRow - TASK_OVERVIEW_FIRST_DATA_ROW + 1, TASK_OVERVIEW_NUM_STATUS_COLS); Logger.log(`タスク状況概観 クリア範囲 (データ無): ${clearRange.getA1Notation()}`); clearRange.clear({contentsOnly: true}); } }
    if (EXPECTED_TASK_OVERVIEW_HEADERS.length === TASK_OVERVIEW_NUM_STATUS_COLS) { const taskHeaderRange = taskOverviewSheet.getRange(TASK_OVERVIEW_HEADER_ROW, TASK_OVERVIEW_NOT_STARTED_COL, 1, TASK_OVERVIEW_NUM_STATUS_COLS); taskHeaderRange.setValues([EXPECTED_TASK_OVERVIEW_HEADERS]); Logger.log(`タスク状況概観シート 固定ヘッダー書き込み完了: ${EXPECTED_TASK_OVERVIEW_HEADERS}`); } else { Logger.log(`警告: EXPECTED_TASK_OVERVIEW_HEADERSの列数(${EXPECTED_TASK_OVERVIEW_HEADERS.length})が期待値(${TASK_OVERVIEW_NUM_STATUS_COLS})と異なります。`); }

// --- 2. プロジェクト状況概観 ---
    // a) データ書き込み (A3以降、ソート済み)
    const numRowsToWriteProject = projectOverviewOutput.length;
    const numColsToWriteTotal = 2 + numColsToWriteProjectData; // A, B + データ列数
    if (numRowsToWriteProject > 0 && numColsToWriteProjectData > 0) {
        // クリア範囲: A3から C列以降の最後の有効列まで、十分な行数
        const clearStartCol = 1; // A列から
        const clearNumCols = projectOverviewSheet.getMaxColumns();
        const clearStartRow = PROJECT_OVERVIEW_FIRST_DATA_ROW;
        const clearNumRows = projectOverviewSheet.getMaxRows() - clearStartRow + 1;
        if (clearNumRows > 0) { const clearRange = projectOverviewSheet.getRange(clearStartRow, clearStartCol, clearNumRows, clearNumCols); Logger.log(`プロジェクト状況概観 クリア範囲 (データ全体): ${clearRange.getA1Notation()}`); clearRange.clear({contentsOnly: true}); } else { Logger.log("プロジェクト状況概観: クリア対象行なし"); }

        // 書き込み範囲: A3から C列以降の最後の有効列まで、書き込む行数分
        const projectWriteRange = projectOverviewSheet.getRange( PROJECT_OVERVIEW_FIRST_DATA_ROW, 1, // A3から
                                                                numRowsToWriteProject, numColsToWriteTotal ); // A,B列 + データ列数
        Logger.log(`プロジェクト状況概観 書き込み範囲 (データ): ${projectWriteRange.getA1Notation()}`);
        if (projectOverviewOutput.length === projectWriteRange.getNumRows()) { projectWriteRange.setValues(projectOverviewOutput); Logger.log(`プロジェクト状況概観シート データ更新完了 (${numRowsToWriteProject}行 x ${numColsToWriteTotal}列)`); }
        else { Logger.log(`★★★ 重大な警告: プロジェクト状況概観 - 書き込み配列の行数(${projectOverviewOutput.length})と書き込み範囲の行数(${projectWriteRange.getNumRows()})が不一致です！`); }
    } else { /* データ無しの場合のクリア */
       Logger.log("プロジェクト状況概観: 書き込むべきプロジェクトまたはメールアドレスデータなし。既存データ範囲をクリアします。");
       const lastDataRow = projectOverviewSheet.getLastRow(); const lastDataCol = projectOverviewSheet.getLastColumn();
       if (lastDataRow >= PROJECT_OVERVIEW_FIRST_DATA_ROW) { const clearRange = projectOverviewSheet.getRange(PROJECT_OVERVIEW_FIRST_DATA_ROW, 1, lastDataRow - PROJECT_OVERVIEW_FIRST_DATA_ROW + 1, projectOverviewSheet.getMaxColumns()); Logger.log(`プロジェクト状況概観 クリア範囲 (データ無): ${clearRange.getA1Notation()}`); clearRange.clear({contentsOnly: true}); }
    }

    // b) ヘッダー処理 (C2以降) - UIの変更を保持する試み (変更なし)
    if (numProjectAssigneeEmails > 0) { const projectHeaderRange = projectOverviewSheet.getRange(PROJECT_OVERVIEW_HEADER_ROW, PRJ_OVERVIEW_FIRST_DATA_COL, 1, numProjectAssigneeEmails); try { Logger.log(`プロジェクト状況概観 処理範囲 (UI維持ヘッダー): ${projectHeaderRange.getA1Notation()}`); const currentHeaderValues = projectHeaderRange.getValues(); Logger.log(`プロジェクト状況概観 読込直後ヘッダー (UI維持試行): ${JSON.stringify(currentHeaderValues)}`); projectHeaderRange.setValues(currentHeaderValues); Logger.log(`プロジェクト状況概観シート UI維持ヘッダー書き戻し完了`); } catch (e) { Logger.log(`エラー: プロジェクト状況概観のUI維持ヘッダー処理中にエラーが発生しました。 ${e}`); } } else { Logger.log(`プロジェクト状況概観: 有効なメールアドレス列がないため、ヘッダー処理はスキップされました。`); }

  } catch (e) { Logger.log(`エラー: 集計結果の書き込み中に失敗しました。 ${e}\n${e.stack}`); }

  Logger.log("概観シート更新処理 完了");
}

/**
 * ヘルパー関数: 概観シートの集計カウント部分をクリアする
 */
function clearOverviewCounts(taskOverviewSheet, projectOverviewSheet, taskOverviewData, projectOverviewHeaderData, projectOverviewProjectData) {
  try {
      // タスク状況概観クリア (D2以降)
      const taskLastRow = taskOverviewSheet.getLastRow();
      if (taskLastRow > 1) {
           taskOverviewSheet.getRange(
               2, TASK_OVERVIEW_NOT_STARTED_COL,
               taskLastRow - 1,
               (TASK_OVERVIEW_TOTAL_COL - TASK_OVERVIEW_NOT_STARTED_COL + 1)
           ).clearContent();
      }
      Logger.log("タスク状況概観シートの集計値をクリアしました。");

      // プロジェクト状況概観クリア (C2以降)
       const projectLastRow = projectOverviewSheet.getLastRow();
       const projectLastCol = projectOverviewSheet.getLastColumn();
       if (projectLastRow > 1 && projectLastCol > PRJ_OVERVIEW_PROJECT_COL) {
          projectOverviewSheet.getRange(
             2, PRJ_OVERVIEW_PROJECT_COL + 1,
             projectLastRow - 1,
             projectLastCol - PRJ_OVERVIEW_PROJECT_COL
          ).clearContent();
       }
      Logger.log("プロジェクト状況概観シートの集計値をクリアしました。");

  } catch (e) {
      Logger.log(`エラー: 概観シートのクリア中に失敗: ${e}`);
  }
}

// ========================================================================
// == 4. トリガー設定 関数 ================================================
// ========================================================================

/** トリガー設定: タスク状況ダイジェスト通知 */
function setupDailyDigestTrigger() {
  const functionName = 'checkTasksAndNotifyDigest'; deleteTriggersByName(functionName); deleteTriggersByName('checkOverdueTasks'); // 古い関数名のトリガーも削除
  try { ScriptApp.newTrigger(functionName).timeBased().everyDays(1).atHour(9).inTimezone(Session.getScriptTimeZone()).create(); Logger.log(`トリガー設定完了: ${functionName} (毎日9時台)`); } catch (e) { Logger.log(`トリガー設定失敗: ${e}`); }
}

/** トリガー設定: 統合シート更新 */
function setupHourlyIntegrationTrigger() { // ★実行頻度(everyHours(1))は要調整
  const functionName = 'integrateSheetsImproved'; deleteTriggersByName(functionName);
  try { ScriptApp.newTrigger(functionName).timeBased().everyMinutes(1).create(); Logger.log(`トリガー設定完了: ${functionName} (1時間ごと)`); } catch (e) { Logger.log(`トリガー設定失敗: ${e}`); }
}
// --- 必要に応じて 10分ごと、15分ごとなどの設定関数もここに用意 ---
/*
function setup10MinuteIntegrationTrigger() { ... everyMinutes(10) ... }
function setup1MinuteIntegrationTrigger() { ... everyMinutes(1) ... }
*/


// ========================================================================
// == 5. 共通 トリガー管理 ユーティリティ ===================================
// ========================================================================

/** 指定関数名のトリガーを削除 */
function deleteTriggersByName(functionName) {
  try { const triggers = ScriptApp.getProjectTriggers(); let deletedCount = 0;
    for (const trigger of triggers) { if (trigger.getHandlerFunction() === functionName) { try { ScriptApp.deleteTrigger(trigger); deletedCount++; Logger.log(`トリガー削除成功: ${trigger.getUniqueId()} (${functionName})`); } catch (e) { Logger.log(`トリガー削除失敗: ${trigger.getUniqueId()} (${functionName}) ${e}`); }}}
  } catch (e) { Logger.log(`トリガー削除処理中エラー: ${e}`); }
}

/** 全トリガー情報をログ出力 */
function logAllTriggers() {
  try { const triggers = ScriptApp.getProjectTriggers(); if (triggers.length === 0) { Logger.log('トリガーなし'); return; } Logger.log(`設定済トリガー (${triggers.length}個):`);
    triggers.forEach((t, i) => { Logger.log(` [${i + 1}] ID:${t.getUniqueId()}, Func:${t.getHandlerFunction()}, Event:${t.getEventType()}, Src:${t.getTriggerSource()}`); });
  } catch (e) { Logger.log(`トリガー情報取得エラー: ${e}`); }
}


