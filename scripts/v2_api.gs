/**
 * GAS v2 - API専用エンジン
 */

// =============================
// ① APIエントリポイント
// =============================

// GETリクエスト：データの取得
function doGet(e) {
  const action = e.parameter.action;
  let response;

  try {
    switch (action) {
      case 'getAppData':
        response = getAppData();
        break;
      case 'getAreaDetails':
        response = getAreaDetails(e.parameter.areaName);
        break;
      case 'getRoster':
        response = getRoster();
        break;
      default:
        response = { success: false, message: 'Invalid GET action: ' + action };
    }
  } catch (err) {
    response = { success: false, message: err.toString() };
  }

  return createJsonResponse(response);
}

// POSTリクエスト：データの登録・更新
function doPost(e) {
  let postData;
  try {
    postData = JSON.parse(e.postData.contents);
  } catch (f) {
    postData = e.parameter;
  }

  const action = postData.action;
  let response;

  try {
    switch (action) {
      case 'submitDistribution':
        response = submitDistribution(
          postData.areaName,
          postData.rowId,
          postData.staffName,
          postData.count,
          postData.isDone,
          postData.staffId
        );
        break;
      case 'registerStaff':
        response = registerStaff(postData.lastName, postData.firstName);
        break;
      default:
        response = { success: false, message: 'Invalid POST action: ' + action };
    }
  } catch (err) {
    response = { success: false, message: err.toString() };
  }

  return createJsonResponse(response);
}

// 共通：JSONレスポンス作成
function createJsonResponse(data) {
  return ContentService.createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
}

// =============================
// ② データ取得ロジック
// =============================

/**
 * モバイルアプリ用：全体サマリー取得
 */
function getAppData() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const guideSheet = ss.getSheetByName(CONFIG.SHEET_GUIDE);
  let totalDistributed = 0;

  if (guideSheet) {
    totalDistributed = guideSheet.getRange("H5").getValue();
    if (typeof totalDistributed === "string") {
      totalDistributed =
        parseInt(totalDistributed.split("/")[0].replace(/,/g, "")) || 0;
    }
  }

  const exclude = [
    CONFIG.SHEET_GUIDE,
    CONFIG.SHEET_ROSTER,
    CONFIG.SHEET_TEMPLATE,
    CONFIG.SHEET_POSTAL,
    CONFIG.SHEET_DISTRICT,
    CONFIG.SHEET_MASTER_EXPORT,
    CONFIG.SHEET_REPORT,
    CONFIG.SHEET_MANUAL,
  ];
  const areas = ss
    .getSheets()
    .filter((s) => !exclude.includes(s.getName()) && !s.isSheetHidden())
    .map((s) => {
      const lastRow = s.getLastRow();
      let done = 0;
      let total = 0;
      if (lastRow >= 2) {
        const data = s.getRange(2, 4, lastRow - 1, 1).getValues();
        total = data.length;
        done = data.filter((r) => r[0] === true).length;
      }
      return {
        name: s.getName(),
        progress: total > 0 ? Math.round((done / total) * 100) : 0,
        count: total,
      };
    });

  return {
    branchName: ss.getName().split(/[ 　]/)[0] || "支部",
    totalDistributed: totalDistributed,
    targetGoal: CONFIG.TARGET_GOAL,
    areas: areas,
    staffList: getRoster(),
  };
}

/**
 * モバイルアプリ用：エリア詳細取得（爆速キャッシュ版）
 */
function getAreaDetails(areaName) {
  const cache = CacheService.getScriptCache();
  const cached = cache.get("DETAILS_" + areaName);
  if (cached) return JSON.parse(cached);

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(areaName);
  if (!sheet) return [];

  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];

  const data = sheet.getRange(2, 1, lastRow - 1, 8).getValues();
  const result = data
    .filter((row) => row[0] !== "")
    .slice(0, 100)
    .map((row, index) => ({
      id: index + 2,
      address: row[0],
      isDone: row[3] === true,
      count: row[5],
      staff: row[6],
      memo: row[2],
    }));

  // 25分間キャッシュ（詳細データは頻繁に変わるため短め）
  cache.put("DETAILS_" + areaName, JSON.stringify(result), 1500);
  return result;
}

/**
 * モバイルアプリ用：配布報告
 */
function submitDistribution(
  areaName,
  rowId,
  staffName,
  count,
  isDone,
  staffId,
) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(areaName);
  if (!sheet) throw new Error("Area not found");

  const now = Utilities.formatDate(new Date(), "JST", "MM/dd HH:mm");

  // キャッシュ更新用の差分計算
  const oldStatus = sheet.getRange(rowId, 4).getValue();
  const isDoneChanged = isDone === oldStatus ? 0 : isDone ? 1 : -1;

  // D列:完了, E列:日時, F列:枚数, G列:名前, H列:ID を一括書き込み
  sheet
    .getRange(rowId, 4, 1, 5)
    .setValues([[isDone, isDone ? now : "", count, staffName, staffId || ""]]);

  updateSheetSummary(sheet);
  updateAreaCache(areaName, isDoneChanged);

  // 詳細データのキャッシュを削除（次回アクセス時に最新を読み込むため）
  CacheService.getScriptCache().remove("DETAILS_" + areaName);

  return { success: true, timestamp: now };
}

function getRoster() {
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(
      CONFIG.SHEET_ROSTER,
    );
    if (!sheet) return [];
    const lastRow = sheet.getLastRow();
    if (lastRow < 2) return [];

    const data = sheet.getRange(2, 1, lastRow - 1, 3).getValues();
    return data
      .map((r) => {
        const idStr = ("000" + r[0]).slice(-3); // 001 形式に整形
        return {
          id: r[0],
          displayName: `${idStr}  ${r[1]} ${r[2]}`.trim(), // スペース2つ
          lastName: r[1],
          firstName: r[2],
        };
      })
      .filter((item) => item.lastName);
  } catch (e) {
    return [];
  }
}

/**
 * 新規スタッフ登録（ロック機能付き）
 */
function registerStaff(lastName, firstName) {
  if (!lastName || !firstName)
    return { success: false, message: "苗字と名前の両方を入力してください" };

  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(30000);
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(CONFIG.SHEET_ROSTER);

    // ID発行ロジック
    const data = sheet.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      if (data[i][1] === lastName && data[i][2] === firstName) {
        const existingId = data[i][0];
        const idStr = ("000" + existingId).slice(-3);
        lock.releaseLock();
        return {
          success: true,
          id: existingId,
          name: `${lastName} ${firstName}`,
          displayName: `${idStr}  ${lastName} ${firstName}`,
          message: "既存の登録情報を復元しました"
        };
      }
    }

    // 新規登録
    const nextId = data.length > 1 ? Math.max(...data.slice(1).map(r => r[0] || 0)) + 1 : 1;
    sheet.appendRow([nextId, lastName, firstName, new Date()]);
    
    lock.releaseLock();
    const idStr = ("000" + nextId).slice(-3);
    return {
      success: true,
      id: nextId,
      name: `${lastName} ${firstName}`,
      displayName: `${idStr}  ${lastName} ${firstName}`,
    };
  } catch (e) {
    if (lock.hasLock()) lock.releaseLock();
    return { success: false, message: "登録エラー: " + e.message };
  }
}

function isNotAdmin() {
  const user = Session.getActiveUser().getEmail();
  const owner = SpreadsheetApp.getActiveSpreadsheet().getOwner().getEmail();
  if (user !== owner) {
    SpreadsheetApp.getUi().alert("❌ 管理者専用です。");
    return true;
  }
  return false;
}

function showAppUrl() {
  const url = ScriptApp.getService().getUrl();
  const html = `<div style="padding:20px; text-align:center;">
    <p>モバイルアプリのURL:</p>
    <input type="text" value="${url}" style="width:100%; padding:10px;" readonly onclick="this.select()">
    <br><br>
    <a href="${url}" target="_blank" style="padding:10px 20px; background:#1a73e8; color:white; text-decoration:none; border-radius:5px;">アプリを開く</a>
  </div>`;
  SpreadsheetApp.getUi().showModalDialog(
    HtmlService.createHtmlOutput(html).setWidth(400).setHeight(200),
    "スマホアプリURL",
  );
}

function openMapDashboard() {
  const apiKey = PropertiesService.getScriptProperties().getProperty(
    "GOOGLE_MAPS_API_KEY",
  );
  if (!apiKey) {
    SpreadsheetApp.getUi().alert("APIキー未設定です。");
    return;
  }
  const html = HtmlService.createTemplateFromFile("map_dashboard");
  html.apiKey = apiKey;
  SpreadsheetApp.getUi().showModalDialog(
    html.evaluate().setWidth(1000).setHeight(700),
    "戦況マップ",
  );
}
