/**
 * GAS v2 - API専用エンジン
 */

// =============================
// ⓪ 基本設定
// =============================
const SPREADSHEET_ID = '1KuA5pN0ItODhwSJph-fwgj_U_ZyHrn9Osew92D99xBs';

function getSS() {
  return SpreadsheetApp.openById(SPREADSHEET_ID);
}

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
    if (e.postData && e.postData.contents) {
      postData = JSON.parse(e.postData.contents);
    } else if (e.parameter.json) {
      postData = JSON.parse(e.parameter.json);
    } else {
      postData = e.parameter;
    }
  } catch (f) {
    postData = e.parameter;
  }

  const action = postData.action || e.parameter.action;
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
  const ss = getSS();
  console.log("Testing Connection to SS: " + ss.getName());
  
  const guideSheet = ss.getSheetByName(CONFIG.SHEET_GUIDE);
  console.log("Guide Sheet Found: " + (guideSheet ? "Yes" : "No (" + CONFIG.SHEET_GUIDE + ")"));
  
  let totalDistributed = 0;
  if (guideSheet) {
    totalDistributed = guideSheet.getRange("H5").getValue();
    console.log("Total Distributed (H5): " + totalDistributed);
  }

  const exclude = [
    CONFIG.SHEET_GUIDE, CONFIG.SHEET_ROSTER, CONFIG.SHEET_TEMPLATE,
    CONFIG.SHEET_POSTAL, CONFIG.SHEET_DISTRICT, CONFIG.SHEET_MASTER_EXPORT,
    CONFIG.SHEET_REPORT, CONFIG.SHEET_MANUAL,
  ];
  
  const allSheets = ss.getSheets();
  console.log("Total Sheets in SS: " + allSheets.length);

  const areas = allSheets
    .filter((s) => !exclude.includes(s.getName()) && !s.isSheetHidden())
    .map((s) => {
      const lastRow = s.getLastRow();
      let done = 0, total = 0;
      if (lastRow >= 2) {
        const data = s.getRange(2, 4, lastRow - 1, 1).getValues();
        total = data.length;
        done = data.filter((r) => r[0] === true).length;
      }
      console.log("Area Processed: " + s.getName() + " (Total: " + total + ", Done: " + done + ")");
      return {
        name: s.getName(),
        progress: total > 0 ? Math.round((done / total) * 100) : 0,
        count: total,
      };
    });

  const roster = getRoster();
  console.log("Staff Roster Count: " + (roster ? roster.length : 0));

  return {
    branchName: ss.getName().split(/[ 　]/)[0] || "支部",
    totalDistributed: totalDistributed,
    targetGoal: CONFIG.TARGET_GOAL,
    areas: areas,
    staffList: roster,
  };
}

/**
 * エリア詳細情報の取得
 */
function getAreaDetails(areaName) {
  const s = getSS().getSheetByName(areaName);
  if (!s) return { success: false, message: "Area not found" };

  const lastRow = s.getLastRow();
  if (lastRow < 2) return { success: true, data: [] };

  const values = s.getRange(2, 1, lastRow - 1, 6).getValues();
  const data = values.map((r, i) => ({
    id: i + 2,
    town: r[0],
    street: r[1],
    houseCount: r[2],
    isDone: r[3],
    staffName: r[4],
    timestamp: r[5]
  }));

  return { success: true, data: data };
}

/**
 * 名簿の取得
 */
function getRoster() {
  const s = getSS().getSheetByName(CONFIG.SHEET_ROSTER);
  if (!s) return [];
  const lastRow = s.getLastRow();
  if (lastRow < 2) return [];
  return s.getRange(2, 1, lastRow - 1, 2).getValues().map(r => ({ id: r[0], name: r[1] }));
}

/**
 * 進捗報告の登録
 */
function submitDistribution(areaName, rowId, staffName, count, isDone, staffId) {
  const ss = getSS();
  const s = ss.getSheetByName(areaName);
  if (!s) return { success: false, message: "Sheet not found: " + areaName };

  const now = new Date();
  s.getRange(rowId, 4, 1, 3).setValues([[isDone, staffName, now]]);
  
  // ログ記録
  const reportSheet = ss.getSheetByName(CONFIG.SHEET_REPORT);
  if (reportSheet) {
    reportSheet.appendRow([now, areaName, rowId, staffName, count, isDone, staffId]);
  }

  return { success: true };
}

/**
 * 配布員の新規登録
 */
function registerStaff(lastName, firstName) {
  const ss = getSS();
  const s = ss.getSheetByName(CONFIG.SHEET_ROSTER);
  if (!s) return { success: false, message: "Roster sheet not found" };

  const name = lastName + " " + firstName;
  const lastRow = s.getLastRow();
  
  // 重複チェック
  const existing = s.getRange(2, 2, lastRow > 1 ? lastRow - 1 : 1, 1).getValues();
  if (existing.some(r => r[0] === name)) {
    // 既存ユーザーを返す
    const rowIndex = existing.findIndex(r => r[0] === name);
    return { success: true, id: s.getRange(rowIndex + 2, 1).getValue(), name: name };
  }

  const newId = "S" + Utilities.formatDate(new Date(), "JST", "yyyyMMddHHmmss");
  s.appendRow([newId, name, new Date()]);
  
  return { success: true, id: newId, name: name };
}
