/**
 * GAS v2 - 純粋 JSON API エンジン
 * UI(HTML)は一切返却せず、ContentService を通じて JSON のみを応答する。
 */

// =============================
// ⓪ 基本設定
// =============================
function getSS() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    if (ss && ss.getId()) return ss;
  } catch (e) {
    // Fallback if not container-bound
  }
  const SPREADSHEET_ID = '1KuA5pN0ItODhwSJph-fwgj_U_ZyHrn9Osew92D99xBs';
  if (!SPREADSHEET_ID) throw new Error("SPREADSHEET_ID is not defined.");
  return SpreadsheetApp.openById(SPREADSHEET_ID);
}

// =============================
// ① APIエントリポイント
// =============================

/**
 * GETリクエスト：JSONデータの取得
 */
function doGet(e) {
  const action = e.parameter.action;
  let response;

  try {
    switch (action) {
      case 'getAppData':
        response = getAppData();
        break;
      case 'getRoster':
        response = getRoster();
        break;
      case 'getAreaDetails':
        response = getAreaDetails(e.parameter.name);
        break;
      default:
        response = { success: true, message: 'POSTING MAP API is online.' };
    }
  } catch (err) {
    response = { success: false, message: err.toString() };
  }

  return createJsonResponse(response);
}

/**
 * POSTリクエスト：データの登録・更新
 */
function doPost(e) {
  let postData;
  try {
    if (e.postData && e.postData.contents) {
      postData = JSON.parse(e.postData.contents);
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
        response = { success: false, message: 'Invalid POST action' };
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

function getAppData() {
  const ss = getSS();
  const sheets = ss.getSheets();
  
  const systemSheets = [
    CONFIG.SHEET_GUIDE, CONFIG.SHEET_ROSTER, CONFIG.SHEET_TEMPLATE, 
    CONFIG.SHEET_POSTAL, CONFIG.SHEET_DISTRICT, CONFIG.SHEET_MASTER_EXPORT, 
    CONFIG.SHEET_REPORT, CONFIG.SHEET_MANUAL, CONFIG.SHEET_SYSTEM_CACHE
  ];

  const areas = [];
  sheets.forEach(s => {
    const name = s.getName();
    if (!systemSheets.includes(name) && !s.isSheetHidden()) {
      // 重い計算は一旦やめて、0%として爆速で返す（後ほど非同期化などで対応）
      areas.push({
        name: name,
        progress: 0 
      });
    }
  });

  return {
    success: true,
    branchName: ss.getName().split(/[ 　]/)[0] || "支部",
    areas: areas
  };
}

function getAreaDetails(areaName) {
  if (!areaName) return { success: false, message: "Area name required" };
  const s = getSS().getSheetByName(areaName);
  if (!s) return { success: false, message: "Area not found" };

  const lastRow = s.getLastRow();
  if (lastRow < 2) return { success: true, points: [] };

  const values = s.getRange(2, 1, lastRow - 1, 6).getValues();
  const points = values.map((r, i) => ({
    rowId: i + 2,
    address: r[0],
    memo: r[2],
    isDone: r[3] === true || r[3] === "TRUE",
    staffName: r[4]
  }));

  return { success: true, points: points };
}

function getRoster() {
  const s = getSS().getSheetByName(CONFIG.SHEET_ROSTER);
  if (!s) return [];
  const lastRow = s.getLastRow();
  if (lastRow < 2) return [];
  return s.getRange(2, 1, lastRow - 1, 2).getValues().map(r => ({ id: r[0], name: r[1] }));
}

function submitDistribution(areaName, rowId, staffName, count, isDone, staffId) {
  const ss = getSS();
  const s = ss.getSheetByName(areaName);
  if (!s) return { success: false, message: "Sheet not found" };

  const now = new Date();
  s.getRange(rowId, 4, 1, 3).setValues([[isDone, staffName, now]]);
  return { success: true };
}

function registerStaff(lastName, firstName) {
  const ss = getSS();
  const s = ss.getSheetByName(CONFIG.SHEET_ROSTER);
  if (!s) return { success: false };

  const name = lastName + " " + firstName;
  const lastRow = s.getLastRow();
  const newId = "S" + Utilities.formatDate(new Date(), "JST", "yyyyMMddHHmmss");
  s.appendRow([newId, name, new Date()]);
  
  return { success: true, id: newId, name: name };
}
