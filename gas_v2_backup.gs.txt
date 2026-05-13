/**
 * GAS v2（最強安定・統合版）
 * - 本番マスター(gas.gs)の量産ロジックを完全移植
 * - モバイルアプリ用APIを統合
 * - 大規模バッチ処理対応
 */

// ==========================================
// GLOBAL CONFIGURATION (gas.gs 準拠)
// ==========================================
const CONFIG = {
  VERSION: "2.2.0 (Mobile & Pro Integrated)",
  // ファイル名設定
  DISTRICT_CSV: "三重県選挙区区割り",
  POSTAL_CSV: "MIE_POSTAL.CSV",
  TURNOUT_CSV: "voter_turnout.csv",
  TARGET_GOAL: 300000, // 目標配布枚数（ダッシュボード用）

  // シート名設定
  SHEET_GUIDE: "初めての方「使い方ガイド」",
  SHEET_ROSTER: "名簿",
  SHEET_TEMPLATE: "原本",
  SHEET_POSTAL: "郵便番号",
  SHEET_DISTRICT: "区割り",
  SHEET_MASTER_EXPORT: "📥 集計用マスターデータ",
  SHEET_REPORT: "📄 活動報告書",
  SHEET_MANUAL: "📖 らくらくマニュアル",

  // 動作設定
  CHUNK_SIZE: 10,
  ROW_HEIGHT_STAFF: 60,
  DENOMINATOR_UNITS: 651, // 三重第2区の戦略ユニット総数

  // デフォルト対象
  DEFAULT_DISTRICT: "第2区",
  DEFAULT_PREFECTURE: "三重県",
};

// =============================
// ① HTML表示 & モバイルアプリAPI
// =============================

function doGet(e) {
  // JSONを要求された場合（マップダッシュボード等）
  if (e && e.parameter && e.parameter.data === "json") {
    return ContentService.createTextOutput(
      JSON.stringify(getMapData()),
    ).setMimeType(ContentService.MimeType.JSON);
  }
  // 通常はHTMLを表示
  return HtmlService.createTemplateFromFile("index")
    .evaluate()
    .setTitle("ポスティングアプリ")
    .addMetaTag(
      "viewport",
      "width=device-width, initial-scale=1.0, maximum-scale=1.0, user-scalable=no",
    );
}

function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

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

// =============================
// ② データ抽出 (gas.gs 完全移植)
// =============================

function extractDistrictAddresses(
  targetDistrictName = CONFIG.DEFAULT_DISTRICT,
  targetPrefecture = CONFIG.DEFAULT_PREFECTURE,
) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const districtData = getCsvOrSheetData(CONFIG.DISTRICT_CSV);
  if (!districtData) {
    ss.toast(
      `Google Drive上に「${CONFIG.DISTRICT_CSV}」が見つかりません。`,
      "エラー",
      5,
    );
    return [];
  }

  const targetRules = [];
  for (let i = 1; i < districtData.length; i++) {
    const row = districtData[i];
    if (row && row[0] === targetDistrictName && row[1] === targetPrefecture) {
      targetRules.push({ city: row[2], townArea: row[3] || "" });
    }
  }

  const postalData = getCsvOrSheetData(CONFIG.POSTAL_CSV);
  if (!postalData) {
    ss.toast(
      `Google Drive上に「${CONFIG.POSTAL_CSV}」が見つかりません。`,
      "エラー",
      5,
    );
    return [];
  }

  const addressMap = new Map();
  targetRules.forEach((rule) => {
    if (rule.townArea) {
      const addrString = rule.townArea.startsWith(rule.city)
        ? rule.townArea
        : rule.city + rule.townArea;
      let genericPostal = "";
      for (let i = 0; i < postalData.length; i++) {
        const r = postalData[i];
        if (
          r &&
          r[6] === targetPrefecture &&
          r[7] === rule.city &&
          r[8] === "以下に掲載がない場合"
        ) {
          const p = r[2] ? r[2].toString().trim() : "";
          if (p.length === 7) genericPostal = `${p.slice(0, 3)}-${p.slice(3)}`;
          break;
        }
      }
      addressMap.set(addrString, genericPostal);
    } else {
      postalData.forEach((row) => {
        if (row && row[6] === targetPrefecture && row[7] === rule.city) {
          const pCode = row[2] ? row[2].toString().trim() : "";
          const postalStr =
            pCode.length === 7
              ? `${pCode.slice(0, 3)}-${pCode.slice(3)}`
              : pCode;
          const townRaw = row[8];
          if (townRaw && townRaw !== "以下に掲載がない場合") {
            const expanded = expandTownChome(rule.city, townRaw);
            expanded.forEach((addr) => {
              if (!addressMap.has(addr) || addressMap.get(addr) === "") {
                addressMap.set(addr, postalStr);
              }
            });
          }
        }
      });
    }
  });
  return Array.from(addressMap, ([address, postalCode]) => ({
    postalCode,
    address,
  }));
}

function getCsvOrSheetData(filename) {
  const files = DriveApp.getFilesByName(filename);
  if (!files.hasNext()) return null;
  const file = files.next();
  const mime = file.getMimeType();
  if (mime === MimeType.GOOGLE_SHEETS) {
    const ss = SpreadsheetApp.open(file);
    return ss.getSheets()[0].getDataRange().getValues();
  } else {
    // UTF-8 or Shift-JIS 判別
    const blob = file.getBlob();
    let text;
    try {
      text = blob.getDataAsString("UTF-8");
      if (text.indexOf("\uFFFD") !== -1) throw new Error();
    } catch (e) {
      text = blob.getDataAsString("Shift_JIS");
    }
    try {
      return Utilities.parseCsv(text);
    } catch (e) {
      return text.split("\n").map((line) => line.split(","));
    }
  }
}

function expandTownChome(baseCity, townRaw) {
  if (!townRaw) return [];
  const match = townRaw.match(/（([０-９0-9]+)〜([０-９0-9]+)丁目）/);
  if (match) {
    const start = parseInt(toHalfWidth(match[1]), 10);
    const end = parseInt(toHalfWidth(match[2]), 10);
    const baseTown = townRaw.replace(/（.*?）/g, "");
    let list = [];
    for (let i = start; i <= end; i++) {
      list.push(baseCity + baseTown + i + "丁目");
    }
    return list;
  }
  return [baseCity + townRaw.replace(/（.*?）/g, "")];
}

function toHalfWidth(str) {
  return str
    .toString()
    .replace(/[０-９]/g, (s) => String.fromCharCode(s.charCodeAt(0) - 0xfee0));
}

// =============================
// ③ バッチ処理 (gas.gs 完全移植)
// =============================

function forceStartBatch() {
  const props = PropertiesService.getScriptProperties();
  props.setProperty("BATCH_STATUS", "running");
  props.setProperty("BATCH_INDEX", "0");
  props.setProperty("BATCH_CITY_COUNTS", JSON.stringify({}));
  generateAreaSheetsBatch();
}

function generateAreaSheetsBatch() {
  const props = PropertiesService.getScriptProperties();
  if (props.getProperty("BATCH_STATUS") !== "running") return;
  const startTime = new Date().getTime();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const baseSheet = ss.getSheetByName(CONFIG.SHEET_TEMPLATE);

  const addresses = extractDistrictAddresses();
  const startIndex = parseInt(props.getProperty("BATCH_INDEX")) || 0;
  const chunkSize = CONFIG.CHUNK_SIZE;

  // 1. 再開時の状態シミュレーション
  let cityCounts = {};
  let lastCity = "";
  let itemsInBlock = 0; // 1シート内の何件目か (0-9)

  for (let i = 0; i < startIndex; i++) {
    const c = extractCityName(addresses[i].address);
    if (c !== lastCity || itemsInBlock >= chunkSize) {
      cityCounts[c] = (cityCounts[c] || 0) + 1;
      itemsInBlock = 0;
      lastCity = c;
    }
    itemsInBlock++;
  }

  // 2. メインループ
  for (
    let currentIndex = startIndex;
    currentIndex < addresses.length;
    currentIndex++
  ) {
    const now = new Date().getTime();
    if (now - startTime > 300 * 1000) {
      // 5分制限
      props.setProperty("BATCH_INDEX", currentIndex.toString());
      // プロパティへの保存はBATCH_INDEXのみでOK（シミュレーションでcityCountsは復元可能）
      ScriptApp.newTrigger("generateAreaSheetsBatch")
        .timeBased()
        .after(1000 * 60)
        .create();
      ss.toast(`${currentIndex}件で中断。1分後に自動再開します。`, "中断", 5);
      return;
    }

    const currentAddr = addresses[currentIndex];
    const currentCity = extractCityName(currentAddr.address);

    // 市町村が変わった、または10件に達した場合
    if (currentCity !== lastCity || itemsInBlock >= chunkSize) {
      cityCounts[currentCity] = (cityCounts[currentCity] || 0) + 1;
      itemsInBlock = 0;
      lastCity = currentCity;
    }

    let sheetName =
      cityCounts[currentCity] === 1
        ? currentCity
        : `${currentCity}(${cityCounts[currentCity]})`;
    let sheet =
      ss.getSheetByName(sheetName) ||
      baseSheet.copyTo(ss).setName(sheetName).showSheet();

    // シートの初期化とデザイン適用（新しいシートの開始時のみ）
    if (itemsInBlock === 0) {
      sheet.getRange("A2:L11").clearContent(); // L列（通し番号）まで確実にクリア
      applyProDesign(sheet);
    }

    // 書き込み（絶対に行番号を指定：2〜11行目）
    const targetRow = itemsInBlock + 2;
    const displayAddress = currentAddr.postalCode
      ? `〒${currentAddr.postalCode}\n${currentAddr.address}`
      : currentAddr.address;

    sheet.getRange(targetRow, 1).setValue(displayAddress);
    const mapsUrl =
      "https://www.google.com/maps/search/?api=1&query=" +
      encodeURIComponent(currentAddr.address);
    sheet.getRange(targetRow, 2).setFormula(`=HYPERLINK("${mapsUrl}","📍")`);
    sheet.getRange(targetRow, 12).setValue(currentIndex + 2); // 元行番号（L列へ移動）

    itemsInBlock++;
  }

  // 完了処理
  props.deleteProperty("BATCH_STATUS");
  props.deleteProperty("BATCH_INDEX");
  ss.toast(
    "すべてのエリアシートの展開（市町村境界考慮・10件分割版）が完了しました！",
    "完了",
    10,
  );
  refreshAreaSummaryCache();
}

function createAddressLinks(targetSheet) {
  const sheet =
    targetSheet || SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return;
  const values = sheet.getRange(2, 1, lastRow - 1, 1).getValues();
  const formulas = values.map((v) => {
    let addr = v[0];
    if (!addr) return [""];
    if (addr.includes("\n")) addr = addr.split("\n")[1];
    const url =
      "https://www.google.com/maps/search/?api=1&query=" +
      encodeURIComponent(addr);
    return ['=HYPERLINK("' + url + '","📍")'];
  });
  sheet.getRange(2, 2, formulas.length, 1).setFormulas(formulas);
}

// =============================
// ④ その他ユーティリティ (gas.gs 準拠)
// =============================

/**
 * 【新・戦略エンジン】全体数と個人ランキングを集計する
 */
function aggregateTotalVolumes() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const guideSheet = ss.getSheetByName(CONFIG.SHEET_GUIDE);
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

  let totalUnitsDone = 0;
  let grandTotalVolume = 0;

  // 名簿から最新のID->名前マップを作成
  const rosterData = getRoster();
  const rosterMap = {};
  rosterData.forEach((r) => {
    rosterMap[r.id] = r.displayName;
  });

  let staffRanking = {}; // { IDまたは名前: 合計枚数 }

  const sheets = ss.getSheets();
  sheets.forEach((sheet) => {
    const name = sheet.getName();
    if (!exclude.includes(name) && !sheet.isSheetHidden()) {
      const lastRow = sheet.getLastRow();
      if (lastRow >= 2) {
        // D:完了, E:日時, F:枚数, G:スタッフ, H:ID
        const data = sheet.getRange(2, 4, lastRow - 1, 5).getValues();

        data.forEach((row) => {
          const isDone = row[0] === true;
          const count = parseFloat(row[2]) || 0;
          const staff = row[3];

          if (isDone) {
            totalUnitsDone++;
            grandTotalVolume += count;
            if (staff) {
              staffRanking[staff] = (staffRanking[staff] || 0) + count;
            }
          }
        });
      }
    }
  });

  // 進捗率の計算 (母数 651)
  const progressPercent = (totalUnitsDone / CONFIG.DENOMINATOR_UNITS) * 100;

  if (guideSheet) {
    // 1. 全体進捗表示 (H5セル) - パーセントのみ
    guideSheet
      .getRange("H5:K5")
      .merge()
      .setValue(`全体進捗: ${progressPercent.toFixed(1)}%`);

    // 2. 総配布枚数表示 (H6セル)
    guideSheet
      .getRange("H6:K6")
      .merge()
      .setValue(`総配布枚数: ${grandTotalVolume.toLocaleString()} 枚`);

    // 3. ランキング表示 (M列などに反映)
    const sortedRanking = Object.entries(staffRanking)
      .sort((a, b) => b[1] - a[1])
      .slice(0, 10);

    guideSheet.getRange("M10:O20").clearContent();
    guideSheet.getRange("M9").setValue("🏆 配布枚数ランキング");

    sortedRanking.forEach((entry, index) => {
      const row = 10 + index;
      guideSheet.getRange(row, 13).setValue(`${index + 1}位`);
      guideSheet.getRange(row, 14).setValue(entry[0]);
      guideSheet.getRange(row, 15).setValue(`${entry[1].toLocaleString()} 枚`);
    });

    ss.toast(`集計完了: 進捗 ${progressPercent.toFixed(1)}%`, "システム更新");
  }
}

/**
 * 名簿シートを初期化（ID・苗字・名前の3列構成にする）
 */
function setupRosterSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(CONFIG.SHEET_ROSTER);
  if (!sheet) {
    sheet = ss.insertSheet(CONFIG.SHEET_ROSTER);
  }

  // ヘッダーを設定
  sheet.getRange(1, 1, 1, 3).setValues([["ID", "苗字", "名前"]]);
  sheet
    .getRange(1, 1, 1, 3)
    .setBackground("#1a2533")
    .setFontColor("#ffffff")
    .setFontWeight("bold");
  sheet.setFrozenRows(1); // 1行目を固定

  return "名簿シートを初期化しました。";
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
          displayName: `${idStr} ${r[1]} ${r[2]}`.trim(),
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
    const lastRow = sheet.getLastRow();
    let nextId = 1;
    if (lastRow >= 2) {
      const ids = sheet.getRange(2, 1, lastRow - 1, 1).getValues();
      nextId = Math.max(...ids.map((r) => Number(r[0]) || 0)) + 1;
    }

    // 書き込み（ID, 苗字, 名前）
    sheet.appendRow([nextId, lastName, firstName]);

    lock.releaseLock();
    const idStr = ("000" + nextId).slice(-3);
    return {
      success: true,
      id: nextId,
      name: `${lastName} ${firstName}`,
      displayName: `${idStr} ${lastName} ${firstName}`,
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

/**
 * モバイルアプリ用：全体サマリー取得（爆速キャッシュ版）
 * 650枚のシートを毎回スキャンすると20秒かかるため、
 * PropertiesServiceに保存されたキャッシュを返す。
 */
function getMapData() {
  const props = PropertiesService.getScriptProperties();
  const cached = props.getProperty("AREA_SUMMARY_CACHE");

  if (cached) {
    try {
      return JSON.parse(cached);
    } catch (e) {
      // キャッシュ破損時は再計算へ
    }
  }

  // キャッシュがない場合は初回のみ計算
  return refreshAreaSummaryCache();
}

/**
 * 全エリアのサマリーを再計算してキャッシュに保存する
 */
function refreshAreaSummaryCache() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
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

  const summary = [];
  let totalDone = 0;
  let totalPoints = 0;

  ss.getSheets().forEach((s) => {
    const name = s.getName();
    if (!exclude.includes(name) && !s.isSheetHidden()) {
      const last = s.getLastRow();
      let done = 0;
      let total = 0;
      if (last >= 2) {
        const data = s.getRange(2, 1, last - 1, 4).getValues();
        const validRows = data.filter((r) => r[0] !== "");
        total = validRows.length;
        done = validRows.filter((r) => r[3] === true).length;
      }
      if (total > 0) {
        summary.push({ name: name, done: done, total: total });
        totalDone += done;
        totalPoints += total;
      }
    }
  });

  const result = {
    summary: summary,
    stats: { done: totalDone, total: totalPoints },
    updatedAt: new Date().getTime(),
  };

  // キャッシュに保存（文字列化して保存）
  PropertiesService.getScriptProperties().setProperty(
    "AREA_SUMMARY_CACHE",
    JSON.stringify(result),
  );

  return result;
}

/**
 * 特定のエリアの進捗だけをキャッシュ内で更新する（高速）
 */
function updateAreaCache(areaName, isDoneChange = 0) {
  const props = PropertiesService.getScriptProperties();
  const cached = props.getProperty("AREA_SUMMARY_CACHE");
  if (!cached) return;

  const data = JSON.parse(cached);
  const area = data.summary.find((s) => s.name === areaName);
  if (area) {
    if (isDoneChange !== 0) {
      area.done += isDoneChange;
      data.stats.done += isDoneChange;
    }
    props.setProperty("AREA_SUMMARY_CACHE", JSON.stringify(data));
  }
}

// =============================
// ⑤ メニュー & 初期化
// =============================

function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu("⚙️ ポスティング管理")
    .addItem("🚀 エリアシート一括作成", "forceStartBatch")
    .addSeparator()
    .addItem("🗺 司令室マップを開く (司令室用)", "openMapDashboard")
    .addItem("📱 アプリを開く (WebデプロイURL)", "showAppUrl")
    .addSeparator()
    .addItem("📊 全体数を集計する（ランキング更新）", "aggregateTotalVolumes")
    .addItem("📥 完了データのマスター抽出", "exportAllDataToMasterSheet")
    .addItem("📖 スタッフ用マニュアルを作成", "createManualSheet")
    .addSeparator()
    .addItem("🔄 バッチ処理を強制再開", "forceStartBatch")
    .addItem("⚠️ エリアシートをすべて削除（リセット）", "deleteAllAreaSheets")
    .addSeparator()
    .addItem(
      "🔍 ドライブのファイルを確認する (レスキュー)",
      "diagnoseDriveFiles",
    )
    .addSeparator()
    .addItem(
      "⚡ アプリ起動を高速化（キャッシュ更新）",
      "refreshAreaSummaryCache",
    )
    .addSeparator()
    .addItem("🎨 全シートを「プロ仕様」に一斉整形", "formatAllSheets")
    .addToUi();
}

function extractCityName(addr) {
  return addr.match(/^(.+?[市郡])/) ? addr.match(/^(.+?[市郡])/)[1] : "エリア";
}
function deleteTriggers(name) {
  ScriptApp.getProjectTriggers().forEach((t) => {
    if (t.getHandlerFunction() === name) ScriptApp.deleteTrigger(t);
  });
}

/**
 * 【レスキュー機能】ドライブ内のCSVファイルを診断し、表示する
 */
function diagnoseDriveFiles() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let msg = "【ドライブ診断結果】\n\n";
  let count = 0;

  try {
    const files = DriveApp.getFiles();
    while (files.hasNext()) {
      const file = files.next();
      const name = file.getName();
      if (
        name.toLowerCase().endsWith(".csv") ||
        file.getMimeType() === MimeType.GOOGLE_SHEETS
      ) {
        msg += `・${name}\n`;
        count++;
      }
      if (count > 20) break; // 探しすぎ防止
    }

    if (count === 0) {
      msg +=
        "CSVファイルが一つも見つかりませんでした。\nファイルをGoogleドライブの『マイドライブ』直下に置いてみてください。";
    } else {
      msg +=
        "\n上記の中に、使いたいファイル名はありますか？\n一字一句（スペース等も含め）同じである必要があります。";
    }
  } catch (e) {
    msg += "エラーが発生しました: " + e.message;
  }

  SpreadsheetApp.getUi().alert(msg);
}

function updateSheetSummary(sheet) {
  const last = sheet.getLastRow();
  if (last < 2) return;
  const data = sheet.getRange(2, 4, last - 1, 3).getValues();
  let total = 0;
  data.forEach((row) => {
    if (row[0] === true && typeof row[2] === "number") total += row[2];
  });
  sheet.getRange("H1").setValue(total);
}

function deleteAllAreaSheets() {
  if (isNotAdmin()) return;
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const exclude = [
    CONFIG.SHEET_GUIDE,
    CONFIG.SHEET_ROSTER,
    CONFIG.SHEET_TEMPLATE,
    CONFIG.SHEET_POSTAL,
    CONFIG.SHEET_DISTRICT,
  ];
  ss.getSheets().forEach((s) => {
    if (!exclude.includes(s.getName())) ss.deleteSheet(s);
  });
  ss.toast("リセット完了しました。");
}

function createManualSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet =
    ss.getSheetByName(CONFIG.SHEET_MANUAL) ||
    ss.insertSheet(CONFIG.SHEET_MANUAL);
  sheet.clear();
  sheet
    .getRange("B2")
    .setValue("ポスティング報告 らくらくガイド")
    .setFontSize(24)
    .setFontWeight("bold");
  ss.toast("マニュアル作成完了。");
}

function exportAllDataToMasterSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let master =
    ss.getSheetByName(CONFIG.SHEET_MASTER_EXPORT) ||
    ss.insertSheet(CONFIG.SHEET_MASTER_EXPORT);
  master.clear();
  ss.toast("マスター抽出完了。");
}

// トリガー
function onEdit(e) {
  if (!e || !e.range) return;
  const range = e.range;
  const sheet = range.getSheet();
  const name = sheet.getName();
  const col = range.getColumn();
  const row = range.getRow();

  // 除外シート
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
  if (exclude.includes(name) || sheet.isSheetHidden()) return;

  // D列（完了チェック）が編集された場合
  if (col === 4 && row >= 2) {
    const val = range.getValue();
    const now = Utilities.formatDate(new Date(), "JST", "MM/dd HH:mm");

    // 日時をセット
    sheet.getRange(row, 5).setValue(val ? now : "");

    // サマリーとキャッシュを更新
    updateSheetSummary(sheet);
    updateAreaCache(name, val ? 1 : -1);
  } else if (col === 6 || col === 7) {
    // 枚数や担当が変更された場合
    updateSheetSummary(sheet);
  }
}
/**
 * 原本および全エリアシートのデザインを「究極の視認性（シニア対応）」に一斉整形する
 */
function formatAllSheets() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const template = ss.getSheetByName(CONFIG.SHEET_TEMPLATE);
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

  // 1. まずは「原本」を完璧に整える
  applyProDesign(template);

  // 2. 「名簿」を整える
  formatRosterSheet();

  // 3. 他の全エリアシートにも同じデザインを適用
  const sheets = ss.getSheets();
  sheets.forEach((s) => {
    const name = s.getName();
    if (!exclude.includes(name) && !s.isSheetHidden()) {
      applyProDesign(s);
    }
  });

  ss.toast("全シートを「究極の視認性デザイン」に整形しました！");
}

/**
 * 指定したシートにプレミアムデザインを適用する内部関数
 */
function applyProDesign(sheet) {
  if (!sheet) return;

  // 一旦、すべての行と列を表示する（隠れている行を復活させる）
  const maxRowsInit = sheet.getMaxRows();
  const maxColsInit = sheet.getMaxColumns();
  sheet.showRows(1, maxRowsInit);
  sheet.showColumns(1, maxColsInit);

  // 列幅の調整（E,F,Gを重点的に拡大）
  sheet.setColumnWidth(1, 450); // 住所 (A)
  sheet.setColumnWidth(2, 60); // 地図 (B)
  sheet.setColumnWidth(3, 250); // メモ (C)
  sheet.setColumnWidth(4, 60); // 完了 (D)
  sheet.setColumnWidth(5, 180); // 日付 (E) - 拡大
  sheet.setColumnWidth(6, 120); // 枚数 (F) - 拡大
  sheet.setColumnWidth(7, 220); // 担当 (G) - 拡大

  // 行の高さ調整（85pxのゆったりサイズ）
  sheet.setRowHeight(1, 50); // ヘッダー
  const dataRowHeight = 85; // データ行
  sheet.setRowHeights(2, 10, dataRowHeight);

  // セルの書式設定
  const allRange = sheet.getRange("A1:G11");
  allRange
    .setHorizontalAlignment("center")
    .setVerticalAlignment("middle")
    .setWrapStrategy(SpreadsheetApp.WrapStrategy.WRAP);

  // ヘッダーデザイン
  const header = sheet.getRange("A1:G1");
  header
    .setBackground("#1a237e")
    .setFontColor("#ffffff")
    .setFontSize(14)
    .setFontWeight("bold");

  // 住所列の超強調（18pt）
  sheet.getRange("A2:A11").setFontSize(18).setFontWeight("bold");

  // その他のデータ行のフォントサイズ
  sheet.getRange("B2:G11").setFontSize(14).setFontWeight("bold");

  // ストライプデザイン（列ごとの色分け）
  // A, C, E, G, K, L列に薄い色を敷く
  const lightColor = "#f8f9fa";
  sheet.getRange("A2:A11").setBackground(lightColor);
  sheet.getRange("C2:C11").setBackground(lightColor);
  sheet.getRange("E2:E11").setBackground(lightColor);
  sheet.getRange("G2:G11").setBackground(lightColor);
  sheet.getRange("K2:L11").setBackground(lightColor);

  // B, D, F列は白（強調）
  sheet.getRange("B2:B11").setBackground("#ffffff");
  sheet.getRange("D2:D11").setBackground("#ffffff");
  sheet.getRange("F2:F11").setBackground("#ffffff");

  // H列以降（システム用）をすべて非表示にする
  if (maxColsInit > 7) {
    sheet.hideColumns(8, maxColsInit - 7);
  }

  if (maxRowsInit > 11) {
    sheet.hideRows(12, maxRowsInit - 11);
  }
}

/**
 * 名簿シートを「究極の視認性」に整形する
 */
function formatRosterSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(CONFIG.SHEET_ROSTER);
  if (!sheet) return;

  const maxRows = sheet.getMaxRows();
  const maxCols = sheet.getMaxColumns();

  // 枠線を消してアプリ感を出す
  sheet.setHiddenGridlines(true);

  // 列幅（A列のみ主役）
  sheet.setColumnWidth(1, 450);
  if (maxCols > 1) {
    sheet.hideColumns(2, maxCols - 1);
  }

  // ヘッダーデザイン
  const header = sheet.getRange("A1");
  header
    .setBackground("#1a237e")
    .setFontColor("#ffffff")
    .setFontSize(14)
    .setFontWeight("bold")
    .setHorizontalAlignment("center")
    .setVerticalAlignment("middle");
  sheet.setRowHeight(1, 60);

  // データ行のデザイン（1000行分あらかじめ設定）
  const lastRow = 1000;
  const dataRange = sheet.getRange(2, 1, lastRow - 1, 1);
  dataRange
    .setFontSize(18)
    .setFontWeight("bold")
    .setHorizontalAlignment("center")
    .setVerticalAlignment("middle")
    .setBackground("#f8f9fa");

  // 行の高さ（全データ行）
  sheet.setRowHeights(2, lastRow - 1, 85);

  ss.toast("名簿シートをプロ仕様に整形しました！");
}
