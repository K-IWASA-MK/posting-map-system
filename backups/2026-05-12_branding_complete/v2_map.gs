/**
 * GAS v2 - マップデータ管理モジュール
 * - 地図表示用データの集計
 * - パフォーマンス向上のためのキャッシュ管理
 */

/**
 * モバイルアプリ用：全体サマリー取得（爆速キャッシュ版）
 * 1. CacheService (10分) をチェック
 * 2. なければ PropertiesService をチェック
 * 3. なければ再計算
 */
function getMapData() {
  const cache = CacheService.getScriptCache();
  const fastCached = cache.get("AREA_SUMMARY_FAST_CACHE");
  if (fastCached) return JSON.parse(fastCached);

  const props = PropertiesService.getScriptProperties();
  const cached = props.getProperty("AREA_SUMMARY_CACHE");

  if (cached) {
    try {
      const data = JSON.parse(cached);
      // PropertiesServiceから読み込んだ場合はCacheServiceに書き戻して次回を高速化
      cache.put("AREA_SUMMARY_FAST_CACHE", cached, 600);
      return data;
    } catch (e) {
      // キャッシュ破損時は再計算へ
    }
  }

  // キャッシュがない場合は初回のみ計算
  return refreshAreaSummaryCache();
}

/**
 * 全エリアのサマリーを再計算してキャッシュに保存する (爆速シャドウシート版)
 */
function refreshAreaSummaryCache() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let shadowSheet = ss.getSheetByName(CONFIG.SHEET_SYSTEM_CACHE);

  // シャドウシートがなければ作成
  if (!shadowSheet) {
    createSystemCacheSheet();
    shadowSheet = ss.getSheetByName(CONFIG.SHEET_SYSTEM_CACHE);
  }

  const lastRow = shadowSheet.getLastRow();
  if (lastRow < 2) return { summary: [], stats: { done: 0, total: 0 } };

  // 1回のAPI通信で全エリアの集計結果を取得 (A:エリア名, B:完了数, C:合計数)
  const data = shadowSheet.getRange(2, 1, lastRow - 1, 3).getValues();
  const summary = [];
  let totalDone = 0;
  let totalPoints = 0;

  data.forEach((row) => {
    const name = row[0];
    const done = Number(row[1]) || 0;
    const total = Number(row[2]) || 0;

    if (name) {
      summary.push({ name: name, done: done, total: total });
      totalDone += done;
      totalPoints += total;
    }
  });

  const result = {
    summary: summary,
    stats: { done: totalDone, total: totalPoints },
    updatedAt: new Date().getTime(),
  };

  const jsonResult = JSON.stringify(result);
  const cache = CacheService.getScriptCache();
  cache.put("AREA_SUMMARY_FAST_CACHE", jsonResult, 600);
  PropertiesService.getScriptProperties().setProperty("AREA_SUMMARY_CACHE", jsonResult);

  return result;
}

/**
 * 集計用シャドウシート (__SYSTEM_CACHE__) を生成/更新する
 * エリアシートが増えた時などに呼び出す
 */
function createSystemCacheSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(CONFIG.SHEET_SYSTEM_CACHE);
  
  if (!sheet) {
    sheet = ss.insertSheet(CONFIG.SHEET_SYSTEM_CACHE);
    sheet.hideSheet();
  }
  
  sheet.clear();
  sheet.getRange(1, 1, 1, 3).setValues([["エリア名", "完了数", "合計数"]]);

  const exclude = [
    CONFIG.SHEET_GUIDE, CONFIG.SHEET_ROSTER, CONFIG.SHEET_TEMPLATE,
    CONFIG.SHEET_POSTAL, CONFIG.SHEET_DISTRICT, CONFIG.SHEET_MASTER_EXPORT,
    CONFIG.SHEET_REPORT, CONFIG.SHEET_MANUAL, CONFIG.SHEET_SYSTEM_CACHE
  ];

  const areaSheets = ss.getSheets().filter(s => !exclude.includes(s.getName()) && !s.isSheetHidden());
  
  if (areaSheets.length === 0) return;

  const areaNames = areaSheets.map(s => [s.getName()]);
  const formulas = areaSheets.map(s => {
    const name = s.getName();
    const escapedName = name.replace(/'/g, "''");
    return [
      `=COUNTIF('${escapedName}'!D:D, TRUE)`,
      `=COUNTA('${escapedName}'!A2:A)`
    ];
  });

  sheet.getRange(2, 1, areaNames.length, 1).setValues(areaNames);
  sheet.getRange(2, 2, formulas.length, 2).setFormulas(formulas);
}

/**
 * 特定のエリアの進捗だけをキャッシュ内で更新する（高速）
 */
function updateAreaCache(areaName, isDoneChange = 0) {
  const props = PropertiesService.getScriptProperties();
  const cache = CacheService.getScriptCache();
  const cached = props.getProperty("AREA_SUMMARY_CACHE");
  if (!cached) return;

  const data = JSON.parse(cached);
  const area = data.summary.find((s) => s.name === areaName);
  if (area) {
    if (isDoneChange !== 0) {
      area.done += isDoneChange;
      data.stats.done += isDoneChange;
    }
    const updatedJson = JSON.stringify(data);
    props.setProperty("AREA_SUMMARY_CACHE", updatedJson);
    cache.put("AREA_SUMMARY_FAST_CACHE", updatedJson, 600);
  }
}
