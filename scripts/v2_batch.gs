/**
 * GAS v2 - バッチ処理モジュール
 * - 大規模データ展開用のバッチエンジン
 * - トリガー管理
 */

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
    
    // シートの取得/作成ロジックを堅牢化
    let sheet = ss.getSheetByName(sheetName);
    if (!sheet) {
      try {
        sheet = baseSheet.copyTo(ss).setName(sheetName);
      } catch (e) {
        // 重複エラーが発生した場合のリカバリ
        sheet = ss.getSheetByName(sheetName);
        if (!sheet) {
          // それでも取得できない場合は名前を少し変えて作成
          sheet = baseSheet.copyTo(ss).setName(sheetName + " ");
        }
      }
    }
    sheet.showSheet();

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
  
  // シャドウシートを最新のリストで更新
  createSystemCacheSheet();
  
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

function deleteTriggers(name) {
  ScriptApp.getProjectTriggers().forEach((t) => {
    if (t.getHandlerFunction() === name) ScriptApp.deleteTrigger(t);
  });
}
