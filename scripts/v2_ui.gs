/**
 * GAS v2 - UI・イベント管理モジュール
 * - メニュー作成 (onOpen)
 * - イベントトリガー (onEdit)
 * - シートデザイン・整形ロジック
 * - 診断機能
 */

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
    .addItem("⏰ 自動集計（1時間ごと）を有効化", "setupHourlyRefreshTrigger")
    .addSeparator()
    .addItem("🎨 全シートを「プロ仕様」に一斉整形", "formatAllSheets")
    .addItem("🔧 名簿シートを初期化・復旧する", "setupRosterSheet")
    .addToUi();
}

/**
 * 定期集計トリガーを設定する
 */
function setupHourlyRefreshTrigger() {
  // 既存のトリガーを掃除
  deleteTriggers("refreshAreaSummaryCache");

  // 1時間おきに実行
  ScriptApp.newTrigger("refreshAreaSummaryCache")
    .timeBased()
    .everyHours(1)
    .create();

  SpreadsheetApp.getUi().alert("1時間おきの自動集計トリガーを設定しました。");
}

/**
 * トリガー
 */
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
  
  createSystemCacheSheet();
  refreshAreaSummaryCache();
  
  ss.toast("リセット完了しました。");
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

  // 全列を表示させてから、D列以降を隠す
  sheet.showColumns(1, maxCols);
  sheet.setHiddenGridlines(true);

  // 列幅設定 (ID: 100, 苗字: 250, 名前: 250)
  sheet.setColumnWidth(1, 100);
  sheet.setColumnWidth(2, 250);
  sheet.setColumnWidth(3, 250);

  if (maxCols > 3) {
    sheet.hideColumns(4, maxCols - 3);
  }

  // ヘッダーデザイン (A1:C1)
  const header = sheet.getRange("A1:C1");
  header
    .setBackground("#1a237e")
    .setFontColor("#ffffff")
    .setFontSize(14)
    .setFontWeight("bold")
    .setHorizontalAlignment("center")
    .setVerticalAlignment("middle");
  sheet.setRowHeight(1, 50);

  // データ行のデザイン（1000行分あらかじめ設定）
  const lastRow = 1000;
  const dataRange = sheet.getRange(2, 1, lastRow - 1, 3);
  dataRange
    .setFontSize(18)
    .setFontWeight("bold")
    .setHorizontalAlignment("center")
    .setVerticalAlignment("middle")
    .setBackground("#ffffff");

  // 行の高さ（全データ行）
  sheet.setRowHeights(2, lastRow - 1, 85);

  ss.toast("名簿シートをプロ仕様に整形しました！");
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

  // ★ 修正: 一旦全データをクリアして真っ新にする
  sheet.clear();
  sheet.getRange(1, 1, 1, 3).setValues([["ID", "苗字", "名前"]]);
  sheet.setFrozenRows(1); // 1行目を固定
  
  formatRosterSheet(); // 整形も同時に行う

  return "名簿シートを真っ新に初期化しました。";
}

/**
 * 管理者権限をチェックする
 */
function isNotAdmin() {
  // 現時点では全員を管理者として扱う（将来的にメールアドレス制限などを入れる場合はここを修正）
  return false;
}
