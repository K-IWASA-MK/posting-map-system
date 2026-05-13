/**
 * GAS v2 - 集計・報告モジュール
 * - 配布枚数の集計
 * - 各種報告用シートの作成・出力
 */

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
