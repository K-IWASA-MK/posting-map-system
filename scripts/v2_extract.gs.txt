/**
 * GAS v2 - 住所抽出モジュール
 * - CSVからの選挙区・住所データの抽出
 * - 住所文字列の正規化
 */

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

function extractCityName(addr) {
  return addr.match(/^(.+?[市郡])/) ? addr.match(/^(.+?[市郡])/)[1] : "エリア";
}
