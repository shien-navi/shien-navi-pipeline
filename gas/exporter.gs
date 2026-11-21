/**
 * 地域資源ナビゲーション - 配信用データ出力（facilities.json / tags.json / facilities.csv）
 * 配置先: gas/exporter.gs
 */

const EXPORT_PROP_KEY = "EXPORT_FOLDER_ID";

/**
 * メイン実行：facilities.json / tags.json / facilities.csv を一括生成してドライブへ保存
 */
function exportAllFeedsToDrive() {
  const ssMap = resolveSpreadsheets_();
  const masterSheet = ssMap.masterSs.getSheetByName(APP_CONFIG.SHEET_NAMES.MASTER);
  const tagDictSheet = ssMap.tagDictSs.getSheetByName(APP_CONFIG.SHEET_NAMES.TAG_DICT);

  if (!masterSheet) {
    throw new Error("「事業所マスタ」シートが見つかりません。");
  }

  const tagsPayload = buildTagsPayload_(tagDictSheet);
  const facilitiesPayload = buildFacilitiesPayload_(masterSheet);
  const facilitiesCsvText = buildFacilitiesCsv_(facilitiesPayload);

  const folder = resolveExportFolder_();
  const stamp = Utilities.formatDate(new Date(), "Asia/Tokyo", "yyyyMMdd_HHmm");

  saveTextFile_(folder, "tags.json", JSON.stringify(tagsPayload, null, 2), "application/json");
  saveTextFile_(folder, "facilities.json", JSON.stringify(facilitiesPayload, null, 2), "application/json");
  saveTextFile_(folder, "facilities.csv", facilitiesCsvText, "text/csv");

  Logger.log(
    "出力完了 | フォルダ: " + folder.getName() +
    " | facilities: " + facilitiesPayload.facilities.length + "件" +
    " | tags: " + tagsPayload.tags.length + "件" +
    " | stamp: " + stamp
  );
}

function buildTagsPayload_(tagDictSheet) {
  const generatedAt = Utilities.formatDate(new Date(), "Asia/Tokyo", "yyyy-MM-dd'T'HH:mm:ss+09:00");
  if (!tagDictSheet || tagDictSheet.getLastRow() < 2) {
    return { generated_at: generatedAt, count: 0, tags: [] };
  }

  const lastRow = tagDictSheet.getLastRow();
  const lastCol = tagDictSheet.getLastColumn();
  const headers = tagDictSheet.getRange(1, 1, 1, lastCol).getValues()[0].map(h => String(h || "").trim());
  const rows = tagDictSheet.getRange(2, 1, lastRow - 1, lastCol).getValues();

  const colIdx = function(name) { return headers.indexOf(name); };
  const iCat = colIdx("カテゴリ");
  const iTag = colIdx("タグ");
  const iDesc = colIdx("説明");
  const iOrder = colIdx("表示順");
  const iPub = colIdx("公開");
  const iTarget = colIdx("適用対象");

  const list = [];
  rows.forEach((r, idx) => {
    const tag = iTag >= 0 ? String(r[iTag] || "").trim() : "";
    if (!tag) return;

    const pubRaw = iPub >= 0 ? String(r[iPub] || "").trim().toUpperCase() : "TRUE";
    const isPublic = !(pubRaw === "FALSE" || pubRaw === "0" || pubRaw === "NO" || pubRaw === "非公開");
    if (!isPublic) return;

    const orderNum = iOrder >= 0 && String(r[iOrder] || "").trim() !== ""
      ? Number(r[iOrder])
      : 9999 + idx;

    const targets = iTarget >= 0
      ? String(r[iTarget] || "").split(/[、,／/]/).map(s => s.trim()).filter(Boolean)
      : [];

    list.push({
      id: "tag_" + Utilities.formatString("%03d", idx + 1),
      tag: tag,
      category: iCat >= 0 ? String(r[iCat] || "").trim() : "",
      description: iDesc >= 0 ? String(r[iDesc] || "").trim() : "",
      display_order: isNaN(orderNum) ? 9999 + idx : orderNum,
      targets: targets
    });
  });

  list.sort((a, b) => a.display_order - b.display_order);

  return {
    generated_at: generatedAt,
    count: list.length,
    tags: list
  };
}

function buildFacilitiesPayload_(masterSheet) {
  const generatedAt = Utilities.formatDate(new Date(), "Asia/Tokyo", "yyyy-MM-dd'T'HH:mm:ss+09:00");
  const lastRow = masterSheet.getLastRow();
  const lastCol = masterSheet.getLastColumn();
  if (lastRow < 2 || lastCol < 1) {
    return { generated_at: generatedAt, count: 0, facilities: [] };
  }

  const headers = masterSheet.getRange(1, 1, 1, lastCol).getValues()[0].map(h => String(h || "").trim());
  const rows = masterSheet.getRange(2, 1, lastRow - 1, lastCol).getValues();

  const idx = function(name) { return headers.indexOf(name); };
  const getVal = function(r, name) {
    const c = idx(name);
    return c >= 0 ? String(r[c] || "").trim() : "";
  };

  const items = [];
  rows.forEach(r => {
    const resourceId = getVal(r, "resource_id");
    const name = getVal(r, "事業所名");
    const status = getVal(r, "status").toLowerCase();

    if (!resourceId && !name) return;
    if (status === "inactive" || status === "closed" || status === "廃止") return;

    const mergedRaw = getVal(r, "統合タグ");
    const memoRaw = getVal(r, "抽出タグメモ");
    const tagSource = mergedRaw || memoRaw;
    const tags = Array.from(new Set(
      tagSource
        .split(/[,、]/)
        .map(s => s.trim())
        .filter(Boolean)
    ));

    const latStr = getVal(r, "緯度");
    const lngStr = getVal(r, "経度");
    const lat = latStr !== "" && !isNaN(Number(latStr)) ? Number(latStr) : null;
    const lng = lngStr !== "" && !isNaN(Number(lngStr)) ? Number(lngStr) : null;

    items.push({
      resource_id: resourceId,
      resource_type: getVal(r, "resource_type") || "facility",
      facility_no: getVal(r, "facility_no"),
      corporation_no: getVal(r, "corporation_no"),
      source: getVal(r, "source"),
      status: getVal(r, "status") || "active",
      wam_url: getVal(r, "wam_url"),
      last_fetch: getVal(r, "last_fetch"),
      last_update: getVal(r, "last_update"),
      no: getVal(r, "No"),
      name: name,
      ward: getVal(r, "区"),
      address: getVal(r, "住所"),
      phone: getVal(r, "電話番号"),
      service_type: getVal(r, "サービス種類"),
      pickup: getVal(r, "送迎"),
      wheelchair: getVal(r, "車いす対応"),
      instructors: getVal(r, "在籍指導員"),
      message: getVal(r, "メッセージ"),
      tags: tags,
      location: {
        lat: lat,
        lng: lng
      }
    });
  });

  return {
    generated_at: generatedAt,
    count: items.length,
    facilities: items
  };
}

function buildFacilitiesCsv_(facilitiesPayload) {
  const csvHeaders = [
    "resource_id",
    "facility_no",
    "name",
    "ward",
    "address",
    "phone",
    "service_type",
    "pickup",
    "wheelchair",
    "instructors",
    "tags",
    "lat",
    "lng",
    "message"
  ];

  const escapeCsv = function(v) {
    const s = v === null || v === undefined ? "" : String(v);
    if (/[",\n\r]/.test(s)) {
      return '"' + s.replace(/"/g, '""') + '"';
    }
    return s;
  };

  const lines = [csvHeaders.map(escapeCsv).join(",")];
  (facilitiesPayload.facilities || []).forEach(f => {
    const row = [
      f.resource_id,
      f.facility_no,
      f.name,
      f.ward,
      f.address,
      f.phone,
      f.service_type,
      f.pickup,
      f.wheelchair,
      f.instructors,
      (f.tags || []).join(" | "),
      f.location ? f.location.lat : "",
      f.location ? f.location.lng : "",
      f.message
    ];
    lines.push(row.map(escapeCsv).join(","));
  });

  return lines.join("\n");
}

function resolveExportFolder_() {
  const folderId = String(PropertiesService.getScriptProperties().getProperty(EXPORT_PROP_KEY) || "").trim();
  if (folderId) {
    return DriveApp.getFolderById(folderId);
  }
  return DriveApp.getRootFolder();
}

function saveTextFile_(folder, fileName, content, mimeType) {
  const existing = folder.getFilesByName(fileName);
  if (existing.hasNext()) {
    const file = existing.next();
    file.setContent(content);
    return file;
  }
  const blob = Utilities.newBlob(content, mimeType, fileName);
  return folder.createFile(blob);
}
