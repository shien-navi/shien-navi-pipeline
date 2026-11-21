/**
 * 地域資源ナビゲーション - WAM取得＆マスタ接続
 * 配置先: gas/wamConnector.gs
 */

const APP_CONFIG = {
  PROP_KEYS: {
    MASTER_SS_ID: "MASTER_SS_ID",
    SYNC_SS_ID: "SYNC_SS_ID",
    SERVICE_DICT_SS_ID: "SERVICE_DICT_ID",
    TAG_DICT_SS_ID: "TAG_DICT_ID",
    RAW_SS_ID: "RAW_SS_ID"
  },
  SHEET_NAMES: {
    MASTER: "事業所マスタ",
    SYNC: "WAM同期管理",
    SERVICE_DICT: "サービス種別辞書",
    TAG_DICT: "タグ辞書",
    RAW: "WAM_RAW"
  },
  WAM: {
    entryUrl: "https://www.wam.go.jp/sfkohyoout/COP000100E0000.do",
    searchUrl: "https://www.wam.go.jp/sfkohyoout/COP000100E12.do",
    baseUrl: "https://www.wam.go.jp",
    userAgent: "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/124.0.0.0 Safari/537.36",
    sleepMs: 800,
    maxRuntimeMs: 4.5 * 60 * 1000
  }
};

/**
 * 初回セットアップ用：各マスタIDを ScriptProperties に登録するヘルパー
 */
function setupSheetIds() {
  const props = PropertiesService.getScriptProperties();
  props.setProperties({
    [APP_CONFIG.PROP_KEYS.MASTER_SS_ID]: "",
    [APP_CONFIG.PROP_KEYS.SYNC_SS_ID]: "",
    [APP_CONFIG.PROP_KEYS.SERVICE_DICT_SS_ID]: "",
    [APP_CONFIG.PROP_KEYS.TAG_DICT_SS_ID]: "",
    [APP_CONFIG.PROP_KEYS.RAW_SS_ID]: ""
  });
  Logger.log("ScriptProperties のキー枠を初期化しました。必要に応じてIDを設定してください。");
}

/**
 * メイン実行：WAM同期管理の未同期行を処理し、WAM_RAW・同期管理・事業所マスタへ反映
 */
function runWamSyncAndConnectMaster(forceAll) {
  const startTime = Date.now();
  const ssMap = resolveSpreadsheets_();

  const syncSheet = getOrCreateSheet_(ssMap.syncSs, APP_CONFIG.SHEET_NAMES.SYNC, [
    "facility_no", "status", "facility_name", "corporation_name", "last_synced_at", "message"
  ]);
  const rawSheet = getOrCreateSheet_(ssMap.rawSs, APP_CONFIG.SHEET_NAMES.RAW, [
    "facility_no", "json", "fetch_date"
  ]);
  const masterSheet = ssMap.masterSs.getSheetByName(APP_CONFIG.SHEET_NAMES.MASTER);
  const tagDictSheet = ssMap.tagDictSs.getSheetByName(APP_CONFIG.SHEET_NAMES.TAG_DICT);

  const tagRules = tagDictSheet ? loadTagRules_(tagDictSheet) : [];
  const masterIndex = masterSheet ? buildFacilityIndex_(masterSheet) : {};

  const lastRow = syncSheet.getLastRow();
  if (lastRow < 2) {
    Logger.log("「WAM同期管理」に処理対象（2行目以降の facility_no）がありません。");
    return;
  }

  const syncRange = syncSheet.getRange(2, 1, lastRow - 1, 6);
  const syncRows = syncRange.getValues();

  let okCount = 0;
  let errCount = 0;
  let stopped = false;

  for (let i = 0; i < syncRows.length; i++) {
    if (Date.now() - startTime > APP_CONFIG.WAM.maxRuntimeMs) {
      stopped = true;
      Logger.log("実行時間上限に近づいたため安全停止しました。再実行で続きから処理できます。");
      break;
    }

    const rowNum = i + 2;
    const facilityNo = String(syncRows[i][0] || "").trim();
    const status = String(syncRows[i][1] || "").trim();

    if (!facilityNo) continue;
    if (!forceAll && status !== "" && status !== "未同期") continue;

    const nowTs = Utilities.formatDate(new Date(), "Asia/Tokyo", "yyyy/MM/dd HH:mm:ss");
    const todayStr = Utilities.formatDate(new Date(), "Asia/Tokyo", "yyyy/MM/dd");

    try {
      const fetched = fetchWamFacilityDetail_(facilityNo);
      const extractedTags = extractTagsFromWamData_(fetched.combinedData, tagRules);

      rawSheet.appendRow([
        facilityNo,
        JSON.stringify(fetched.combinedData),
        todayStr
      ]);

      syncRows[i][1] = "完了";
      syncRows[i][2] = fetched.facilityName;
      syncRows[i][3] = fetched.corporationName;
      syncRows[i][4] = nowTs;
      syncRows[i][5] = "OK | tags:" + extractedTags.length;

      if (masterSheet && masterIndex[facilityNo]) {
        updateMasterRowWithWam_(
          masterSheet,
          masterIndex[facilityNo],
          fetched,
          extractedTags,
          nowTs
        );
      }

      okCount++;
      Logger.log("[" + rowNum + "行目] " + facilityNo + " (" + fetched.facilityName + ") -> 完了 / タグ: " + extractedTags.join(", "));
    } catch (e) {
      const msg = String(e && e.message ? e.message : e);
      syncRows[i][1] = "エラー";
      syncRows[i][4] = nowTs;
      syncRows[i][5] = msg;
      errCount++;
      Logger.log("[" + rowNum + "行目] " + facilityNo + " -> エラー: " + msg);
    }

    syncSheet.getRange(rowNum, 1, 1, 6).setValues([syncRows[i]]);
    SpreadsheetApp.flush();

    if (APP_CONFIG.WAM.sleepMs > 0) {
      Utilities.sleep(APP_CONFIG.WAM.sleepMs);
    }
  }

  Logger.log("完了サマリ | 成功: " + okCount + " / エラー: " + errCount + " / 中断: " + (stopped ? "あり" : "なし"));
}

function runWamSyncAll() {
  runWamSyncAndConnectMaster(true);
}

function resolveSpreadsheets_() {
  const props = PropertiesService.getScriptProperties();
  const active = SpreadsheetApp.getActiveSpreadsheet();

  const openOrActive = function(propKey) {
    const id = String(props.getProperty(propKey) || "").trim();
    return id ? SpreadsheetApp.openById(id) : active;
  };

  return {
    masterSs: openOrActive(APP_CONFIG.PROP_KEYS.MASTER_SS_ID),
    syncSs: openOrActive(APP_CONFIG.PROP_KEYS.SYNC_SS_ID),
    serviceDictSs: openOrActive(APP_CONFIG.PROP_KEYS.SERVICE_DICT_SS_ID),
    tagDictSs: openOrActive(APP_CONFIG.PROP_KEYS.TAG_DICT_SS_ID),
    rawSs: openOrActive(APP_CONFIG.PROP_KEYS.RAW_SS_ID)
  };
}

function loadTagRules_(tagDictSheet) {
  const lastRow = tagDictSheet.getLastRow();
  const lastCol = tagDictSheet.getLastColumn();
  if (lastRow < 2) return [];

  const headers = tagDictSheet.getRange(1, 1, 1, lastCol).getValues()[0].map(h => String(h || "").trim());
  const rows = tagDictSheet.getRange(2, 1, lastRow - 1, lastCol).getValues();

  const idxTag = headers.indexOf("タグ");
  const idxPub = headers.indexOf("公開");
  const idxWamLink = headers.indexOf("WAM連携");
  const idxWamItem = headers.indexOf("WAM項目名");

  const rules = [];
  rows.forEach(r => {
    const tag = idxTag >= 0 ? String(r[idxTag] || "").trim() : "";
    if (!tag) return;

    const pubVal = idxPub >= 0 ? String(r[idxPub] || "").trim().toUpperCase() : "TRUE";
    if (pubVal === "FALSE" || pubVal === "0" || pubVal === "NO") return;

    const wamLinkRaw = idxWamLink >= 0 ? String(r[idxWamLink] || "").trim() : "";
    const wamItem = idxWamItem >= 0 ? String(r[idxWamItem] || "").trim() : "";

    const linkUpper = wamLinkRaw.toUpperCase();
    const isWamEnabled = (linkUpper === "TRUE" || linkUpper === "1" || linkUpper === "○" || linkUpper === "YES" || wamItem !== "");

    rules.push({
      tag: tag,
      isWamEnabled: isWamEnabled,
      wamItemNames: wamItem ? wamItem.split(/[、,／/]/).map(s => s.trim()).filter(Boolean) : [tag]
    });
  });

  return rules;
}

function extractTagsFromWamData_(combinedData, tagRules) {
  if (!combinedData) return [];
  const staff = combinedData.staff || {};
  const service = combinedData.service || {};
  const operation = combinedData.operation || {};
  const yesNoPool = Object.assign({}, operation, service);

  const matched = new Set();

  tagRules.forEach(rule => {
    if (!rule.isWamEnabled) return;

    for (let i = 0; i < rule.wamItemNames.length; i++) {
      const keyName = rule.wamItemNames[i];

      if (keyName in yesNoPool) {
        if (isYesValue_(yesNoPool[keyName])) {
          matched.add(rule.tag);
          break;
        }
      }

      if (keyName in staff) {
        const arr = Array.isArray(staff[keyName]) ? staff[keyName] : [staff[keyName]];
        const hasPositive = arr.some(v => {
          const num = parseFloat(String(v || "").replace(/[^\d.-]/g, ""));
          return !isNaN(num) && num > 0;
        });
        if (hasPositive) {
          matched.add(rule.tag);
          break;
        }
      }
    }
  });

  Object.keys(yesNoPool).forEach(k => {
    if (isYesValue_(yesNoPool[k])) {
      if (k.indexOf("喀痰吸引") >= 0 || k.indexOf("医療的ケア") >= 0) {
        matched.add("医療的ケア");
      }
    }
  });

  return Array.from(matched);
}

function isYesValue_(val) {
  if (Array.isArray(val)) {
    return val.some(v => isYesValue_(v));
  }
  const s = String(val || "").trim();
  return s === "あり" || s === "有" || s === "TRUE" || s === "true";
}

function buildFacilityIndex_(masterSheet) {
  const lastRow = masterSheet.getLastRow();
  const lastCol = masterSheet.getLastColumn();
  if (lastRow < 2 || lastCol < 1) return {};

  const headers = masterSheet.getRange(1, 1, 1, lastCol).getValues()[0].map(h => String(h || "").trim());
  const facNoColIdx = headers.indexOf("facility_no");
  if (facNoColIdx < 0) return {};

  const rows = masterSheet.getRange(2, facNoColIdx + 1, lastRow - 1, 1).getValues();
  const index = {};
  for (let i = 0; i < rows.length; i++) {
    const no = String(rows[i][0] || "").trim();
    if (no) {
      index[no] = i + 2;
    }
  }
  return index;
}

function updateMasterRowWithWam_(masterSheet, rowNum, fetched, extractedTags, nowTs) {
  const lastCol = masterSheet.getLastColumn();
  const headers = masterSheet.getRange(1, 1, 1, lastCol).getValues()[0].map(h => String(h || "").trim());
  const rowVals = masterSheet.getRange(rowNum, 1, 1, lastCol).getValues()[0];

  const setByHeader = function(colName, newVal) {
    const idx = headers.indexOf(colName);
    if (idx >= 0) {
      rowVals[idx] = newVal;
    }
  };

  setByHeader("last_fetch", nowTs);
  setByHeader("status", "active");
  if (fetched.facilityName) {
    const nameIdx = headers.indexOf("事業所名");
    if (nameIdx >= 0 && !String(rowVals[nameIdx] || "").trim()) {
      rowVals[nameIdx] = fetched.facilityName;
    }
  }

  const tagStr = extractedTags.join(",");
  setByHeader("抽出タグメモ", tagStr);

  const mergedIdx = headers.indexOf("統合タグ");
  if (mergedIdx >= 0) {
    const existing = String(rowVals[mergedIdx] || "")
      .split(",")
      .map(s => s.trim())
      .filter(Boolean);
    const combinedSet = new Set(existing.concat(extractedTags));
    rowVals[mergedIdx] = Array.from(combinedSet).join(",");
  }

  masterSheet.getRange(rowNum, 1, 1, lastCol).setValues([rowVals]);
}

function fetchWamFacilityDetail_(facilityNo) {
  const ua = APP_CONFIG.WAM.userAgent;

  const res1 = UrlFetchApp.fetch(APP_CONFIG.WAM.entryUrl, {
    headers: { "User-Agent": ua },
    muteHttpExceptions: true
  });
  if (res1.getResponseCode() !== 200) throw new Error("初期画面取得失敗: " + res1.getResponseCode());
  const html1 = res1.getContentText("UTF-8");
  const cookieHeader = extractCookieHeader_(res1.getAllHeaders());
  const token1 = extractToken_(html1);
  if (!token1) throw new Error("_TOKEN 取得失敗");

  const res2 = UrlFetchApp.fetch(APP_CONFIG.WAM.searchUrl, {
    method: "post",
    payload: {
      "_FRAMEID": "root", "_TARGETID": "root", "_LUID": "",
      "_TOKEN": token1, "_FORMID": "COP000100", "_SUBINDEX": "",
      "vo_headVO_adress": "", "vo_headVO_corporation_name": "",
      "vo_headVO_facility_name": "", "vo_headVO_no": facilityNo,
      "vo_headVO_townCode": "", "vo_headVO_prefCode": ""
    },
    headers: { "User-Agent": ua, "Referer": APP_CONFIG.WAM.entryUrl, "Cookie": cookieHeader },
    muteHttpExceptions: true
  });
  if (res2.getResponseCode() !== 200) throw new Error("検索POST失敗: " + res2.getResponseCode());
  const html2 = res2.getContentText("UTF-8");

  const linkMatch = html2.match(/href="(\/sfkohyoout\/COP020100E00\.do\?[^"]+)"/);
  if (!linkMatch) throw new Error("詳細リンクなし");
  const detailUrl = APP_CONFIG.WAM.baseUrl + linkMatch[1].replace(/&amp;/g, "&");

  const res3 = UrlFetchApp.fetch(detailUrl, {
    headers: { "User-Agent": ua, "Referer": APP_CONFIG.WAM.searchUrl, "Cookie": cookieHeader },
    muteHttpExceptions: true
  });
  if (res3.getResponseCode() !== 200) throw new Error("詳細画面取得失敗: " + res3.getResponseCode());
  const html3 = res3.getContentText("UTF-8");

  const titleMatch = html3.match(/<title>([\s\S]*?)<\/title>/i);
  const facilityName = titleMatch
    ? decodeHtmlEntities_(titleMatch[1]).replace(/事業所詳細情報[ \s]*/g, "").trim()
    : "";

  const corpMatch = html3.match(
    />法人等の名称<\/th>[\s\S]*?<td[^>]*>([\s\S]*?)<\/td>|>法人等の名称<\/td>[\s\S]*?<td[^>]*>([\s\S]*?)<\/td>/i
  );
  const corporationName = cleanHtmlText_(corpMatch ? (corpMatch[1] || corpMatch[2] || "") : "");

  const hiddens = extractHiddenInputs_(html3);
  const token3 = hiddens["_TOKEN"] || extractToken_(html3);
  const formId3 = hiddens["_FORMID"] || "COP020201E00";

  const fetchTab = function(actionCode) {
    const r = UrlFetchApp.fetch(APP_CONFIG.WAM.baseUrl + "/sfkohyoout/" + actionCode + ".do", {
      method: "post",
      payload: Object.assign({}, hiddens, { "_TOKEN": token3, "_FORMID": formId3 }),
      headers: { "User-Agent": ua, "Referer": detailUrl, "Cookie": cookieHeader },
      muteHttpExceptions: true
    });
    if (r.getResponseCode() !== 200) throw new Error(actionCode + " 取得失敗: " + r.getResponseCode());
    return r.getContentText("UTF-8");
  };

  return {
    facilityName: facilityName,
    corporationName: corporationName,
    combinedData: {
      staff: parseStaffTab_(fetchTab("COP020203E00")),
      service: parseYesNoTab_(fetchTab("COP020204E00")),
      operation: parseYesNoTab_(fetchTab("COP020206E00"))
    }
  };
}

function parseStaffTab_(tabHtml) {
  const map = {};
  for (const tr of tabHtml.matchAll(/<tr[^>]*>([\s\S]*?)<\/tr>/gi)) {
    const cells = [...tr[1].matchAll(/<td[^>]*class=["'][^"']*tb03_td_up[^"']*["'][^>]*>([\s\S]*?)<\/td>/gi)]
      .map(c => cleanHtmlText_(c[1]));
    if (cells.length >= 2) {
      const role = cells[0];
      if (role && !/^[0-9.]+$/.test(role)) {
        map[role] = cells.slice(1);
      }
    }
  }
  return map;
}

function parseYesNoTab_(tabHtml) {
  const afterMenu = tabHtml.split('id="big08"')[1] || tabHtml.split('id="big01"')[1] || tabHtml;
  const map = {};

  for (const tr of afterMenu.matchAll(/<tr[^>]*>([\s\S]*?)<\/tr>/gi)) {
    const row = tr[1];
    if (!row.includes("radio-parts-checked")) continue;

    const checkedVals = [...row.matchAll(/<span[^>]*class=["'][^"']*radio-parts-checked[^"']*["'][^>]*>([\s\S]*?)<\/span>/gi)]
      .map(m => cleanHtmlText_(m[1]))
      .filter(Boolean);

    const withoutRadios = row.replace(
      /<span[^>]*class=["'][^"']*radio-parts(?:-checked)?[^"']*["'][^>]*>[\s\S]*?<\/span>/gi,
      ""
    );
    const label = cleanHtmlText_(withoutRadios);

    if (!label || checkedVals.length === 0) continue;
    const val = checkedVals.length === 1 ? checkedVals[0] : checkedVals;

    if (!(label in map)) {
      map[label] = val;
    } else if (Array.isArray(map[label])) {
      map[label].push(val);
    } else {
      map[label] = [map[label], val];
    }
  }
  return map;
}

function getOrCreateSheet_(ss, name, headers) {
  let sh = ss.getSheetByName(name);
  if (!sh) {
    sh = ss.insertSheet(name);
    if (headers && headers.length) sh.appendRow(headers);
  }
  return sh;
}

function extractCookieHeader_(headers) {
  const raw = headers["Set-Cookie"] || headers["set-cookie"];
  if (!raw) return "";
  return Array.isArray(raw)
    ? raw.map(c => String(c).split(";")[0]).join("; ")
    : String(raw).split(";")[0];
}

function extractToken_(html) {
  const m = html.match(/name="_TOKEN"\s+value="([^"]*)"/);
  return m ? m[1] : "";
}

function extractHiddenInputs_(html) {
  const hiddens = {};
  for (const m of html.matchAll(/<input[^>]*type=["']?hidden["']?[^>]*name=["']?([^"'\s>]+)["']?[^>]*value=["']?([^"'\s>]*)["']?[^>]*>/gi)) {
    hiddens[m[1]] = m[2];
  }
  return hiddens;
}

function cleanHtmlText_(htmlStr) {
  return decodeHtmlEntities_(
    String(htmlStr || "")
      .replace(/<[^>]+>/g, " ")
      .replace(/&nbsp;/gi, " ")
  )
    .replace(/\s+/g, " ")
    .trim();
}

function decodeHtmlEntities_(str) {
  return String(str || "")
    .replace(/&#x([0-9a-fA-F]+);/g, (_, hex) => String.fromCharCode(parseInt(hex, 16)))
    .replace(/&#([0-9]+);/g, (_, dec) => String.fromCharCode(parseInt(dec, 10)))
    .replace(/&amp;/g, "&")
    .replace(/&lt;/g, "<")
    .replace(/&gt;/g, ">")
    .replace(/&quot;/g, '"');
}
